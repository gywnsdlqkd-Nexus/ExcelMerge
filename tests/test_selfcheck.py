# -*- coding: utf-8 -*-
"""자체 점검(--selfcheck) — 화면 없이 사용자 흐름을 몰아 보는 릴리스 게이트.

이 점검의 값은 **실제로 났던 버그를 잡느냐**로 정해진다. 그래서 건강한 앱에서 통과하는
것만이 아니라, 실기에서야 드러났던 두 회귀를 되살렸을 때 **실패하는지**를 못 박는다.

  · v219 — 셀값란에서 고른 글자를 Ctrl+C 하면 표(또는 경로)가 복사됐다
  · v221 — 저장을 한 번 한 뒤 Ctrl+S 를 다시 누르면 오류 창이 떴다

그리고 사용자의 것을 건드리지 않는다는 약속(원본 파일 · 설정 · 로그 · 알림 창 가로채기
원상복구)을 고정한다. 저장소가 공개라 실데이터 대신 합성 xlsx 를 쓴다.
"""
import ast
import hashlib
import json
import os

import openpyxl
import pytest

from excelmerge import selfcheck

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    for r in rows:
        ws.append(r)
    wb.save(str(path))
    return str(path)


@pytest.fixture
def pair(tmp_path):
    """값이 다른 칸과 한쪽에만 있는 행이 섞인 한 쌍 — 실제 비교처럼."""
    src = tmp_path / "src"
    src.mkdir()
    a = _xlsx(src / "a.xlsx", [["ID", "Desc", "Num"],
                               ["k1", "대여 만료 아이템 보상", 10],
                               ["k2", "입문 보상", 20],
                               ["k3", "같은 값", 30]])
    b = _xlsx(src / "b.xlsx", [["ID", "Desc", "Num"],
                               ["k1", "기간 만료 아이템 보상", 11],
                               ["k2", "입문 보상", 20],
                               ["k3", "같은 값", 30],
                               ["k4", "새 행", 40]])
    return a, b


def _run(qapp, pair, tmp_path):
    work = tmp_path / "work"
    work.mkdir()
    checks = selfcheck.run_checks(pair[0], pair[1], str(work), timeout=30)
    return {c["name"]: c for c in checks}


def _sha(p):
    with open(p, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


# ── 건강한 앱은 통과한다 ────────────────────────────────────────────────────

def test_a_healthy_app_passes_every_check(qapp, pair, tmp_path):
    res = _run(qapp, pair, tmp_path)
    bad = {k: v["detail"] for k, v in res.items() if not v["ok"]}
    assert not bad, bad
    # 흐름의 단계가 빠지지 않았는지 — 앞에서 멈추면 뒤 항목이 아예 없다
    for name in ("비교 완료", "셀값란 Ctrl+C 가 그 값을 복사", "병합 준비됨 (A → B)",
                 "B 가 저장됨", "A 는 그대로", "저장 뒤 Ctrl+S 에 오류 없음",
                 "처리되지 않은 예외 없음"):
        assert name in res, f"'{name}' 단계까지 가지 않았다: {list(res)}"


# ── 실제로 났던 회귀를 잡는다 ───────────────────────────────────────────────

def test_it_catches_the_cell_box_copy_regression(qapp, pair, tmp_path, monkeypatch):
    """v219 — 셀값란이 Ctrl+C 를 가져가지 않으면 표의 복사가 이긴다."""
    from PyQt5.QtWidgets import QPlainTextEdit
    from excelmerge.widgets import CellEditWidget
    monkeypatch.setattr(CellEditWidget, "event", QPlainTextEdit.event)
    res = _run(qapp, pair, tmp_path)
    assert not res["셀값란 Ctrl+C 가 그 값을 복사"]["ok"]


def test_it_catches_the_dead_worker_regression(qapp, pair, tmp_path, monkeypatch):
    """v221 — 끝난 저장 워커(C++ 쪽이 지워짐)에 isRunning() 을 부르면 터진다."""
    from excelmerge.diff_view import DiffView

    def v220(self):
        w = self._staged_merge_worker
        return w is not None and w.isRunning()

    monkeypatch.setattr(DiffView, "_saving_now", v220)
    res = _run(qapp, pair, tmp_path)
    assert not res["저장 뒤 Ctrl+S 에 오류 없음"]["ok"]
    assert not res["처리되지 않은 예외 없음"]["ok"]


# ── 첫 변경 셀의 A 쪽이 비어도 앱 탓으로 실패하지 않는다 ────────────────────
# 예전 점검은 Alt+Down 한 번에 선 칸의 셀값란을 봤다. 그 칸이 B 에만 있는 열·행이면 A 쪽이
# 비어 '셀값란에 값이 뜸'이 실패했다 — 앱은 멀쩡한데 release.py 가 막혔다.

_MODIFIED_A = "대여 만료 아이템 보상"     # 아래 쌍에서 A 쪽 값이 있는 유일한 변경 셀

LAYOUTS = {
    # 열을 이름으로 맞춰 New 가 가운데 열이 된다 → 머리글부터 매 행 첫 변경이 B 전용 칸
    "B 전용 열": ([["ID", "Desc"], ["k1", "같은 값"], ["k2", _MODIFIED_A], ["k3", "입문 보상"]],
                 [["ID", "New", "Desc"], ["k1", "n1", "같은 값"],
                  ["k2", "n2", "기간 만료 아이템 보상"], ["k3", "n3", "입문 보상"]]),
    # B 전용 행 둘 — 비교 직후 현재 셀은 본체에 보이는 첫 칸(첫 B 전용 행)이라, Alt+Down 이
    # 다음으로 가는 칸도 B 전용 행이다. 하나만 끼우면 그 다음이 바로 A 값 있는 칸이라 못 잡는다.
    "B 전용 행": ([["ID", "Desc"], ["k1", "같은 값"], ["k2", _MODIFIED_A]],
                 [["ID", "Desc"], ["k1", "같은 값"], ["n1", "새 행"], ["n2", "새 행 2"],
                  ["k2", "기간 만료 아이템 보상"]]),
    # 변경 셀이 전부 A 쪽이 비었다 — Alt+Down 으로 닿을 칸이 없다
    "B 에 행만 더함": ([["ID", "Desc"], ["k1", "같은 값"]],
                     [["ID", "Desc"], ["k1", "같은 값"], ["k2", "새 행"]]),
}


def _layout(tmp_path, name):
    src = tmp_path / "layout"
    src.mkdir()
    rows_a, rows_b = LAYOUTS[name]
    return _xlsx(src / "a.xlsx", rows_a), _xlsx(src / "b.xlsx", rows_b)


def _old_landing(view, table):
    """고치기 전의 착지 — 포커스를 주고 Alt+Down 한 번."""
    from PyQt5.QtCore import Qt
    table.setFocus()
    selfcheck._press(table, Qt.Key_Down, Qt.AltModifier)
    return table._current_cell()


@pytest.mark.parametrize("name", ["B 전용 열", "B 전용 행"])
def test_the_pair_trips_the_old_landing(qapp, tmp_path, monkeypatch, name):
    """이 쌍들이 실제로 그 버그를 재현하는지 — 아니면 아래 통과 테스트가 아무것도 증명하지 않는다."""
    monkeypatch.setattr(selfcheck, "_land_on_a_value", _old_landing)
    res = _run(qapp, _layout(tmp_path, name), tmp_path)
    assert not res["셀값란에 값이 뜸"]["ok"]


@pytest.mark.parametrize("name", ["B 전용 열", "B 전용 행"])
def test_a_b_only_first_change_still_passes(qapp, tmp_path, name):
    res = _run(qapp, _layout(tmp_path, name), tmp_path)
    bad = {k: v["detail"] for k, v in res.items() if not v["ok"]}
    assert not bad, bad
    # A 쪽 값이 있는 **변경** 셀에 섰다 — 아무 칸이 아니라
    assert res["셀값란에 값이 뜸"]["detail"] == _MODIFIED_A
    assert res["셀값란 Ctrl+C 가 그 값을 복사"]["detail"] == repr(_MODIFIED_A)


def test_changes_only_on_the_b_side_use_an_unchanged_cell(qapp, tmp_path):
    """변경 셀이 모두 A 쪽이 비면 A 쪽 값이 있는 첫 칸으로 — 셀값란 복사는 변경 여부와 무관하다."""
    res = _run(qapp, _layout(tmp_path, "B 에 행만 더함"), tmp_path)
    bad = {k: v["detail"] for k, v in res.items() if not v["ok"]}
    assert not bad, bad
    assert "셀값란 Ctrl+C 가 그 값을 복사" in res


@pytest.mark.parametrize("name", [None, "B 전용 열"])
def test_the_landing_still_goes_through_alt_down(qapp, pair, tmp_path, monkeypatch, name):
    """Alt+Down 이 안 먹으면 셀값란이 빈 채로 남아 실패해야 한다 — 착지는 실제 키 경로다.

    None = 첫 변경 셀에 A 값이 있는 기본 쌍(아무 칸에도 서지 않고 누름),
    'B 전용 열' = A 쪽이 빈 앞 변경 셀에 서 두고 누름.
    """
    from excelmerge.diff_view import DiffView
    monkeypatch.setattr(DiffView, "_on_next_diff_shortcut", lambda self: None)
    res = _run(qapp, pair if name is None else _layout(tmp_path, name), tmp_path)
    assert not res["셀값란에 값이 뜸"]["ok"]


def test_a_b_only_first_change_still_catches_the_copy_regression(qapp, tmp_path, monkeypatch):
    """새 착지에서도 v219 를 잡는다."""
    from PyQt5.QtWidgets import QPlainTextEdit
    from excelmerge.widgets import CellEditWidget
    monkeypatch.setattr(CellEditWidget, "event", QPlainTextEdit.event)
    res = _run(qapp, _layout(tmp_path, "B 전용 열"), tmp_path)
    assert res["셀값란에 값이 뜸"]["ok"]
    assert not res["셀값란 Ctrl+C 가 그 값을 복사"]["ok"]


# ── 사용자의 것을 건드리지 않는다 ───────────────────────────────────────────

def test_the_originals_are_only_read(qapp, pair, tmp_path):
    before = [_sha(p) for p in pair]
    _run(qapp, pair, tmp_path)
    assert [_sha(p) for p in pair] == before


def test_message_boxes_are_restored_afterwards(qapp, pair, tmp_path):
    """알림 창을 가로챈 채 남기면 같은 프로세스의 다른 코드가 알림을 못 띄운다."""
    from PyQt5.QtWidgets import QMessageBox
    before = {k: QMessageBox.__dict__[k] for k in ("information", "warning", "critical", "question")}
    _run(qapp, pair, tmp_path)
    after = {k: QMessageBox.__dict__[k] for k in before}
    assert after == before


def test_isolation_moves_settings_and_logs_to_a_temp_folder(monkeypatch):
    monkeypatch.setenv("APPDATA", "C:/사용자설정")
    monkeypatch.setenv("QT_QPA_PLATFORM", "windows")
    selfcheck.isolate_environment()
    try:
        assert os.path.basename(os.environ["APPDATA"]).startswith("excelmerge_selfcheck_")
        assert os.environ["QT_QPA_PLATFORM"] == "offscreen", "화면에 창이 뜬다"
    finally:
        os.rmdir(os.environ["APPDATA"])


def test_the_entry_isolates_before_the_app_is_imported():
    """진단 로그 경로는 import 순간 정해진다 — 격리가 그 뒤면 사용자 로그에 쓴다."""
    with open(os.path.join(ROOT, "excel_diff_merge.py"), encoding="utf-8") as f:
        src = f.read()
    assert src.index("_selfcheck.isolate_environment()") < \
        src.index("from excelmerge.main_window import MainWindow")


def test_the_module_imports_nothing_from_the_app_at_the_top():
    """selfcheck 를 import 하는 것만으로 로그가 열리면 격리가 늦는다."""
    with open(os.path.join(ROOT, "excelmerge", "selfcheck.py"), encoding="utf-8") as f:
        tree = ast.parse(f.read())
    top = [n for n in tree.body if isinstance(n, (ast.Import, ast.ImportFrom))]
    names = [a.name for n in top for a in n.names] + \
            [n.module or "" for n in top if isinstance(n, ast.ImportFrom)]
    assert not [n for n in names if n.startswith(("excelmerge", "PyQt5"))], names


# ── 진입점 ──────────────────────────────────────────────────────────────────

def test_main_writes_the_result_and_the_exit_code(qapp, pair, tmp_path):
    """창 없는(windowed) exe 라 표준 출력이 없다 — 게이트는 JSON 과 종료 코드로 판단한다."""
    from excelmerge import __version__
    out = tmp_path / "r.json"
    code = selfcheck.main(["ExcelMerge.exe", "--selfcheck", pair[0], pair[1], str(out)])
    res = json.loads(out.read_text(encoding="utf-8"))
    assert code == 0 and res["ok"] is True
    assert res["version"] == __version__


def test_main_without_arguments_is_an_error():
    assert selfcheck.main(["ExcelMerge.exe", "--selfcheck"]) == 2
