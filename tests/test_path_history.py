# -*- coding: utf-8 -*-
"""A/B 경로칸 히스토리 드롭다운 — 최근에 연 경로를 골라 바로 불러온다.

  저장   쪽(a/b)마다 최신순 20개, 같은 경로는 하나(대소문자 무시), 임시 폴더 경로는 안 남김
         (P4V 가 %TEMP%/p4v 에 내려받는 리비전은 곧 지워진다)
  기록   실제로 불러올 때 — P4V 명령줄(load_pair) · 드롭/찾아보기/Enter · 폴더 비교
  펼침   ▼ 클릭 · F4 · Alt+↓. Alt+↓ 는 '다음 변경' 단축키라 **경로칸에서만** 가져간다
  보임   지금 있는 경로만, 그 화면이 다루는 종류만(폴더 비교는 폴더만)

전부 화면 없이 돈다. 설정은 conftest 가 테스트마다 임시 APPDATA 로 격리한다.
"""
import json
import os

import pytest
from PyQt5.QtCore import Qt
from PyQt5.QtTest import QTest
from PyQt5.QtWidgets import QApplication, QMenu

from excelmerge import prefs


@pytest.fixture
def not_transient(monkeypatch, tmp_path):
    """테스트 파일은 임시 폴더(tmp_path)에 있다 — '곧 지워질 경로' 판정에서 빼 준다."""
    other = tmp_path / "__진짜_임시_아님__"
    monkeypatch.setattr(prefs, "_transient_roots", lambda: [str(other)])


def _files(tmp_path, *names):
    out = []
    for n in names:
        p = tmp_path / n
        if n.endswith("/"):
            p.mkdir(parents=True, exist_ok=True)
        else:
            # 진짜 엑셀 파일이어야 한다 — 고르면 실제로 불러오므로, 가짜면 오류 창이 뜬다
            # (화면 없는 테스트에서 그 모달 창이 접근 위반으로 프로세스를 죽였다).
            import openpyxl
            p.parent.mkdir(parents=True, exist_ok=True)
            wb = openpyxl.Workbook()
            wb.active.append(["ID", "V"])
            wb.active.append(["k1", n])
            wb.save(str(p))
        out.append(str(p).rstrip("\\/"))
    return out


# ── 저장 ─────────────────────────────────────────────────────────────────────

def test_newest_first_and_no_duplicates(not_transient, tmp_path):
    a1, a2 = _files(tmp_path, "a1.xlsx", "a2.xlsx")
    prefs.add_path_history("a", a1)
    prefs.add_path_history("a", a2)
    prefs.add_path_history("a", a1.upper())            # 윈도우는 대소문자를 안 가린다
    h = prefs.load_path_history("a")
    assert [p.lower() for p in h] == [a1.lower(), a2.lower()]
    assert h[0] == os.path.abspath(a1.upper()), "다시 연 경로의 표기가 앞에 와야 한다"


def test_sides_are_separate(not_transient, tmp_path):
    a, b = _files(tmp_path, "a.xlsx", "b.xlsx")
    prefs.add_path_history("a", a)
    prefs.add_path_history("b", b)
    assert prefs.load_path_history("a") == [a]
    assert prefs.load_path_history("b") == [b]


def test_keeps_only_the_newest_twenty(not_transient, tmp_path):
    for i in range(25):
        prefs.add_path_history("a", str(tmp_path / f"f{i}.xlsx"))
    h = prefs.load_path_history("a")
    assert len(h) == 20 and h[0].endswith("f24.xlsx") and h[-1].endswith("f5.xlsx")


def test_temporary_paths_are_not_kept(tmp_path):
    """P4V 가 비교용으로 내려받는 리비전 — 짧은 이름(GYWNSD~1) 표기로 와도 걸러야 한다."""
    import tempfile
    t = tempfile.gettempdir()
    prefs.add_path_history("a", os.path.join(t, "p4v", "NX", "DT_ContentsList#5.uasset"))
    assert prefs.load_path_history("a") == []
    assert prefs._is_transient_path(os.path.join(t, "p4v", "x.xlsx"))
    assert not prefs._is_transient_path(r"D:\Build\GameData\ExcelData\x.xlsx")


def test_broken_settings_do_not_crash(not_transient, tmp_path):
    """손으로 고쳤거나 옛 형식인 prefs 도 견딘다 — 설정은 부가 기능이다."""
    p = prefs._prefs_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    for bad in ({"path_history": "문자열"}, {"path_history": {"a": "문자열"}},
                {"path_history": {"a": [1, None, "", "C:/ok.xlsx"]}}):
        with open(p, "w", encoding="utf-8") as f:
            json.dump(bad, f)
        h = prefs.load_path_history("a")
        assert all(isinstance(x, str) and x for x in h)
        prefs.add_path_history("a", str(tmp_path / "n.xlsx"))
        assert prefs.load_path_history("a")[0].endswith("n.xlsx")


def test_clear(not_transient, tmp_path):
    a, = _files(tmp_path, "a.xlsx")
    prefs.add_path_history("a", a)
    prefs.add_path_history("b", a)
    prefs.clear_path_history("a")
    assert prefs.load_path_history("a") == [] and prefs.load_path_history("b") == [a]


# ── 기록: 실제로 불러올 때 ──────────────────────────────────────────────────

@pytest.fixture
def win(qapp):
    from excelmerge.main_window import MainWindow
    w = MainWindow()
    w.show()
    yield w
    w.close(); w.deleteLater(); qapp.processEvents()


def test_p4v_command_line_pair_is_recorded(not_transient, tmp_path, win, qapp):
    a, b = _files(tmp_path, "A/Data_X.xlsx", "B/Data_X.xlsx")
    win.open_file_compare(a, b)
    qapp.processEvents()
    assert prefs.load_path_history("a")[0] == a
    assert prefs.load_path_history("b")[0] == b


def test_p4v_temp_revisions_are_not_recorded(tmp_path, win, qapp):
    """P4V 가 넘긴 임시 리비전은 기록하지 않는다(진짜 임시 폴더 판정 그대로)."""
    a, b = _files(tmp_path, "A/x.xlsx", "B/x.xlsx")      # tmp_path 는 시스템 임시 폴더 아래
    win.open_file_compare(a, b)
    qapp.processEvents()
    assert prefs.load_path_history("a") == [] and prefs.load_path_history("b") == []


def test_a_dropped_file_is_recorded(not_transient, tmp_path, win, qapp):
    a, = _files(tmp_path, "drop/a.xlsx")
    view = win.tabs.currentWidget()
    view.panel_a._on_file_dropped(a)                     # 드롭·찾아보기·Enter 가 모두 여기로
    qapp.processEvents()
    assert prefs.load_path_history("a") == [a]


def test_folder_compare_is_recorded(not_transient, tmp_path, win, qapp):
    da, db = _files(tmp_path, "FA/", "FB/")
    win.open_folder_compare(da, db)
    qapp.processEvents()
    assert prefs.load_path_history("a")[0] == da
    assert prefs.load_path_history("b")[0] == db


# ── 펼침 · 고르기 ───────────────────────────────────────────────────────────

def _open_menu(hist):
    hist.show_menu()
    QApplication.instance().processEvents()
    m = QApplication.activePopupWidget()
    assert isinstance(m, QMenu), "메뉴가 안 떴다"
    return m


def _labels(menu):
    return [a.text() for a in menu.actions() if not a.isSeparator()]


def _close(menu):
    menu.close()
    QApplication.instance().processEvents()


def test_the_list_shows_existing_paths_newest_first(not_transient, tmp_path, win):
    a1, a2, gone = _files(tmp_path, "a1.xlsx", "a2.xlsx", "gone.xlsx")
    for p in (gone, a1, a2):
        prefs.add_path_history("a", p)
    os.remove(gone)
    m = _open_menu(win.tabs.currentWidget().panel_a.history)
    try:
        assert _labels(m) == [a2, a1, "목록 지우기"], "없는 경로는 숨기고 최신순"
        assert m.activeAction() is m.actions()[0], "펼치자마자 Enter 가 가장 최근 경로여야 한다"
    finally:
        _close(m)
    assert gone in prefs.load_path_history("a"), "숨기기만 하고 지우지는 않는다(드라이브가 잠시 빠졌을 수 있다)"


def test_picking_a_file_loads_it(not_transient, tmp_path, win, qapp, monkeypatch):
    a, = _files(tmp_path, "pick/a.xlsx")
    prefs.add_path_history("a", a)
    view = win.tabs.currentWidget()
    loaded = []
    view.panel_a.file_loaded.connect(loaded.append)
    m = _open_menu(view.panel_a.history)
    QTest.keyClick(m, Qt.Key_Return)                     # 가장 최근 경로
    qapp.processEvents()
    assert view.panel_a.get_path() == a and loaded == [a]


def test_picking_a_folder_in_the_file_view_opens_folder_compare(not_transient, tmp_path, win, qapp):
    d, = _files(tmp_path, "FA/")
    prefs.add_path_history("a", d)
    view = win.tabs.currentWidget()
    asked = []
    view.panel_a.folder_loaded.connect(asked.append)
    m = _open_menu(view.panel_a.history)
    m.actions()[0].trigger()
    qapp.processEvents()
    assert asked == [d]


def test_the_folder_view_lists_only_folders(not_transient, tmp_path, win, qapp):
    f, d = _files(tmp_path, "f.xlsx", "FA/")
    da, db = _files(tmp_path, "X/", "Y/")
    prefs.add_path_history("a", f)
    prefs.add_path_history("a", d)
    fv = win.open_folder_compare(da, db) or win.tabs.currentWidget()
    qapp.processEvents()
    m = _open_menu(fv.history_a)
    try:
        labels = _labels(m)
        assert f not in labels and d in labels and da in labels
    finally:
        _close(m)


def test_empty_history_says_so(win):
    m = _open_menu(win.tabs.currentWidget().panel_b.history)
    try:
        acts = [a for a in m.actions() if not a.isSeparator()]
        assert len(acts) == 1 and not acts[0].isEnabled() and "없습니다" in acts[0].text()
    finally:
        _close(m)


def test_clear_from_the_menu(not_transient, tmp_path, win, qapp):
    a, = _files(tmp_path, "a.xlsx")
    prefs.add_path_history("a", a)
    m = _open_menu(win.tabs.currentWidget().panel_a.history)
    [x for x in m.actions() if x.text() == "목록 지우기"][0].trigger()
    qapp.processEvents()
    assert prefs.load_path_history("a") == []


# ── 키보드: F4 · Alt+↓ — 그리고 표의 Alt+↓ 는 그대로 '다음 변경' ─────────────

def _next_diff_fired(view):
    """'다음 변경'(Alt+↓) 단축키가 실제로 발동했는지 센다 — 단축키 객체의 신호를 직접 본다.
    (슬롯 메서드를 바꿔 끼우면 안 된다. 생성 때 바운드 메서드로 이미 연결돼 있다.)"""
    from PyQt5.QtGui import QKeySequence
    from PyQt5.QtWidgets import QShortcut
    scs = [sc for sc in view.findChildren(QShortcut) if sc.key() == QKeySequence("Alt+Down")]
    assert scs, "전제: 비교 화면에 Alt+↓ 단축키가 있어야 한다"
    fired = []
    for sc in scs:
        sc.activated.connect(lambda: fired.append(1))
    return fired


def _alt_down(w):
    QTest.keyClick(w, Qt.Key_Down, Qt.AltModifier)
    QTest.keyRelease(w, Qt.Key_Alt, Qt.NoModifier)
    QApplication.instance().processEvents()


@pytest.mark.parametrize("key,mod", [(Qt.Key_F4, Qt.NoModifier), (Qt.Key_Down, Qt.AltModifier)])
def test_keys_open_the_list_from_the_path_field(win, qapp, key, mod):
    view = win.tabs.currentWidget()
    fired = _next_diff_fired(view)
    edit = view.panel_a.path_edit
    edit.setFocus()
    qapp.processEvents()
    QTest.keyClick(edit, key, mod)
    if mod == Qt.AltModifier:
        QTest.keyRelease(edit, Qt.Key_Alt, Qt.NoModifier)
    qapp.processEvents()
    m = QApplication.activePopupWidget()
    try:
        assert isinstance(m, QMenu), "경로칸에서 눌렀는데 목록이 안 떴다"
        assert not fired, "경로칸의 Alt+↓ 가 '다음 변경'으로 새어 나갔다"
    finally:
        if m is not None:
            _close(m)


def test_alt_down_in_the_table_still_jumps_to_the_next_change(win, qapp):
    """경로칸이 Alt+↓ 를 가져가는 건 경로칸에 커서가 있을 때뿐이다."""
    view = win.tabs.currentWidget()
    fired = _next_diff_fired(view)
    t = view.panel_a.table
    t.setFocus()
    qapp.processEvents()
    _alt_down(t)
    assert fired, "표에서 Alt+↓ 가 '다음 변경'으로 가지 않았다"
    assert QApplication.activePopupWidget() is None, "표에서 Alt+↓ 로 경로 목록이 떴다"


def test_the_arrow_is_on_both_path_fields(win, qapp, tmp_path):
    view = win.tabs.currentWidget()
    for panel in (view.panel_a, view.panel_b):
        assert panel.history.action in panel.path_edit.actions()


# ── 검토에서 확인된 문제들 ───────────────────────────────────────────────────

_CHILD = r"""
import os, sys, time
sys.path.insert(0, sys.argv[1])
from excelmerge import prefs
prefs._transient_roots = lambda: [os.path.join(sys.argv[1], "__없음__")]
go = sys.argv[2]
while not os.path.exists(go):
    time.sleep(0.001)
prefs.add_path_history("b", sys.argv[3])
"""


def test_processes_starting_together_keep_every_setting(tmp_path):
    """P4V 에서 여러 파일을 한꺼번에 Diff 하면 ExcelMerge 가 거의 동시에 여러 개 뜬다.

    예전 쓰기(파일을 비우고 쓰기 + 잠금 없음)는 서로 덮어써 히스토리를 잃고, 비워진 파일을
    읽은 쪽이 {} 로 다시 써서 **파일별 키·제외 열·시트 기억까지 통째로** 날렸다(검토 재현:
    6개 동시 시작 10회 중 2회). 진짜 프로세스로 똑같이 몰아 본다.
    """
    import subprocess
    import sys
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    seed = {"key_row": 2, "key_col": 3,
            "last_keys": {f"C:/k{i}.xlsx": [0, 1, "ID"] for i in range(50)},
            "last_sheets": {f"C:/s{i}.xlsx": "Data" for i in range(50)},
            "last_excluded": {f"C:/e{i}.xlsx": ["#Desc"] for i in range(50)}}
    child = tmp_path / "child.py"
    child.write_text(_CHILD, encoding="utf-8")
    env = dict(os.environ)                 # conftest 가 격리한 APPDATA 를 그대로 물려준다
    for round_ in range(3):
        p = prefs._prefs_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(seed, f)
        go = tmp_path / f"go{round_}"
        paths = [str(tmp_path / f"r{round_}_{i}.xlsx") for i in range(6)]
        procs = [subprocess.Popen([sys.executable, str(child), root, str(go), q], env=env)
                 for q in paths]
        go.write_text("go", encoding="utf-8")      # 동시에 출발
        for pr in procs:
            assert pr.wait(timeout=60) == 0
        with open(p, encoding="utf-8") as f:
            got = json.load(f)
        assert (got.get("key_row"), got.get("key_col")) == (2, 3), f"{round_}회차: 전역 키가 날아갔다"
        for k in ("last_keys", "last_sheets", "last_excluded"):
            assert len(got.get(k, {})) == 50, f"{round_}회차: {k} 가 날아갔다"
        hist = [e[0] for e in got["path_history"]["b"]]
        missing = [q for q in paths if os.path.abspath(q) not in hist]
        assert not missing, f"{round_}회차: 히스토리 유실 {len(missing)}/6"
    assert not [n for n in os.listdir(os.path.dirname(p)) if n.endswith(".tmp")], "임시 파일이 남았다"


def test_a_corrupt_settings_file_is_kept_aside_not_silently_lost(tmp_path):
    """정말 깨진 prefs.json 은 원본을 남겨 두고 새로 쓴다 — 저장을 영영 막지도, 조용히
    날리지도 않는다."""
    p = prefs._prefs_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write("{ 반쯤 쓰인 파일")
    prefs.save_key_prefs(1, 4)
    assert prefs.load_key_prefs() == (1, 4)
    backups = [n for n in os.listdir(os.path.dirname(p)) if n.startswith("prefs.json.corrupt-")]
    assert len(backups) == 1
    with open(os.path.join(os.path.dirname(p), backups[0]), encoding="utf-8") as f:
        assert f.read() == "{ 반쯤 쓰인 파일"


def test_history_writes_keep_the_other_settings(not_transient, tmp_path):
    a, = _files(tmp_path, "a.xlsx")
    prefs.save_key_prefs(1, 2)
    prefs.save_last_key(a, 0, 3, "ID")
    prefs.save_last_sheet(a, "Data")
    prefs.save_last_excluded(a, ["#Desc"])
    prefs.add_path_history("a", a)
    prefs.clear_path_history("a")
    assert prefs.load_key_prefs() == (1, 2)
    assert prefs.load_last_key(a) == (0, 3)
    assert prefs.load_last_sheet(a) == "Data"
    assert prefs.load_last_excluded(a) == ["#Desc"]


def test_reopening_the_newest_path_does_not_rewrite(not_transient, tmp_path, monkeypatch):
    """켤 때마다 쓰는 일을 줄인다 — 이미 맨 앞이면 파일을 건드리지 않는다."""
    a, = _files(tmp_path, "a.xlsx")
    prefs.add_path_history("a", a)
    writes = []
    real = prefs._write_prefs
    monkeypatch.setattr(prefs, "_write_prefs", lambda d: (writes.append(1), real(d)))
    prefs.add_path_history("a", a)
    assert writes == []


def test_files_cannot_push_folders_out(not_transient, tmp_path):
    """종류마다 20개 — 파일을 많이 연 날에도 폴더 비교의 ▼ 가 비지 않는다."""
    dirs = _files(tmp_path, "D1/", "D2/")
    for d in dirs:
        prefs.add_path_history("a", d)
    for i in range(30):
        prefs.add_path_history("a", str(tmp_path / f"f{i}.xlsx"))
    assert prefs.load_path_history("a", ("d",)) == dirs[::-1]
    assert len(prefs.load_path_history("a", ("f",))) == 20


def test_files_opened_inside_folder_compare_are_not_recorded(not_transient, tmp_path, win, qapp):
    """폴더를 훑으며 연 파일들이 사용자가 직접 연 파일을 밀어내지 않게."""
    da, db = _files(tmp_path, "FA/", "FB/")
    fa, fb = _files(tmp_path, "FA/x.xlsx", "FB/x.xlsx")
    fv = win.open_folder_compare(da, db)
    qapp.processEvents()
    fv.open_file_pair.emit(fa, fb, "x.xlsx")
    qapp.processEvents()
    assert prefs.load_path_history("a", ("f",)) == [] and prefs.load_path_history("b", ("f",)) == []
    assert prefs.load_path_history("a", ("d",)) == [da]


def test_clear_in_the_folder_view_keeps_the_file_entries(not_transient, tmp_path, win, qapp):
    f, = _files(tmp_path, "f.xlsx")
    da, db = _files(tmp_path, "FA/", "FB/")
    prefs.add_path_history("a", f)
    fv = win.open_folder_compare(da, db)
    qapp.processEvents()
    m = _open_menu(fv.history_a)
    [x for x in m.actions() if x.text() == "목록 지우기"][0].trigger()
    qapp.processEvents()
    assert prefs.load_path_history("a", ("d",)) == []
    assert prefs.load_path_history("a", ("f",)) == [f], "보이지도 않던 파일 기록까지 지웠다"


def test_clicking_the_arrow_opens_and_a_second_click_closes(not_transient, tmp_path, win, qapp):
    """▼ 를 다시 누르면 닫혀야 한다. Qt 는 메뉴 밖 클릭으로 메뉴를 닫은 뒤 그 클릭을 ▼ 에
    다시 전달해 곧바로 또 열었다(검토에서 확인)."""
    a, = _files(tmp_path, "a.xlsx")
    prefs.add_path_history("a", a)
    hist = win.tabs.currentWidget().panel_a.history
    hist.action.trigger()                               # 첫 클릭 — 열린다
    qapp.processEvents()
    m = QApplication.activePopupWidget()
    assert isinstance(m, QMenu)
    m.hide()                                            # 메뉴 밖(▼) 클릭이 먼저 메뉴를 닫고
    qapp.processEvents()
    hist.action.trigger()                               # 그 클릭이 ▼ 에 다시 전달된다
    qapp.processEvents()
    assert QApplication.activePopupWidget() is None, "닫으려고 누른 ▼ 가 다시 열었다"
    hist._hidden_at -= 1.0                              # 시간이 지난 뒤의 새 클릭은 연다
    hist.action.trigger()
    qapp.processEvents()
    m = QApplication.activePopupWidget()
    assert isinstance(m, QMenu)
    _close(m)


def test_an_ampersand_in_a_path_is_shown_and_loaded_as_is(not_transient, tmp_path, win, qapp):
    """'&' 는 메뉴에서 단축 글자 표시다 — 이스케이프하지 않으면 'R&D' 가 'RD' 로 보였다."""
    a, = _files(tmp_path, "R&D/a.xlsx")
    prefs.add_path_history("a", a)
    view = win.tabs.currentWidget()
    m = _open_menu(view.panel_a.history)
    act = m.actions()[0]
    assert act.text() == a.replace("&", "&&")
    loaded = []
    view.panel_a.file_loaded.connect(loaded.append)
    act.trigger()
    qapp.processEvents()
    assert loaded == [a]


_WRITER = r"""
import os, sys
sys.path.insert(0, sys.argv[1])
from excelmerge import prefs
for i in range(300):
    prefs.save_key_prefs(2, 3)
"""


def test_a_reader_never_sees_a_half_written_file(tmp_path):
    """비교를 열 때 키 위치는 잠금 없이 읽는다(load_last_key 등). 다른 프로세스가 쓰는 도중에
    읽어도 빈/반쪽 파일을 보면 안 된다 — 보면 기본 키(A1)로 비교해 행이 조용히 빠진다.
    파일을 비우고 쓰던 예전 방식은 여기서 빈 설정을 읽혔다. 원자적 교체가 그것을 막는다."""
    import subprocess
    import sys
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    p = prefs._prefs_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"key_row": 2, "key_col": 3,
                   "last_keys": {f"C:/k{i}.xlsx": [0, 1, "ID"] for i in range(400)}}, f)
    w = tmp_path / "writer.py"
    w.write_text(_WRITER, encoding="utf-8")
    proc = subprocess.Popen([sys.executable, str(w), root], env=dict(os.environ))
    reads = bad = partial = 0
    while proc.poll() is None:
        c = prefs._read_prefs()
        reads += 1
        if len(c.get("last_keys", {})) != 400 or (c.get("key_row"), c.get("key_col")) != (2, 3):
            bad += 1
        # 디스크의 파일 자체도 본다 — 재시도 없이 한 번 읽는 쪽(옛 버전 앱)이 반쪽 내용을
        # 보면 안 된다. 교체 순간 열기가 거부되는 것은 괜찮다(내용을 본 게 아니다).
        try:
            with open(p, encoding="utf-8") as f:
                text = f.read()
        except OSError:
            continue
        try:
            json.loads(text)
        except ValueError:
            partial += 1
    assert proc.wait(timeout=60) == 0
    assert reads > 20, f"전제: 쓰는 동안 충분히 읽어야 한다({reads}회)"
    assert bad == 0, f"쓰는 도중 {bad}/{reads}회 빈·반쪽 설정을 읽었다"
    assert partial == 0, f"디스크에서 반쪽 파일이 {partial}/{reads}회 보였다 — 원자적으로 바뀌지 않는다"


def test_a_corrupt_settings_file_does_not_stall_every_read():
    """깨진 prefs.json 을 매번 여러 번 다시 읽으며 기다리면, 고쳐질 때까지 실행할 때마다 1초 넘게
    멈췄다(한 번 실행에 읽기가 9번쯤 일어난다). 깨진 내용은 한 번만 다시 본다."""
    import time
    p = prefs._prefs_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write("{ 깨짐")
    t0 = time.perf_counter()
    for _ in range(9):
        assert prefs.load_key_prefs() == (0, 0)
    took = time.perf_counter() - t0
    assert took < 0.6, f"깨진 설정을 9번 읽는 데 {took:.2f}초"


def test_a_read_only_settings_file_gives_up_at_once():
    """읽기 전용이면 기다려도 안 풀린다 — 저장할 때마다 UI 를 멈추지 않고 바로 포기한다."""
    import stat
    import time
    prefs.save_key_prefs(1, 1)
    p = prefs._prefs_path()
    os.chmod(p, stat.S_IREAD)
    try:
        t0 = time.perf_counter()
        prefs.save_key_prefs(2, 2)
        took = time.perf_counter() - t0
        assert took < 0.15, f"읽기 전용 설정 저장에 {took:.2f}초"
        assert prefs.load_key_prefs() == (1, 1)
        assert not [n for n in os.listdir(os.path.dirname(p)) if n.endswith(".tmp")]
    finally:
        os.chmod(p, stat.S_IREAD | stat.S_IWRITE)
