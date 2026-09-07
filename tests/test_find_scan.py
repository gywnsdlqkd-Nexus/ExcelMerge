# -*- coding: utf-8 -*-
"""찾기 스캔 회귀 — **화면에 보이는 텍스트만** 찾아야 한다.

diff 모드 찾기는 성능 때문에 _diff_matrix 튜플을 직접 읽는다
([_iter_find_matches_diff](../excelmerge/diff_view.py)) — 셀당 모델 display_text() 를
양쪽에 호출하면 449K셀 표에서 파이썬 메서드 호출이 90만 회가 되어 F3 한 번에 640ms 가
들었다(직접 읽기 94ms).

그 대가로 "무엇이 화면에 보이는 텍스트인가"를 스캔이 직접 알아야 한다:

  · 일반 셀   → A 패널은 a_val, B 패널은 b_val → 둘 중 하나라도 일치하면 매치
  · staged 셀 → display_text 가 **양쪽 패널에 staged 방향의 값 하나만** 준다
                → 그 값으로만 판정해야 한다

두 번째를 빼먹으면 화면에 없는 쪽 값에 일치해 '보이지 않는 글자를 찾는' 오동작이 된다
(실제 파일에서 확인: 우회 제거 시 13개 셀 오검출).
"""
import pytest
from PyQt5.QtWidgets import QApplication

from excelmerge.constants import DIR_A2B, DIR_B2A

HEADER = ["ID", "V"]
# display 행 1 = 키 "1" → V 열이 ("modified", "alpha", "beta")
A_DATA = [HEADER, ["1", "alpha"], ["2", "gamma"]]
B_DATA = [HEADER, ["1", "beta"], ["2", "gamma"]]
CELL = (1, 1)


def _wait_diff(dv, timeout_ms=5000):
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    QApplication.instance().processEvents()


@pytest.fixture
def dv(qapp):
    """A/B 가 V 열 한 칸만 다른 활성 DiffView."""
    from excelmerge.main_window import MainWindow
    win = MainWindow()
    win.show()
    try:
        view = win.tabs.currentWidget()
        view.panel_a.set_path("a.xlsx")
        view.panel_b.set_path("b.xlsx")
        view._key_col, view._key_row = 0, 0
        for p in (view.panel_a, view.panel_b):
            p.table.set_key_col(0)
            p.table.set_key_row(0)
        view._on_loaded(A_DATA, B_DATA)
        _wait_diff(view)
        assert view._diff_matrix[CELL[0]][CELL[1]][1:] == ("alpha", "beta"), \
            f"전제 불일치: {view._diff_matrix[CELL[0]][CELL[1]]}"
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _find(dv, term):
    return set(dv._iter_find_matches(dv._make_find_matcher(term)))


def test_unstaged_cell_matches_both_sides(dv):
    """병합 준비 전 — A 패널은 alpha, B 패널은 beta 를 보여주므로 둘 다 찾힌다."""
    assert CELL in _find(dv, "alpha")
    assert CELL in _find(dv, "beta")


def test_staged_a2b_searches_only_a_value(dv):
    """A→B 준비 후 양쪽 패널이 'alpha' 를 보여준다 → 'beta' 는 찾히면 안 된다."""
    dv._staged[CELL] = DIR_A2B
    ma, mb = dv.panel_a.table.model(), dv.panel_b.table.model()
    assert ma.display_text(*CELL) == mb.display_text(*CELL) == "alpha", \
        "전제: staged 셀은 양쪽 패널이 같은 값을 보여준다"
    assert CELL in _find(dv, "alpha")
    assert CELL not in _find(dv, "beta"), \
        "화면에 없는 B 쪽 값('beta')에 일치했다 — staged 우회가 빠졌다"


def test_staged_b2a_searches_only_b_value(dv):
    """대칭 — B→A 준비 후에는 'beta' 만 찾힌다."""
    dv._staged[CELL] = DIR_B2A
    assert CELL in _find(dv, "beta")
    assert CELL not in _find(dv, "alpha"), \
        "화면에 없는 A 쪽 값('alpha')에 일치했다"


def test_real_stage_path_agrees(dv):
    """우클릭 병합 준비 경로(_stage_selected)를 실제로 타도 같은 결과인지."""
    dv.panel_a.table._select_range(*CELL, *CELL)
    QApplication.instance().processEvents()
    dv._stage_selected(DIR_A2B)
    QApplication.instance().processEvents()
    assert dv._staged.get(CELL) == DIR_A2B, dv._staged
    assert CELL not in _find(dv, "beta")


def _oracle(dv, match):
    """수정 전 구현 — 모델 display_text 로 판정하는 정본."""
    tbl_a = dv.panel_a.table
    models = (tbl_a.model(), dv.panel_b.table.model())
    rows = max(m.data_rows for m in models)
    cols = max(m.data_cols for m in models)
    out = []
    for r in range(rows):
        if tbl_a.isRowHidden(r):
            continue
        for c in range(cols):
            if c in dv._excluded_cols:
                continue
            for model in models:
                if match(model.display_text(r, c)):
                    out.append((r, c))
                    break
    return out


@pytest.mark.parametrize("term", ["alpha", "beta", "gamma", "1", "a", "ID", "없는값"])
@pytest.mark.parametrize("staged", [None, DIR_A2B, DIR_B2A])
@pytest.mark.parametrize("excluded", [(), (1,)])
@pytest.mark.parametrize("diff_only", [True, False])
def test_matches_display_text_oracle(dv, term, staged, excluded, diff_only):
    """숨김 행·제외 열·staged 조합 전부에서 구 구현과 결과가 같아야 한다."""
    dv.diff_only_btn.setChecked(diff_only)
    QApplication.instance().processEvents()
    dv._excluded_cols.clear()
    dv._excluded_cols.update(excluded)
    dv._staged.clear()
    if staged is not None:
        dv._staged[CELL] = staged
    match = dv._make_find_matcher(term)
    assert list(dv._iter_find_matches(match)) == _oracle(dv, match)


def test_staged_stays_shared_with_models_during_recompute(dv):
    """리셋은 .clear() 여야 한다 — 모델이 이 dict 를 **참조로** 들고 있기 때문.

    새 객체로 갈아치우면 다음 populate 까지 모델이 옛 dict 를 들고 있어, 찾기(뷰의
    _staged 를 읽는다)와 표시 텍스트(모델의 _staged)가 서로 다른 상태를 본다.
    _recompute_diff 는 DiffWorker(비동기)를 띄우므로 그 창이 실제로 열린다 —
    따라서 워커 완료를 **기다리지 않고** 그 창 안에서 검증해야 한다.
    (기다리면 populate 가 참조를 다시 맞춰 재바인딩 버그가 가려진다.)
    """
    for model in (dv.panel_a.table.model(), dv.panel_b.table.model()):
        assert model._staged is dv._staged, "populate 가 참조를 넘기지 않았다"

    dv._staged[CELL] = DIR_A2B
    match = dv._make_find_matcher("beta")
    assert CELL not in set(dv._iter_find_matches(match)), "전제: staged 셀은 'beta' 를 숨긴다"

    dv._recompute_diff()          # 워커 시작 — done 시그널은 아직 처리되지 않았다
    for model in (dv.panel_a.table.model(), dv.panel_b.table.model()):
        assert model._staged is dv._staged, (
            "재계산이 _staged 를 재바인딩해 모델과 갈라졌다 — 이 창에서 찾기와 "
            "표시 텍스트가 서로 다른 staged 를 본다")
    assert list(dv._iter_find_matches(match)) == _oracle(dv, match),         "재계산 진행 중 창에서 찾기 결과가 표시 텍스트와 어긋났다"

    _wait_diff(dv)                # 뒷정리 — 워커를 남기지 않는다
