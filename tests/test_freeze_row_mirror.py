# -*- coding: utf-8 -*-
"""틀 고정 좌측 오버레이의 행 숨김 미러 — 델타 최적화가 정합을 깨면 안 된다.

left 오버레이는 본체와 같은 행을 같은 순서로 보여야 한다. 예전엔 갱신 때마다 전 행을
훑어 맞췄는데, 그 전량 미러가 '변경점만 보기' 토글·열 제외 비용의 40~60% 였다
(실측 6,328행: 열 제외 38.5ms 중 23.4ms). **열 제외는 행 가시성을 전혀 바꾸지 않는데도**
매번 전 행을 돌았다.

지금은 _apply_diff_filter 가 실제로 뒤집힌 행만 넘긴다. 그 대가로 'left 가 본체와
동기'라는 전제가 생겼으므로, 여기서 두 가지를 고정한다:
  ① 어떤 경로로 갱신하든 미러 뒤 host/left 의 행 숨김이 **완전히 일치**한다.
  ② 그 전제가 깨질 수 있는 지점(모델 리셋·freeze 해제)에서는 델타를 무시하고 전 행을 훑는다.
"""
import pytest
from PyQt5.QtWidgets import QApplication

HEADER = ["ID", "V", "W"]
A_DATA = [HEADER] + [[str(i), "x" if i % 3 == 0 else "s", "p"] for i in range(1, 30)]
B_DATA = [HEADER] + [[str(i), "y" if i % 3 == 0 else "s", "p"] for i in range(1, 30)]


def _wait_diff(dv, timeout_ms=5000):
    w = getattr(dv, "_diff_worker", None)
    if w is not None:
        w.wait(timeout_ms)
    QApplication.instance().processEvents()


@pytest.fixture
def dv(qapp, monkeypatch):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
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
        assert view._freeze["a"].active
        yield view
    finally:
        win.close()
        win.deleteLater()
        qapp.processEvents()


def _assert_parity(dv, note=""):
    """host 와 left 의 행 숨김이 전 행에서 일치하는지 — 델타가 놓친 행이 없다는 증거."""
    for side in ("a", "b"):
        fc = dv._freeze[side]
        host = fc.host
        n = host.model().rowCount()
        mismatched = [r for r in range(n)
                      if host.isRowHidden(r) != fc.left.isRowHidden(r)]
        assert not mismatched, f"{side}: 행 {mismatched[:10]} 어긋남 {note}"


class _Spy:
    """left 오버레이가 **훑은 행 수**와 실제로 바꾼 행 수를 센다.

    setRowHidden 만 세면 전 행 스캔과 델타를 구분할 수 없다 — 전 행을 훑어도 상태가
    같은 행은 호출하지 않으므로 호출 수가 같아진다. 그래서 루프가 행마다 부르는
    isRowHidden 을 세어 '몇 행을 봤는지'를 본다.
    """

    def __init__(self, view):
        self.view = view
        self.examined = 0
        self.changed = 0
        self._orig_get = view.isRowHidden
        self._orig_set = view.setRowHidden

    def __enter__(self):
        def get(r):
            self.examined += 1
            return self._orig_get(r)

        def set_(r, h):
            self.changed += 1
            return self._orig_set(r, h)
        self.view.isRowHidden = get
        self.view.setRowHidden = set_
        return self

    def __exit__(self, *exc):
        del self.view.isRowHidden
        del self.view.setRowHidden


def test_parity_after_initial_load(dv):
    _assert_parity(dv, "(최초 로드)")


def test_column_exclude_costs_no_row_mirror(dv):
    """열 제외/해제는 행 가시성을 바꾸지 않는다 — 미러가 한 행도 건드리면 안 된다."""
    fc = dv._freeze["a"]
    with _Spy(fc.left) as spy:
        dv._on_columns_exclude_set([2], True)
    assert spy.examined == 0, f"열 제외에 행 {spy.examined}개를 훑었다"
    _assert_parity(dv, "(열 제외 ON)")

    with _Spy(fc.left) as spy:
        dv._on_columns_exclude_set([2], False)
    assert spy.examined == 0
    _assert_parity(dv, "(열 제외 OFF)")


def test_parity_across_filter_toggles(dv):
    for state in (False, True, False, True):
        dv.diff_only_btn.setChecked(state)
        QApplication.instance().processEvents()
        _assert_parity(dv, f"(변경점만 보기={state})")


def test_filter_toggle_mirrors_only_flipped_rows(dv):
    """토글은 실제로 뒤집힌 행 수만큼만 미러한다(전 행이 아니라)."""
    dv.diff_only_btn.setChecked(True)
    QApplication.instance().processEvents()
    fc = dv._freeze["a"]
    host = fc.host
    n = host.model().rowCount()
    before = [host.isRowHidden(r) for r in range(n)]

    with _Spy(fc.left) as spy:
        dv.diff_only_btn.setChecked(False)
        QApplication.instance().processEvents()
    flipped = sum(1 for r in range(n) if host.isRowHidden(r) != before[r])

    assert 0 < flipped < n, "이 데이터로는 델타/전량을 구분할 수 없다 — 테스트가 무의미"
    assert spy.changed == flipped, f"미러 {spy.changed}회 vs 실제 뒤집힘 {flipped}행"
    assert spy.examined == flipped,         f"전 행({n})을 훑었다 — 델타가 적용되지 않았다(훑은 행 {spy.examined})"
    _assert_parity(dv)


def test_delta_ignored_while_out_of_sync(dv):
    """동기 보장이 없으면(모델 리셋 직후 등) 델타를 무시하고 전 행을 훑어 복구한다."""
    fc = dv._freeze["a"]
    target = next(r for r in range(fc.host.model().rowCount())
                  if not fc.host.isRowHidden(r))
    fc.left.setRowHidden(target, True)          # 강제로 어긋나게
    assert fc.host.isRowHidden(target) != fc.left.isRowHidden(target)

    fc._rows_synced = False                     # 모델 리셋이 세우는 플래그와 동일
    fc.refresh(set())                           # 빈 델타 — 그래도 복구돼야 한다
    _assert_parity(dv, "(out-of-sync 복구)")


def test_model_reset_invalidates_sync(dv):
    """populate(모델 리셋)는 델타 신뢰를 해제해야 한다."""
    fc = dv._freeze["a"]
    assert fc._rows_synced
    dv.panel_a.table.model().set_diff_data(
        dv._diff_matrix, dv._diff_row_meta, dv._staged,
        dv._merged_cells, dv._excluded_cols)
    assert not fc._rows_synced


def test_refresh_without_args_still_full_scans(dv):
    """기존 호출부(테스트 포함)의 무인자 refresh() 는 전 행 스캔 그대로."""
    fc = dv._freeze["a"]
    target = next(r for r in range(fc.host.model().rowCount())
                  if not fc.host.isRowHidden(r))
    fc.left.setRowHidden(target, True)
    fc.refresh()
    _assert_parity(dv, "(무인자 refresh)")
