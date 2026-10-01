# -*- coding: utf-8 -*-
"""열 매칭 2단계 — 화면에 실제로 켠다.

1단계(tests/test_column_matching.py)는 엔진만이었다. 여기서는 워커가 col_meta 를 만들어
화면까지 흘려보내고, 헤더가 **side 별 자기 열 문자**를 보여 주는 것을 고정한다. 세로
헤더가 row_meta 로 side 별 원본 행 번호를 보여 주는 것과 같은 규칙이다.

3단계부터는 저장도 col_meta 로 좌표를 옮겨 쓴다. 다만 **대상 파일에 없는 열**은 병합할
수 없다 — 쓰려면 열을 새로 끼워야 하는데, 열 삽입은 그 오른쪽 모든 셀 참조·수식·서식을
밀어서 행 삽입과는 비교가 안 되게 위험하다. 그래서 준비 단계에서 걸러내고, 몇 개를
뺐는지 말해 준다(조용히 빼면 '준비했는데 저장이 안 됐다' 가 된다).
"""
import pytest
from PyQt5.QtCore import Qt

from excelmerge.theme import HEADER_ONESIDE_BG

# A: ID | NAME | VALUE        B: ID | GRADE | NAME | VALUE   (GRADE 가 중간에 끼었다)
A_DATA = [["ID", "NAME", "VALUE"], ["k1", "칼", "10"], ["k2", "방패", "20"]]
B_DATA = [["ID", "GRADE", "NAME", "VALUE"],
          ["k1", "A", "칼", "10"], ["k2", "S", "방패", "20"]]
SAME_A = [["ID", "NAME"], ["k1", "칼"], ["k2", "방패"]]
SAME_B = [["ID", "NAME"], ["k1", "칼"], ["k2", "몽둥이"]]


@pytest.fixture
def make_view(qapp, monkeypatch, tmp_path):
    from excelmerge import diff_view as dv_mod
    from excelmerge.main_window import MainWindow
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: None))
    monkeypatch.setenv("APPDATA", str(tmp_path))
    wins = []

    def _make(a=A_DATA, b=B_DATA):
        win = MainWindow()
        win.show()
        wins.append(win)
        view = win.tabs.currentWidget()
        view.panel_a.set_path(str(tmp_path / "a.xlsx"))
        view.panel_b.set_path(str(tmp_path / "b.xlsx"))
        view._on_loaded(a, b)
        w = getattr(view, "_diff_worker", None)
        if w is not None:
            w.wait(8000)
        for _ in range(20):
            qapp.processEvents()
        return view

    yield _make
    for win in wins:
        win.close()
        win.deleteLater()
    qapp.processEvents()


def _letters(view, side, n):
    m = (view.panel_a if side == "a" else view.panel_b).table.model()
    return [m.headerData(c, Qt.Horizontal, Qt.DisplayRole) for c in range(n)]


# ── 워커 → 화면 배선 ─────────────────────────────────────────────────────────

def test_col_meta_reaches_the_view(make_view):
    v = make_view()
    assert v._diff_col_meta == [(0, 0), (None, 1), (1, 2), (2, 3)]


def test_the_matrix_width_follows_col_meta(make_view):
    v = make_view()
    assert len(v._diff_matrix[0]) == 4


def test_matching_is_off_when_headers_are_identical(make_view):
    """맞추나 마나 같은 경우에는 화면에 아무 표시도 뜨면 안 된다."""
    v = make_view(SAME_A, SAME_B)
    assert v._diff_col_meta == [(0, 0), (1, 1)]
    assert "이름으로 맞춤" not in v.status.currentMessage()


def test_a_new_comparison_clears_the_old_col_meta(make_view, qapp):
    v = make_view()
    assert v._diff_col_meta is not None
    v._run_compare()                       # 경로가 가짜라 로드가 실패한다 — 그래도 비워야 한다
    for _ in range(10):
        qapp.processEvents()
    assert v._diff_col_meta is None, "이전 비교의 열 매핑이 남았다"


# ── 헤더: side 별 자기 열 문자 ───────────────────────────────────────────────

def test_each_panel_shows_its_own_column_letters(make_view):
    v = make_view()
    assert _letters(v, "a", 4) == ["A", "-", "B", "C"], "A 에 없는 열은 '-' 여야 한다"
    assert _letters(v, "b", 4) == ["A", "B", "C", "D"]


def test_letters_are_plain_when_not_matched(make_view):
    v = make_view(SAME_A, SAME_B)
    assert _letters(v, "a", 2) == ["A", "B"]
    assert _letters(v, "b", 2) == ["A", "B"]


def test_own_col_maps_to_the_file_column(make_view):
    v = make_view()
    ma, mb = v.panel_a.table.model(), v.panel_b.table.model()
    assert [ma.own_col(c) for c in range(4)] == [0, None, 1, 2]
    assert [mb.own_col(c) for c in range(4)] == [0, 1, 2, 3]


# ── 한쪽에만 있는 열 표시 ────────────────────────────────────────────────────

def test_one_sided_column_is_marked_on_both_panels(make_view):
    v = make_view()
    for panel in (v.panel_a, v.panel_b):
        m = panel.table.model()
        assert [c for c in range(4) if m.one_sided_col(c)] == [1]


def test_one_sided_header_asks_for_the_tint(make_view):
    """모델이 색을 **요청**하는지까지만 본다.

    ★ 이게 화면에 칠해지는지는 여기서 알 수 없다. 실제로 한동안 한 픽셀도 칠해지지
    않았는데 이 테스트는 통과했다(스타일시트가 헤더 배경을 먼저 그려 버린다).
    그려진 픽셀은 tests/test_header_paint.py 가 본다.
    """
    m = make_view().panel_a.table.model()
    assert m.headerData(1, Qt.Horizontal, Qt.BackgroundRole) == HEADER_ONESIDE_BG
    assert m.headerData(2, Qt.Horizontal, Qt.BackgroundRole) != HEADER_ONESIDE_BG


def test_one_sided_header_explains_itself(make_view):
    m = make_view().panel_a.table.model()
    tip = m.headerData(1, Qt.Horizontal, Qt.ToolTipRole)
    assert tip and "한쪽 파일에만" in tip
    assert m.headerData(2, Qt.Horizontal, Qt.ToolTipRole) is None


def test_the_key_column_keeps_its_own_colour(make_view):
    """키 열이 우선 — 키 표시를 한쪽 전용 색이 덮으면 안 된다."""
    from excelmerge.theme import HEADER_KEY_BG
    m = make_view().panel_a.table.model()
    assert m.headerData(0, Qt.Horizontal, Qt.BackgroundRole) == HEADER_KEY_BG


# ── 사용자에게 알리기 ────────────────────────────────────────────────────────

def test_the_status_says_columns_were_matched(make_view):
    msg = make_view().status.currentMessage()
    assert "열을 이름으로 맞춤" in msg, msg
    assert "B 전용 1열" in msg, msg


def test_a_only_columns_are_counted_too(make_view):
    a = [["ID", "MEMO", "VALUE"], ["k1", "메모", "10"]]
    b = [["ID", "VALUE"], ["k1", "10"]]
    msg = make_view(a, b).status.currentMessage()
    assert "A 전용 1열" in msg, msg


# ── 준비: 반대쪽에 없는 열은 뺀다 ───────────────────────────────────────────
# 쓰려면 열을 새로 끼워야 하는데, 열 삽입은 그 오른쪽 모든 셀 참조·수식·서식을 밀어서
# 행 삽입과는 비교가 안 되게 위험하다. 그래서 준비 단계에서 걸러낸다.

def test_one_sided_columns_are_skipped_in_both_directions(make_view):
    """행이 양쪽에 있으면 한쪽 전용 열(GRADE)은 어느 방향으로도 건드리지 않는다.

    B→A 는 쓸 자리가 없고(A 에 그 열이 없다), A→B 는 쓸 값이 "" 라 B 의 멀쩡한 열을
    비우게 된다 — 열 하나를 통째로 지우는 셈이라 둘 다 손대지 않는다.
    """
    from excelmerge.constants import DIR_A2B, DIR_B2A
    v = make_view()
    cells = {(1, c) for c in range(4)}
    assert v._unmergeable_cells(cells, DIR_A2B) == {(1, 1)}
    assert v._unmergeable_cells(cells, DIR_B2A) == {(1, 1)}


def test_a_row_only_in_the_target_clears_the_one_sided_column_too(make_view):
    """행 삭제 병합은 예외 — 한쪽 전용 열까지 비워야 행이 통째로 사라진다.

    여기서 빼면 그 열만 값이 남아 반쪽짜리 빈 행이 생긴다(v206 과 같은 종류).
    """
    from excelmerge.constants import DIR_A2B
    a = [["ID", "NAME"], ["k1", "칼"]]
    b = [["ID", "GRADE", "NAME"], ["k1", "A", "칼"], ["k9", "S", "창"]]
    v = make_view(a, b)
    r = next(i for i, (ar, _br) in enumerate(v._diff_row_meta) if ar is None)
    cells = {(r, c) for c in range(3)}
    assert v._unmergeable_cells(cells, DIR_A2B) == set(), (
        "행을 지우는 병합인데 한쪽 전용 열을 뺐다")


# 행 1 에 '쓸 수 있는 변경'(NAME)과 '쓸 수 없는 변경'(B 전용 GRADE)이 함께 있다.
MIX_A = [["ID", "NAME"], ["k1", "칼"]]
MIX_B = [["ID", "GRADE", "NAME"], ["k1", "A", "검"]]


def test_staging_keeps_the_writable_cells_and_drops_the_rest(make_view, qapp):
    from excelmerge.constants import DIR_B2A
    v = make_view(MIX_A, MIX_B)
    assert v._diff_col_meta == [(0, 0), (None, 1), (1, 2)]
    v.panel_b.table.selectRow(1)
    for _ in range(10):
        qapp.processEvents()
    v._stage_selected(DIR_B2A)
    for _ in range(10):
        qapp.processEvents()
    assert set(v._staged) == {(1, 2)}, f"준비 대상이 틀렸다: {v._staged}"
    assert "제외됨" in v.status.currentMessage(), v.status.currentMessage()


def test_staging_says_nothing_when_nothing_is_skipped(make_view, qapp):
    """뺄 게 없으면 군더더기 문구를 붙이지 않는다."""
    from excelmerge.constants import DIR_A2B
    v = make_view(SAME_A, SAME_B)          # 열 구성이 같아 뺄 열이 없다
    v.panel_b.table.selectRow(2)
    for _ in range(10):
        qapp.processEvents()
    v._stage_selected(DIR_A2B)
    for _ in range(10):
        qapp.processEvents()
    assert v._staged, "전제: 준비된 셀이 있어야 한다"
    assert "제외됨" not in v.status.currentMessage()


def test_staging_only_unwritable_cells_stages_nothing(make_view, qapp, monkeypatch):
    """고른 게 전부 반대쪽에 없는 열이면 아무것도 준비되지 않고 이유를 알려 준다."""
    from excelmerge import diff_view as dv_mod
    from excelmerge.constants import DIR_B2A
    told = []
    monkeypatch.setattr(dv_mod.QMessageBox, "information",
                        staticmethod(lambda *a, **k: told.append(a[2] if len(a) > 2 else "")))
    v = make_view()
    v.panel_b.table.selectRow(1)
    for _ in range(10):
        qapp.processEvents()
    v._stage_selected(DIR_B2A)
    assert v._staged == {}, v._staged
    assert told and "열을 새로 만드는" in told[-1], told


class _FakeSignal:
    def connect(self, *a, **k):
        pass


class _FakeWorker:
    """StagedMergeWorker 대역 — 파일을 쓰지 않고 받은 인자만 기록한다.

    QThread 흉내를 blockSignals/isRunning 까지 내야 한다. 탭을 닫을 때
    DiffView.shutdown() 이 실행 중 워커를 정리하면서 이 객체를 건드리는데, 거기서
    예외가 나면 closeEvent 안에서 터져 프로세스가 통째로 죽는다(실제로 겪었다:
    0xC0000409 fail-fast, 테스트는 전부 통과한 뒤 요약 직전에 사망).
    """
    last = None

    def __init__(self, *args, **kw):
        _FakeWorker.last = self
        self.args = args
        self.done = _FakeSignal()
        self.error = _FakeSignal()
        self.finished = _FakeSignal()

    def start(self):
        self.started = True

    def blockSignals(self, _b):
        pass

    def isRunning(self):
        return False

    def deleteLater(self):
        pass


def test_saving_is_no_longer_blocked_when_columns_are_matched(make_view, monkeypatch):
    """3단계부터 저장은 col_meta 로 좌표를 옮겨 쓴다 — 더 이상 막지 않는다."""
    from excelmerge import diff_view as dv_mod
    from excelmerge.constants import DIR_A2B
    v = make_view()
    warned = []
    monkeypatch.setattr(dv_mod.QMessageBox, "warning",
                        staticmethod(lambda *a, **k: warned.append(a)))
    monkeypatch.setattr(dv_mod, "StagedMergeWorker", _FakeWorker)
    monkeypatch.setattr(dv_mod, "_is_file_locked", lambda p: False)
    v._staged[(1, 2)] = DIR_A2B
    v._save_staged("b")
    assert warned == [], f"저장이 막혔다: {warned}"
    assert _FakeWorker.last is not None, "저장 워커가 시작되지 않았다"
    assert _FakeWorker.last.args[-1] == v._diff_col_meta, (
        "워커에 col_meta 가 전달되지 않았다")
