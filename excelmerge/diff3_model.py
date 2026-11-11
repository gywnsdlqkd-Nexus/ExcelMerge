"""Diff3TableModel — 3파일 비교 표 모델. 패널(왼쪽·가운데·오른쪽)마다 하나씩 쓴다.

2파일 모델(DiffTableModel)을 이어받아 **값을 고르는 자리와 색 규칙만** 바꾼다. 표·틀 고정·
델리게이트·미니맵은 모델의 조회 API(display_text·cell_kind·diff_char_ranges·
is_added_placeholder·own_col·orig_row …)만 보므로 그대로 동작한다.

셀은 (status, l_val, b_val, r_val) — diff3_engine 참고. 행·열 대응은 3원소 (l, b, r).
병합 준비(staged)·병합 완료(merged)는 {(r, c): SRC_LEFT|SRC_RIGHT} — 가운데(결과)에 어느
쪽 값을 쓰는지. 둘 다 Diff3View 가 소유하고 세 모델에 **참조로** 넘긴다.

색 규칙(패널별로 '그 파일이 무엇을 바꿨나'만 칠한다)
  왼쪽   기준과 다른 칸 — 왼쪽만 바뀜(노랑)·양쪽이 같게(보라)·충돌(빨강). 오른쪽만 바뀐 칸은
         왼쪽 파일 입장에선 기준 그대로라 흰색.
  오른쪽 대칭.
  가운데 바뀐 칸 전부(어느 쪽이 바꿨든) — 병합 결과가 들어갈 자리다.
  준비·완료 칸은 가운데와 **값을 가져온 쪽**에 준비(주황)·병합(파랑).
  그 파일에 없는 행·열은 빗금(2파일의 '반대쪽에만 있는 행'과 같은 뜻).
"""
from PyQt5.QtCore import Qt

from .constants import MODE_DIFF, MODE_PREVIEW
from .diff3_engine import S3_BOTH, S3_CONFLICT, S3_LEFT, S3_RIGHT, S3_SAME, SRC_LEFT, SRC_RIGHT
from .diff_model import DiffTableModel, _char_diff_ranges
from .theme import DIFF_COLORS, EXCLUDED_CELL_BG

ROLE_LEFT, ROLE_BASE, ROLE_RIGHT = "l", "b", "r"
_ROLE_IDX = {ROLE_LEFT: 0, ROLE_BASE: 1, ROLE_RIGHT: 2}
# 이 패널 파일이 기준과 다르게 만든 상태 — 패널 색칠 대상
_OWN_CHANGES = {
    ROLE_LEFT: (S3_LEFT, S3_BOTH, S3_CONFLICT),
    ROLE_RIGHT: (S3_RIGHT, S3_BOTH, S3_CONFLICT),
    ROLE_BASE: (S3_LEFT, S3_RIGHT, S3_BOTH, S3_CONFLICT),
}
_SRC_OF_ROLE = {ROLE_LEFT: SRC_LEFT, ROLE_RIGHT: SRC_RIGHT}
_STATUS_BG = {
    S3_LEFT: DIFF_COLORS["modified"],
    S3_RIGHT: DIFF_COLORS["modified"],
    S3_BOTH: DIFF_COLORS["both"],
    S3_CONFLICT: DIFF_COLORS["conflict"],
}


class Diff3TableModel(DiffTableModel):
    def __init__(self, role: str, parent=None):
        # side 는 2파일 코드가 'a'/'b' 로 분기하는 곳에서만 쓰인다 — 여기선 그 분기를 전부
        # 덮어쓰므로 역할 문자 그대로 둔다(잘못 새면 바로 드러나게).
        super().__init__(role, parent)
        self.role = role
        self._idx = _ROLE_IDX[role]      # row_meta3 · col_meta3 의 내 자리
        self._vi = self._idx + 1         # 셀 튜플의 내 값 자리

    # ── 결정(준비·완료) ──────────────────────────────────────────────────────
    def decision(self, r: int, c: int):
        """(r, c) 에 걸린 결정 — (출처, 'staged'|'merged') 또는 None."""
        d = self._staged.get((r, c))
        if d is not None:
            return d, "staged"
        d = self._merged.get((r, c)) if isinstance(self._merged, dict) else None
        if d is not None:
            return d, "merged"
        return None

    def staged_coords(self):
        """'병합 준비 취소' 메뉴·Alt+Backspace 가 보는 칸 — 3파일은 저장된 병합(파랑)도 풀 수
        있다(다음 저장에서 결과에서 빠진다). 2파일은 저장된 칸을 되돌리지 않으므로 준비만 본다."""
        if not self._merged:
            return self._staged.keys()
        return self._staged.keys() | self._merged.keys()

    def _shows_decision(self, r: int, c: int):
        """이 패널이 결정 색을 보이는가 — 가운데는 언제나, 양옆은 값을 가져온 쪽일 때만."""
        dec = self.decision(r, c)
        if dec is None:
            return None
        if self.role == ROLE_BASE or _SRC_OF_ROLE.get(self.role) == dec[0]:
            return dec
        return None

    # ── 값 ─────────────────────────────────────────────────────────────────
    def _compute_col_has_values(self, c: int) -> bool:
        if self._mode == MODE_DIFF:
            for row in self._diff_matrix:
                if c < len(row):
                    _st, lv, bv, rv = row[c]
                    if lv != "" or bv != "" or rv != "":
                        return True
            return False
        return super()._compute_col_has_values(c)

    def _compute_row_has_values(self, r: int) -> bool:
        if self._mode == MODE_DIFF:
            if not (0 <= r < self._data_rows):
                return False
            return any(v != "" for cell in self._diff_matrix[r] for v in cell[1:])
        return super()._compute_row_has_values(r)

    def display_text(self, r: int, c: int) -> str:
        if not self.is_data_cell(r, c):
            return ""
        if self._mode == MODE_PREVIEW:
            row = self._preview[r]
            return row[c] if c < len(row) else ""
        cell = self._diff_matrix[r][c]
        if self.role == ROLE_BASE:
            dec = self.decision(r, c)
            if dec is not None:
                return cell[1] if dec[0] == SRC_LEFT else cell[3]
        return cell[self._vi]

    def cell_kind(self, r: int, c: int) -> str:
        """병합 자격 — merged > staged > 변경. 패널과 무관하게 같다(선택이 세 패널에 미러되므로
        어느 패널에서 우클릭해도 같은 메뉴가 떠야 한다)."""
        if self._mode != MODE_DIFF or not self.is_data_cell(r, c):
            return S3_SAME
        if (r, c) in self._merged:
            return "merged"
        if (r, c) in self._staged:
            return "staged"
        return "changed" if self._diff_matrix[r][c][0] != S3_SAME else S3_SAME

    def _formula_orig_row(self, r: int):
        if self._mode == MODE_DIFF:
            try:
                return self._row_meta[r][self._idx]
            except (IndexError, TypeError):
                return None
        return r

    def is_formula(self, r: int, c: int) -> bool:
        if not self._formula_flags or not self.is_data_cell(r, c):
            return False
        if self._mode == MODE_DIFF and self.role == ROLE_BASE and self.decision(r, c):
            return False   # 가운데가 남의 값을 보이는 칸
        orig = self._formula_orig_row(r)
        col = self.own_col(c)
        return orig is not None and col is not None and (orig, col) in self._formula_flags

    def _absent(self, r: int, c: int) -> bool:
        """이 패널 파일에 그 행이나 열이 없다."""
        try:
            if self._row_meta[r][self._idx] is None:
                return True
        except (IndexError, TypeError):
            pass
        return self.own_col(c) is None

    def diff_char_ranges(self, r: int, c: int) -> list:
        """내 값에서 기준(가운데는 바꾼 쪽)과 다른 문자 구간."""
        if self._mode != MODE_DIFF or not self.is_data_cell(r, c):
            return []
        if c in self._excluded or self.decision(r, c) is not None:
            return []
        st, lv, bv, rv = self._diff_matrix[r][c]
        if st not in _OWN_CHANGES[self.role]:
            return []
        if self.role == ROLE_LEFT:
            own, other = lv, bv
        elif self.role == ROLE_RIGHT:
            own, other = rv, bv
        else:
            own, other = bv, (lv if lv != bv else rv)
        if own == "" or other == "" or own == other:
            return []      # 한쪽이 비었으면 글자 비교가 아니라 '생김/지움' — 빗금·색으로 충분
        cached = self._char_range_cache.get((r, c))
        if cached is None:
            cached = _char_diff_ranges(own, other)
            self._char_range_cache[(r, c)] = cached
        return cached

    def is_added_placeholder(self, r: int, c: int) -> bool:
        """빗금 — 이 패널 파일에 없는 행·열. 가운데가 결정된 값을 보이면 빗금을 걷는다."""
        if self._mode != MODE_DIFF or not self.is_data_cell(r, c):
            return False
        if c in self._excluded:
            return False
        if self.role == ROLE_BASE and self.decision(r, c) is not None:
            return False
        return self._absent(r, c)

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.BackgroundRole and index.isValid() and self._mode == MODE_DIFF:
            r, c = index.row(), index.column()
            if not self.is_data_cell(r, c):
                return None
            dec = self._shows_decision(r, c)
            if dec is not None:
                return DIFF_COLORS["merged" if dec[1] == "merged" else "staged"]
            if c in self._excluded:
                return EXCLUDED_CELL_BG
            st = self._diff_matrix[r][c][0]
            if st in _OWN_CHANGES[self.role] and not self._absent(r, c):
                return _STATUS_BG[st]
            return DIFF_COLORS[S3_SAME]
        return super().data(index, role)

    # ── 헤더 ───────────────────────────────────────────────────────────────
    def own_col(self, section: int):
        if not self._col_meta:
            return section
        if 0 <= section < len(self._col_meta):
            return self._col_meta[section][self._idx]
        return None

    def one_sided_col(self, section: int) -> bool:
        """세 파일 중 일부에만 있는 열."""
        if not self._col_meta or not (0 <= section < len(self._col_meta)):
            return False
        return None in self._col_meta[section]

    def orig_row(self, r: int):
        """이 패널 파일의 행 번호. 없으면 가운데, 그래도 없으면 남은 쪽 — 키 행 지정용."""
        if 0 <= r < len(self._row_meta):
            meta = self._row_meta[r]
            for i in (self._idx, 1, 0, 2):
                if meta[i] is not None:
                    return meta[i]
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if orientation == Qt.Horizontal and role == Qt.ToolTipRole and self._mode == MODE_DIFF:
            if self.one_sided_col(section):
                return "이 열은 세 파일 중 일부에만 있습니다 — 없는 파일에는 병합할 수 없습니다."
            return None
        if (orientation == Qt.Vertical and role == Qt.DisplayRole
                and self._mode == MODE_DIFF and self._row_meta):
            if section < self._data_rows:
                if section < len(self._row_meta):
                    orig = self._row_meta[section][self._idx]
                    return str(orig + 1) if orig is not None else "-"
                return "-"
            return str(section + 1)
        return super().headerData(section, orientation, role)
