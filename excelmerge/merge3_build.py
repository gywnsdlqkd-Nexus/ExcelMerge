"""3파일 병합 패치 구성 — 결정 {(r, c): SRC_LEFT|SRC_RIGHT} 에서 가운데(기준)에 쓸 패치를 만든다.

Qt·파일 접근이 없는 순수 로직이다. 쓸 자리 계산은 2파일의 build_side_patches 를 **그대로**
쓴다 — 3파일 매트릭스를 (왼쪽, 가운데) 또는 (오른쪽, 가운데) 2파일 모양으로 비춰서 넘긴다.
행 삭제(빈 행 승격)·새 행 삽입·한쪽에만 있는 열 규칙이 2파일과 똑같이 적용되고, 거기서 이미
검증된 규칙을 두 번 구현하지 않는다.

서식 좌표는 (태그, 좌표) 로 돌려준다 — 태그는 SRC_LEFT/SRC_RIGHT. 저장 엔진은 소스를
{태그: (경로, 시트)} 로 받아 각 칸의 서식을 제 출처 파일에서 가져온다.
"""
from .constants import DIR_A2B
from .diff3_engine import SRC_LEFT, SRC_RIGHT
from .merge_build import build_side_patches

_SIDE_IDX = {SRC_LEFT: 0, SRC_RIGHT: 2}


class _RowsAs2:
    """3파일 매트릭스를 (소스, 가운데) 2파일 셀로 보이게 — 필요한 행만 그때그때 만든다.
    20만 셀짜리 시트에서 결정이 몇 칸뿐이면 몇 행만 만든다."""

    def __init__(self, matrix, vi):
        self._m = matrix
        self._vi = vi

    def __len__(self):
        return len(self._m)

    def __getitem__(self, r):
        vi = self._vi
        return [(cell[0], cell[vi], cell[2]) for cell in self._m[r]]


class _MetaAs2:
    def __init__(self, meta, i):
        self._m = meta
        self._i = i

    def __len__(self):
        return len(self._m)

    def __getitem__(self, r):
        m = self._m[r]
        return (m[self._i], m[1])


def build_merge3_patches(matrix: list, row_meta3: list, col_meta3: list | None,
                         decisions: dict):
    """반환: (patches, insert_rows, style_src)
      patches     : {cell_ref: value}                         가운데의 기존 셀 덮어쓰기
      insert_rows : {display_r: [(col, value, (tag, ref)|None)]}  가운데에 없는 행(화면 순서)
      style_src   : {target_ref: (tag, source_ref)}
    """
    patches: dict = {}
    inserts: dict = {}
    style_src: dict = {}
    for src, i in _SIDE_IDX.items():
        # 가져올 값이 가운데 값과 같은 칸은 **쓰지 않는다** — 결과가 같으니 쓸 이유가 없고, 쓰면
        # 그 칸의 수식이 계산값 상수로 바뀐다(행 단위 충돌에서 남긴 쪽 행을 통째로 고를 때
        # 아무도 안 바꾼 수식 칸이 상수가 됐다 — 검토 실측). 비교도 값 기준이라 결과 값은 같다.
        vi = i + 1
        staged = {k: DIR_A2B for k, v in decisions.items()
                  if v == src and matrix[k[0]][k[1]][vi] != matrix[k[0]][k[1]][2]}
        if not staged:
            continue
        cm2 = [(t[i], t[1]) for t in col_meta3] if col_meta3 else None
        p, ins, ss = build_side_patches(
            DIR_A2B, _RowsAs2(matrix, i + 1), _MetaAs2(row_meta3, i), staged, cm2)
        patches.update(p)
        style_src.update({t: (src, s) for t, s in ss.items()})
        for r, cells in ins.items():
            inserts.setdefault(r, []).extend(
                (col, val, (src, ref) if ref is not None else None) for col, val, ref in cells)
    # 화면 순서 — 양쪽에서 온 새 행이 섞여도 파일 끝에 화면에서 보던 순서대로 붙는다.
    inserts = {r: sorted(inserts[r], key=lambda t: t[0]) for r in sorted(inserts)}
    return patches, inserts, style_src


def expand_row_removals(matrix: list, row_meta3: list, cells, src: str) -> set:
    """한쪽이 **지운 행**을 그쪽 값으로 병합하면 그 행의 바뀐 칸을 전부 함께 고른다.

    지운 쪽 값은 모두 "" 라, 일부 칸만 가져오면 가운데에 반쪽짜리 빈 행이 남는다(2파일 v206 과
    같은 종류의 사고). 행이 통째로 비어야 저장 엔진이 '행 삭제'로 올린다.
    """
    i = _SIDE_IDX[src]
    out = set(cells)
    for r in {r for r, _c in cells}:
        meta = row_meta3[r] if r < len(row_meta3) else None
        if meta is None or meta[i] is not None or meta[1] is None:
            continue
        out.update((r, c) for c, cell in enumerate(matrix[r]) if cell[0] != "same")
    return out
