"""병합 패치 구성 — staged(병합 준비) 셀에서 파일 패치를 만드는 순수 로직.

Qt/파일 접근이 없어 단위테스트가 쉽다. 기존엔 StagedMergeWorker.run() 안에서 a2b/b2a
두 블록이 거의 동일하게 반복됐는데, direction 파라미터로 통합했다.
파일 쓰기(빈열 승격·XML 패치)는 여전히 xlsx_writer가 담당한다(여기선 '무엇을 쓸지'만 계산).
"""
from .constants import DIR_A2B
from .xlsx_writer import _cell_ref


def build_side_patches(direction: str, diff_matrix: list, row_meta: list, staged: dict):
    """한 방향(direction)에 대한 파일 패치 구성물을 만든다.

    반환: (patches, insert_rows, style_src)
      patches     : {cell_ref: value}                 기존 셀 덮어쓰기(값)
      insert_rows : {display_r: [(col, value, src_ref)]}  대상에 행이 없어 새로 삽입할 행
      style_src   : {target_ref: source_ref}          덮어쓰기 셀의 서식 소스 좌표

    a_to_b: 대상=B(b_orig), 소스=A(a_orig), 값=a_val
    b_to_a: 대상=A(a_orig), 소스=B(b_orig), 값=b_val
    대상 원본 행이 없으면(반대쪽에 그 행이 없음) patches 대신 insert_rows에 넣는다.
    값으로 병합 — 수식이 아닌 diff_matrix의 계산값을 기록.
    """
    a2b = (direction == DIR_A2B)
    cells = {k for k, v in staged.items() if v == direction}
    patches: dict[str, str] = {}
    insert_rows: dict[int, list[tuple]] = {}
    style_src: dict[str, str] = {}

    for r in {rr for (rr, _c) in cells}:
        row = diff_matrix[r]
        a_orig, b_orig = row_meta[r] if r < len(row_meta) else (None, None)
        tgt_orig = b_orig if a2b else a_orig   # 쓸 대상 원본 행
        src_orig = a_orig if a2b else b_orig   # 서식 소스 원본 행
        row_cells = [c for c in range(len(row)) if (r, c) in cells]
        for c in row_cells:
            _, a_val, b_val = diff_matrix[r][c]
            val = a_val if a2b else b_val
            if tgt_orig is not None:
                tref = _cell_ref(tgt_orig, c)
                patches[tref] = val
                if src_orig is not None:
                    style_src[tref] = _cell_ref(src_orig, c)
            else:
                sref = _cell_ref(src_orig, c) if src_orig is not None else None
                insert_rows.setdefault(r, []).append((c, val, sref))

    return patches, insert_rows, style_src
