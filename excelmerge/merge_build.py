"""병합 패치 구성 — staged(병합 준비) 셀에서 파일 패치를 만드는 순수 로직.

Qt/파일 접근이 없어 단위테스트가 쉽다. 기존엔 StagedMergeWorker.run() 안에서 a2b/b2a
두 블록이 거의 동일하게 반복됐는데, direction 파라미터로 통합했다.
파일 쓰기(빈열 승격·XML 패치)는 여전히 xlsx_writer가 담당한다(여기선 '무엇을 쓸지'만 계산).
"""
from .constants import DIR_A2B
from .xlsx_writer import _cell_ref


def build_side_patches(direction: str, diff_matrix: list, row_meta: list,
                       staged: dict, col_meta: list | None = None):
    """한 방향(direction)에 대한 파일 패치 구성물을 만든다.

    반환: (patches, insert_rows, style_src)
      patches     : {cell_ref: value}                 기존 셀 덮어쓰기(값)
      insert_rows : {display_r: [(col, value, src_ref)]}  대상에 행이 없어 새로 삽입할 행
      style_src   : {target_ref: source_ref}          덮어쓰기 셀의 소스 서식 좌표

    a_to_b: 대상=B(b_orig), 소스=A(a_orig), 값=a_val
    b_to_a: 대상=A(a_orig), 소스=B(b_orig), 값=b_val
    대상 원본 행이 없으면(반대쪽에 그 행이 없음) patches 대신 insert_rows에 넣는다.
    값으로 병합 — 수식이 아닌 diff_matrix의 계산값을 기록.

    ★ 행과 마찬가지로 **열도 화면 번호가 곧 파일 번호가 아니다.** col_meta 가 있으면
    화면 열 → 그 side 의 파일 열로 옮겨서 좌표를 만든다. 이 변환을 빼먹으면 A 의 값이
    B 의 엉뚱한 열에 들어간다 — 그리고 '의도한 자리에 의도한 값' 이라 셀 단위 검증은
    통과한다(시트 전수 대조가 잡는다). col_meta 가 없으면 화면 열 = 파일 열(예전 그대로).

    **한쪽에만 있는 열** 은 이렇게 다룬다.
      · 대상에 그 열이 없으면 언제나 건너뛴다 — 쓸 자리가 없다. 열 삽입은 그 오른쪽
        모든 셀 참조·수식·서식을 밀어서 행 삽입과는 비교가 안 되게 위험하므로 지원하지
        않는다.
      · 소스에 그 열이 없으면 **행이 양쪽에 있을 때만** 건너뛴다. 쓸 값이 "" 라 대상의
        멀쩡한 열을 비우게 되는데, 열은 그 결과가 열 하나를 통째로 지우는 것이라
        되돌리기 어렵다. 손대지 않고 눈에 보이게 둔다(헤더가 연초록으로 뜬다).
      · ★ 그 행이 **대상에만 있으면**(소스에 없는 행) 이 병합은 곧 '행 삭제' 다. 그때는
        한쪽 전용 열까지 비워야 행이 통째로 비고 삭제로 승격된다. 여기서 건너뛰면 그
        열만 값이 남아 **반쪽짜리 빈 행**이 남는다 — v206 과 같은 종류의 사고다.
    """
    a2b = (direction == DIR_A2B)

    def cols_of(c):
        """화면 열 c → (대상 파일 열, 소스 파일 열). 없으면 None."""
        if not col_meta:
            return c, c
        if not (0 <= c < len(col_meta)):
            return None, None
        a_col, b_col = col_meta[c]
        return (b_col, a_col) if a2b else (a_col, b_col)

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
        # 소스에 그 행이 없다 = 이 병합은 '대상에서 행을 지운다' 는 뜻.
        row_removal = (src_orig is None and tgt_orig is not None)
        for c in row_cells:
            tgt_col, src_col = cols_of(c)
            if tgt_col is None:
                continue               # 대상에 없는 열 — 쓸 자리가 없다
            if src_col is None and not row_removal:
                continue               # 소스에 없는 열 — 멀쩡한 열을 비우지 않는다
            _, a_val, b_val = diff_matrix[r][c]
            val = a_val if a2b else b_val
            sref = (_cell_ref(src_orig, src_col)
                    if (src_orig is not None and src_col is not None) else None)
            if tgt_orig is not None:
                tref = _cell_ref(tgt_orig, tgt_col)
                patches[tref] = val
                if sref is not None:
                    style_src[tref] = sref
            else:
                insert_rows.setdefault(r, []).append((tgt_col, val, sref))

    return patches, insert_rows, style_src
