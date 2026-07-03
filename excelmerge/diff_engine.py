"""diff 엔진 — 키 열 기반/행 순서 기반 비교 (excel_diff_merge.py에서 분리)."""


def _cell_status(a_val: str, b_val: str) -> str:
    if a_val == "" and b_val != "":
        return "added"
    if a_val != "" and b_val == "":
        return "modified"
    if a_val != b_val:
        return "modified"
    return "same"


def _compute_diff_row_order(
    a_data: list[list], b_data: list[list],
) -> tuple[list[list], list[tuple[int | None, int | None]]]:
    """키 없음 — 행 순서 그대로 1:1 매칭."""
    cols = max(
        (max(len(r) for r in a_data) if a_data else 0),
        (max(len(r) for r in b_data) if b_data else 0),
    )
    n = max(len(a_data), len(b_data))
    diff_matrix: list[list] = []
    row_meta: list[tuple] = []
    for i in range(n):
        a_row = a_data[i] if i < len(a_data) else []
        b_row = b_data[i] if i < len(b_data) else []
        row = []
        for c in range(cols):
            av = a_row[c] if c < len(a_row) else ""
            bv = b_row[c] if c < len(b_row) else ""
            row.append((_cell_status(av, bv), av, bv))
        diff_matrix.append(row)
        a_idx = i if i < len(a_data) else None
        b_idx = i if i < len(b_data) else None
        row_meta.append((a_idx, b_idx))
    return diff_matrix, row_meta


def compute_diff(
    a_data: list[list], b_data: list[list], key_col: int = 0
) -> tuple[list[list], list[tuple[int | None, int | None]]]:
    """
    key_col 열 값을 키로 행을 매칭하여 diff를 계산한다.
    key_col == -1 이면 행 순서 기반(ROW order) 비교를 수행한다.
    반환값:
      diff_matrix : list of rows, 각 row = [(status, a_val, b_val), ...]
      row_meta    : [(orig_a_row, orig_b_row), ...]  — None 은 해당 파일에 없는 행
    """
    if not a_data and not b_data:
        return [], []
    if key_col == -1:
        return _compute_diff_row_order(a_data, b_data)

    cols = max(
        (max(len(r) for r in a_data) if a_data else 0),
        (max(len(r) for r in b_data) if b_data else 0),
    )

    def get_key(row):
        return row[key_col] if row and key_col < len(row) else ""

    # 헤더 행(첫 행)은 키 매칭 없이 그대로 1:1
    a_header = a_data[0] if a_data else None
    b_header = b_data[0] if b_data else None

    # 나머지 행을 첫 번째 열 값으로 인덱싱
    a_body = a_data[1:] if len(a_data) > 1 else []
    b_body = b_data[1:] if len(b_data) > 1 else []

    # 키 → (원본 인덱스, 행 데이터) 매핑 (원본 파일 기준 행 번호는 1-based body index + 1)
    a_map: dict[str, tuple[int, list]] = {}
    for i, row in enumerate(a_body):
        key = get_key(row)
        if key == "":
            continue   # 키가 없는 행(빈 행 등)은 키 비교 대상에서 제외
        if key not in a_map:   # 중복 키는 첫 번째 행 사용
            a_map[key] = (i + 1, row)   # +1: 헤더 row가 row 0

    b_map: dict[str, tuple[int, list]] = {}
    for i, row in enumerate(b_body):
        key = get_key(row)
        if key == "":
            continue   # 키가 없는 행(빈 행 등)은 키 비교 대상에서 제외
        if key not in b_map:
            b_map[key] = (i + 1, row)

    # 표시 순서: A의 순서를 기준으로, A에 없고 B에만 있는 키는 뒤에 추가
    all_keys: list[str] = []
    seen: set[str] = set()
    for row in a_body:
        k = get_key(row)
        if k == "" or k in seen:
            continue
        all_keys.append(k)
        seen.add(k)
    for row in b_body:
        k = get_key(row)
        if k == "" or k in seen:
            continue
        all_keys.append(k)
        seen.add(k)

    def make_row(a_row, b_row):
        row = []
        for c in range(cols):
            av = a_row[c] if c < len(a_row) else ""
            bv = b_row[c] if c < len(b_row) else ""
            row.append((_cell_status(av, bv), av, bv))
        return row

    diff_matrix: list[list] = []
    row_meta: list[tuple] = []

    # 헤더 행
    if a_header is not None or b_header is not None:
        ah = a_header or []
        bh = b_header or []
        diff_matrix.append(make_row(ah, bh))
        row_meta.append((0 if a_header is not None else None,
                         0 if b_header is not None else None))

    # 본문 행
    for key in all_keys:
        a_idx, a_row = a_map[key] if key in a_map else (None, [])
        b_idx, b_row = b_map[key] if key in b_map else (None, [])
        diff_matrix.append(make_row(a_row, b_row))
        row_meta.append((a_idx, b_idx))

    return diff_matrix, row_meta
