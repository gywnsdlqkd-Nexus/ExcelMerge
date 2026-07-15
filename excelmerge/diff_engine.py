"""diff 엔진 — 키 열 기반/행 순서 기반 비교 (excel_diff_merge.py에서 분리)."""
from .constants import STATUS_SAME, STATUS_ADDED, STATUS_MODIFIED


def count_changed(diff_matrix: list, excluded_cols=None) -> int:
    """diff 매트릭스에서 'same'이 아니고 제외 열이 아닌 셀 수를 센다(O(R×C)).
    diff_view._count_changed와 DiffWorker가 공유하는 순수 함수 — 무거운 스캔을
    UI 스레드 밖(워커)에서도 돌릴 수 있게 분리했다."""
    excl = excluded_cols or set()
    return sum(
        1
        for row in diff_matrix
        for c, (st, *_) in enumerate(row)
        if st != STATUS_SAME and c not in excl
    )


def count_dropped_key_rows(a_data: list, b_data: list, key_col: int,
                           key_row: int = 0) -> int:
    """키 열 기반 비교에서 매칭 대상에서 제외되는 본문 행 수(공백 키 + 중복 키).
    compute_diff는 공백 키 행을 건너뛰고 중복 키는 첫 행만 쓰므로, 사용자가 '행이 줄었다'를
    인지할 수 있도록 그 수를 센다. ROW 순서(key_col == -1)면 드롭 없음(0).
    key_row: 헤더 행 인덱스 — 행 0..key_row(프리앰블+헤더)는 본문에서 제외한다."""
    if key_col is None or key_col < 0:
        return 0

    start = (key_row if key_row and key_row > 0 else 0) + 1  # 본문 시작 = 헤더 다음 행

    def _dropped(rows: list) -> int:
        seen: set = set()
        dropped = 0
        for row in rows[start:]:   # 프리앰블+헤더 제외
            key = row[key_col] if row and key_col < len(row) else ""
            if key == "" or key in seen:
                dropped += 1
            else:
                seen.add(key)
        return dropped

    return _dropped(a_data or []) + _dropped(b_data or [])


def _cell_status(a_val: str, b_val: str) -> str:
    """added: 한쪽 파일에만 값이 있음 (A 전용/B 전용 모두) / modified: 양쪽 값이 다름."""
    if (a_val == "") != (b_val == ""):
        return STATUS_ADDED
    if a_val != b_val:
        return STATUS_MODIFIED
    return STATUS_SAME


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
    a_data: list[list], b_data: list[list], key_col: int = 0, key_row: int = 0
) -> tuple[list[list], list[tuple[int | None, int | None]]]:
    """
    key_col 열 값을 키로 행을 매칭하여 diff를 계산한다.
    key_col == -1 이면 행 순서 기반(ROW order) 비교를 수행한다(key_row 무시).

    key_row: 헤더 행 인덱스(기본 0). 행 0..key_row(프리앰블 + 헤더)는 키 매칭 없이
    위치 기준 1:1로 상단에 그대로 방출하고, 행 key_row+1 이후만 key_col 값으로 매칭한다.
    key_row 가 데이터 범위를 벗어나면 본문 없이 전 행이 1:1이 된다(안전 처리).

    반환값:
      diff_matrix : list of rows, 각 row = [(status, a_val, b_val), ...]
      row_meta    : [(orig_a_row, orig_b_row), ...]  — None 은 해당 파일에 없는 행.
                    값은 항상 원본 파일 기준 0-based 행 번호(병합 저장 좌표와 일치).
    """
    if not a_data and not b_data:
        return [], []
    if key_col == -1:
        return _compute_diff_row_order(a_data, b_data)

    kr = key_row if (key_row and key_row > 0) else 0
    n_head = kr + 1   # 프리앰블+헤더 행 수 (행 0..kr)

    cols = max(
        (max(len(r) for r in a_data) if a_data else 0),
        (max(len(r) for r in b_data) if b_data else 0),
    )

    def get_key(row):
        return row[key_col] if row and key_col < len(row) else ""

    def make_row(a_row, b_row):
        row = []
        for c in range(cols):
            av = a_row[c] if c < len(a_row) else ""
            bv = b_row[c] if c < len(b_row) else ""
            row.append((_cell_status(av, bv), av, bv))
        return row

    # 본문(헤더 다음 행부터)만 키 인덱싱. 원본 행 번호 = n_head + body index.
    a_body = a_data[n_head:] if len(a_data) > n_head else []
    b_body = b_data[n_head:] if len(b_data) > n_head else []

    a_map: dict[str, tuple[int, list]] = {}
    for i, row in enumerate(a_body):
        key = get_key(row)
        if key == "":
            continue   # 키가 없는 행(빈 행 등)은 키 비교 대상에서 제외
        if key not in a_map:   # 중복 키는 첫 번째 행 사용
            a_map[key] = (i + n_head, row)

    b_map: dict[str, tuple[int, list]] = {}
    for i, row in enumerate(b_body):
        key = get_key(row)
        if key == "":
            continue
        if key not in b_map:
            b_map[key] = (i + n_head, row)

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

    diff_matrix: list[list] = []
    row_meta: list[tuple] = []

    # 프리앰블 + 헤더 행: 위치 기준 1:1 (행 0..kr). 한쪽에만 있는 행은 그쪽만.
    head_n = min(n_head, max(len(a_data), len(b_data)))
    for i in range(head_n):
        a_present = i < len(a_data)
        b_present = i < len(b_data)
        diff_matrix.append(make_row(a_data[i] if a_present else [],
                                    b_data[i] if b_present else []))
        row_meta.append((i if a_present else None, i if b_present else None))

    # 본문 행
    for key in all_keys:
        a_idx, a_row = a_map[key] if key in a_map else (None, [])
        b_idx, b_row = b_map[key] if key in b_map else (None, [])
        diff_matrix.append(make_row(a_row, b_row))
        row_meta.append((a_idx, b_idx))

    return diff_matrix, row_meta
