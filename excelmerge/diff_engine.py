"""diff 엔진 — 키 열 기반/행 순서 기반 비교 (excel_diff_merge.py에서 분리)."""
from .constants import STATUS_SAME, STATUS_ADDED, STATUS_MODIFIED


def count_changed(diff_matrix: list, excluded_cols=None) -> int:
    """diff 매트릭스에서 'same'이 아니고 제외 열이 아닌 셀 수를 센다(O(R×C)).

    ★ 실행 경로에서는 더 이상 쓰지 않는다 — 세는 일은 count_changed_masked 가
      비트마스크로 처리한다(6328행 x 71열: 47ms -> 0.5ms). 이 함수는 '변경 셀'의
      **정의를 그대로 적어 둔 정본**으로 남겨 마스크 경로를 대조하는 데 쓴다
      (tests/test_row_change_masks.py). 정의를 바꿀 일이 생기면 여기와
      row_change_masks 를 함께 고쳐야 한다.
    """
    excl = excluded_cols or set()
    return sum(
        1
        for row in diff_matrix
        for c, (st, *_) in enumerate(row)
        if st != STATUS_SAME and c not in excl
    )


def row_change_masks(diff_matrix: list) -> list:
    """행별 '변경된 열' 비트마스크 — 비트 c 가 1이면 그 행 c 열이 same 이 아니다.

    변경 행 판정(변경점만 보기 필터 / 미니맵 / 변경점 이동 / 변경 셀 수)이 모두 행마다
    `any(st != SAME for c, (st, *_) in enumerate(row) if c not in excl)` 로 O(C) 튜플
    언패킹을 되풀이했다. 마스크를 한 번 만들어 두면 제외 열을 반영한 판정이
    `masks[r] & keep_mask(...)` 정수 연산 하나로 끝난다 — 실측 94~126배
    (6328행 x 71열: 미니맵 122ms→1.3ms, 필터 스캔 66ms→0.5ms).

    O(R x C) 스캔이므로 DiffWorker(백그라운드)가 매트릭스와 함께 만든다(6328x71 에서 32ms).

    ★ 매트릭스의 파생 상태다 — status 를 바꾸는 쪽은 마스크도 함께 갱신해야 한다.
      diff_view 는 그 갱신을 _set_matrix_cell() 한 곳에 묶어 둔다.
    """
    masks = []
    for row in diff_matrix:
        m = 0
        for c, cell in enumerate(row):
            if cell[0] != STATUS_SAME:
                m |= 1 << c
        masks.append(m)
    return masks


def keep_mask(cols: int, excluded_cols=None) -> int:
    """제외 열을 뺀 '검사 대상 열' 비트마스크 — row_change_masks 결과와 AND 해서 쓴다."""
    m = (1 << cols) - 1
    for c in (excluded_cols or ()):
        if 0 <= c < cols:
            m &= ~(1 << c)
    return m


def count_changed_masked(masks: list, keep: int) -> int:
    """마스크 기반 변경 셀 수 — count_changed 와 결과가 같아야 한다(테스트로 고정)."""
    return sum((m & keep).bit_count() for m in masks)


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


def dropped_key_rows(a_data: list, b_data: list, key_col: int,
                     key_row: int = 0, limit: int = 500) -> dict:
    """비교에서 빠진 행의 '왜'를 돌려준다 — 개수만 세는 count_dropped_key_rows 의 확장.

    반환:
      {"blank": n, "dup": n, "rows": [(side, 원본행(1-based), 키, 사유), ...(limit 까지)],
       "dup_keys": [(키, A건수, B건수), ...최대 10], "per_side": {"A": n, "B": n}}

    사유는 두 가지뿐이다. 빈 키(키 칸이 비어 매칭 불가)와 중복 키(같은 키의 두 번째
    이후 — compute_diff 는 첫 행만 쓴다). 실측하면 현장 데이터는 거의 전부 중복 키다
    (Data_SkillModule_CS.xlsx 를 B열 키로 비교하면 14,659행이 전부 중복).
    """
    empty = {"blank": 0, "dup": 0, "rows": [], "dup_keys": [], "per_side": {"A": 0, "B": 0}}
    if key_col is None or key_col < 0:
        return empty
    start = (key_row if key_row and key_row > 0 else 0) + 1
    dup_count: dict = {"A": {}, "B": {}}
    out = dict(empty)
    out["rows"] = []
    for side, rows in (("A", a_data or []), ("B", b_data or [])):
        seen: set = set()
        for i, row in enumerate(rows[start:], start=start):
            key = row[key_col] if row and key_col < len(row) else ""
            if key == "":
                reason = "빈 키"
                out["blank"] += 1
            elif key in seen:
                reason = "중복 키"
                out["dup"] += 1
                dup_count[side][key] = dup_count[side].get(key, 0) + 1
            else:
                seen.add(key)
                continue
            out["per_side"][side] += 1
            if len(out["rows"]) < limit:
                out["rows"].append((side, i + 1, key, reason))
    keys = set(dup_count["A"]) | set(dup_count["B"])
    out["dup_keys"] = sorted(
        ((k, dup_count["A"].get(k, 0), dup_count["B"].get(k, 0)) for k in keys),
        key=lambda t: -(t[1] + t[2]))[:10]
    return out


def unique_key_candidates(a_data: list, b_data: list, key_row: int = 0,
                          max_cols: int = 40) -> list:
    """키로 쓸 수 있는 열 — **양쪽 파일 모두에서** 빈 값 없고 중복도 없는 열.

    한쪽에서만 유니크한 열을 고르면 반대쪽에서 다시 행이 빠지므로 둘 다 본다.
    반환: [(열 인덱스, 헤더 텍스트), ...] — 왼쪽 열부터.
    비용이 O(행×열)이라 비교할 때가 아니라 **진단 창을 열 때만** 부른다.
    """
    start = (key_row if key_row and key_row > 0 else 0) + 1
    sides = [r for r in (a_data or [], b_data or []) if len(r) > start]
    if not sides:
        return []
    ncols = min(max_cols, max(max((len(r) for r in rows), default=0) for rows in sides))
    out = []
    for c in range(ncols):
        ok = True
        for rows in sides:
            vals = [(row[c] if c < len(row) else "") for row in rows[start:]]
            if not vals or any(v == "" for v in vals) or len(set(vals)) != len(vals):
                ok = False
                break
        if ok:
            header_row = sides[0][start - 1] if start - 1 < len(sides[0]) else []
            out.append((c, header_row[c] if c < len(header_row) else ""))
    return out


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

    행 순서: A 의 순서를 뼈대로 하고, B 에만 있는 키(신규 행)는 B 에서의 자리에 끼운다
    (자세한 규칙은 아래 본문 주석 참고). B 끝에 추가된 신규 행만 표 맨 아래로 간다.

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

    # 표시 순서: A 의 순서가 뼈대. B 에만 있는 키(= 신규 행)는 **B 에서의 제자리**에 끼운다.
    #
    # 끼우는 자리는 'B 에서 그 행 다음에 나오는, 양쪽에 다 있는 키' 바로 앞이다(고전적
    # diff 정렬과 같은 규칙). 그래서 B 중간에 새로 넣은 행이 표에서도 그 이웃 사이에 보이고,
    # 예전처럼 맨 아래로 밀려 원래 위치를 잃지 않는다.
    #  · A 에만 있는 행(삭제)과 자리가 겹치면 삭제 → 신규 순서로 놓인다.
    #  · B 끝에 붙은 신규 행은 뒤에 공통 키가 없으므로 예전과 같이 표 맨 아래로 간다.
    #  · 공통 키의 순서가 A/B 에서 다르면 A 순서를 따른다(뼈대가 A 이므로).
    inserts: dict[str, list[str]] = {}   # 공통 키 → 그 앞에 끼워 넣을 B 전용 키들
    pending: list[str] = []
    seen_b: set[str] = set()
    for row in b_body:
        k = get_key(row)
        if k == "" or k in seen_b:
            continue
        seen_b.add(k)
        if k in a_map:
            if pending:
                inserts.setdefault(k, []).extend(pending)
                pending = []
        else:
            pending.append(k)
    tail = pending   # 뒤에 공통 키가 없는 신규 행 = B 의 맨 끝에 추가된 행

    all_keys: list[str] = []
    seen: set[str] = set()
    for row in a_body:
        k = get_key(row)
        if k == "" or k in seen:
            continue
        seen.add(k)
        all_keys.extend(inserts.get(k, ()))
        all_keys.append(k)
    all_keys.extend(tail)

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
