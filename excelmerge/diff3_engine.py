"""3파일 비교 엔진 — 가운데(기준) 대비 왼쪽·오른쪽 변경을 나눈다. Qt 에 의존하지 않는다.

WinMerge · P4V 의 3-way 관례를 따른다: 가운데 = 공통 기준(base), 양옆 = 그 기준에서 갈라진 두
버전(P4V 병합이면 왼쪽 theirs, 오른쪽 yours). 병합 결과는 가운데에 쓴다.

**2파일 엔진(diff_engine)은 건드리지 않는다.** 셀·행·열 대응이 3원소가 되므로 2파일 경로
(diff_model · staging · merge_build)에 섞으면 언패킹 오류가 나거나, 더 나쁘게는 오류 없이
틀린 값이 나온다(구조 조사에서 실측). 이 모듈의 결과는 3파일 전용 모델·뷰만 쓴다.

출력
  matrix    : [[(status, l_val, b_val, r_val), ...], ...] — status 가 0번 자리라 diff_engine 의
              row_change_masks / keep_mask / count_changed_masked 를 그대로 쓴다('same' 동일).
  row_meta3 : [(l_row, b_row, r_row), ...] — 원본 파일 기준 0-based 행 번호, 없으면 None.
  col_meta3 : [(l_col, b_col, r_col), ...] — 화면 열 ↔ 파일 열, 없으면 None. None 이면 위치 기준.

상태
  same      세 값이 같다
  left      왼쪽만 기준과 다르다(오른쪽 = 기준)
  right     오른쪽만 기준과 다르다
  both      양쪽이 **같은 값으로** 바뀌었다
  conflict  양쪽이 **서로 다른 값으로** 바뀌었다 — 또는 아래 행 단위 충돌

행 단위 충돌. 한쪽이 지운 행을 다른 쪽이 고쳤으면 그 행의 변경 칸을 모두 충돌로 둔다.
셀 단위로만 나누면 '지움' 칸(왼쪽만)과 '고침' 칸(충돌)이 갈라지고, 충돌 없는 칸만 자동
병합하면 반쪽만 빈 행이 남는다(2파일에서 v206 에 같은 종류의 사고가 있었다).
"""
import bisect
import re
from collections import Counter
from difflib import SequenceMatcher
from operator import eq, itemgetter

from .constants import STATUS_SAME
from .diff_engine import match_columns

S3_SAME = STATUS_SAME          # 'same' — 마스크·필터 공용 함수가 그대로 동작하도록 같은 값
S3_LEFT = "left"
S3_RIGHT = "right"
S3_BOTH = "both"
S3_CONFLICT = "conflict"
S3_CHANGED = (S3_LEFT, S3_RIGHT, S3_BOTH, S3_CONFLICT)

# 병합 출처 — 가운데(결과)에 어느 쪽 값을 쓰나. 2파일의 DIR_A2B/DIR_B2A 와 섞지 않는다.
SRC_LEFT = "from_left"
SRC_RIGHT = "from_right"


def classify(l_val: str, b_val: str, r_val: str) -> str:
    """셀 하나의 3-way 상태(값만 본다 — 행 단위 충돌은 compute_diff3 가 덧씌운다)."""
    if l_val == b_val:
        return S3_SAME if r_val == b_val else S3_RIGHT
    if r_val == b_val:
        return S3_LEFT
    return S3_BOTH if l_val == r_val else S3_CONFLICT


def _header_names(data: list, kr: int) -> list:
    row = data[kr] if (data and kr < len(data)) else []
    return [str(v).strip() for v in row]


def match_columns3(left: list, base: list, right: list, key_row: int = 0) -> list | None:
    """헤더 이름으로 맞춘 3파일 열 대응 [(l, b, r)]. 이름으로 맞출 수 없으면 None(= 위치 기준).

    (기준, 왼쪽) 과 (기준, 오른쪽) 을 2파일 match_columns 로 따로 맞춘 뒤 기준 열을 뼈대로
    합친다. 한쪽에만 생긴 열은 그 파일에서의 제자리(다음 기준 열 바로 앞)에 끼운다. 양쪽이
    **같은 자리에 같은 이름의 열**을 새로 만들었으면 하나로 합친다(두 번 보이지 않게).
    """
    bl = match_columns(base, left, key_row)
    br = match_columns(base, right, key_row)
    if bl is None or br is None:
        return None
    kr = key_row if (key_row and key_row > 0) else 0
    l_names, r_names = _header_names(left, kr), _header_names(right, kr)

    def slots(pairs):
        """기준 열 b → 그 앞에 끼울 새 열들, 그리고 꼬리. pairs 는 [(b, other)]."""
        before: dict = {}
        mapped: dict = {}
        pending: list = []
        for b, o in pairs:
            if b is None:
                pending.append(o)
            else:
                mapped[b] = o
                if pending:
                    before[b] = pending
                    pending = []
        return before, mapped, pending

    l_before, l_map, l_tail = slots(bl)
    r_before, r_map, r_tail = slots(br)

    def added(ls, rs):
        """같은 자리에 새로 생긴 열들 — 같은 이름은 하나로."""
        out, used = [], set()
        for lc in ls:
            name = l_names[lc] if lc < len(l_names) else None
            match = next((rc for rc in rs if rc not in used
                          and rc < len(r_names) and r_names[rc] == name), None)
            if match is not None:
                used.add(match)
            out.append((lc, None, match))
        out.extend((None, None, rc) for rc in rs if rc not in used)
        return out

    n_base = max((len(r) for r in base), default=0) if base else 0
    base_cols = sorted(set(l_map) | set(r_map)) or list(range(n_base))
    meta = []
    for b in base_cols:
        meta.extend(added(l_before.get(b, []), r_before.get(b, [])))
        meta.append((l_map.get(b), b, r_map.get(b)))
    meta.extend(added(l_tail, r_tail))
    meta.extend(_unnamed_tail(left, base, right, kr))
    return meta


def _unnamed_tail(left: list, base: list, right: list, kr: int) -> list:
    """이름 있는 헤더 오른쪽의 **이름 없는 열**(값만 있는 꼬리 열)을 위치로 짝지어 덧붙인다.

    match_columns 는 꼬리의 빈 헤더를 열로 보지 않는다. 2파일은 그래도 괜찮다 — 고른 칸만
    덧쓰므로 비교에서 빠진 열은 대상 파일에 그대로 남는다. 3파일 저장은 '가운데 원본에서
    다시 만든다' 라 비교에서 빠진 열의 양쪽 수정이 **기준값으로 되돌아간다**(검토 실측: 이름
    없는 메모 열의 theirs·yours 수정이 둘 다 사라짐). 그래서 보이게 한다 — 이름 있는 마지막
    열 다음부터 k 번째 열끼리 짝짓는다.
    """
    def named(data):
        names = _header_names(data, kr)
        while names and names[-1] == "":
            names.pop()
        return len(names)

    datas = (left, base, right)
    starts = [named(d) for d in datas]
    extra = [max(0, max((len(r) for r in d), default=0) - s) for d, s in zip(datas, starts)]
    return [tuple(starts[i] + k if k < extra[i] else None for i in range(3))
            for k in range(max(extra))]


def usable_col_meta3(col_meta3: list | None, key_col: int) -> list | None:
    """키 열이 **세 파일 모두에** 있어야 행을 맞출 수 있다 — 아니면 None(= 위치 기준)."""
    if not col_meta3 or key_col == -1 or not (0 <= key_col < len(col_meta3)):
        return None
    return col_meta3 if None not in col_meta3[key_col] else None


def _weave(seq: list, other_keys: list, present: set) -> list:
    """seq 에 없는 other 의 키를 other 에서의 제자리(다음 공통 키 바로 앞)에 끼운다.
    뒤에 공통 키가 없으면 맨 뒤 — 2파일 compute_diff 의 신규 행 규칙과 같다."""
    inserts: dict = {}
    pending: list = []
    for k in other_keys:
        if k in present:
            if pending:
                inserts.setdefault(k, []).extend(pending)
                pending = []
        else:
            pending.append(k)
    out: list = []
    for k in seq:
        out.extend(inserts.get(k, ()))
        out.append(k)
    out.extend(pending)
    return out


def _row_cells(cols, l_row, b_row, r_row):
    def pick(row, c):
        return row[c] if (row is not None and c is not None and c < len(row)) else ""
    lv = [pick(l_row, lc) for lc, _b, _r in cols]
    bv = [pick(b_row, bc) for _l, bc, _r in cols]
    rv = [pick(r_row, rc) for _l, _b, rc in cols]
    return lv, bv, rv


def _make_row(cols, l_row, b_row, r_row):
    """한 행의 셀들. 기준에 있는 행을 한쪽이 지우고 다른 쪽이 고쳤으면 행 단위 충돌."""
    lv, bv, rv = _row_cells(cols, l_row, b_row, r_row)
    if lv == bv and rv == bv:                       # 대부분의 행 — 빠른 길
        return [(S3_SAME, v, v, v) for v in bv]
    cells = [(classify(l, b, r), l, b, r) for l, b, r in zip(lv, bv, rv)]
    if b_row is not None and ((l_row is None) != (r_row is None)):
        ki = 2 if l_row is None else 0              # 지우지 않은 쪽
        kept = rv if l_row is None else lv
        # '고쳤나' 는 기준과 남긴 쪽 **둘 다에 있는 열**에서만 본다. 남긴 쪽이 열을 새로
        # 만들었거나 지운 것은 이 행을 고친 게 아니다 — 그걸 세면 지운 행이 모두 거짓 충돌이
        # 돼 자동 병합이 멀쩡한 행 삭제를 건너뛴다(검토 실측).
        edited = any(kept[i] != bv[i] for i, t in enumerate(cols)
                     if t[1] is not None and t[ki] is not None)
        if edited:                                  # 지운 행을 다른 쪽이 고쳤다
            cells = [((S3_CONFLICT if st != S3_SAME else st), l, b, r)
                     for st, l, b, r in cells]
    return cells


def compute_diff3(left: list, base: list, right: list, key_col: int = 0, key_row: int = 0,
                  col_meta3: list | None = None, ambiguous: set | None = None):
    """세 파일 비교 → (matrix, row_meta3). key_col 은 **화면 열** 번호, -1 이면 행 순서.

    ambiguous 에 집합을 주면, 키가 비었거나 겹쳐 **내용으로 짝지은 행**의 화면 행 번호를 채운다.
    그 짝은 틀릴 수 있으므로 자동 병합에서 뺀다(사람이 보고 고른다).
    """
    if not left and not base and not right:
        return [], []
    wide = max((max((len(r) for r in d), default=0) for d in (left, base, right)), default=0)
    cm = usable_col_meta3(col_meta3, key_col)
    cols = cm if cm else [(c, c, c) for c in range(wide)]

    if key_col == -1:
        n = max(len(left), len(base), len(right))
        matrix, meta = [], []
        for i in range(n):
            rows = [d[i] if i < len(d) else None for d in (left, base, right)]
            matrix.append(_make_row(cols, *rows))
            meta.append(tuple(i if i < len(d) else None for d in (left, base, right)))
        return matrix, meta

    kr = key_row if (key_row and key_row > 0) else 0
    n_head = kr + 1
    lk, bk, rk = cols[key_col] if 0 <= key_col < len(cols) else (key_col,) * 3

    rows3 = [_body_rows(d, kc, n_head) for d, kc in ((left, lk), (base, bk), (right, rk))]
    heads = tuple(_head_rows(d, kr) for d in (left, base, right))
    key_t = cols[key_col] if 0 <= key_col < len(cols) else None
    # 키 행 바로 아래 글자 행 묶음(2·3단 머리글의 설명·형식 행 — '정렬'·'int32')이 세 파일 모두 같은 줄 수로
    # 있고 번호 열을 설명하면, 머리글 행처럼 그 자리끼리 짝짓고 내용 짝짓기에서는 뺀다 — 본문 첫 행으로 남으면
    # 번호 열 첫 값이 '정렬'이 되어 다시 매김 감지가 꺼지고 복사본 끼움을 원본과 짝지었다(검토 실측). 한쪽이 열을
    # 더하거나 설명 칸을 고쳐도 그 자리의 칸 변경으로 보인다(검토 실측). 한쪽만 줄 수가 다르면(설명 행을 지우거나
    # 본문 첫 행이 바뀜) 지금처럼 본문으로 둔다.
    # 세 파일 모두에 있는 앞쪽 줄까지만 본다(한쪽만 주석 행을 더하거나 빼도 앞쪽 설명 행은 그대로 — 검토 실측).
    # 세 파일에서 똑같은 줄은 그대로 짝짓는다. 다른 줄은 **머리글처럼 칸을 넓게 채운 행**(머리글 이름 칸의 절반
    # 이상)이고 키가 아닌 채운 칸이 양쪽 모두 절반 넘게 같을 때만 — 설명 칸·키 칸 수정, 열 추가는 칸 변경으로
    # 보이고(검토 실측), 상수 표·겹친 키 묶음의 맨 위 행 순서 바꿈·바꿔 넣기(다른 칸도 다름)와 성긴 장 구분·주석
    # 행('1장 숲의 입구', '// 1장')은 짝짓지 않는다 — 자리로 잘못 짝지으면 표시 없이 자동 병합됐다(검토 실측).
    # 다른 줄은 키가 아닌 채운 칸이 아래 본문(같은 열 앞쪽 20행)에 같은 값으로 나오면 데이터 행이다('Hero'·
    # 'Normal' — 보상 묶음 맨 위 행을 끼우거나 맞바꾸면 엉뚱한 행에 병합됐다, 검토 실측). 그래도 자리로 짝지은
    # 다른 줄은 '불확실'로 표시해 자동 병합에서 뺀다 — 짝이 틀려도 표시 없이 쓰이지 않게(검토 실측).
    datas = (left, base, right)
    named = sum(1 for v in (base[kr] if kr < len(base) else []) if str(v).strip() != "")
    sub, same_rows, proj_sets = 0, [], []
    for j in range(min(len(h) for h in heads) - (kr + 1)):
        pr = [_project(d[kr + 1 + j], cols, f) for f, d in enumerate(datas)]
        same = pr[0] == pr[1] == pr[2]
        if not same:
            filled = [c for c, v in enumerate(pr[1]) if v != "" and c != key_col]
            if sum(1 for v in pr[1] if v != "") * 2 < max(named, 2) or not filled:
                break
            if not all(2 * sum(1 for c in filled if pr[f][c] == pr[1][c]) > len(filled) for f in (0, 2)):
                break
            below = [_project(r, cols, 1) for r in base[kr + 1 + j + 1:kr + 1 + j + 21]]
            if any(any(b[c] == pr[1][c] for b in below) for c in filled):
                break
            # 옮긴 흔적 — 가운데 줄(키나 내용)이 그쪽 다른 자리에 있거나, 그쪽 줄이 가운데 다른 자리에 있으면 그
            # 행은 옮겨졌다(상수 표 맨 위 두 행 맞바꿈·끼움 — 검토 실측). 키 칸만 고친 설명 행('아이디'→'고유 ID')은
            # 옛 값이 다른 곳에 없으므로 그대로 둔다.
            if not proj_sets:
                for f, d in enumerate(datas):
                    rows_f = d[kr + 1:]
                    proj_sets.append(({_project(r, cols, f): i for i, r in enumerate(rows_f, kr + 1)},
                                      {_project(r, cols, f)[key_col]: i for i, r in enumerate(rows_f, kr + 1)}
                                      if 0 <= key_col < len(cols) else {}))
            at = kr + 1 + j
            moved = False
            for f in (0, 2):
                for a, b in ((1, f), (f, 1)):
                    whole, keyed = proj_sets[b]
                    if whole.get(pr[a], at) != at:
                        moved = True
                    k = pr[a][key_col] if 0 <= key_col < len(cols) else ""
                    if k != "" and keyed.get(k, at) != at:
                        moved = True
            if moved:
                break
        same_rows.append(same)
        sub = j + 1
    if sub and not _describes_numbers(base, kr + 1, sub, [t[1] for t in cols]):
        sub = 0
    if sub:
        rows3 = [[e for e in rows if not kr < e[0] <= kr + sub] for rows in rows3]
    (l_map, l_order), (b_map, b_order), (r_map, r_order) = _row_ids(rows3, cols, heads, key_t)
    if sub:
        hks = [f"\x00sub-head{j}" if same_rows[j] else ("\x00sub-head", j) for j in range(sub)]
        for mp, order, d in ((l_map, l_order, left), (b_map, b_order, base), (r_map, r_order, right)):
            for j, hk in enumerate(hks):
                mp[hk] = (kr + 1 + j, d[kr + 1 + j])
            order[0:0] = hks

    keys = list(b_order)
    present = set(keys)
    keys = _weave(keys, l_order, present)
    present = set(keys)
    keys = _weave(keys, r_order, present)

    matrix, meta = [], []
    for i in range(min(n_head, max(len(left), len(base), len(right)))):
        rows = [d[i] if i < len(d) else None for d in (left, base, right)]
        matrix.append(_make_row(cols, *rows))
        meta.append(tuple(i if i < len(d) else None for d in (left, base, right)))
    for k in keys:
        li, lrow = l_map.get(k, (None, None))
        bi, brow = b_map.get(k, (None, None))
        ri, rrow = r_map.get(k, (None, None))
        if ambiguous is not None and isinstance(k, tuple):
            ambiguous.add(len(matrix))
        matrix.append(_make_row(cols, lrow, brow, rrow))
        meta.append((li, bi, ri))
    return matrix, meta


def _describes_numbers(data, i, m, idxs) -> bool:
    """data[i:i+m] 가 설명·형식 행 묶음인가 — 어떤 열에서 그 칸은 글자인데 묶음 아래 본문(앞쪽 20개)은 모두
    정수다('정렬'·'int32' 아래 1, 2, 3 …). 모든 칸이 글자인 본문 첫 행(문자열 표)은 설명 행이 아니다."""
    for c in idxs:
        if c is None or not any(c < len(r) and r[c] != "" and not _is_number(r[c]) for r in data[i:i + m]):
            continue
        below = []
        for r in data[i + m:]:
            v = r[c] if c < len(r) else ""
            if v != "":
                below.append(v)
                if len(below) == 20:
                    break
        if len(below) >= 3 and all(_is_int(v) for v in below):
            return True
    return False


def _body_rows(data, kc, n_head):
    """본문 행 [(원본 행 번호, 행, 키)] — 값이 하나도 없는 행은 뺀다."""
    out = []
    for i, row in enumerate(data[n_head:], start=n_head):
        k = row[kc] if (kc is not None and kc < len(row)) else ""
        if k == "" and not any(v != "" for v in row):
            continue
        out.append((i, row, k))
    return out


def _project(row, cols, idx):
    """한 파일의 행을 주어진 열 순서의 값으로 — 세 파일의 행을 같은 자로 견주기 위해."""
    return tuple(row[t[idx]] if (t[idx] is not None and t[idx] < len(row)) else ""
                 for t in cols)


_PAIR_DP_MAX = 10_000       # 바뀐 구간을 모든 쌍으로 계산하는 크기(a×b) — 넘으면 띠로
_PAIR_WORK_MAX = 5_000_000  # 〃 칸 비교 수(a×b×열) 상한
_GROUP_MAX = 300            # 공통 앞뒤를 뗀 가운데 부분 행 수 — 넘으면 빠른 정렬(autojunk)로
_DIRECT_MAX = 2_500         # 이보다 작은 가운데(a×b)만 닻 없이 바로 정렬 — 크면 닻을 먼저 잡는다


def _cell_sim(p, q) -> float:
    """두 칸 값이 얼마나 비슷한가(0~1). 같으면 1, 한쪽만 비었으면 0, 아니면 앞뒤로 같은 글자 비율의
    절반(동점을 가르는 정도로만 — '100'→'120' 은 'gold'→'potion' 보다 비슷하다). 칸마다 문자열
    정렬을 돌리면 2천 행 묶음에서 수십 초가 걸렸다(검토 실측) — 앞뒤 공통 길이만 본다."""
    if p == q:
        return 1.0
    if p == "" or q == "":
        return 0.0
    n = min(len(p), len(q))
    pre = 0
    while pre < n and p[pre] == q[pre]:
        pre += 1
    suf = 0
    while suf < n - pre and p[-1 - suf] == q[-1 - suf]:
        suf += 1
    return 0.5 * (pre + suf) / max(len(p), len(q))


_BAND_CELLS_MAX = 2_000_000  # 띠 DP 칸 수 상한 — 넘으면(행 수가 크게 다를 때) 앞에서부터


def _pair_block(b_seg, x_seg, b_tags=None, x_tags=None, weak=()):
    """바뀐 구간 안에서 **비슷한 행끼리** 순서를 지키며 짝짓는다 → [(i, j)].

    앞에서부터 짝지으면, 한쪽이 행 하나를 끼우고 바로 다음 행도 고쳤을 때 기존 행이 끼운 행과
    짝지어져 다른 쪽 수정이 엉뚱한 행에 붙었다(검토 실측).
      · 행 유사도 = 칸 유사도의 평균(같은 칸 1, 다른 칸은 앞뒤 공통 글자 비율의 절반 — 동점 가르기).
        **구간 안에서 값이 달라지는 열만** 센다 — 빈 열·늘 같은 값 열이 많으면 모든 쌍이 비슷해져
        끼운 행과 고친 행을 가르지 못했다(검토 실측: 그런 열 4개부터 틀림).
      · 모든 열의 같은 칸이 절반 이상인 쌍만 후보. 모든 칸이 같은 짝(그대로인 행)에는 덤을 준다.
      · 짝 위치를 옮길 때마다(건너뛰기를 새로 시작할 때) 비용을 뺀다(어파인 갭) — 짝 수가 많을수록
        유리해지지 않게(약한 짝 둘이 강한 짝 하나를 이기던 문제), 거의 같은 행이 많을 때 지그재그로
        짝짓지 않게.
      · 구간이 크면 **대각선 근처 띠**만 계산한다(띠 폭 = 행 수 차이 + 40). 예전엔 상한을 넘으면
        앞에서부터 짝지어 한 행만 끼워도 뒤가 모두 밀렸고(검토 실측: 360행×40열), 모든 쌍을 계산하면
        수 초가 걸렸다(검토 실측: 기본값 많은 표 10초).
      · weak 열(다시 매겼을 수 있는 순번 열)은 짝 자격에는 넣되 **고르는 점수에서는 뺀다** — 끝에서 두 번째
        보상을 지우고 Seq 를 다시 매기면 'Seq 가 같은 짝'과 'RewardID 가 같은 짝'이 동점이 되고, 글자
        유사도(10500↔10400)가 엉뚱한 쪽을 골랐다(검토 실측).
    """
    a, b = len(b_seg), len(x_seg)
    if not a or not b:
        return []
    ncol = max(len(r) for r in b_seg + x_seg)

    def at(r, c):
        return r[c] if c < len(r) else ""
    first = b_seg[0]
    vary = [c for c in range(ncol)
            if any(at(r, c) != at(first, c) for r in b_seg) or
            any(at(r, c) != at(first, c) for r in x_seg)] or list(range(ncol))
    bs = [tuple(at(r, c) for c in vary) for r in b_seg]
    xs = [tuple(at(r, c) for c in vary) for r in x_seg]
    width = max(1, len(vary))
    n_const = max(0, ncol - len(vary))       # 구간 전체에서 늘 같은 열 — 모든 쌍에서 같다
    weak_set = set(weak or ())
    strong = [k for k, c in enumerate(vary) if c not in weak_set]
    if len(strong) == len(vary):
        s_get = None
    else:
        s_width = max(1, len(strong))
        s_get = (itemgetter(*strong) if len(strong) > 1 else
                 (lambda t, k=strong[0]: (t[k],)) if strong else (lambda t: ()))

    # 100×100 을 넘는 구간은 대각선 근처 띠만 계산한다(띠 폭 = 행 수 차이 + 40) — 모든 쌍을 계산하면
    # 기본값('0')이 대부분인 표에서 수 초가 걸렸다(검토 실측: 250행 묶음 8개 7.9초).
    band = None if a * b <= _PAIR_DP_MAX else abs(a - b) + 40
    if band is not None and a * (2 * band + 1) > _BAND_CELLS_MAX:
        # 너무 커서 정렬하지 못하면 짝짓지 않는다(지운 행 + 새 행) — 예전엔 n번째끼리 짝지어, 행을 끼운
        # 자리 아래가 모두 밀린 채 한쪽 수정이 충돌 없이 다른 행에 저장됐다(검토 실측). 짝이 없으면 다른
        # 쪽 수정은 '지운 행을 고침' 충돌로 드러난다.
        return []

    def in_band(i, j):
        if band is None:
            return True
        return abs(j - i * b / a) <= band + 1

    sims: dict = {}
    # 같은 내용의 행이 많은 표(기본값 '0' 이 대부분)는 같은 행 쌍을 몇 번이고 다시 견준다 — 내용별 번호를
    # 매겨 행 쌍 점수를 한 번만 계산한다.
    intern: dict = {}
    bid = [intern.setdefault(t, len(intern)) for t in bs]
    xid = [intern.setdefault(t, len(intern)) for t in xs]
    by_content: dict = {}
    cell_cache: dict = {}

    def with_tag(v, i, j):
        """다른 구간(표식)이면 감점 — 내용이 같은 행이 여러 구간에 있을 때 제 구간끼리 짝짓게 한다.
        자격까지 막지는 않는다 — 고유 키 행 하나만 옮겨도 아래 묶음 표식이 바뀌어 같은 행끼리도 짝을
        못 짓고 묶음이 통째로 '지움+새 행'이 됐다(실측). 같은 구간에 가산점을 주면 짝마다 덤이 붙어
        약한 짝 둘이 강한 짝 하나를 다시 이겼다(실측) — 그래서 감점으로."""
        if v >= 0 and b_tags is not None and b_tags[i] != x_tags[j]:
            v -= 0.3
            return v if v > 0 else -1.0
        return v

    def sim(i, j):
        key = (i, j)
        v = sims.get(key)
        if v is None:
            ck = (bid[i], xid[j])
            v = by_content.get(ck)
            if v is not None:
                v = sims[key] = with_tag(v, i, j)
                return v
            # 후보 자격은 **모든 열**로(절반 이상 비슷), 고르는 점수는 **달라지는 열**로 본다.
            # 모든 열로 점수를 매기면 늘 같은 열이 많을 때 아무 쌍이나 비슷해 약한 짝 둘이 강한 짝
            # 하나를 이겼고, 달라지는 열로 자격까지 보면 한 칸만 고친 1행 대 1행도 짝을 못 지었다.
            # 자격은 **같은 칸 수**로만 — 글자 유사도까지 자격에 넣으면 거의 모든 쌍에 문자열 계산이 돌아
            # 50행 묶음 40개가 16초 걸렸다(실측). 글자 유사도는 자격을 얻은 쌍의 점수(동점 가르기)에만.
            bi, xj = bs[i], xs[j]
            same = sum(map(eq, bi, xj))
            if s_get is not None:
                pi, qj = s_get(bi), s_get(xj)
                ss = sum(map(eq, pi, qj))
                # 자격도 다시 매긴 순번 열을 빼고 센다 — 번호가 바뀐 데다 스탯 셋을 올린 행이 '절반 이상 같음'을
                # 못 넘어 짝을 잃고 '지움 + 새 행'으로 갈렸다(검토 실측)
                elig = (n_const + ss) * 2 >= ncol - (len(vary) - len(strong))
            else:
                elig = (n_const + same) * 2 >= ncol
            if not elig:
                v = -1.0
            elif s_get is not None:
                # 점수는 순번 열을 뺀 열로 — 나머지가 모두 같으면(번호만 다시 매긴 행) 같은 행 덤
                extra = 0.0
                if 0 < s_width - ss <= 3:
                    for p, q in zip(pi, qj):
                        if p != q:
                            cs = cell_cache.get((p, q))
                            if cs is None:
                                cs = cell_cache[(p, q)] = _cell_sim(p, q)
                            extra += cs
                v = (ss + extra) / s_width + 1e-6 + (0.5 if ss == s_width else 0.0)
            else:
                # 동점 가르기(글자 유사도)는 다른 칸이 몇 개뿐인 쌍에만 — 많이 다르면 의미가 없고, 모든
                # 쌍에 돌리면 기본값 많은 2천 행 표에서 340만 번 불렸다(실측). 같은 값 쌍은 한 번만.
                extra = 0.0
                if 0 < width - same <= 3:
                    for p, q in zip(bi, xj):
                        if p != q:
                            cs = cell_cache.get((p, q))
                            if cs is None:
                                cs = cell_cache[(p, q)] = _cell_sim(p, q)
                            extra += cs
                # 모든 칸이 같은 짝(그대로인 행)에는 덤을 준다 — 없으면 '복사해 옆에 넣은 새 행 + 지운
                # 이웃 행'의 약한 짝 둘이 그대로인 행 하나를 이겨, 다른 쪽 수정이 새 행에 붙었다(검토 실측).
                v = (same + extra) / width + 1e-6 + (0.5 if same == width else 0.0)
            by_content[ck] = v
            v = sims[key] = with_tag(v, i, j)
        return v

    # 정렬 — 짝 점수의 합에서 **짝 위치를 옮길 때마다**(건너뛰기를 새로 시작할 때) 비용을 뺀다(어파인 갭).
    # 내용이 거의 같은 행이 많으면 '조금 더 비슷한 다른 행'으로 지그재그 짝을 지어 양쪽 짝이 엇갈리고
    # 거짓 충돌이 수천 칸 났다(검토 실측). 행 하나를 끼운 것처럼 한 번 옮기는 것은 싸게, 자꾸 옮기는
    # 것은 비싸게. 건너뛴 행 하나하나에는 비용을 매기지 않는다 — 그러면 짝 수가 많을수록 유리해져
    # 약한 짝 둘이 강한 짝 하나를 이겼다(검토 실측).
    GAP = 0.15
    NEG = -1e18
    M, X, Y = 0, 1, 2           # 직전 이동: 짝(또는 시작) / 가운데 행 건너뜀 / 한쪽 행 건너뜀
    dp: dict = {}

    def val(i, j, st):
        if i == a or j == b:
            if i == a and j == b:
                return 0.0
            if i == a:                       # 남은 한쪽 행을 건너뛴다
                return 0.0 if st == Y else -GAP
            return 0.0 if st == X else -GAP
        if not in_band(i, j):
            return NEG
        return dp[(i, j, st)]

    for i in range(a - 1, -1, -1):
        if band is None:
            js = range(b - 1, -1, -1)
        else:
            c = i * b / a
            js = range(min(b - 1, int(c) + band + 1), max(-1, int(c) - band - 2), -1)
        for j in js:
            if not in_band(i, j):
                continue
            sv = sim(i, j)
            pair = sv + val(i + 1, j + 1, M) if sv >= 0 else NEG
            skip_b = val(i + 1, j, X)
            skip_x = val(i, j + 1, Y)
            dp[(i, j, M)] = max(pair, skip_b - GAP, skip_x - GAP)
            dp[(i, j, X)] = max(pair, skip_b, skip_x - GAP)
            dp[(i, j, Y)] = max(pair, skip_b - GAP, skip_x)

    pairs, i, j, st = [], 0, 0, M
    eps = 1e-9
    while i < a and j < b:
        if not in_band(i, j):                # 띠 밖 — 띠 쪽으로 돌아간다
            if j < i * b / a:
                j += 1
            else:
                i += 1
            continue
        here = dp[(i, j, st)]
        sv = sims.get((i, j), -1.0)
        if sv >= 0 and abs(here - (sv + val(i + 1, j + 1, M))) < eps:
            pairs.append((i, j))
            i, j, st = i + 1, j + 1, M
        elif abs(here - (val(i + 1, j, X) - (0 if st == X else GAP))) < eps:
            i, st = i + 1, X
        else:
            j, st = j + 1, Y
    return pairs


def _step_of(vals):
    """1, 2, 3 … 처럼 **빠짐없이 일정한 간격**으로 늘거나 주는 정수 열이면 그 간격, 아니면 None."""
    if len(vals) < 3:
        return None
    try:
        nums = [int(v) for v in vals]
    except (TypeError, ValueError):
        return None
    step = nums[1] - nums[0]
    if step == 0 or any(b - a != step for a, b in zip(nums, nums[1:])):
        return None
    return step


_SEQ_WINDOW = 8          # 순번 열로 짝지은 행은 앞뒤 이만큼 안의 행보다 비슷해야 닻이 된다
_ONLY_HERE = object()    # 투영한 행이 한쪽에서만 한 번 나오지 않을 때 끼우는 표식(어떤 값과도 다르다)


def _unique_projected_runs(bs2, xs2, drop) -> list:
    """drop 열을 뺀 나머지가 똑같은 행으로 이어지는 3행 이상 덩어리 [(i, j, 길이)].

    다시 매긴 순번 열(Seq·No)만 다르고 나머지가 그대로인 행을 짝짓는다 — 순번을 다시 매기면 끼운 자리
    아래가 모두 '바뀐 행'이라 똑같은 행 덩어리가 없고, 순번 값으로는 k칸 밀린 짝만 나와 큰 시트가 띠
    정렬(수 초)이나 위치 짝(다른 보상에 저장)으로 떨어졌다(검토 실측).
    투영한 행이 **양쪽에서 한 번씩만** 나올 때만 견준다 — 값 종류가 적은 열만 남으면(일괄 수정 뒤 등급
    1~5 같은) 우연히 같은 행이 여럿 이어져 엉뚱한 자리를 닻으로 잡는다."""
    ncol = len(bs2[0])
    keep = [k for k in range(1, ncol) if k not in drop]
    if not keep:
        return []
    get = itemgetter(*keep)
    pb = [get(t) for t in bs2]
    px = [get(t) for t in xs2]
    cb, cx = Counter(pb), Counter(px)
    sb = [t if cb[t] == 1 and cx.get(t, 0) == 1 else (_ONLY_HERE, 0, i) for i, t in enumerate(pb)]
    sx = [t if cx[t] == 1 and cb.get(t, 0) == 1 else (_ONLY_HERE, 1, j) for j, t in enumerate(px)]
    return [(i, j, n) for i, j, n in SequenceMatcher(None, sb, sx, autojunk=False).get_matching_blocks()
            if n >= 3]


def _id_anchor_runs(bs, xs) -> list:
    """거의 고유한 열 값으로 짝지은 닻 [(i, j, 길이)] — 길이 3 이상, 짝지은 행이 **이웃보다** 비슷한 구간.

    후보 열은 가운데 쪽에서 빈 칸이 거의 없고(90% 이상) 값이 거의 다른(90% 이상) 열. 열마다 값만으로
    정렬해 보고, 그렇게 짝지은 행 (i, j) 가 옆 짝 (i, j±w) 보다 더 비슷할 때만(후보 열과 조각 전체에서
    늘 같은 값인 열을 뺀 나머지 열의 같은 칸 수로) 이어지는 행으로 본다. 값으로 짝지은 쌍 가운데 이웃
    검사를 넘는 비율이 모자란 열은 버린다 — 우연히 이어진 구간만 남은 열이다.
      · **순번 열**(바뀐 쪽 파일에서 같은 간격으로 빠짐없이 이어지는 정수 열 — 다시 매긴 Seq·No·Sort, 연속
        번호 ID)은 값이 행 위치를 따라갈 수 있다. 행을 여럿 끼우거나 지우고 번호를 다시 매기면 값으로 지은
        짝이 k칸 밀리는데, 이웃도 모두 틀린 행이라 ±1 검사를 우연히 넘었다(검토 실측). 다시 매긴 쪽은 늘
        빈틈이 없다 — 가운데만 연속인 열은 바뀐 쪽이 행을 지우거나 끼운 진짜 ID 라 순번으로 보지 않는다
        (그렇게 보면 일괄 수정 표에서 ID 짝을 버려 모든 행이 밀렸다 — 검토 실측). 이런 열은 앞뒤
        _SEQ_WINDOW 칸 안의 모든 행과 견주고, 넘는 비율도 90% 이상이어야 한다. 진짜 ID 열을 먼저 본다.
        동점 판에서 두 파일 행 수가 다르면 값으로 가린다 — 값 짝이 곧 행 수 차이만큼 밀린 짝이면 진짜 ID(믿음),
        같은 자리 값이 같고 밀린 짝이 더 비슷한 행이 10% 이상이면 다시 매긴 순번(쓰지 않음), 그 밖(끝에 붙임·
        끝을 지움)은 믿는다. 믿을 때 통과 비율 기준은 50%.
    순서: ① 이웃보다 **더** 비슷한 짝(엄격) — 한 열이 모든 행을 덮으면 바로 끝. ② 순번 열이 있으면 그 열을
    빼고 똑같은 행 덩어리(_unique_projected_runs). ③ 그래도 없으면 **동점도** 이어지는 행으로 본다 — ID
    말고는 행마다 달라지는 열이 없는 표(스탯만 있는 표를 일괄 수정)에서는 진짜 짝도 이웃과 동점이라 닻이
    하나도 안 잡혀 위치로 짝지어졌다(검토 실측). 동점 판에서는 모든 열을 앞뒤 _SEQ_WINDOW 칸과 견준다.
    맨 앞 칸은 구간 표식이라 열로 보지 않는다. 행은 비교할 열로 한 번만 투영해 C 수준으로 견준다
    (쌍마다 모든 열을 파이썬으로 훑으면 2만 행 표가 수 초씩 더 걸렸다 — 검토 실측)."""
    if not bs or not xs:
        return []
    ncol = max(len(t) for t in bs + xs)

    def pad(t):
        t = tuple(t)
        return t if len(t) >= ncol else t + ("",) * (ncol - len(t))
    bs2 = [pad(t) for t in bs]
    xs2 = [pad(t) for t in xs]
    nx = len(xs2)
    full = min(len(bs2), nx)

    cands, constant, positional = [], set(), set()
    for c in range(1, ncol):
        get = itemgetter(c)
        vals = list(map(get, bs2))
        xv = list(map(get, xs2))
        if len(set(vals).union(xv)) == 1:
            constant.add(c)                  # 모든 짝에 똑같이 더해져 이웃과 견주면 상쇄된다
            continue
        filled = [v for v in vals if v != ""]
        if len(filled) < 0.9 * len(vals) or len(set(filled)) < 0.9 * len(filled):
            continue
        pos = _step_of(xv) is not None
        if pos:
            positional.add(c)
        cands.append((c, vals, xv, pos))
    # 진짜 ID 열을 먼저 보되, 상한 때문에 순번 열이 통째로 잘리지 않게 따로 센다 — 거의 고유한 스탯 열이
    # 8개 넘는 표를 일괄 수정하면 연번 ID 가 후보에서 빠져 닻이 하나도 없었다(검토 실측).
    cands = [t for t in cands if not t[3]][:8] + [t for t in cands if t[3]][:2]

    def runs_for(c, vals, xv, w, allow_tie, need):
        cmp_cols = [k for k in range(1, ncol) if k != c and k not in constant]
        if not cmp_cols:
            pb = px = None
        else:
            get = itemgetter(*cmp_cols)
            pb = [get(t) for t in bs2]
            px = [get(t) for t in xs2]
            if len(cmp_cols) == 1:
                pb = [(v,) for v in pb]
                px = [(v,) for v in px]

        def same(i, j):
            return sum(map(eq, pb[i], px[j])) if pb is not None else 0

        if allow_tie and w > 1 and c in positional:
            # 행 수가 다를 때 이 연속 번호가 무엇인지 값으로 가린다(d = 행 수 차이).
            #   · 값 짝이 곧 d 만큼 밀린 짝이면(앞을 지우거나 앞에 붙인 진짜 연번 ID) ID 로 믿는다.
            #   · 같은 자리 값이 같고(번호가 위치를 따름) 밀린 짝이 더 비슷한 행이 10% 이상이면 중간에 끼우거나
            #     지우고 다시 매긴 순번 — 쓰지 않는다.
            #   · 같은 자리 값이 같고 밀린 짝이 낫지 않으면(끝에 붙이거나 끝을 지움) 값 짝이 맞다.
            # 위치만 보면 앞을 지운 진짜 ID 를 순번으로 버려, 수 초 정렬하거나 시트 전체가 '지움+새 행'이
            # 됐다(검토 실측). 믿는 경우는 통과 비율 기준을 진짜 ID 와 같게(50%) 둔다 — 등급을 일부 바꾼 행이
            # 이웃에 져 90% 를 못 넘기고 ID 를 버렸다(검토 실측).
            d = nx - len(bs2)
            idx = [i for i in range(len(bs2)) if i < nx and 0 <= i + d < nx]
            m0 = min(len(bs2), nx)
            if m0 and sum(1 for i in range(m0) if vals[i] == xv[i]) < 0.5 * m0:
                # 같은 자리 값이 절반도 안 맞는다 — 번호가 위치를 따르지 않는 진짜 연번 ID(맨 위를 지우고 끝에
                # 같은 수만큼 붙여 행 수가 같은 경우 포함 — 검토 실측)
                need = min(need, 0.5)
            elif d and idx:
                shifted = sum(1 for i in idx if vals[i] == xv[i + d])
                atpos = sum(1 for i in idx if vals[i] == xv[i])
                if shifted >= 0.9 * len(idx):
                    need = min(need, 0.5)
                elif atpos >= 0.5 * len(idx):
                    better = sum(1 for i in idx if same(i, i + d) > same(i, i))
                    if better >= 0.1 * len(idx):
                        return [], 0.0
                    need = min(need, 0.5)

        def keeps(i, j):
            """(i, j) 가 앞뒤 w 칸 안의 다른 짝보다 비슷한가(allow_tie 면 같아도 된다), 그리고 그 차이."""
            own = same(i, j)
            nb = -1
            for jj in range(max(0, j - w), min(nx, j + w + 1)):
                if jj == j:
                    continue
                v = same(i, jj)
                if v > nb:
                    nb = v
                    if v > own or (v == own and not allow_tie):
                        return None
            return own - max(nb, 0)
        runs, score, matched, passed = [], 0.0, 0, 0
        for i0, j0, size in SequenceMatcher(None, vals, xv, autojunk=False).get_matching_blocks():
            k = 0
            while k < size:
                start, margin = k, 0.0
                while k < size:
                    matched += 1
                    gap = keeps(i0 + k, j0 + k)
                    if gap is None:
                        break
                    passed += 1
                    margin += gap + (1.0 if allow_tie else 0.0)
                    k += 1
                if k - start >= 3:
                    runs.append((i0 + start, j0 + start, k - start))
                    score += margin
                k += 1
        if matched and passed < need * matched:
            return [], 0.0
        return runs, score

    def best_of(allow_tie):
        best, best_score = [], 0.0
        for c, vals, xv, pos in cands:
            w = _SEQ_WINDOW if (pos or allow_tie) else 1
            runs, score = runs_for(c, vals, xv, w, allow_tie, 0.9 if pos else 0.5)
            if not allow_tie and runs and sum(n for _i, _j, n in runs) >= full:
                return runs                  # 모든 행이 이웃보다 나은 짝 — 다른 열이 더 낫지 않다
            if score > best_score:
                best, best_score = runs, score
        return best

    best = best_of(False)
    if best:
        return best
    if positional:
        best = _unique_projected_runs(bs2, xs2, positional | constant)
        if best:
            return best
    return best_of(True)


_JOIN_STATS_MIN = 20    # 열 통계(순번 열·일괄 수정 열·ID 열)를 믿을 만한 가운데 크기(양쪽 행 수)


def _join_pairs(mb, mx, weak=frozenset(), b_rows=None, x_rows=None, id_cols=()) -> dict:
    """순서와 상관없이 **확실한** 짝 {한쪽 행: 가운데 행}.

    순서를 지키며 정렬하면 묶음을 정렬하거나 두 행을 맞바꿨을 때 다른 보상끼리 짝지어 한쪽 수정이 충돌
    없이 다른 행에 저장됐고(검토 실측), 번호를 다시 매기거나 한 열을 모든 행에서 고치면 같은 행이 하나도
    없어 큰 시트가 수 초 정렬하거나 밀려 짝지어졌다(검토 실측).
      ⓪ **모든 칸이 같고** 가운데와 그쪽에 한 번씩만 나오는 행끼리(자리를 옮긴 행).
      ① 가운데가 크면(양쪽 _JOIN_STATS_MIN 행 이상) 짝의 근거에서 두 가지 열을 빼고 한 번 더:
         바뀐 쪽에서 같은 간격으로 빠짐없이 이어지는 정수 열(다시 매긴 Seq·No 일 수 있다), 두 파일 값
         묶음이 10%(그리고 3행) 넘게 다른 열(많은 행에서 고친 열 — 일괄 수정은 짝의 근거가 못 된다).
      ② 남은 행을 **ID 열**의 값이 가운데 양쪽에 한 번씩만 있고 쓰는 열의 절반 이상이 같을 때 짝 — ID
         말고 다른 열도 고친 행(등급 변경·확률 조정·다음 ID 로 복사한 새 행 옆의 원래 행)이 정렬에서 이웃과
         동점이 돼 엇갈렸다(검토 실측). ID 열은 **머리글이 ID 이름**(_is_id_name)인 열 가운데 시트 전체 값
         종류가 많은 순으로, **묶음 전체**에서 빈칸이 거의 없고 값이 절반 이상 다른 첫 열 하나뿐이다(0·1 부터
         이어지는 SlotID 같은 묶음 안 순번은 뺀다). 이름이 맞는 열이 없으면 키 열 바로 오른쪽의 고유 정수 열
         (Index·No)을 쓴다. 내용만으로 고르면 Weight·HP 가 ID 노릇을 해 지운 보상과 같은 확률의 새 보상, HP 가
         같은 복사본을 같은 행으로 짝지었다(검토 실측). 쓸 열이 없으면 ② 는 하지 않는다.
    몇 행짜리 가운데에서는 열 통계가 뜻이 없다 — 한 칸 바뀐 열이 '일괄 수정'으로, 모든 열이 '거의 고유'로,
    짧은 RewardID 열이 '순번'으로 보여 지운 보상과 새 보상, 맞바꾼 확률, 다시 매긴 Seq 가 엉뚱하게
    짝지어졌다(검토 실측). 그래서 작은 가운데는 ⓪만 쓴다.
    한 행짜리 짝은 그 행에 **그 열에서 양쪽 모두 한 번만 나오는 값**이 있을 때만(①에서는 뺀 순번 열 값도
    같을 때만 — 다시 매김은 덩어리째 일어난다), 아니면 3행 이상 이어지는 짝만 쓴다 — 기본값이 많은 표에서
    우연히 같아진 행끼리 엇갈려 짝지어지지 않게. 맨 앞 칸(구간 표식)은 늘 근거로 쓴다."""
    na, nx = len(mb), len(mx)
    if not na or not nx:
        return {}
    ncol = max(len(t) for t in mb + mx)

    def pad(t):
        t = tuple(t)
        return t if len(t) >= ncol else t + ("",) * (ncol - len(t))
    b2 = [pad(t) for t in mb]
    x2 = [pad(t) for t in mx]
    big = min(na, nx) >= _JOIN_STATS_MIN
    counts, keep, seq, ident = {}, [0], [], []
    for c in range(1, ncol):
        get = itemgetter(c)
        vb = list(map(get, b2))
        vx = list(map(get, x2))
        cvx = Counter(vx)
        cvb = cvx if vb == vx else Counter(vb)
        counts[c] = (cvb, cvx)
        if c in weak:
            seq.append(c)
            continue
        if not big:
            keep.append(c)                   # 작은 가운데는 열 통계를 쓰지 않는다
            continue
        if c not in id_cols and len(cvx) == nx and _step_of(vx) is not None:
            seq.append(c)                    # 바뀐 쪽의 연속 번호 — 다시 매겼을 수 있다
            continue
        if vb != vx:
            # 많은 행에서 바뀐 열은 짝의 근거가 아니다. 거의 고유한 열(ID)은 절반 넘게 바뀌어야 뺀다 — 보상
            # 몇 개를 바꿨다고 RewardID 를 빼면 남은 Weight 로 지운 보상과 새 보상을 짝지었다(검토 실측).
            filled = na - cvb.get("", 0)
            idlike = filled >= 0.9 * na and len(cvb) - ("" in cvb) >= 0.9 * filled
            differ = min(na, nx) - sum((cvb & cvx).values())
            if differ > max(2, (0.5 if idlike else 0.1) * min(na, nx)):
                continue
        keep.append(c)

    def rare(i, cols):
        """가운데 행 i 에 그 열에서 양쪽 모두 한 번만 나오는 값이 있는가."""
        for c in cols:
            v = b2[i][c]
            if v != "":
                cvb, cvx = counts[c]
                if cvb[v] == 1 and cvx.get(v, 0) == 1:
                    return True
        return False

    out, used = {}, set()

    def join(cols, must_equal):
        get = itemgetter(*cols)
        pb = [get(t) for t in b2]
        px = [get(t) for t in x2]
        cb, cx = Counter(pb), Counter(px)
        xpos = {t: j for j, t in enumerate(px) if cx[t] == 1}
        raw = {}
        for i, t in enumerate(pb):
            if cb[t] == 1:
                j = xpos.get(t)
                if j is not None:
                    raw[j] = i
        run_of, start = {}, None             # 이어진 길이를 한 번에 센다
        for j in sorted(raw):
            if start is None or raw.get(j - 1) != raw[j] - 1:
                start = j
            run_of[j] = start
        run_len = Counter(run_of.values())
        rc = [c for c in cols if c]
        for j, i in raw.items():
            if j in out or i in used:
                continue
            if run_len[run_of[j]] >= 3 or (
                    rare(i, rc) and all(b2[i][c] == x2[j][c] for c in must_equal)):
                out[j] = i
                used.add(i)

    join(list(range(ncol)), [])                                    # ⓪
    if weak:
        # ⓪′ 다시 매긴 순번만 다르고 나머지가 모두 같은 행 — 같은 자리 행보다 같은 칸이 많을 때만(묶음을
        # 정렬하거나 두 행을 맞바꾸고 Seq 를 다시 매기면 Seq 로 자리끼리 짝지어 다른 보상에 저장됐다 — 검토
        # 실측. 연번 ID + 확률 맞바꿈처럼 같은 자리 짝과 동점이면 짓지 않는다)
        rest = [c for c in range(ncol) if c not in weak]
        get = itemgetter(*rest)
        pb = [get(t) for t in b2]
        px = [get(t) for t in x2]
        cb, cx = Counter(pb), Counter(px)
        xpos = {t: j for j, t in enumerate(px) if cx[t] == 1}

        def eqn(i, j):
            return sum(map(eq, b2[i], x2[j]))
        for i, t in enumerate(pb):
            if i in used or cb[t] != 1:
                continue
            j = xpos.get(t)
            if j is None or j in out:
                continue
            own = eqn(i, j)
            if (i < nx and eqn(i, i) >= own) or (j < na and eqn(j, j) >= own):
                continue
            out[j] = i
            used.add(i)
    if big and 1 < len(keep) < ncol:
        join(keep, seq)                                            # ①
    cols = keep[1:]
    p = _id_col(b_rows if b_rows is not None else b2, x_rows if x_rows is not None else x2,
                [c for c in id_cols if c in cols])
    if p is None:
        return out
    cvb, cvx = counts[p]
    xat = {}
    for j, t in enumerate(x2):
        v = t[p]
        if v != "" and j not in out and cvx[v] == 1:
            xat[v] = j
    for i, t in enumerate(b2):
        v = t[p]
        if i in used or v == "" or cvb[v] != 1:
            continue
        j = xat.get(v)
        if j is None or j in out:
            continue
        xr = x2[j]
        if 2 * sum(1 for k in cols if t[k] == xr[k]) >= len(cols):
            out[j] = i
            used.add(i)
    return out


def _id_col(b_rows, x_rows, cols):
    """묶음의 ID 열 — cols(ID 열 후보, 우선순위 순) 가운데 두 파일 묶음 전체에서 빈칸이 10% 미만이고 값이 절반
    이상 다른 첫 열(Gold·Gem 행의 1·2 가 겹치는 RewardID 도 ID 다)."""
    for c in cols:
        ok = True
        for rows in (b_rows, x_rows):
            vals = [t[c] if c < len(t) else "" for t in rows]
            filled = [v for v in vals if v != ""]
            if not vals or len(filled) < 0.9 * len(vals) or len(set(filled)) < 0.5 * len(filled):
                ok = False
                break
        if ok:
            return c
    return None


def _chain(pairs: dict) -> list:
    """짝 {j: i} 가운데 순서를 지키는 가장 긴 짝 열 [(i, j)] — 나머지는 자리를 옮긴 행이다."""
    import bisect
    items = sorted((j, i) for j, i in pairs.items())
    tails, tail_idx, prev = [], [], [-1] * len(items)
    for n, (j, i) in enumerate(items):
        k = bisect.bisect_left(tails, i)
        if k == len(tails):
            tails.append(i)
            tail_idx.append(n)
        else:
            tails[k] = i
            tail_idx[k] = n
        prev[n] = tail_idx[k - 1] if k else -1
    chain = []
    n = tail_idx[-1] if tail_idx else -1
    while n >= 0:
        j, i = items[n]
        chain.append((i, j))
        n = prev[n]
    return chain[::-1]


def _renumbered_cols(x_rows, mb, mx) -> set:
    """다시 매긴 순번으로 보이는 열 — 정렬 점수·확실한 짝의 근거에서 뺄 열(맨 앞 표식 칸 제외 번호).

    후보는 **묶음 전체**의 바뀐 쪽에서 같은 간격으로 빠짐없이 이어지는 정수 열(가운데만 보면 끝 근처를 고친
    묶음에서 1~2행뿐이라 알 수 없다 — 검토 실측). 다시 매긴 Seq 와 연번 ID(RewardID 1001~1004, 몬스터 ID)는
    열 모양만으로는 가를 수 없어 값으로 가린다(머리글이 ID 이름인 열은 _align 이 처음부터 뺀다 — ID 는 다시
    매기지 않는다):
      · 가운데가 작으면 묶음이 0·1 부터 시작하거나(Seq·No 모양) 가운데의 값 묶음이 양쪽에서 다를 때(실제로
        번호가 바뀜)만 순번이다 — 연속 모양만으로 보면 연번 RewardID 를 버려 이웃한 두 행의 확률만 맞바꿔도
        행이 엇갈렸고(검토 실측), 0·1 시작만 보면 구간마다 이어지는 No 를 놓쳤다(검토 실측).
      · 가운데가 크면 같은 자리 값이 절반 이상 같고(번호가 위치를 따름) **행 수 차이만큼 밀린 짝**이 같은 자리
        짝보다 비슷한 행이 10% 이상일 때만 순번이다 — 끝에 붙이거나(밀린 짝이 낫지 않음) 앞을 지운(같은 자리
        값이 다름) 진짜 연번 ID 를 버리면 이웃과 엇갈리거나 수 초 정렬했다(검토 실측). 근거가 없으면 ID 로 둔다."""
    nx = len(x_rows)
    if nx < 3:
        return set()
    ncol = max(len(t) for t in x_rows)
    flat = all(len(t) == ncol for t in x_rows)
    cands = []
    for c in range(1, ncol):
        if flat and _step_of([x_rows[0][c], x_rows[1][c], x_rows[2][c]]) is None:
            continue                         # 앞 세 칸부터 등차가 아니면 볼 필요 없다(넓은 시트 속도)
        col = (list(map(itemgetter(c), x_rows)) if flat
               else [t[c] if c < len(t) else "" for t in x_rows])
        if _step_of(col) is not None:
            cands.append(c)
    if not cands:
        return set()

    def at(t, c):
        return t[c] if c < len(t) else ""
    if min(len(mb), len(mx)) < _JOIN_STATS_MIN:
        out = set()
        for c in cands:
            if int(at(x_rows[0], c)) in (0, 1) or \
                    Counter(at(t, c) for t in mb) != Counter(at(t, c) for t in mx):
                out.add(c)
        return out
    d = len(mx) - len(mb)
    if d == 0:
        return set()
    better, atpos, total = Counter(), Counter(), 0
    for i in range(len(mb)):
        j1 = i + d
        if i >= len(mx) or not 0 <= j1 < len(mx):
            continue
        b, x0, x1 = mb[i], mx[i], mx[j1]
        s0 = sum(map(eq, b, x0))
        s1 = sum(map(eq, b, x1))
        total += 1
        for c in cands:
            e0 = at(b, c) == at(x0, c)
            e1 = at(b, c) == at(x1, c)
            atpos[c] += e0
            if s1 - e1 > s0 - e0:
                better[c] += 1
    return {c for c in cands
            if total and atpos[c] >= 0.5 * total and better[c] >= 0.1 * total}


def _seq_like(x_rows, c) -> bool:
    """바뀐 쪽 묶음에서 처음부터 매긴 번호(1,2,3 / 10,20,30)인 열 — 이름이 SlotID·SubID·SortKey·Sort_No 여도
    묶음 안 순번이다(ID 로 쓰면 다시 매긴 번호끼리, 곧 위치로 짝지었다 — 검토 실측)."""
    col = [t[c] if c < len(t) else "" for t in x_rows]
    return _is_numbering(col)


def _is_numbering(col) -> bool:
    """처음부터 매긴 번호인가 — 정수이고 같은 간격으로 빠짐없이 이어지며 0·1 또는 그 간격에서 시작
    (1,2,3 / 0,1,2 / 10,20,30). 한 행뿐이면 0·1 일 때만."""
    if not col:
        return False
    head = col[:3]
    if not all(_is_int(v) for v in head):
        return False
    h = [int(v) for v in head]
    if len(h) >= 2 and (h[1] - h[0] <= 0 or h[0] not in (0, 1, h[1] - h[0])):
        return False                          # 앞 칸부터 번호 모양이 아니면 끝까지 볼 필요 없다(넓은 시트 속도)
    if not all(_is_int(v) for v in col):
        return False
    nums = [int(v) for v in col]
    if len(nums) == 1:
        return nums[0] in (0, 1)
    step = nums[1] - nums[0]
    if step <= 0 or nums[0] not in (0, 1, step):
        return False
    return all(b - a == step for a, b in zip(nums, nums[1:]))


def _is_int(v) -> bool:
    try:
        int(v)
        return True
    except (TypeError, ValueError):
        return False


def _int_unique(rows, c, ascending=True) -> bool:
    """정수이고 고유 — 가운데 쪽은 오름차순(Index·No 모양)까지 본다. 확률·스탯 열은 대개 정렬돼 있지 않다.
    바뀐 쪽은 새 ID 한 행이 순서 밖에 끼어도 되므로 고유만 본다(오름차순까지 보면 대신 쓰는 ID 가 꺼져 지운
    몬스터와 이웃이 엇갈렸다 — 검토 실측). 구간 표시 행의 빈칸은 뺀다."""
    vals = [t[c] if c < len(t) else "" for t in rows]
    vals = [v for v in vals if v != ""]
    if not vals or not all(_is_int(v) for v in vals) or len(set(vals)) != len(vals):
        return False
    if not ascending:
        return True
    nums = [int(v) for v in vals]
    return all(a < b for a, b in zip(nums, nums[1:]))


def _align(b_rows, x_rows, g, tag, id_cols=(), fallback=None, renum=frozenset(), renum_w=frozenset(),
           sheet_unique=frozenset()):
    """같은 묶음의 가운데 행들과 한쪽 행들을 **내용으로** 짝짓는다 → 한쪽 행마다 식별자.

    같은 행은 같은 행끼리, 바뀐 구간은 비슷한 행끼리(_pair_block), 남는 것은 그 파일에서 새로
    생긴 행. 순서(n번째)로 짝지으면 한 행만 끼워도 그 뒤가 모두 한 칸씩 밀렸다(검토 실측).
    공통 앞뒤를 뗀 가운데는 먼저 **확실한 짝**(_join_pairs — 자리를 옮긴 행도)을 짓고, 그 가운데 순서를
    지키는 짝을 경계로 나눈 사이만 정렬한다. 너무 커서 정렬하지 못하는 조각은 짝짓지 않는다(지운 행 + 새 행).
    """
    out = [None] * len(x_rows)
    nb, nx = len(b_rows), len(x_rows)
    # 공통 앞뒤는 그대로 짝 — 대부분의 묶음은 몇 행만 바뀐다. 남은 가운데만 정렬해 빠르게.
    pre = 0
    while pre < nb and pre < nx and b_rows[pre] == x_rows[pre]:
        out[pre] = (g, "b", pre)
        pre += 1
    suf = 0
    while (suf < nb - pre and suf < nx - pre
           and b_rows[nb - 1 - suf] == x_rows[nx - 1 - suf]):
        out[nx - 1 - suf] = (g, "b", nb - 1 - suf)
        suf += 1
    mb, mx = b_rows[pre:nb - suf], x_rows[pre:nx - suf]
    if mb and mx:
        # 묶음 안 순번(SlotID 1..n)은 ID 가 아니다 — 단 시트 전체에서 고유한 열(1..N 으로 매긴 MonsterID)은 빈 키
        # 시트에서 묶음이 곧 시트라 순번처럼 보일 뿐이다(빼면 다음 ID 로 복사한 새 행과 원래 행을 엇갈렸다 — 검토 실측).
        ids = [c for c in id_cols if c not in renum and (c in sheet_unique or not _seq_like(x_rows, c))]
        if not id_cols and fallback is not None and fallback not in renum \
                and not _seq_like(x_rows, fallback) \
                and _int_unique(b_rows, fallback, True) and _int_unique(x_rows, fallback, False):
            ids = [fallback]
        weak = (_renumbered_cols(x_rows, mb, mx) | set(renum_w)) - set(ids)
        joined = _join_pairs(mb, mx, weak, b_rows, x_rows, ids)
        for j, i in joined.items():
            out[pre + j] = (g, "b", pre + i)
        used = set(joined.values())
        bounds = [(-1, -1)] + _chain(joined) + [(len(mb), len(mx))]
        for (i1, j1), (i2, j2) in zip(bounds, bounds[1:]):
            bi = [i for i in range(i1 + 1, i2) if i not in used]
            xj = [j for j in range(j1 + 1, j2) if j not in joined]
            if bi and xj:
                for i, j in _align_middle([mb[i] for i in bi], [mx[j] for j in xj], weak):
                    out[pre + xj[j]] = (g, "b", pre + bi[i])
    for j in range(nx):
        if out[j] is None:
            out[j] = (g, tag, j)
    return out


def _align_middle(mb, mx, weak=frozenset()) -> list:
    """가운데 조각 정렬(순서를 지킴) → [(i, j)]. 맨 앞 칸은 구간 표식이다 — 같은 구간끼리만 짝이 되고
    (다르면 감점), 유사도에는 넣지 않는다(넣으면 늘 같은 값이라 모든 쌍이 비슷해진다)."""
    pairs = []
    na, nb2 = len(mb), len(mx)

    def pair(seg_b, seg_x, ob, ox):
        """조각 하나를 정렬(어파인 갭 DP) — 맨 앞 칸은 구간 표식이라 떼어 따로 넘긴다."""
        if not seg_b or not seg_x:
            return
        for i, j in _pair_block([t[1:] for t in seg_b], [t[1:] for t in seg_x],
                                [t[0] for t in seg_b], [t[0] for t in seg_x],
                                weak=[c - 1 for c in weak]):
            pairs.append((ob + i, ox + j))

    def pair_big(seg_b, seg_x, ob, ox):
        """큰 조각 — 거의 고유한 열(ItemID 등)의 값으로 닻을 한 번 더 잡고 그 사이만 정렬한다.
        블록을 옮기면서 모든 행을 함께 고치면(정렬 번호 다시 매김·일괄 수정) '똑같은 행' 덩어리가 없어
        닻이 안 잡히고, 띠 밖으로 옮긴 블록이 한 칸씩 밀려 짝지어졌다(검토 실측)."""
        if len(seg_b) * len(seg_x) <= _PAIR_DP_MAX:
            pair(seg_b, seg_x, ob, ox)
            return
        qi = qj = 0
        for i0, j0, size in _id_anchor_runs(seg_b, seg_x) + [(len(seg_b), len(seg_x), 0)]:
            pair(seg_b[qi:i0], seg_x[qj:j0], ob + qi, ox + qj)
            for d in range(size):
                pairs.append((ob + i0 + d, ox + j0 + d))
            qi, qj = i0 + size, j0 + size

    if na * nb2 <= _DIRECT_MAX:
        # 작은 가운데는 바로 정렬한다 — '똑같은 행'끼리 먼저 묶으면(SequenceMatcher) 일괄 수정으로
        # 우연히 다른 행과 똑같아진 행을 먼저 짝지어 거짓 충돌이 났다(검토 실측). 상한을 넘는 것까지
        # 바로 정렬하면 100행 묶음 표가 2초 넘게 걸렸다(검토 실측) — 그 위는 닻으로.
        pair(mb, mx, 0, 0)
        return pairs
    # 큰 가운데는 **3행 이상 똑같이 이어지는 덩어리**를 닻으로 먼저 고정하고 그 사이만 정렬한다.
    #   · 통째로 정렬(띠 DP)하면 1~2만 행 표에서 10초 넘게 걸리고 메모리를 1GB 넘게 썼으며, 띠 밖으로
    #     옮긴 블록(45행 이상)은 한 칸씩 밀려 짝지어졌다(검토 실측).
    #   · 한 행짜리 일치는 닻으로 쓰지 않는다 — 우연히 같아진 행을 먼저 묶는 문제가 그것이었다.
    # 같은 내용이 많으면 정확한 정렬(autojunk=False)이 세제곱 가까이 느려지므로 큰 것은 빠른 정렬로.
    big = na > _GROUP_MAX or nb2 > _GROUP_MAX
    sm = SequenceMatcher(None, mb, mx, autojunk=big)
    pi = pj = 0
    for i0, j0, size in sm.get_matching_blocks():
        if size < 3 and size != 0:
            continue                       # 짧은 일치는 닻이 아니다 — 사이 조각에서 함께 정렬
        pair_big(mb[pi:i0], mx[pj:j0], pi, pj)
        for d in range(size):
            pairs.append((i0 + d, j0 + d))
        pi, pj = i0 + size, j0 + size
    return pairs


def _head_rows(data, kr) -> list:
    """ID 이름을 찾을 머리글 행들 — 키 행까지의 모든 행, 그리고 바로 아래 행이 숫자 없는 글자 행이면 그 행도
    (위에 한글 설명 행이 있는 2단 머리글을 기본 키 행 1행으로 열면 영문 머리글이 본문 첫 행이 된다 — 검토 실측)."""
    rows = [r for r in data[:kr + 1]]

    def text_only(i):
        vals = [v for v in data[i] if v != ""] if i < len(data) else []
        return bool(vals) and not any(_is_number(v) for v in vals)
    m = 0
    while m < 3 and text_only(kr + 1 + m):
        m += 1
    # 둘째 줄부터는 그 아래가 숫자 있는 본문일 때만(이름·설명·형식 3단 머리글 — 검토 실측). 모든 행이 글자인
    # 문자열 표는 지금처럼 한 줄만.
    if m > 1 and not (kr + 1 + m < len(data) and any(_is_number(v) for v in data[kr + 1 + m] if v != "")):
        m = 1
    return rows + list(data[kr + 1:kr + 1 + m])


def _is_number(v) -> bool:
    try:
        float(v)
        return True
    except (TypeError, ValueError):
        return False


def _is_id_name(name) -> bool:
    """머리글이 ID 열 이름인가 — …ID·…Key·…Code(대소문자 무시). 내용만 보고 ID 열을 고르면 Weight·HP 처럼
    행마다 값이 다른 열이 ID 노릇을 해, 지운 보상과 같은 확률의 새 보상이나 HP 가 같은 복사본을 같은 행으로
    짝지었다(검토 실측). 게임 데이터의 ID 열은 거의 늘 이렇게 이름 붙는다."""
    n = str(name or "").strip().lower().rstrip(".#:)] ")
    if not n:
        return False
    # Idx·Index·No·번호·코드 도 ID 이름이다(RewardIdx·보상 번호) — 묶음 안 순번(SlotNo 등)은 _seq_like 가 거른다
    return (n.endswith(("id", "key", "code", "idx", "index", "번호", "코드", "아이디"))
            or n in ("no", "num", "number") or n.endswith(("_no", " no", "number", "_num")))


def _row_ids(rows3, cols, heads=None, key_t=None):
    """세 파일의 행 식별자 — [(식별자 → (원본 행 번호, 행), 파일 순서의 식별자 목록)] × 3.

    키가 세 파일 모두에서 한 번씩만 나오면 식별자 = 키(2파일과 같은 짝짓기). 그 밖의 행은 묶음
    안에서 가운데 기준으로 내용 정렬해 짝짓는다(_align).
      · 키가 **겹치는** 행은 그 키끼리 한 묶음(보상 그룹처럼 같은 ID 가 이어지는 행).
      · 키가 **빈** 행은 (직전 구간 키, 빈 키) 묶음. 시트 전체를 한 묶음으로 잡으면 내용이 같은
        빈 키 행(구분선, 블록의 둘째 행 등)이 다른 구간 것과 짝지어졌다(검토 실측). 구간 키는
        **세 파일 모두에 한 번씩 있는 키**만 쓴다 — 한쪽이 그 행을 지우거나 키를 바꾸면 그 파일에서만
        구간이 갈려 아래 행이 모두 '지움 + 새 행'과 거짓 충돌이 됐다(검토 실측).
    내용 비교는 **세 파일 모두에 있는 열**로만 한다 — 한쪽이 값 있는 열을 더하거나 지우면 모든 행이
    달라 보여 정렬이 무너졌다(검토 실측).
    양쪽이 **똑같은 행**을 새로 넣었으면 하나로 묶는다(두 번 들어가지 않게).

    2파일은 이런 행을 비교에서 빼도 고른 칸만 덧쓰므로 대상에 그대로 남는다. 3파일 저장은
    가운데 원본에서 다시 만들기 때문에 비교에서 빠진 행의 양쪽 수정이 기준값으로 되돌아간다
    (검토 실측) — 그래서 빼지 않는다.
    """
    counts = [Counter(k for _i, _r, k in rows) for rows in rows3]
    keys = set().union(*counts)
    unique = {k for k in keys if k != "" and all(c.get(k, 0) <= 1 for c in counts)}
    anchors = {k for k in unique if all(c.get(k) for c in counts)}
    common = [t for t in cols if None not in t] or list(cols)
    # 투영한 행에서 ID 열 후보의 자리(맨 앞 표식 칸 때문에 +1).
    #   · 머리글(2단 머리글 포함) 이름이 ID 인 열 — 시트 전체 가운데 쪽 값 종류가 많은 순서로. 가장 왼쪽을
    #     고르면 GradeID 같은 분류 열이 ItemID 보다 먼저 뽑혀 같은 등급의 다른 아이템끼리 짝지었다(검토 실측).
    #   · 그런 이름이 없으면 키 열 바로 오른쪽 열을 대신 후보로(Index·No·번호 같은 ID 열) — 이름 규칙에 걸리지
    #     않는다고 ID 짝을 끄면 몬스터 표의 밸런스 패치 + 행 지움이 이웃과 엇갈렸다(검토 실측). 이 열은 _align
    #     에서 정수이고 고유할 때만 쓴다.
    # 시트 전체 행 번호를 **다시 매긴** 열 — 가운데와 그쪽 파일 모두 본문 전체가 처음부터 매긴 번호(No·번호·Index
    # 에 1..N, =ROW()-1)이고 값 목록이 달라졌으면 그쪽에서는 행 자리를 가리킬 뿐이다. 한 행만 끼워도 다시 매겨져,
    # ID 로 쓰면 아래 모든 묶음이 한 칸씩 엇갈렸다(검토 실측). 그쪽에서는 ID 후보에서 뺀다(순번인지는
    # _renumbered_cols 가 내용으로 가린다). 어느 한 파일만 보고 빼면, 1..N 으로 매긴 진짜 ID 를 다시 매기지 않고
    # 행을 지우거나 스탯만 고친 표에서 ID 짝이 꺼져 이웃 몬스터와 엇갈렸다(검토 실측). 구간 표시 행의 빈칸은 뺀다.
    def _numbered(rows, c):
        if c is None or len(rows) < 3:
            return None
        head = []
        for _i, row, _k in rows:                 # 앞쪽 값 세 개로 먼저 거른다(넓은 시트 속도)
            v = row[c] if c < len(row) else ""
            if v != "":
                head.append(v)
                if len(head) == 3:
                    break
        if len(head) < 3 or not _is_numbering(head):
            return None
        col = [row[c] if c < len(row) else "" for _i, row, _k in rows]
        col = [v for v in col if v != ""]
        if len(col) < 3 or not _is_numbering(col[:3]) or not _is_numbering(col):
            return None
        return col
    trans_cache = {}

    def _cols_of(f):
        """그 파일 본문을 공통 열 단위로 뒤집은 목록(열마다 값 튜플) — 한 번만 만든다."""
        if f not in trans_cache:
            rows = _rows_of(f)
            trans_cache[f] = list(zip(*rows)) if rows else [()] * len(common)
        return trans_cache[f]

    def _proj_index(rows, idx, drop, keep_col):
        """행 → (번호 열 drop 을 모두 뺀 나머지 값, keep_col 값). 나머지 값이 한 번씩만 나오는 행만.
        열 단위로 뒤집어 둔 값을 다시 묶어 만든다(넓은 시트에서 행마다 투영하면 수 초 걸렸다)."""
        tc = _cols_of(idx)
        keep = [tc[c] for c in range(len(common)) if c not in drop]
        keys = list(zip(*keep)) if keep else [()] * len(tc[keep_col])
        pairs = list(zip(keys, tc[keep_col]))
        cnt = Counter(keys)
        return {key: v for key, v in pairs if cnt[key] == 1}

    def _id_named(k):
        t = common[k]
        if not heads or t == key_t:
            return False
        return any(_is_id_name(h[t[f]]) for f, hrows in enumerate(heads) for h in hrows
                   if t[f] is not None and t[f] < len(h))

    proj_cache = {}

    def _rows_of(f):
        """그 파일 본문 행을 공통 열로 투영한 목록(한 번만 만든다)."""
        if f not in proj_cache:
            cols = [t[f] for t in common]
            if any(c is None for c in cols):
                proj_cache[f] = [_project(row, common, f) for _i, row, _k in rows3[f]]
            else:
                width = max(cols) + 1
                get = itemgetter(*cols) if len(cols) > 1 else (lambda r, c=cols[0]: (r[c],))
                proj_cache[f] = [get(row if len(row) >= width else list(row) + [""] * (width - len(row)))
                                 for _i, row, _k in rows3[f]]
        return proj_cache[f]

    bulk_cache = {}

    def _same_in_place(vb, vx):
        """두 열 값이 공통 길이만큼 앞쪽 또는 뒤쪽에서 그대로인가(끝에만 붙이거나 자른 경우 포함)."""
        n = min(len(vb), len(vx))
        if len(vb) == len(vx):
            return vb == vx
        return vb[:n] == vx[:n] or vb[len(vb) - n:] == vx[len(vx) - n:]

    def _bulk_cols(f):
        """가운데와 그쪽 값 묶음이 10%(그리고 3행) 넘게 다른 열 — 한쪽이 많은 행에서 고친 열."""
        if f not in bulk_cache:
            pb, px = _rows_of(1), _rows_of(f)
            n = min(len(pb), len(px))
            out = set()
            tb, tx = _cols_of(1), _cols_of(f)
            for c in range(min(len(tb), len(tx))):
                vb, vx = tb[c], tx[c]
                if _same_in_place(vb, vx):
                    continue
                differ = n - sum((Counter(vb) & Counter(vx)).values())
                if differ > max(2, 0.1 * n):
                    out.add(c)
            bulk_cache[f] = out
        return bulk_cache[f]

    def _name_like(c):
        """머리글이 Name·이름을 담은 열인가(세 파일 머리글 가운데 하나라도)."""
        t = common[c]
        if not heads:
            return False
        for f, hrows in enumerate(heads):
            for h in hrows:
                if t[f] is not None and t[f] < len(h):
                    n = str(h[t[f]] or "").strip().lower()
                    if "name" in n or "이름" in n or any(_ko_name_split(w) for w in _name_tokens(n)):
                        return True
        return False

    # '…명' 으로 끝나도 이름이 아닌 낱말
    _NOT_NAME_KO = ("설명", "생명", "수명", "투명", "조명", "증명", "운명", "분명", "발명", "변명", "해명", "규명",
                    "판명", "선명", "혁명", "사명", "서명", "표명", "익명", "숙명", "광명", "인명", "구명", "치명",
                    "공명", "청명", "자명", "결명")

    def _name_tokens(n):
        """이름 판정용 머리글 낱말 — 인원 단위 '(명)'·'[명]'·'(단위: 명)' 은 지우고, 칸 전체가 단위 '명'·'단위: 명' 이면
        낱말 없음(단위 행·'출전 인원(명)' 을 이름 열로 봤다 — 검토 실측)."""
        s = re.sub(r"[(\[（]\s*(단위\s*[:：]?\s*)?명\s*[)\]）]", " ", str(n or ""))
        s = re.sub(r"(\d)\s*명(?![가-힣])", r"\1 ", s)    # '최대 5명'·'(1~5명)' 의 인원 단위(검토 실측)
        toks = _tokens(s)
        return [] if toks in (["명"], ["단위", "명"]) else toks

    def _ko_name_split(w):
        """한글 낱말이 이름 머리글이면 [앞부분, '이름'] (앞부분 없으면 ['이름']), 아니면 None — '몬스터이름'·
        '몬스터명'·'몬스터명칭'·'명'. 한국어 시트는 '…명' 표기가 더 흔한데 이름 열로 안 봤다(검토 실측)."""
        if not w or w.isascii():
            return None
        for suf in ("이름", "명칭", "명"):
            if w.endswith(suf):
                if suf == "명" and any(w.endswith(x) for x in _NOT_NAME_KO):
                    return None
                head = w[:-len(suf)]
                return [head, "이름"] if head else ["이름"]
        return None

    _RES_WORDS = ("icon", "prefab", "model", "sprite", "effect", "file", "res", "resource", "path", "key", "text",
                  "texture", "dev", "anim", "animation", "sound", "image", "atlas", "material", "mesh", "particle",
                  "portrait", "asset", "skin", "fx", "sfx", "vfx", "bgm", "audio", "thumbnail", "thumb",
                  "아이콘", "프리팹", "모델", "스프라이트", "이펙트", "리소스", "파일", "경로", "이미지", "사운드",
                  "텍스트", "애니", "초상화", "에셋", "머티리얼", "텍스처")
    _LANG_SUFFIX = ("kr", "ko", "kor", "korean", "en", "eng", "english", "jp", "ja", "jpn", "cn", "zh", "tw", "th",
                    "de", "fr", "es", "pt", "ru", "vi", "it", "us", "gb", "hans", "hant", "chs", "cht",
                    "한글", "영문", "한국어", "영어")
    _LANG_PREFIX = ("kr", "kor", "korean", "en", "eng", "english", "jp", "jpn", "cn", "zh", "tw",
                    "한글", "영문", "한국어", "영어", "일본어", "중국어", "국문")
    _NAME_REF = ("key", "text", "id", "tid", "textid", "textkey", "string", "str", "sid", "키", "텍스트", "아이디")
    _CAT_WORDS = ("skill", "class", "element", "type", "스킬", "클래스", "속성", "종류")
    # 묶음 낱말 — StageName·MapName·GroupName 은 값이 되풀이되면(스폰 표에서 여러 행을 묶음) 표시 이름보다 아래,
    # 행마다 다르면(스테이지 표 자기 이름) 표시 이름이다 — 머리글만으로 내리면 스테이지 표의 '어려움' 복사본 끼움을
    # 복사로 봤다(검토 실측)
    _GROUP_WORDS = ("stage", "map", "group", "chapter", "wave", "area", "zone", "region",
                    "스테이지", "맵", "그룹", "챕터", "웨이브", "지역")

    def _tokens(text):
        """머리글 글을 낱말로 — CamelCase·_·공백·괄호로 나눈 소문자 낱말들(PresetName → preset, name)."""
        import re
        parts = re.split(r"[\s_\-()\[\]/.:]+", str(text or "").strip())
        out = []
        for p in parts:
            out += [w.lower() for w in re.findall(r"[A-Z]+(?![a-z])|[A-Z]?[a-z]+|\d+|[^\x00-\x7f]+", p)]
        return out

    def _has_word(text, words):
        """낱말 단위로 맞춘다(영문은 낱말이 같거나 +s, 한글은 부분 문자열) — 'Preset' 의 res, 'Result' 의 res,
        'Profile' 의 file 처럼 낱말 안에 든 글자에 걸리지 않게(검토 실측)."""
        toks = _tokens(text)
        low = str(text or "").lower()
        for w in words:
            if w.isascii():
                if any(t == w or t == w + "s" for t in toks):
                    return True
            elif w in low:
                return True
        return False

    _disp_cache = {}

    def _group_head(toks):
        """name·이름 바로 앞 머리 낱말이 묶음 낱말(Stage·Map·스테이지 …)인가 — 표시 수식어(StageDisplayName 의
        Display)는 건너뛰고, 한글은 낱말 끝으로 본다('보스스테이지'). WaveMonsterName·'웨이브몬스터이름'은 머리 낱말이
        Monster·몬스터라 아니다(검토 실측)."""
        head = [w for w in toks[:-1] if w not in ("display", "title", "show", "표시")]
        if not head:
            return False
        h = head[-1]
        if h.isascii():
            return _has_word(h, _GROUP_WORDS)
        return any(h.endswith(w) for w in _GROUP_WORDS if not w.isascii())

    def _display_name(c):
        return _display_info(c)[0]

    def _display_info(c):
        """(등급, 핵심 낱말) — 핵심 낱말은 언어 앞뒤 표기를 뗀 머리글(MonsterName_KR·MonsterName_EN → 'monster name').
        표시 이름 등급 — 머리글 행마다 따로 보고 가장 높은 등급을 쓴다. name·이름 으로 끝나는 행만 본다(설명
        행 '이름 (텍스트 키)'·형식 행 'FText' 가 Name 을 깎지 않게 — 검토 실측). 그 행에 리소스 낱말이 있으면 0
        (IconName·'아이콘 이름'), 분류 낱말만 있으면 1(SkillName·ClassName), 그 밖은 2(Name·이름·MonsterName)."""
        if c in _disp_cache:
            return _disp_cache[c]
        t = common[c]
        if not heads:
            return 0, ""
        best, best_core = 0, ""
        for f, hrows in enumerate(heads):
            for h in hrows:
                if t[f] is None or t[f] >= len(h):
                    continue
                n = str(h[t[f]] or "").strip()
                # 대소문자 구분 — 'FoodName'·'FloorName' 같은 진짜 머리글까지 형식 칸으로 버렸다(검토 실측)
                if (re.fullmatch(r"F(Name|Text|String|Vector\w*|Rotator|(Linear)?Color|GameplayTag\w*|Soft\w*|Guid"
                                 r"|DateTime|Timespan|Transform|IntPoint)", n)
                        or re.fullmatch(r"[TU][A-Z]\w*\s*<.*>", n)   # TArray<FName>·TSet<FName>(검토 실측)
                        or re.fullmatch(r"u?int\d*|float|double|bool|byte|string", n, re.I)):
                    continue                 # 형식 행(FName·FText·int32 …)은 이름 머리글이 아니다(검토 실측)
                toks = []
                for w in _name_tokens(n):
                    # 붙여 쓴 한글 언어 앞말(한글이름·영문이름)을 낱말로 나눈다
                    lp = next((p for p in _LANG_PREFIX if not p.isascii() and w.startswith(p) and len(w) > len(p)),
                              None)
                    if lp:
                        toks.append(lp)
                        w = w[len(lp):]
                    # 붙여 쓴 '스테이지이름'·'몬스터명'도 띄어 쓴 것처럼, '…명'·'…명칭'은 '이름'과 같게(검토 실측)
                    toks += _ko_name_split(w) or [w]
                while toks and toks[-1] in _LANG_SUFFIX:
                    toks = toks[:-1]         # Name_KR·이름(한글)·DisplayName_ko 의 언어 꼬리를 뗀다(검토 실측)
                while len(toks) > 1 and toks[0] in _LANG_PREFIX:
                    toks = toks[1:]          # KorName·EN_Name·'영문 이름' 의 언어 앞말도 뗀다(검토 실측)
                for i in range(len(toks) - 1):
                    if toks[i] in ("name", "이름") and all(w in _NAME_REF for w in toks[i + 1:]):
                        # NameKey·NameTextID·HeroNameTextID·MonsterNameKey — 텍스트 테이블을 가리키는 이름 열(검토 실측)
                        toks = toks[:i + 1]
                        break
                if not toks or not (toks[-1] == "name" or toks[-1].endswith("이름")):
                    continue
                core = " ".join(toks)
                if _has_word(core, _RES_WORDS):
                    g = 0
                elif (_has_word(core, _CAT_WORDS)
                      or _group_head(toks) and not _uniq90(c)):
                    # 묶음 낱말은 name 바로 앞 머리 낱말일 때만(WaveMonsterName·StageBossName 은 몬스터 이름 — 검토
                    # 실측), 값이 거의(90%) 고유하지 않을 때만 — 스테이지당 1~3행인 스폰 표의 StageName 도 묶음이다
                    # (검토 실측)
                    g = 1
                elif toks in (["name"], ["이름"]) or toks[-1] == "이름" and len(toks) == 1:
                    g = 3                    # 딱 Name·이름 — 다른 …Name 보다 먼저
                else:
                    g = 2
                if g > best:
                    best, best_core = g, core
        _disp_cache[c] = (best, best_core)
        return best, best_core

    col_stat = {}

    def _col_stat(c):
        """가운데 열 c 의 (우연 일치 확률 < 0.5, 값이 절반 넘게 다름) — 열마다 한 번만 센다(투영·번호 열마다 다시
        세어 큰 시트에서 1~2초 늘었다 — 검토 실측)."""
        if c not in col_stat:
            vals = _cols_of(1)[c]
            n = len(vals)
            cnt = Counter(vals)
            col_stat[c] = (bool(n) and sum(m * m for m in cnt.values()) * 2 < n * n, len(cnt) * 2 > n)
        return col_stat[c]

    row_fill = []

    def _uniq90(c):
        """가운데 열 c 의 값이 90% 이상 고유한가 — 묶음 열(StageName) 판정용. 장 구분 주석 행처럼 그 칸 밖에 값이
        거의 없는 행은 빼고 센다(주석 행 빈칸이 되풀이 값이 되어 스테이지 표의 StageName 을 묶음 열로 내렸다 —
        검토 실측). 데이터 행의 빈칸은 되풀이 값으로 센다 — 묶음 첫 행에만 이름을 적은 스폰 표·통째로 빈 열을
        고유로 봤다(검토 실측)."""
        key = ("u", c)
        if key not in col_stat:
            rows = _rows_of(1)
            if not row_fill:
                row_fill.append([sum(1 for v in r if v != "") for r in rows])
            vals = [r[c] for r, k in zip(rows, row_fill[0]) if k - (r[c] != "") >= 2]
            # 빈칸이 값 사이사이 여러 번 끊겨 나올 때(묶음 첫 행에만 이름)만 되풀이 값으로 센다 — 끝이나 한 장에 몰린
            # 미작성 이름(다음 장 행만 먼저 만듦)까지 세면 스테이지 표의 StageName 을 묶음 열로 내렸다(검토 실측)
            filled = sum(1 for v in vals if v != "")
            runs = sum(1 for i in range(1, len(vals)) if vals[i] == "" and vals[i - 1] != "")
            # 이름 대부분이 바로 뒤에 빈칸을 거느리면(묶음이 1~2개뿐인 표 포함) 되풀이로 센다 — 빈칸 비율로 가르면 이름을
            # 절반 넘게 아직 안 적은 스테이지 표를 묶음으로 봤다(검토 실측)
            if filled and not (runs * 2 >= filled or runs >= 3 and runs * 5 >= filled):
                vals = [v for v in vals if v != ""]
            col_stat[key] = bool(vals) and len(set(vals)) * 10 >= len(vals) * 9
        return col_stat[key]

    def _same_row_left(rb, xo, cmp_cols):
        """옛 자리 행 xo 가 원래 행 rb 와 '절반 넘게 같은가' — 마지막 행 백업 판정용.
          · 서로 다른 두 행이 우연히 같을 확률이 절반 이상인 열(모두 0 인 Flag, 0/1 Bool, 거의 한 값인 기본값 열)은
            빼고 센다 — 그런 열이 많은 표에서 마지막 행 앞에 끼운 전혀 다른 새 행도 절반 넘게 같아 보였다(검토 실측).
          · 행마다 거의 다른 열(이름·아이콘·스탯)이 둘 이상이면 그 가운데 하나는 같아야 한다 — 기본값 enum 열이 많은
            표에서 같은 회귀가 남았다(검토 실측)."""
        varying = [c for c in cmp_cols if _col_stat(c)[0]] or cmp_cols
        if 2 * sum(1 for c in varying if rb[c] == xo[c]) <= len(varying):
            return False
        hient = [c for c in varying if _col_stat(c)[1]]
        return len(hient) < 2 or any(rb[c] == xo[c] for c in hient)

    def _changed_share(c, f):
        """그 열에서 가운데와 그쪽 값 묶음이 다른 비율."""
        tb, tx = _cols_of(1), _cols_of(f)
        n = min(len(tb[c]), len(tx[c])) or 1
        return (n - sum((Counter(tb[c]) & Counter(tx[c])).values())) / n

    def _between_in_place(v, b_by_num, x_by_num, kget):
        """옛 번호 v 다음부터 가운데 끝 번호까지의 행이 그쪽에서 같은 번호에 (거의) 그대로 있는가."""
        try:
            lo = int(v)
        except (TypeError, ValueError):
            return False
        nums = sorted((int(n), n) for n in b_by_num if _is_int(n) and int(n) > lo)
        if not nums:
            return False
        kept = sum(1 for _i, n in nums if n in x_by_num and kget(x_by_num[n]) == kget(b_by_num[n]))
        return kept * 2 > len(nums)

    def _moved_number(f, k, drop):
        """번호 목록은 같아도 행이 다른 번호로 옮겨 갔는가 — 행을 끼우거나 지우거나 맞바꾸거나 정렬하고 번호를
        자리대로 다시 매긴 경우. **옮겨 간 번호가 가운데에 있던 번호**일 때만 센다 — 끝에 붙인 새 번호는 이동이
        아니다(다음 ID 로 끝에 복사한 새 행이 '옮겨 간 행'으로 보여, 진짜 ID 를 빼고 원래 행을 복사본과 짝지었다 —
        검토 실측). 두 가지로 본다.
          · 번호 열을 모두 빼면 똑같은 행(양쪽에 한 번씩)이 다른 번호에 있다.
          · ID 이름 열(RewardID 등)의 값(양쪽에 한 번씩)이 다른 번호에 붙어 있고, 그 행의 다른 칸도 함께 옮겨 갔다
            (새 번호에 원래 있던 가운데 행보다 그 값의 가운데 행과 더 비슷하다) — 옮기면서 Weight 를 함께 고쳐도
            잡되, 두 몬스터의 DropID 만 맞바꾼 것은 행 이동이 아니다(검토 실측)."""
        bset = {r[k] for r in _rows_of(1) if r[k] != ""}
        pb_all, px_all = _rows_of(1), _rows_of(f)
        b_by_num = {}
        for r in pb_all:
            if r[k] != "":
                b_by_num.setdefault(r[k], r)
        x_by_num = {}
        for r in px_all:
            if r[k] != "":
                x_by_num.setdefault(r[k], r)
        # 한쪽이 모든(거의 모든) 행에서 고친 열은 빼고도 한 번 더 본다 — 밸런스 패치(HP ×1.2)와 끼움·다시 매김을
        # 함께 하면 똑같은 행이 하나도 없었고(검토 실측), 빼기만 하면 몹 묶음을 새 몹으로 바꾼 표에서 근거가
        # 사라졌다(검토 실측). 두 투영 가운데 하나라도 옮긴 행이 있으면 된다.
        bulk = _bulk_cols(f) - drop
        tb, tx = _cols_of(1), _cols_of(f)
        if tb[k] == tx[k] and all(tb[c] == tx[c] for c in range(len(common)) if c not in drop | bulk):
            # 번호도, 일괄 수정 밖의 내용도 자리마다 그대로 — 그 투영에서는 옮긴 행이 없다. 일괄 수정 열을 넣은 투영은
            # 계속 본다(몹 20행을 빼고 다른 곳에 새 몹 20행을 넣으면 이름 열까지 '일괄 수정'으로 빠져 근거가
            # 사라졌다 — 검토 실측).
            projs = (drop,) if bulk else ()
        else:
            projs = (drop,) if not bulk else (drop | bulk, drop)
        # 거의 모든 값이 바뀐 열만 뺀 투영도 본다 — 몹 묶음 교체(이름 열이 '일괄 수정'으로 빠짐)와 모든 행 HP 패치가
        # 함께 오면 위 두 투영 모두 근거가 없었다(검토 실측)
        heavy = {c for c in bulk if _changed_share(c, f) > 0.5}
        if heavy and heavy != bulk:
            projs = projs + (drop | heavy,)
        for proj in projs:
            cmp_cols = [c for c in range(len(common)) if c not in proj]
            if tb[k] == tx[k]:
                # 번호가 자리마다 그대로면, 옮겨 간 행은 내용이 달라진 자리에만 있을 수 있다 — 그 자리만 본다
                # (넓은 시트에서 모든 행을 다시 묶으면 수 초 걸렸다)
                diff = set()
                for c in cmp_cols:
                    if tb[c] != tx[c]:
                        diff.update(i for i, (p, q) in enumerate(zip(tb[c], tx[c])) if p != q)
                if not diff:
                    continue
                if len(diff) < 2:
                    continue                 # 한 자리만 바뀌었으면 옮길 수 없다(맞바꿈도 두 자리)
            b = _proj_index(rows3[1], 1, proj, k)
            x = _proj_index(rows3[f], f, proj, k)
            if not cmp_cols:
                continue
            kget = itemgetter(*cmp_cols) if len(cmp_cols) > 1 else (lambda r, c=cmp_cols[0]: (r[c],))
            base_keys = {kget(r) for r in pb_all}
            # 이름처럼 행마다 다른 글자 열 — 있으면 '원래 행이 남아 있는가'를 이 열로 본다(열이 몇 개 없는 표에서
            # 절반 기준만 쓰면 끝에 끼운 새 몬스터를 원래 몬스터로 봐 이동을 놓쳤다 — 검토 실측)
            ident = [c for c in cmp_cols
                     if len(set(tb[c])) == len(tb[c]) and sum(1 for v in tb[c] if not _is_number(v)) * 2 > len(tb[c])]
            # 이름 계열 글자 열(머리글에 Name·이름 — IconName·PrefabName 도 포함, 고유하지 않아도 됨)을 모두 모은다
            # 표시 이름(Name·MonsterName·이름)은 값이 숫자(텍스트 테이블 TextID)여도 이름 열이다(검토 실측)
            names = [c for c in cmp_cols
                     if _name_like(c) and (sum(1 for v in tb[c] if not _is_number(v)) * 2 > len(tb[c])
                                           or _display_name(c) >= 2)]
            # 표시 이름(Name·이름·MonsterName — 겹쳐도 됨)이 있으면 그것으로, 없으면 범주형(ElementName 처럼 값이
            # 되풀이되는 열)을 뺀 이름 계열 열로
            # 표시 이름은 가장 왼쪽 하나만 — 목록에 없는 리소스 이름(PortraitName·FXName)이 표시 이름으로 섞여도 판정이
            # 흔들리지 않게(검토 실측)
            grades = {c: _display_name(c) for c in names}
            # 딱 Name·이름(언어 꼬리 뗀 Name_KR 포함)은 모두, 그 밖의 …Name 은 가장 왼쪽 하나 — 다국어 이름 열 가운데
            # 왼쪽이 아닌 열만 바꾼 복사본을 복사로 봤다(검토 실측)
            display = [c for c in names if grades[c] == 3]
            if not display:
                # 그 밖의 …Name 은 가장 왼쪽 열과 핵심 낱말이 같은 열 모두(MonsterName_KR·MonsterName_EN)
                # 값이 되풀이되는 범주형(StageName·MapName — 몇 행씩 같은 값)은 먼저 뺀다 — 왼쪽의 StageName 이
                # 뽑혀 몬스터 이름을 바꾼 복사본을 복사로 봤다(검토 실측)
                g2 = [c for c in names if grades[c] == 2]
                g2 = [c for c in g2 if len(set(tb[c])) * 2 > len(tb[c])] or g2
                display = [c for c in g2 if _display_info(c)[1] == _display_info(g2[0])[1]] if g2 else []
            if not display:
                # 표시 이름이 없으면 분류 단어만 붙은 이름 열(SkillName·ClassName) 가운데 값이 절반 넘게 다른 가장
                # 왼쪽 열 하나, 그것도 없으면 범주형을 뺀 이름 계열 열 모두
                second = [c for c in names if grades[c] == 1 and len(set(tb[c])) * 2 > len(tb[c])]
                display = second[:1]
            if not display:
                names = [c for c in names if len(set(tb[c])) * 2 > len(tb[c])]
            if display:
                names = display
            nonbulk = [c for c in cmp_cols if c not in bulk]
            base_cnt = Counter(kget(r) for r in pb_all)

            def placeholder(n):
                """가운데의 그 번호 행이 자리 표시 행(번호 밖 칸이 모두 비었거나 같은 내용이 여럿 — Dummy)인가."""
                r = b_by_num.get(n)
                if r is None:
                    return False
                return all(r[c] == "" for c in cmp_cols) or base_cnt[kget(r)] > 2
            # 끝 너머 복사본 판정용 — 가운데 번호 최댓값, 번호순 목록과 '그쪽 같은 번호에 그대로' 뒤쪽 누적 수를 한 번만
            # (복사본마다 시트 전체를 다시 훑어 2만 행에서 수십 초 걸렸다 — 검토 실측)
            ints = sorted((int(n), n) for n in b_by_num if _is_int(n))
            bmax = ints[-1][0] if ints else None
            nkeys = [i for i, _n in ints]
            suffix = [0] * (len(ints) + 1)
            for j in range(len(ints) - 1, -1, -1):
                n = ints[j][1]
                # 제자리인지는 일괄 수정 열을 빼고 본다 — 모든 행 HP 패치와 끝 백업 복사가 겹치면 모든 행이
                # '제자리 아님'이 되어 복사본을 이동으로 봤다(검토 실측)
                same_here = n in x_by_num and all(x_by_num[n][c] == b_by_num[n][c] for c in nonbulk)
                suffix[j] = suffix[j + 1] + (1 if same_here else 0)

            def between_in_place(v):
                if not _is_int(v):
                    return False
                j = bisect.bisect_right(nkeys, int(v))
                rest = len(nkeys) - j
                return rest > 0 and suffix[j] * 2 > rest
            for key, v in b.items():
                if v == "" or key not in x or x[key] in ("", v):
                    continue
                # 복사인가 이동인가 — 그쪽의 옛 번호 자리를 본다. 비었거나, 다른 가운데 행이 와 있거나, 원래 행과
                # 절반도 안 같은 행(새 행)이 와 있으면 이동이다. 원래 행이 (고쳐졌어도) 남아 있으면 복사다 — 다음
                # ID 로 끝에 복사한 새 행, 이미 있는 다른 ID 자리에 붙여 넣은 행(검토 실측). 새 번호가 가운데에
                # 없던 번호여도(맨 끝 행이 N+1 로 밀림) 센다.
                rb = b_by_num.get(v)
                xo = x_by_num.get(v)
                if rb is not None and xo is not None:
                    ko = kget(xo)
                    if ko != kget(rb) and ko in base_keys:
                        return True          # 옛 자리에 다른 가운데 행이 와 있다 — 밀리거나 맞바꿈
                    nx = x[key]
                    beyond = bmax is not None and _is_int(nx) and int(nx) > bmax
                    if beyond and _is_int(v) and int(v) == bmax and _same_row_left(rb, xo, cmp_cols):
                        # 마지막 행을 끝 너머 다음 번호로 복사했고 옛 자리 행이 원본과 절반 넘게 같으면 백업(복사)으로
                        # 본다 — '마지막 행 바로 위에 이름 바꾼 복사본을 끼움'과 파일 모양이 같아 머리글 이름으로 가르려
                        # 했지만, 머리글 표기(FName 형식 행, Name_KR, ResourceName …)마다 판정이 갈려 원본 수정이 백업에
                        # 저장됐다(검토 실측). 끝에 백업을 붙이는 쪽이 훨씬 흔하므로 이쪽으로 정한다. 옛 자리에 전혀
                        # 다른 새 행이 왔으면(마지막 행 앞에 끼우고 다시 매김) 이동이다.
                        continue
                    if (beyond or placeholder(nx)) and between_in_place(v):
                        # 끝에 붙였거나 비어 있던 예약 ID 행(빈칸·Dummy)에 넣은 복사본, 사이 행은 제자리 — 다시 매김이
                        # 아니다(원본을 고쳐도). 예약 행에 붙인 백업을 이동으로 봐 원본과 짝지었다(검토 실측).
                        continue
                    if names or ident:
                        # 원래 행이 남았는지: 이름 계열 열이 있으면 그 열이 **모두** 같을 때(IconName 처럼 앞에 있는
                        # 이름 같은 열 하나만 보면 이름만 바꾼 복사본을 놓쳤다), 없으면 고유 글자 열 가운데 하나라도
                        # 같을 때(모두 같아야 하면 원본의 Desc 한 칸만 고쳐도 이동으로 봤다) — 검토 실측.
                        if (all(rb[c] == xo[c] for c in names) if names
                                else any(rb[c] == xo[c] for c in ident)):
                            continue         # 원래 행(이름)이 남아 있다 — 복사
                    elif 2 * sum(1 for c in cmp_cols if rb[c] == xo[c]) >= len(cmp_cols):
                        continue             # 원래 행이 남아 있다 — 복사
                return True
        named = [c for c in range(len(common)) if c not in drop and _id_named(c)]
        if not named:
            return False
        pb, px = _rows_of(1), _rows_of(f)
        by_num = {}
        for r in pb:
            if r[k] != "":
                by_num.setdefault(r[k], r)
        tb, tx = _cols_of(1), _cols_of(f)
        for c in named:
            if _same_in_place(tb[c], tx[c]):
                continue                     # 값이 모두 제자리(끝에만 붙이거나 자름) — 옮겨 간 행이 없다(큰 시트 속도)
            cb = Counter(r[c] for r in pb)
            cx = Counter(r[c] for r in px)
            xv = {r[c]: r for r in px if r[c] != "" and cx[r[c]] == 1}
            for rb in pb:
                v = rb[c]
                if v == "" or cb[v] != 1 or v not in xv:
                    continue
                rx = xv[v]
                if rx[k] == rb[k] or rx[k] not in bset:
                    continue
                other = by_num.get(rx[k])
                if other is None:
                    continue
                skip = drop | {c}

                def sim(a, b2):
                    return sum(1 for n, (p, q) in enumerate(zip(a, b2)) if n not in skip and p == q)
                if sim(rb, rx) > sim(other, rx):
                    return True
        return False

    def _renums():
        """쪽별 번호 열 → (renum, renum_w). 내용으로 짝짓는 묶음이 있을 때만 계산한다(키가 모두 고유한 큰 시트에서
        시트 전체를 몇 번씩 투영해 비교가 몇 배 느려졌다 — 검토 실측).
          renum  : 그쪽 ID 후보에서 뺄 열 — 행 번호 이름(No·Index·번호 등)이거나 ID 이름이 아닌 번호 열, 그리고
                   다시 매김이 확인된 열. …ID·…Key·…Code 이름의 1..N 진짜 ID 는 확인되기 전에는 ID 로 둔다 — 빼면
                   다음 ID 로 끝에 복사한 새 행과 원래 행을 엇갈려 짝지었다(검토 실측).
          renum_w: 정렬 점수·확실한 짝 근거에서도 뺄 열 — **행이 다른 번호로 옮겨 간** 것이 확인된 열. '목록이
                   달라짐'은 근거가 아니다 — 끝에 붙이거나 끝을 잘라도 목록이 달라져, 1..N 진짜 ID 가 점수에서 빠지고
                   밸런스 패치와 겹쳐 이웃 몬스터끼리 엇갈렸다(검토 실측). 가운데에 끼우고 다시 매긴 경우는 옮겨 간
                   행으로 잡힌다."""
        sides = ({}, None, {})
        for k, t in enumerate(common):
            if t == key_t:
                continue
            bcol = _numbered(rows3[1], t[1])
            if bcol is None:
                continue
            for f in (0, 2):
                if _numbered(rows3[f], t[f]) is not None:
                    sides[f][k] = True
        renum = ({}, None, {})
        renum_w = ({}, None, {})
        for f in (0, 2):
            drop = set(sides[f])
            for k in sides[f]:
                moved = _moved_number(f, k, drop)
                if moved:
                    renum_w[f][k + 1] = True
                if moved or not _strong_id_name(k):
                    renum[f][k + 1] = True
        return renum, renum_w

    def _strong_id_name(k):
        """…ID·…Key·…Code·코드·아이디 이름 — No·Num·Index·Idx·번호 처럼 행 번호일 수도 있는 이름은 아니다."""
        t = common[k]
        if not heads or t == key_t:
            return False
        for f, hrows in enumerate(heads):
            for h in hrows:
                if t[f] is not None and t[f] < len(h):
                    n = str(h[t[f]] or "").strip().lower().rstrip(".#:)] ")
                    if n.endswith(("id", "key", "code", "코드", "아이디")):
                        return True
        return False
    id_cols = []
    if heads:
        named = []
        for k, t in enumerate(common):
            if t == key_t:
                continue                     # 키 열은 묶음 안에서 늘 같은 값(또는 빈 값)이라 ID 가 될 수 없다
            names = [h[t[f]] if (t[f] is not None and t[f] < len(h)) else ""
                     for f, rows in enumerate(heads) for h in rows]
            if any(_is_id_name(nm) for nm in names):
                bcol = t[1]
                kinds = len({row[bcol] for _i, row, _k in rows3[1] if bcol is not None and bcol < len(row)})
                named.append((-kinds, k + 1))
        id_cols = [c for _n, c in sorted(named)]
    fallback = None
    if not id_cols and key_t is not None and key_t in common:
        ki = common.index(key_t)
        if ki + 1 < len(common):
            # 대신 쓰는 열은 가운데 파일 **시트 전체**에서 고유해야 한다 — 묶음 안에서만 고유하면 확률(Weight)
            # 같은 열이 ID 노릇을 해, 확률을 맞바꾸거나 같은 확률의 새 보상으로 바꾸면 다른 보상끼리 짝지었다
            # (검토 실측). _align 에서 정수이고 오름차순인지도 본다.
            bcol = common[ki + 1][1]
            vals = [row[bcol] if bcol is not None and bcol < len(row) else "" for _i, row, _k in rows3[1]]
            vals = [v for v in vals if v != ""]      # 구간 표시 행의 빈칸은 세지 않는다
            if vals and len(set(vals)) == len(vals):
                fallback = ki + 2

    groups = [{}, {}, {}]              # 파일별 묶음 → [(원본 행 번호, 행)]
    orders = [[], [], []]              # 파일별 (묶음, 몇 번째)
    near = [{}, {}, {}]                # 파일별 (묶음, 몇 번째) → 그 파일에서 바로 앞의 고유 키
    tags = [{}, {}, {}]                # 파일별 (묶음, 몇 번째) → 그 행의 구간(세 파일 공통 고유 키)
    for f, rows in enumerate(rows3):
        anchor = None
        last_unique = None
        for i, row, k in rows:
            if k in unique:
                g = k
                last_unique = k
                if k in anchors:
                    anchor = k
            else:
                # 빈 키는 구간별로 묶는다(내용이 같은 빈 키 행이 다른 구간 것과 짝지어졌다 — 검토 실측).
                # 겹치는 키는 **키끼리** 묶고 구간은 표식으로만 쓴다 — 구간으로 묶으면 고유 키 행 하나만
                # 옮겨도 그 뒤 묶음이 통째로 '지움+새 행'으로 갈렸다(검토 실측). 표식이 같은 행끼리
                # 먼저 같다고 보므로 떨어진 곳에 되풀이되는 행('#' 구분 행)도 제 구간끼리 짝지어진다.
                g = ("§", anchor if k == "" else None, k)
            lst = groups[f].setdefault(g, [])
            orders[f].append((g, len(lst)))
            near[f][(g, len(lst))] = last_unique
            tags[f][(g, len(lst))] = anchor
            lst.append((i, row))

    ids = [{}, {}, {}]                 # 파일별 (묶음, 몇 번째) → 식별자
    sheet_unique = set()
    for c in id_cols:
        bc = common[c - 1][1]
        vals = [row[bc] for _i, row, _k in rows3[1] if bc is not None and bc < len(row) and row[bc] != ""]
        if vals and len(set(vals)) == len(vals):
            sheet_unique.add(c)
    renums = None
    for g in set(groups[0]) | set(groups[1]) | set(groups[2]):
        L, B, R = (groups[f].get(g, []) for f in range(3))
        if not isinstance(g, tuple):
            for f, rows in enumerate((L, B, R)):
                if rows:
                    ids[f][(g, 0)] = g
            continue
        if renums is None:
            renums = _renums()
        renum, renum_w = renums
        bp = [(tags[1][(g, n)],) + _project(row, common, 1) for n, (_i, row) in enumerate(B)]
        lp = [(tags[0][(g, n)],) + _project(row, common, 0) for n, (_i, row) in enumerate(L)]
        rp = [(tags[2][(g, n)],) + _project(row, common, 2) for n, (_i, row) in enumerate(R)]
        lid = _align(bp, lp, g, "l", id_cols, fallback, set(renum[0]), set(renum_w[0]), sheet_unique)
        rid = _align(bp, rp, g, "r", id_cols, fallback, set(renum[2]), set(renum_w[2]), sheet_unique)
        # 양쪽이 똑같이 새로 넣은 행 — 오른쪽 것을 왼쪽 것의 식별자로. **같은 자리**(그 파일에서
        # 바로 앞의 고유 키가 같음)일 때만 — 양쪽이 서로 다른 새 부모 행 밑에 같은 내용의 빈 키
        # 자식을 달았을 때 둘이 하나로 합쳐져 한 부모의 자식이 사라졌다(검토 실측).
        spare = {}
        for j, x in enumerate(lid):
            if x[1] == "l":
                spare.setdefault((lp[j], near[0].get((g, j))), []).append(x)
        for j, x in enumerate(rid):
            key = (rp[j], near[2].get((g, j)))
            if x[1] == "r" and spare.get(key):
                rid[j] = spare[key].pop(0)
        for n in range(len(B)):
            ids[1][(g, n)] = (g, "b", n)
        for n, x in enumerate(lid):
            ids[0][(g, n)] = x
        for n, x in enumerate(rid):
            ids[2][(g, n)] = x
    out = []
    for f in range(3):
        m = {ids[f][gn]: groups[f][gn[0]][gn[1]] for gn in orders[f]}
        out.append((m, [ids[f][gn] for gn in orders[f]]))
    return out


def count_dropped3(left: list, base: list, right: list, col_meta3, key_col: int,
                   key_row: int) -> int:
    """키가 비었거나 중복인 행 수(세 파일 합). 이 행들은 빠지지 않고 **내용으로** 짝지어진다 —
    짝이 틀릴 수 있어 알리고, 자동 병합에서도 뺀다(compute_diff3 의 ambiguous)."""
    if key_col == -1:
        return 0
    cm = usable_col_meta3(col_meta3, key_col)
    kr = key_row if (key_row and key_row > 0) else 0
    total = 0
    for i, data in enumerate((left, base, right)):
        kc = cm[key_col][i] if cm else key_col
        seen = set()
        for row in data[kr + 1:]:
            k = row[kc] if (kc is not None and kc < len(row)) else ""
            if k == "":
                if any(v != "" for v in row):
                    total += 1
                continue
            if k in seen:
                total += 1
            seen.add(k)
    return total


def perspective(matrix: list, row_meta3: list, col_meta3: list | None, perm: tuple):
    """매트릭스를 **다른 파일을 가운데(쓰는 곳)에 둔 모양**으로 비춘다 — (matrix, row_meta3, col_meta3).

    perm = (왼쪽 자리에 올 파일, 가운데 자리에 올 파일, 오른쪽 자리에 올 파일) — 원래 자리 번호
    (0 왼쪽 · 1 가운데 · 2 오른쪽). 행·열 짝은 그대로 두고(화면 좌표 (r, c) 가 같다) 값 자리와 상태만
    새 가운데 기준으로 다시 매긴다. 상태 규칙은 _make_row 와 같다 — 값은 classify, 새 가운데에 있는
    행을 한쪽이 지우고 다른 쪽이 고쳤으면 행 단위 충돌.

    3파일 화면이 가운데가 아닌 파일에 저장할 때 쓴다: 병합 준비·저장 규칙(행 삭제·새 행·한쪽에만
    있는 열)은 모두 '가운데에 쓴다'로 짜여 있으므로, 대상 파일을 가운데에 둔 모양을 만들어 같은
    규칙을 그대로 태운다. perm 이 (0, 1, 2) 면 원래 매트릭스와 같다.
    """
    a, t, z = perm
    cols = [(cm[a], cm[t], cm[z]) for cm in col_meta3] if col_meta3 else None
    metas = [(m[a], m[t], m[z]) for m in row_meta3]
    va, vt, vz = a + 1, t + 1, z + 1
    out = []
    for row, m in zip(matrix, metas):
        if all(cell[0] == S3_SAME for cell in row):
            out.append(row)                         # 세 값이 같은 칸은 어느 모양에서나 같다
            continue
        cells = [(classify(c[va], c[vt], c[vz]), c[va], c[vt], c[vz]) for c in row]
        if m[1] is not None and ((m[0] is None) != (m[2] is None)):
            ki = 2 if m[0] is None else 0           # 지우지 않은 쪽
            edited = any(cell[ki + 1] != cell[2] for i, cell in enumerate(cells)
                         if cols is None or (cols[i][1] is not None and cols[i][ki] is not None))
            if edited:
                cells = [((S3_CONFLICT if st != S3_SAME else st), l, b, r)
                         for st, l, b, r in cells]
        out.append(cells)
    return out, metas, cols


def conflict_masks(matrix: list) -> list:
    """행별 '충돌 열' 비트마스크 — row_change_masks 와 같은 모양(다음/이전 충돌, 충돌 수)."""
    out = []
    for row in matrix:
        m = 0
        for c, cell in enumerate(row):
            if cell[0] == S3_CONFLICT:
                m |= 1 << c
        out.append(m)
    return out


def auto_merge_plan(matrix: list, row_meta3: list | None = None, skip_rows=()) -> dict:
    """충돌 없는 변경을 가운데로 가져올 계획 {(r, c): SRC_LEFT|SRC_RIGHT} — WinMerge 의 Auto Merge.

    왼쪽만 바뀐 칸은 왼쪽에서, 오른쪽만 바뀐 칸은 오른쪽에서, 양쪽이 같게 바뀐 칸은 (어느
    쪽이든 같은 값이라) 왼쪽에서 가져온다. 충돌은 사람이 고른다.

    row_meta3 를 주면, **가운데에 없는 행**(양쪽이 새로 만든 행)에 충돌이 하나라도 있으면 그
    행은 통째로 뺀다. 일부 칸(보통 같은 키)만 가져가면 충돌을 열어 둔 채 저장했을 때 키만 있는
    빈 행이 생긴다(검토 실측) — 기존 행이면 열린 충돌은 기준값으로 남지만 새 행엔 기준값이 없다.
    """
    plan = {}
    for r, row in enumerate(matrix):
        if r in skip_rows:              # 내용으로 짝지은 행 — 짝이 틀릴 수 있어 사람이 고른다
            continue
        if (row_meta3 is not None and r < len(row_meta3) and row_meta3[r][1] is None
                and any(cell[0] == S3_CONFLICT for cell in row)):
            continue
        for c, (st, *_v) in enumerate(row):
            if st in (S3_LEFT, S3_BOTH):
                plan[(r, c)] = SRC_LEFT
            elif st == S3_RIGHT:
                plan[(r, c)] = SRC_RIGHT
    return plan
