# -*- coding: utf-8 -*-
"""열 번호 ↔ 열 문자 변환 — openpyxl 을 시작 경로에서 떼어 내기 위한 지역 구현.

**왜 여기 있나.** `from openpyxl.utils import get_column_letter` 한 줄이 openpyxl 패키지
전체를 끌어온다(실측: 모듈 430개, 1.14초 — numpy·chart·drawing·workbook·worksheet 까지).
우리가 쓰는 건 이 변환 함수 둘뿐인데, 그 둘 때문에 프로그램이 켜질 때마다 1초 넘게
기다리게 했다. 네 모듈(diff_model·diff_view·widgets·xlsx_writer)이 그랬다.

실제 파일을 읽을 때는 openpyxl 이 여전히 필요하지만(calamine 실패 시 폴백), 그건
**파일을 열 때** 필요하지 **켤 때** 필요한 게 아니다 — loaders 에서 함수 안으로 옮겼다.

**동작은 openpyxl 과 같다.** 좌표는 저장할 셀 자리를 정하는 데 쓰이므로, 다르게
동작하면 엉뚱한 칸에 쓰게 된다. 범위·오류까지 맞춰 두고 tests/test_colref.py 에서
openpyxl 과 전수 대조한다(1~18278 전부).
"""

_ALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
_MAX_COL = 18278          # 'ZZZ' — openpyxl 과 같은 상한
_TO_IDX = {c: i + 1 for i, c in enumerate(_ALPHA)}


def get_column_letter(col_idx: int) -> str:
    """1-based 열 번호 → 열 문자. 1 ≤ col_idx ≤ 18278('ZZZ') 밖이면 ValueError.

    'A'=1, 'Z'=26, 'AA'=27, 'ZZ'=702.
    """
    if not 1 <= col_idx <= _MAX_COL:
        raise ValueError(f"Invalid column index {col_idx}")
    letters = []
    while col_idx > 0:
        col_idx, rem = divmod(col_idx - 1, 26)
        letters.append(_ALPHA[rem])
    return "".join(reversed(letters))


def column_index_from_string(col: str) -> int:
    """열 문자 → 1-based 열 번호. 대소문자를 가리지 않는다.

    'A'~'ZZZ' 밖이거나 글자가 아니면 ValueError — openpyxl 과 같은 조건이다.
    """
    error = (f"'{col}' is not a valid column name. "
             f"Column names are from A to ZZZ")
    if not col or len(col) > 3:
        raise ValueError(error)
    idx = 0
    for ch in col.upper():
        pos = _TO_IDX.get(ch)
        if pos is None:
            raise ValueError(error)
        idx = idx * 26 + pos
    if not 0 < idx <= _MAX_COL:
        raise ValueError(error)
    return idx
