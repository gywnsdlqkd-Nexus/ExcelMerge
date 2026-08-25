"""병합 준비(스테이징) 판단 로직 — Qt 비의존 순수 함수 모음.

**왜 분리했나.** 이 판단들은 원래 DiffView(1300여 줄, 68 메서드) 안에서 UI 효과
(QMessageBox·셀값란 위젯·상태바·패널 재채색)와 한 덩어리로 얽혀 있었다. 그래서
"어떤 셀이 병합 대상인가" 같은 **순수 규칙조차 QApplication + MainWindow 를 띄워야만**
검증할 수 있었다. 규칙만 여기로 빼내 단위 테스트가 가능해졌고, DiffView 는 규칙 호출 +
UI 효과 적용만 남는다.

**상태 소유권은 옮기지 않았다(의도).** `_staged` dict 는 DiffView·DiffTableModel·FilePanel
이 **참조로 공유**한다(같은 객체를 보고 있어야 재채색이 일관된다). 소유권을 이 모듈로
옮기면 공유 참조가 조용히 갈라질 위험이 있어, 여기서는 상태를 인자로 받아 계산만 한다.
"""
from .constants import STATUS_SAME, DIR_A2B


def stageable_cells(diff_matrix, cells) -> set:
    """cells 중 실제로 병합 준비할 수 있는 셀만 남긴다.

    탈락 기준: 격자 범위 밖 / 변경 없음(same). 그것뿐이다.
    키 보충으로 들어온 셀 중 '매칭 행의 동일한 키 셀'은 same 이라 여기서 자동 탈락하고,
    신규 행의 키 셀만 남는다.

    ★ '변경 검사에서 제외'한 열은 **탈락시키지 않는다.** 제외는 차이를 표시·집계하지
    않겠다는 뜻이고, 병합 준비는 사용자의 명시적 지시다. 예전엔 제외 열을 버려서, B 에만
    있는 행을 A→B 로 지울 때 제외 열(en/ja/zh_Hans 등)의 텍스트가 B 에 그대로 남았다.
    (diff_matrix 는 제외와 무관하게 실제 status 를 담고 있다 — compute_diff 는 제외를
    아예 모른다. 그래서 이 필터만 제거하면 실제 차이가 그대로 병합 대상이 된다.)
    """
    if not diff_matrix:
        return set()
    out = set()
    for (r, c) in cells:
        if not (0 <= r < len(diff_matrix)):
            continue
        row = diff_matrix[r]
        if not (0 <= c < len(row)):
            continue
        if row[c][0] == STATUS_SAME:
            continue
        out.add((r, c))
    return out


def staged_display_value(diff_matrix, r: int, c: int, direction: str) -> str:
    """staged 셀의 셀값란 표시값(= 병합될 값).

    값으로 병합하므로 소스 side 의 계산값을 양쪽 셀값란에 **동일하게** 표시한다.
    격자 범위를 벗어나거나 셀 구조가 예상과 다르면 빈 문자열.
    """
    try:
        _status, a_val, b_val = diff_matrix[r][c]
    except (IndexError, TypeError, ValueError):
        return ""
    return a_val if direction == DIR_A2B else b_val


def excludable_cols(cols, key_col, excluded_cols) -> list:
    """제외 '추가' 대상으로 유효한 열만 — 키 열은 제외 불가, 이미 제외된 열은 중복 제외 안 함."""
    return [c for c in cols if c != key_col and c not in excluded_cols]
