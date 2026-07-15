"""공용 문자열 상수 — 여러 모듈에 흩어져 있던 매직 스트링의 단일 출처.

값은 기존 리터럴과 100% 동일하게 유지한다(diff_matrix 직렬화·기존 비교·테스트 호환).
오타(예: "a_to_b" vs "a2b")로 인한 불일치 버그를 막고 의미를 명시하기 위한 상수화.
"""

# diff 셀 상태 — compute_diff가 diff_matrix 각 셀 튜플의 첫 요소로 채운다.
STATUS_SAME = "same"
STATUS_ADDED = "added"
STATUS_MODIFIED = "modified"

# 병합 방향 — staged 값 / 저장·복사 방향.
DIR_A2B = "a_to_b"
DIR_B2A = "b_to_a"

# DiffTableModel 표시 모드.
MODE_EMPTY = "empty"
MODE_DIFF = "diff"
MODE_PREVIEW = "preview"
