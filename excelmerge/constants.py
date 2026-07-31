"""공용 문자열 상수 + 앱 데이터 경로 — 흩어져 있던 값들의 단일 출처.

값은 기존 리터럴과 100% 동일하게 유지한다(diff_matrix 직렬화·기존 비교·테스트 호환).
오타(예: "a_to_b" vs "a2b")로 인한 불일치 버그를 막고 의미를 명시하기 위한 상수화.
"""
import os

# 앱 데이터 디렉터리 이름 — %APPDATA%/<이 이름>/ 아래에 로그·설정을 둔다.
APP_DIR_NAME = "ExcelMerge"


def appdata_path(*parts: str) -> str:
    """%APPDATA%/ExcelMerge/<parts...> 경로. APPDATA 없으면 홈 디렉터리 기준.

    crashlog/logutil/prefs/updater 가 각자 동일 구현을 갖고 있던 것을 통합했다.
    호출 시점에 환경변수를 읽으므로 테스트가 APPDATA 를 임시 폴더로 바꿔치기하는
    방식(_with_tmp_appdata)이 그대로 동작한다 — import 시점에 고정하면 안 된다.
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, APP_DIR_NAME, *parts)


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
