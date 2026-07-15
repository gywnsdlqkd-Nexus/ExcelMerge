"""진단 로깅 — 광범위 except 로 '조용히 삼킨' 오류를 파일에 남겨 사후 진단을 돕는다.

기록 위치: %APPDATA%/ExcelMerge/excelmerge.log (RotatingFileHandler, 1MB × 백업 2).
로거 설정/기록이 실패해도 앱 동작에 절대 영향을 주지 않는다(NullHandler 폴백).
graceful degradation(폴백) 경로에서 `log.debug(..., exc_info=True)` 로 남기는 것이 목적이며,
사용자에게 보이는 흐름은 그대로 둔다(로그는 진단용일 뿐 동작을 바꾸지 않는다).
"""
import logging
import logging.handlers
import os

_LOGGER_NAME = "excelmerge"
_configured = False


def log_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ExcelMerge", "excelmerge.log")


def get_logger() -> logging.Logger:
    """패키지 공용 로거(1회 설정). import 시 파일 핸들러를 붙인다."""
    global _configured
    logger = logging.getLogger(_LOGGER_NAME)
    if _configured:
        return logger
    _configured = True
    logger.setLevel(logging.DEBUG)
    logger.propagate = False   # 루트로 전파 안 함(중복/콘솔 오염 방지)
    try:
        p = log_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        h = logging.handlers.RotatingFileHandler(
            p, maxBytes=1024 * 1024, backupCount=2, encoding="utf-8")
        h.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)s %(name)s: %(message)s"))
        logger.addHandler(h)
    except Exception:
        logger.addHandler(logging.NullHandler())
    return logger


log = get_logger()
