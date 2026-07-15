# -*- coding: utf-8 -*-
"""진단 로깅 — 로거 설정 + 삼킴 경로가 실제로 로그를 남기는지."""
import logging
import struct

from excelmerge.logutil import get_logger, log_path


def _capture():
    """excelmerge 로거(propagate=False)에 임시 핸들러를 달아 레코드를 수집."""
    logger = get_logger()
    records = []

    class _H(logging.Handler):
        def emit(self, r):
            records.append(r)

    h = _H()
    logger.addHandler(h)
    return logger, h, records


def test_logger_configured():
    logger = get_logger()
    assert logger.name == "excelmerge"
    assert logger.handlers, "핸들러 미설정"
    assert not logger.propagate           # 루트 전파 안 함
    assert log_path().endswith("excelmerge.log")


def test_uasset_fallback_logs(tmp_path):
    from excelmerge.uasset_parser import load_uasset_as_matrix, _UASSET_MAGIC
    logger, h, records = _capture()
    try:
        p = tmp_path / "garbage.uasset"
        p.write_bytes(struct.pack("<I", _UASSET_MAGIC) + b"\x00" * 200)
        load_uasset_as_matrix(str(p))     # 본격 파싱 실패 → 폴백 + log.warning
        assert any("uasset" in r.getMessage() for r in records), \
            "uasset 폴백 경로가 로그를 남기지 않음"
    finally:
        logger.removeHandler(h)


def test_calamine_fallback_logs(tmp_path, monkeypatch):
    import openpyxl
    from excelmerge import loaders
    p = tmp_path / "v.xlsx"
    wb = openpyxl.Workbook(); wb.active.append(["ID", "V"]); wb.save(str(p))
    loaders.clear_values_cache()
    monkeypatch.setattr(loaders, "_get_calamine", lambda: None)   # calamine 강제 실패
    logger, h, records = _capture()
    try:
        loaders.load_values_any(str(p), None, "Sheet")
        assert any("calamine" in r.getMessage() for r in records), \
            "calamine 폴백 경로가 로그를 남기지 않음"
    finally:
        logger.removeHandler(h)
