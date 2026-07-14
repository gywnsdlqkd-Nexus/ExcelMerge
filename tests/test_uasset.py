# -*- coding: utf-8 -*-
"""uasset_parser 견고성 테스트.

663줄의 바이너리 파서로 커버리지가 0이었다. 유효한 UE5 DataTable 픽스처를 합성하긴
어려우므로, 여기서는 파서의 핵심 계약 — **어떤 입력에도 절대 예외를 던지지 않고
[["field","value"], ...] 폴백 매트릭스를 반환한다** — 를 검증한다.
"""
import struct

from excelmerge.uasset_parser import load_uasset_as_matrix, _UASSET_MAGIC


def _is_matrix(m):
    return isinstance(m, list) and all(isinstance(r, list) for r in m)


def test_too_short_file(tmp_path):
    p = tmp_path / "tiny.uasset"
    p.write_bytes(b"\x00\x01")
    m = load_uasset_as_matrix(str(p))
    assert _is_matrix(m) and m[0] == ["field", "value"]
    assert any("너무 짧" in "".join(row) for row in m), m


def test_non_ue_magic_falls_back(tmp_path):
    p = tmp_path / "notue.uasset"
    p.write_bytes(b"XXXX" + b"\x00" * 64)          # 잘못된 magic
    m = load_uasset_as_matrix(str(p))
    assert _is_matrix(m) and m[0] == ["field", "value"]
    assert any("시그니처" in "".join(row) for row in m), m


def test_valid_magic_garbage_body_no_crash(tmp_path):
    """UE magic 은 맞지만 본문이 쓰레기 → 본격 파싱 실패해도 폴백으로 안전 반환."""
    p = tmp_path / "magic.uasset"
    p.write_bytes(struct.pack("<I", _UASSET_MAGIC) + b"\x00" * 256)
    m = load_uasset_as_matrix(str(p))
    assert _is_matrix(m) and m[0] == ["field", "value"]
    # magic 이 매트릭스에 기록돼야 한다
    assert any(row and row[0] == "magic" for row in m), m
