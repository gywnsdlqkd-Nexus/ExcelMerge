# -*- coding: utf-8 -*-
"""DataTable 의 소프트 경로(SoftObjectProperty)를 패키지 버전별 형식대로 읽는다.

UE5 1008(ADD_SOFTOBJECTPATH_LIST)부터 소프트 경로는 패키지 헤더의 테이블에 모이고,
프로퍼티 값에는 그 테이블의 **int32 인덱스 4바이트**만 남는다. 파서가 이걸 예전처럼
FName(8바이트)으로 읽어, 이름 테이블의 엉뚱한 경로에 **다음 프로퍼티의 이름 번호**가
`_N` 접미사로 붙어 나왔다. 실제 제보: DT_ContentsList 의 ContentsBGM 전 행이 틀렸고,
두 리비전의 이름 수가 하나 달라 접미사가 `_83`/`_82` 로 갈려 **모든 행이 변경**으로 떴다.

저장소가 공개라 실제 게임 DT 는 못 쓴다 — 아래 빌더로 바이트를 직접 짠다.
"""
import struct

import pytest

from excelmerge import uasset_parser as P
from excelmerge.uasset_parser import load_uasset_as_matrix


# ── 합성 DT 빌더 ────────────────────────────────────────────────────────────

def _fstr(s: str) -> bytes:
    """FString — 빈 문자열은 길이 0 만 쓴다(언리얼과 같다)."""
    if not s:
        return struct.pack("<i", 0)
    b = s.encode("ascii") + b"\0"
    return struct.pack("<i", len(b)) + b


class _DT:
    """UE5 DataTable 한 장을 바이트로 짠다 — 파서가 밟는 구조만 정확히.

    soft_mode
      "index"  값 = 헤더 테이블 인덱스 (UE5 1008+, 실제 빌드 DT 의 형식)
      "top"    값 = 패키지 FName + 에셋 FName + SubPath (UE5 1007)
      "legacy" 값 = 경로 FName + SubPath (UE5 1006 이하)
    """

    def __init__(self, soft_mode="index", extra_names=()):
        self.mode = soft_mode
        self.ue5 = {"index": 1009, "top": 1007, "legacy": 1006}[soft_mode]
        self.names = ["None", "IntProperty", "SoftObjectProperty", "ArrayProperty",
                      *extra_names]
        self.soft: list[tuple[str, str, str]] = []
        self.rows: list[tuple[str, list[bytes]]] = []

    def n(self, s: str) -> int:
        if s not in self.names:
            self.names.append(s)
        return self.names.index(s)

    def fname(self, s: str) -> bytes:
        return struct.pack("<ii", self.n(s), 0)

    def _tag(self, name, typ, value: bytes, extra=b"") -> bytes:
        return (self.fname(name) + self.fname(typ)
                + struct.pack("<ii", len(value), 0) + extra + b"\0" + value)

    def _soft_value(self, pkg: str, asset: str, sub: str = "") -> bytes:
        if self.mode == "index":
            key = (pkg, asset, sub)
            if key not in self.soft:
                self.soft.append(key)
            return struct.pack("<i", self.soft.index(key))
        if self.mode == "top":
            return self.fname(pkg) + self.fname(asset) + _fstr(sub)
        path = pkg if pkg == "None" else f"{pkg}.{asset}"
        return self.fname(path) + _fstr(sub)

    def int_(self, name, v):
        return self._tag(name, "IntProperty", struct.pack("<i", v))

    def soft_(self, name, pkg, asset="None", sub=""):
        return self._tag(name, "SoftObjectProperty", self._soft_value(pkg, asset, sub))

    def soft_array(self, name, paths):
        body = struct.pack("<i", len(paths)) + b"".join(
            self._soft_value(p, a) for p, a in paths)
        return self._tag(name, "ArrayProperty", body,
                         extra=self.fname("SoftObjectProperty"))

    def row(self, key, *tags):
        self.n(key)
        self.rows.append((key, list(tags)))

    def build(self) -> bytes:
        body = struct.pack("<i", len(self.rows))
        for key, tags in self.rows:
            body += self.fname(key) + b"".join(tags) + self.fname("None")
        # 이름은 몸통·소프트 테이블을 짜면서 늘어나므로, 그 둘을 먼저 짠다
        soft_tbl = b"".join(self.fname(p) + self.fname(a) + _fstr(s)
                            for p, a, s in self.soft)
        name_tbl = b"".join(_fstr(s) + b"\0\0\0\0" for s in self.names)

        def head(total, name_off, soft_off):
            h = struct.pack("<Iiiiii", P._UASSET_MAGIC, -8, 864, 522, self.ue5, 0)
            h += struct.pack("<ii", 0, total) + _fstr("None")
            h += struct.pack("<iii", 0, len(self.names), name_off)
            if self.ue5 >= P.UE5_SOFTPATH_LIST:
                h += struct.pack("<ii", len(self.soft), soft_off)
            return h + b"\0" * 64

        name_off = len(head(0, 0, 0))
        soft_off = name_off + len(name_tbl)
        total = soft_off + len(soft_tbl)
        out = head(total, name_off, soft_off) + name_tbl + soft_tbl
        assert len(out) == total
        return out + b"\0\0\0\0" + body

    def save(self, tmp_path, name="DT_T.uasset") -> str:
        p = tmp_path / name
        p.write_bytes(self.build())
        return str(p)


def _col(path, col):
    m = load_uasset_as_matrix(path)
    assert m[0][:2] != ["field", "value"], f"본격 파싱 실패 — 폴백 덤프: {m[:6]}"
    i = m[0].index(col)
    return {r[0]: r[i] for r in m[1:]}


def _contents(dt: _DT) -> _DT:
    """제보 파일과 같은 모양 — 소프트 경로 열 바로 뒤에 다른 열이 온다."""
    dt.row("Title", dt.int_("Id", 1),
           dt.soft_("Bgm", "/Game/Audio/Seq/Intro_Cue", "Intro_Cue"),
           dt.int_("Flag", 1))
    dt.row("Prologue", dt.int_("Id", 2),
           dt.soft_("Bgm", "/Game/Audio/BGM/Theme_A", "Theme_A"),
           dt.int_("Flag", 1))
    dt.row("Lobby", dt.int_("Id", 3), dt.soft_("Bgm", "None"), dt.int_("Flag", 0))
    return dt


EXPECTED = {
    "Title": "/Game/Audio/Seq/Intro_Cue.Intro_Cue",
    "Prologue": "/Game/Audio/BGM/Theme_A.Theme_A",
    "Lobby": "None",
}


# ── UE5 1008+ : 헤더 테이블 인덱스 ──────────────────────────────────────────

def test_index_values_resolve_through_the_header_table(tmp_path):
    """제보 그대로 — 고치기 전에는 이름 테이블의 엉뚱한 이름 + `_N` 이 나왔다."""
    assert _col(_contents(_DT()).save(tmp_path), "Bgm") == EXPECTED


def test_the_next_column_is_untouched(tmp_path):
    """값을 8바이트로 읽으면 다음 프로퍼티의 이름 인덱스를 먹는다 — 이웃 열이 멀쩡해야 한다."""
    p = _contents(_DT()).save(tmp_path)
    assert _col(p, "Flag") == {"Title": "1", "Prologue": "1", "Lobby": "0"}
    assert _col(p, "Id") == {"Title": "1", "Prologue": "2", "Lobby": "3"}


def test_a_different_name_count_does_not_fake_a_change(tmp_path):
    """두 리비전의 이름 수가 달라도 같은 경로는 같은 값이어야 한다.

    고치기 전에는 접미사가 다음 열의 이름 번호라, 이름이 하나만 늘어도 **모든 행**이
    변경으로 떴다(실측 `_83` 대 `_82`).
    """
    a = _contents(_DT()).save(tmp_path, "A.uasset")
    b = _contents(_DT(extra_names=("Unrelated_A", "Unrelated_B"))).save(tmp_path, "B.uasset")
    assert _col(a, "Bgm") == _col(b, "Bgm") == EXPECTED


def test_subpath_is_kept(tmp_path):
    dt = _DT()
    dt.row("R1", dt.soft_("Bgm", "/Game/Audio/Bank", "Bank", "Track.Inner"))
    assert _col(dt.save(tmp_path), "Bgm") == {"R1": "/Game/Audio/Bank.Bank:Track.Inner"}


def test_soft_paths_inside_an_array(tmp_path):
    dt = _DT()
    dt.row("R1", dt.soft_array("Bgms", [("/Game/Audio/BGM/Theme_A", "Theme_A"),
                                        ("None", "None")]))
    assert _col(dt.save(tmp_path), "Bgms") == {
        "R1": "[/Game/Audio/BGM/Theme_A.Theme_A, None]"}


def test_an_unreadable_table_shows_a_placeholder_not_a_wrong_path(tmp_path):
    """테이블 위치가 깨졌으면 '그럴듯한 엉뚱한 경로' 대신 자리표시를 보여야 한다."""
    raw = bytearray(_contents(_DT()).build())
    summ = P._parse_package_summary(bytes(raw))
    pos = bytes(raw).index(struct.pack("<ii", summ["soft_count"], summ["soft_offset"]))
    struct.pack_into("<i", raw, pos + 4, len(raw) + 100)   # 파일 밖을 가리키게
    p = tmp_path / "broken.uasset"
    p.write_bytes(bytes(raw))
    vals = _col(str(p), "Bgm")
    assert set(vals.values()) == {"soft(4B)"}, vals


def test_summary_reports_the_table():
    summ = P._parse_package_summary(_contents(_DT()).build())
    assert summ["file_version_ue5"] == 1009
    assert summ["soft_count"] == 3          # Intro, Theme_A, None


# ── 1008 이전 : 값에 경로가 그대로 ──────────────────────────────────────────

@pytest.mark.parametrize("mode", ["top", "legacy"])
def test_inline_paths_before_the_table_existed(tmp_path, mode):
    """테이블이 생기기 전 형식도 계속 읽어야 한다 — 1007 은 FName 2개, 그 전은 1개."""
    dt = _contents(_DT(soft_mode=mode))
    assert P._parse_package_summary(dt.build())["soft_count"] == 0
    assert _col(dt.save(tmp_path), "Bgm") == EXPECTED
