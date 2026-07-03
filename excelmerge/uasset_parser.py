"""UE5 DataTable .uasset 바이너리 파서 (excel_diff_merge.py에서 분리)."""
import struct


_UASSET_MAGIC = 0x9E2A83C1


def _read_fstring(buf: bytes, offset: int) -> tuple[str, int]:
    """UE FString 디코드 — Int32 length + (ASCII | UTF-16LE) + null terminator.
    음수 length는 UTF-16, 양수는 ASCII. 길이에 null terminator 포함.
    반환: (문자열, 다음 offset). 실패 시 ("", offset+4)로 진행해 풀 손상 시에도
    가능한 한 많은 엔트리를 노출.
    """
    if offset + 4 > len(buf):
        return "", offset + 4
    (n,) = struct.unpack_from("<i", buf, offset)
    offset += 4
    if n == 0:
        return "", offset
    if n > 0:
        # ASCII (length 에 null 포함)
        end = offset + n
        if end > len(buf):
            return "", offset
        raw = buf[offset:end - 1] if n >= 1 else b""
        try:
            s = raw.decode("ascii", errors="replace")
        except Exception:
            s = ""
        return s, end
    # UTF-16
    count = -n
    end = offset + count * 2
    if end > len(buf):
        return "", offset
    raw = buf[offset:end - 2] if count >= 1 else b""
    try:
        s = raw.decode("utf-16-le", errors="replace")
    except Exception:
        s = ""
    return s, end


class _UAssetParseError(Exception):
    """uasset DataTable 본격 파싱 실패 — 호출자가 폴백 경로로 전환."""


_KNOWN_PROP_TYPES = {
    "IntProperty", "Int8Property", "Int16Property", "Int64Property",
    "UInt16Property", "UInt32Property", "UInt64Property", "ByteProperty",
    "BoolProperty", "FloatProperty", "DoubleProperty",
    "NameProperty", "StrProperty", "TextProperty", "ObjectProperty",
    "SoftObjectProperty", "EnumProperty", "StructProperty",
    "ArrayProperty", "MapProperty", "SetProperty",
}


def _try_decode_name_table(
    buf: bytes, off: int, count: int, hash_len: int,
) -> tuple[list[str], int] | None:
    """주어진 (offset, count, hash_len) 조합으로 NameTable 엄격 디코드.
    각 엔트리: FString(4B len + bytes + null) + hash_len bytes.
    잘못된 length가 1개라도 나오면 None — drift 누적으로 흡수되는 가짜 통과 차단.
    """
    decoded: list[str] = []
    cur = off
    n_buf = len(buf)
    for _ in range(count):
        if cur + 4 > n_buf:
            return None
        n = struct.unpack_from("<i", buf, cur)[0]
        cur += 4
        if n > 0:
            byte_len = n
            if byte_len > 8192 or cur + byte_len > n_buf:
                return None
            try:
                s = buf[cur:cur + byte_len - 1].decode("utf-8", errors="replace")
            except Exception:
                return None
            cur += byte_len
        elif n < 0:
            byte_len = (-n) * 2
            if byte_len > 16384 or cur + byte_len > n_buf:
                return None
            try:
                s = buf[cur:cur + byte_len - 2].decode("utf-16-le", errors="replace")
            except Exception:
                return None
            cur += byte_len
        else:
            s = ""
        cur += hash_len
        if cur > n_buf:
            return None
        decoded.append(s)
    return decoded, cur


def _parse_package_summary(buf: bytes) -> dict | None:
    """UE5 PackageFileSummary 정확 파싱 — LegacyFileVersion <= -8 대상.
    반환: {name_count, name_offset, total_header_size, ...} 또는 None (실패).
    """
    try:
        if len(buf) < 0x40:
            return None
        off = 0
        magic = struct.unpack_from("<I", buf, off)[0]; off += 4
        if magic != _UASSET_MAGIC:
            return None
        legacy = struct.unpack_from("<i", buf, off)[0]; off += 4
        if legacy > 0 or legacy < -10:
            return None
        # legacy_ue3_version (항상 존재, -7 이전엔 없지만 보통은 있음)
        off += 4  # legacy_ue3_version
        off += 4  # file_version_ue4
        if legacy <= -8:
            off += 4  # file_version_ue5
        off += 4  # file_version_licensee_ue4
        # CustomVersions Array (count + count * 20bytes)
        custom_count = struct.unpack_from("<i", buf, off)[0]; off += 4
        if custom_count < 0 or custom_count > 200:
            return None
        off += custom_count * 20
        if off + 8 > len(buf):
            return None
        total_header_size = struct.unpack_from("<i", buf, off)[0]; off += 4
        if total_header_size <= 0 or total_header_size > len(buf):
            return None
        # FolderName FString
        fn_len = struct.unpack_from("<i", buf, off)[0]; off += 4
        if fn_len > 0:
            if fn_len > 4096 or off + fn_len > len(buf):
                return None
            off += fn_len
        elif fn_len < 0:
            sz = (-fn_len) * 2
            if sz > 8192 or off + sz > len(buf):
                return None
            off += sz
        if off + 12 > len(buf):
            return None
        off += 4  # package_flags
        name_count = struct.unpack_from("<i", buf, off)[0]; off += 4
        name_offset = struct.unpack_from("<i", buf, off)[0]; off += 4
        if name_count <= 0 or name_count > 1_000_000:
            return None
        if name_offset <= 0 or name_offset >= len(buf):
            return None
        return {
            "name_count": name_count,
            "name_offset": name_offset,
            "total_header_size": total_header_size,
            "legacy": legacy,
        }
    except Exception:
        return None


def _scan_name_table(buf: bytes) -> tuple[list[str], int, int]:
    """NameTable 디코드 — 1차로 정확 파싱, 2차로 휴리스틱(1바이트 step).
    반환: (names, name_offset, name_data_end). 실패 시 _UAssetParseError.
    """
    if len(buf) < 0x40:
        raise _UAssetParseError("buffer too short")

    summary = _parse_package_summary(buf)
    if summary is not None:
        cnt = summary["name_count"]
        off = summary["name_offset"]
        for hash_len in (8, 4, 0, 16):
            res = _try_decode_name_table(buf, off, cnt, hash_len)
            if res is not None:
                decoded, end = res
                return decoded, off, end

    best: tuple[list[str], int, int] | None = None
    scan_end = min(len(buf), 0x400)
    for probe in range(0x20, scan_end - 8):
        cnt, off = struct.unpack_from("<ii", buf, probe)
        if cnt <= 0 or cnt > 200000:
            continue
        if off <= probe or off >= len(buf):
            continue
        for hash_len in (8, 4, 0, 16):
            res = _try_decode_name_table(buf, off, cnt, hash_len)
            if res is None:
                continue
            decoded, end = res
            if best is None or len(decoded) > len(best[0]):
                best = (decoded, off, end)
            break

    if best is None or not best[0]:
        raise _UAssetParseError("name table decode failed")
    return best


def _read_fname(buf: bytes, off: int, names: list[str]) -> tuple[str, int]:
    """FName 8B (NameIndex int32 + Number int32) → 문자열.
    Number > 0 이면 'Name_{Number-1}' 형식.
    """
    if off + 8 > len(buf):
        raise _UAssetParseError("fname truncated")
    idx, num = struct.unpack_from("<ii", buf, off)
    if idx < 0 or idx >= len(names):
        raise _UAssetParseError(f"name index out of range: {idx}")
    name = names[idx]
    if num > 0:
        name = f"{name}_{num - 1}"
    return name, off + 8


def _read_property_tag(
    buf: bytes, off: int, names: list[str],
) -> tuple[dict | None, int, int]:
    """UPropertyTagged 헤더 1개 파싱.
    반환: (tag_dict | None, value_off, value_end). tag_dict가 None이면 'None' 종료자.
    """
    name, off = _read_fname(buf, off, names)
    if name == "None":
        return None, off, off

    type_name, off = _read_fname(buf, off, names)
    if type_name not in _KNOWN_PROP_TYPES:
        raise _UAssetParseError(f"unknown property type: {type_name}")

    if off + 8 > len(buf):
        raise _UAssetParseError("size/index truncated")
    size, array_index = struct.unpack_from("<ii", buf, off)
    off += 8
    if size < 0 or size > 100 * 1024 * 1024:
        raise _UAssetParseError(f"unreasonable size: {size}")

    tag: dict = {"name": name, "type": type_name, "size": size,
                 "array_index": array_index}

    if type_name == "StructProperty":
        struct_name, off = _read_fname(buf, off, names)
        tag["struct_name"] = struct_name
        if off + 16 > len(buf):
            raise _UAssetParseError("struct guid truncated")
        off += 16
    elif type_name == "BoolProperty":
        if off + 1 > len(buf):
            raise _UAssetParseError("bool truncated")
        tag["bool_value"] = buf[off] != 0
        off += 1
    elif type_name in ("ByteProperty", "EnumProperty"):
        enum_name, off = _read_fname(buf, off, names)
        tag["enum_name"] = enum_name
    elif type_name in ("ArrayProperty", "SetProperty"):
        inner_type, off = _read_fname(buf, off, names)
        tag["inner_type"] = inner_type
    elif type_name == "MapProperty":
        key_type, off = _read_fname(buf, off, names)
        val_type, off = _read_fname(buf, off, names)
        tag["key_type"] = key_type
        tag["value_type"] = val_type

    if off + 1 > len(buf):
        raise _UAssetParseError("hasguid truncated")
    has_guid = buf[off]
    off += 1
    if has_guid:
        if off + 16 > len(buf):
            raise _UAssetParseError("propguid truncated")
        off += 16

    value_off = off
    value_end = value_off + size
    if value_end > len(buf):
        raise _UAssetParseError("value truncated")
    return tag, value_off, value_end


def _fmt_num(v) -> str:
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return str(int(v))
    return str(v)


def _read_struct_value(buf, off, end, names) -> str:
    """중첩 struct = TaggedProperties 시퀀스. {k=v, k=v} 표기."""
    parts = []
    cur = off
    safety = 0
    while cur < end and safety < 256:
        safety += 1
        try:
            tag, vo, ve = _read_property_tag(buf, cur, names)
        except Exception:
            break
        if tag is None:
            break
        val = _read_property_value(buf, vo, ve, tag, names)
        parts.append(f"{tag['name']}={val}")
        cur = ve
    return "{" + ", ".join(parts) + "}"


def _read_array_value(buf, off, end, tag, names) -> str:
    """ArrayProperty/SetProperty 본문. 스칼라 inner는 inline, 그 외는 count 표기."""
    try:
        count = struct.unpack_from("<i", buf, off)[0]
        if count < 0 or count > 1000000:
            return f"array:{end-off}B"
        cur = off + 4
        inner = tag.get("inner_type", "")
        if inner == "StructProperty":
            return f"[{count} structs]"
        items: list[str] = []
        for _ in range(min(count, 5)):
            if cur >= end:
                break
            if inner == "IntProperty":
                if cur + 4 > end: break
                items.append(str(struct.unpack_from("<i", buf, cur)[0])); cur += 4
            elif inner == "Int64Property":
                if cur + 8 > end: break
                items.append(str(struct.unpack_from("<q", buf, cur)[0])); cur += 8
            elif inner == "FloatProperty":
                if cur + 4 > end: break
                items.append(_fmt_num(struct.unpack_from("<f", buf, cur)[0])); cur += 4
            elif inner == "DoubleProperty":
                if cur + 8 > end: break
                items.append(_fmt_num(struct.unpack_from("<d", buf, cur)[0])); cur += 8
            elif inner == "ByteProperty":
                if cur >= end: break
                items.append(str(buf[cur])); cur += 1
            elif inner == "BoolProperty":
                if cur >= end: break
                items.append("true" if buf[cur] else "false"); cur += 1
            elif inner == "NameProperty":
                if cur + 8 > end: break
                n, cur = _read_fname(buf, cur, names)
                items.append(n)
            elif inner == "StrProperty":
                s, cur = _read_fstring(buf, cur)
                items.append(s)
            elif inner == "ObjectProperty":
                if cur + 4 > end: break
                items.append(f"Obj({struct.unpack_from('<i', buf, cur)[0]})"); cur += 4
            else:
                break
        suffix = ", …" if count > len(items) else ""
        return f"[{', '.join(items)}{suffix}]"
    except Exception:
        return f"array:{end-off}B"


def _read_map_value(buf, off, end, tag, names) -> str:
    try:
        if off + 8 <= end:
            _, n_entries = struct.unpack_from("<ii", buf, off)
            return f"map(count={n_entries})"
    except Exception:
        pass
    return f"map:{end-off}B"


def _read_property_value(buf, value_off, value_end, tag, names) -> str:
    """tag 타입별로 값 디코드. 실패 시 hex: 폴백."""
    t = tag["type"]
    try:
        if t == "IntProperty":
            return str(struct.unpack_from("<i", buf, value_off)[0])
        if t == "Int8Property":
            return str(struct.unpack_from("<b", buf, value_off)[0])
        if t == "Int16Property":
            return str(struct.unpack_from("<h", buf, value_off)[0])
        if t == "Int64Property":
            return str(struct.unpack_from("<q", buf, value_off)[0])
        if t == "UInt16Property":
            return str(struct.unpack_from("<H", buf, value_off)[0])
        if t == "UInt32Property":
            return str(struct.unpack_from("<I", buf, value_off)[0])
        if t == "UInt64Property":
            return str(struct.unpack_from("<Q", buf, value_off)[0])
        if t == "FloatProperty":
            return _fmt_num(struct.unpack_from("<f", buf, value_off)[0])
        if t == "DoubleProperty":
            return _fmt_num(struct.unpack_from("<d", buf, value_off)[0])
        if t == "BoolProperty":
            return "true" if tag.get("bool_value") else "false"
        if t == "ByteProperty":
            if tag.get("enum_name") and tag["enum_name"] != "None":
                name, _ = _read_fname(buf, value_off, names)
                return name
            return str(buf[value_off])
        if t == "EnumProperty":
            name, _ = _read_fname(buf, value_off, names)
            return name
        if t == "NameProperty":
            name, _ = _read_fname(buf, value_off, names)
            return name
        if t == "StrProperty":
            s, _ = _read_fstring(buf, value_off)
            return s
        if t == "TextProperty":
            return f"text({value_end - value_off}B)"
        if t == "ObjectProperty":
            idx = struct.unpack_from("<i", buf, value_off)[0]
            return f"Obj({idx})"
        if t == "SoftObjectProperty":
            try:
                name, off2 = _read_fname(buf, value_off, names)
                sub, _ = _read_fstring(buf, off2)
                return f"{name}{':' + sub if sub else ''}"
            except Exception:
                return f"soft({value_end - value_off}B)"
        if t == "StructProperty":
            return _read_struct_value(buf, value_off, value_end, names)
        if t in ("ArrayProperty", "SetProperty"):
            return _read_array_value(buf, value_off, value_end, tag, names)
        if t == "MapProperty":
            return _read_map_value(buf, value_off, value_end, tag, names)
    except Exception:
        pass
    raw = buf[value_off:value_end]
    return f"hex:{raw[:16].hex()}{'…' if len(raw) > 16 else ''}"


def _row_name_is_valid(name: str) -> bool:
    """RowName이 정상 키 형태인가. 패키지 경로/Property 타입명/None은 거부."""
    if not name or name == "None":
        return False
    if name.startswith("/") or name.endswith("Property"):
        return False
    return True


def _try_count_rows(
    buf: bytes, row_map_off: int, num_rows: int, names: list[str], cap: int,
) -> int:
    """후보 위치에서 실제로 몇 row를 안전하게 디코드할 수 있는지 카운트.
    cap 만큼만 시도해 빠르게 비교."""
    cur = row_map_off
    n_buf = len(buf)
    decoded = 0
    target = min(num_rows, cap)
    for _ in range(target):
        if cur + 8 > n_buf:
            break
        try:
            key, cur = _read_fname(buf, cur, names)
        except Exception:
            break
        if not _row_name_is_valid(key):
            break
        safety = 0
        ok = True
        while safety < 1000:
            safety += 1
            try:
                tag, vo, ve = _read_property_tag(buf, cur, names)
            except Exception:
                ok = False
                break
            if tag is None:
                cur = vo
                break
            cur = ve
        if not ok:
            break
        decoded += 1
    return decoded


def _find_row_map_start(
    buf: bytes, scan_from: int, names: list[str],
) -> tuple[int, int] | None:
    """Export 데이터 영역에서 RowMap 시작점 탐색.
    검증: int32 NumRows + FName RowName(키 형태) + 첫 PropertyTag 정상 +
          후보 위치에서 실제 row 디코드 카운트 산정.
    가장 많이 디코드되는 후보 채택.
    반환: (row_map_offset = NumRows 직후 = 첫 RowName 시작, num_rows) | None
    """
    n_buf = len(buf)
    name_count = len(names)
    if scan_from < 0:
        scan_from = 0
    if scan_from >= n_buf - 24:
        return None

    best: tuple[int, int, int] | None = None  # (off, num_rows, score)
    pos = scan_from
    end = n_buf - 24
    while pos < end:
        num_rows, name_idx, name_num = struct.unpack_from("<iii", buf, pos)
        if 1 <= num_rows <= 100000 and 0 <= name_idx < name_count and 0 <= name_num <= 1024:
            row_name = names[name_idx]
            if _row_name_is_valid(row_name):
                try:
                    tag, _, _ = _read_property_tag(buf, pos + 12, names)
                    if tag is not None and tag.get("type") in _KNOWN_PROP_TYPES \
                            and 0 < tag.get("size", 0) <= 1024 * 1024:
                        cap = min(num_rows, 64)
                        score = _try_count_rows(buf, pos + 4, num_rows, names, cap)
                        if score >= max(1, cap // 2):
                            if best is None or score > best[2]:
                                best = (pos + 4, num_rows, score)
                                if score >= cap:
                                    return best[0], best[1]
                except Exception:
                    pass
        pos += 1
    if best is not None:
        return best[0], best[1]
    return None


def _parse_row_map(
    buf: bytes, row_map_off: int, num_rows: int, names: list[str],
) -> tuple[list[str], list[dict]]:
    """RowMap 전체 파싱. row 단위 try/except로 일부 실패는 빈 row로 진행."""
    row_keys: list[str] = []
    rows: list[dict[str, str]] = []
    cur = row_map_off

    for _ in range(num_rows):
        if cur + 8 > len(buf):
            break
        try:
            key, cur = _read_fname(buf, cur, names)
        except Exception:
            break
        row_props: dict[str, str] = {}
        safety = 0
        row_failed = False
        while safety < 1000:
            safety += 1
            try:
                tag, vo, ve = _read_property_tag(buf, cur, names)
            except Exception:
                row_failed = True
                break
            if tag is None:
                cur = vo
                break
            try:
                val = _read_property_value(buf, vo, ve, tag, names)
            except Exception:
                val = ""
            row_props[tag["name"]] = val
            cur = ve
        row_keys.append(key)
        rows.append(row_props)
        if row_failed:
            break

    if not row_keys:
        raise _UAssetParseError("no rows decoded")
    return row_keys, rows


def _build_datatable_matrix(
    row_keys: list[str], rows: list[dict[str, str]],
) -> list[list[str]]:
    """헤더=['RowName'] + 컬럼 합집합(첫 등장 순서 보존)."""
    cols: list[str] = []
    seen: set[str] = set()
    for r in rows:
        for k in r.keys():
            if k not in seen:
                seen.add(k)
                cols.append(k)
    matrix: list[list[str]] = [["RowName"] + cols]
    for key, props in zip(row_keys, rows):
        matrix.append([key] + [props.get(c, "") for c in cols])
    return matrix


def _load_uasset_fallback_matrix(buf: bytes) -> list[list[str]]:
    """본격 파싱 실패 시 [field, value] 헤더 + NameTable 덤프."""
    matrix: list[list[str]] = [["field", "value"]]

    if len(buf) < 8:
        matrix.append(["error", "파일이 너무 짧습니다"])
        return matrix

    (magic,) = struct.unpack_from("<I", buf, 0)
    matrix.append(["magic", f"0x{magic:08X}"])
    if magic != _UASSET_MAGIC:
        matrix.append(["error", "UE 패키지 시그니처가 아닙니다"])
        return matrix

    (legacy,) = struct.unpack_from("<i", buf, 4)
    matrix.append(["legacy_file_version", str(legacy)])

    try:
        names, name_offset, _ = _scan_name_table(buf)
    except _UAssetParseError:
        names, name_offset = [], 0

    matrix.append(["name_count", str(len(names))])
    matrix.append(["name_offset", f"0x{name_offset:08X}" if name_offset else "0"])
    matrix.append(["file_size", str(len(buf))])

    if names:
        matrix.append(["--- name_table ---", ""])
        for i, s in enumerate(names):
            matrix.append([f"name[{i}]", s])
    else:
        matrix.append(["warning", "Name Table 디코드 실패 — 헤더만 표시"])
    return matrix


def load_uasset_as_matrix(path: str) -> list[list[str]]:
    """UE5 DataTable .uasset → RowName + RowStruct 프로퍼티 매트릭스.
    본격 파싱 흐름:
      1) magic 검증
      2) NameTable 디코드 (hash 길이 자동 판별)
      3) Export 영역에서 RowMap 시작점 탐지
      4) RowMap 파싱 (row 단위 안전망)
      5) 컬럼 합집합으로 매트릭스 빌드
    어느 단계든 실패하면 기존 [field/value] 헤더 덤프 폴백 — 앱 크래시 방지.
    """
    with open(path, "rb") as f:
        buf = f.read()

    if len(buf) < 8:
        return _load_uasset_fallback_matrix(buf)
    (magic,) = struct.unpack_from("<I", buf, 0)
    if magic != _UASSET_MAGIC:
        return _load_uasset_fallback_matrix(buf)

    try:
        names, _, name_end = _scan_name_table(buf)
        summary = _parse_package_summary(buf)
        scan_from = summary["total_header_size"] if summary else name_end
        hit = _find_row_map_start(buf, scan_from, names)
        if hit is None:
            raise _UAssetParseError("row map not found")
        row_map_off, num_rows = hit
        row_keys, rows = _parse_row_map(buf, row_map_off, num_rows, names)
        return _build_datatable_matrix(row_keys, rows)
    except Exception:
        return _load_uasset_fallback_matrix(buf)
