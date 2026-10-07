"""xlsx 저장 — sheet XML 직접 패치로 수식 보존 기록 (excel_diff_merge.py에서 분리)."""
import bisect
import os
import posixpath
import re
import zipfile
from copy import deepcopy
from collections import defaultdict

from lxml import etree
from .colref import get_column_letter, column_index_from_string

from . import ooxml
from .logutil import log


_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_COL_RE = re.compile(r"([A-Z]+)(\d+)")

# 자주 쓰는 네임스페이스 태그(Clark 표기) — 반복 f-string 생성 제거.
_TAG_SHEETDATA = f"{{{_NS}}}sheetData"
_TAG_ROW = f"{{{_NS}}}row"
_TAG_C = f"{{{_NS}}}c"
_TAG_F = f"{{{_NS}}}f"
_TAG_V = f"{{{_NS}}}v"

# 저장하며 패키지에서 빼는 부품. 여기 넣은 것은 그것을 가리키는 **선언·관계까지** 함께
# 걷어낸다 — 부품만 빼고 참조를 남기면 Excel 이 열 때마다 "내용에 문제가 있습니다" 복구
# 창을 띄운다. 실제로 그렇게 나갔다(40_Build 의 5개 파일, .xlsx 3 / .xlsm 2).
_DROP_PARTS = frozenset({"xl/calcChain.xml"})


def _cell_ref(r: int, c: int) -> str:
    return f"{get_column_letter(c + 1)}{r + 1}"


def _promote_empty_cols_to_delete(
    patches: dict[str, str],
    delete_row_nums: set[int],
    path: str | None,
    sheet_name=None,
) -> tuple[dict[str, str], set[int], set[str]]:
    """
    1단계: 패치 적용 후 모든 셀이 빈값인 행 → delete_row_nums로 승격 (patches에서 제거)
    2단계: 행 삭제 반영 후 모든 셀이 빈값인 열 → delete_col_letters 반환

    ★ 1단계의 '빈값'은 **화면에 보이는 값**(수식이면 계산 결과)으로 판단한다.
    한쪽에만 있는 행을 지울 때, 그 행의 어떤 셀이 양쪽 모두 빈값이면 변경이 아니라
    '같음'이라 병합 준비에서 빠진다 → 패치가 안 붙는다. 그 셀이 마침 **결과가 빈
    수식**(`<f>…</f><v/>`)이면 수식 원문을 값으로 쳤을 때 '이 행엔 아직 내용이 있다'로
    읽혀 행 삭제 승격이 막히고, 나머지 열만 빈값으로 덮여 **빈 행**이 남았다.
    (실제 사례: Data_HelpPopUp_C.xlsx 의 TID 245102 — L열 `#Desc` 가 결과가 빈 VLOOKUP.)

    2단계(열 삭제)는 그대로 수식 원문을 값으로 본다 — 지금 결과가 비어 있다고 해서
    수식이 든 열을 통째로 지우면 안 되기 때문이다.
    """
    if not path:
        return patches, delete_row_nums, set()

    try:
        with zipfile.ZipFile(path, "r") as zin:
            sheet_path = _resolve_sheet_path(zin, sheet_name)
            xml_data = zin.read(sheet_path)
        tree = etree.fromstring(xml_data)
        ns = _NS
        sheetdata = tree.find(f"{{{ns}}}sheetData")
        file_cells: dict[str, str] = {}
        file_disp: dict[str, str] = {}
        file_row_refs: dict[int, set[str]] = defaultdict(set)
        if sheetdata is not None:
            for row_el in sheetdata:
                rn = int(row_el.get("r", 0))
                for c_el in row_el:
                    ref = c_el.get("r", "")
                    if not ref:
                        continue
                    v_el  = c_el.find(f"{{{ns}}}v")
                    f_el  = c_el.find(f"{{{ns}}}f")
                    is_el = c_el.find(f"{{{ns}}}is")
                    if v_el is not None and v_el.text:
                        disp = v_el.text
                    elif is_el is not None:
                        t_el = is_el.find(f"{{{ns}}}t")
                        disp = t_el.text if (t_el is not None and t_el.text) else ""
                    else:
                        disp = ""
                    val = ("=" + f_el.text) if (f_el is not None and f_el.text) else disp
                    file_cells[ref] = val
                    file_disp[ref] = disp
                    file_row_refs[rn].add(ref)
    except Exception:
        # 빈 열/행 감지 실패 → 승격 없이 진행(안전). 원인은 진단 로그로만 남긴다.
        log.warning("빈 열/행 감지 실패(승격 건너뜀): %s", path, exc_info=True)
        return patches, delete_row_nums, set()

    merged: dict[str, str] = {**file_cells, **patches}
    # 행이 비었는지 볼 때만 쓰는 '보이는 값' 시야 — 수식은 **계산 결과**로 본다.
    # (열 삭제 판정은 아래 merged 를 계속 쓴다 — 수식이 든 열은 지우지 않는다.)
    merged_disp: dict[str, str] = {**file_disp, **patches}

    new_patches = dict(patches)
    new_deletes = set(delete_row_nums)

    # ── 1단계: 빈 행 → 행 삭제 승격 ──────────────────────────────────────────
    patch_rows: dict[int, set[str]] = defaultdict(set)
    for ref in patches:
        m = _COL_RE.match(ref)
        if m:
            patch_rows[int(m.group(2))].add(ref)

    for row_num, patched_refs in patch_rows.items():
        if row_num in new_deletes:
            continue
        # 이 행의 모든 패치값이 빈값인지 확인
        if any(merged.get(ref, "") != "" for ref in patched_refs):
            continue
        # 파일의 이 행에 패치 외 다른 값이 있는지 확인
        all_refs_in_row = file_row_refs.get(row_num, set())
        non_patched_refs = all_refs_in_row - patched_refs
        if any(merged_disp.get(ref, "") != "" for ref in non_patched_refs):
            continue
        # 행 전체가 빈값 → 행 삭제로 전환
        for ref in patched_refs:
            new_patches.pop(ref, None)
        new_deletes.add(row_num)

    # ── 2단계: 행 삭제 반영 후 빈 열 감지 ────────────────────────────────────
    col_vals: dict[str, list[str]] = defaultdict(list)
    for ref, val in merged.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        row_num = int(m.group(2))
        if row_num in new_deletes:
            continue
        col_vals[m.group(1)].append(val)

    delete_col_letters: set[str] = set()
    for col_letter, vals in col_vals.items():
        if all(v == "" for v in vals):
            delete_col_letters.add(col_letter)

    return new_patches, new_deletes, delete_col_letters


def _is_file_locked(path: str) -> bool:
    """다른 프로세스가 잡고 있어 쓸 수 없는 상태인지 — **파일을 만들지 않고** 확인한다.

    예전엔 open(path, "a") 였다. append 모드는 없는 경로에 **빈 파일을 만들어 버려서**,
    '잠겼는지 보기만' 하는 호출이 없던 파일을 생성하는 부작용이 있었다.
    없는 파일은 잠김이 아니다 — 실제 저장 단계에서 제대로 된 오류로 드러난다.
    """
    try:
        with open(path, "r+b"):
            return False
    except FileNotFoundError:
        return False
    except OSError:          # PermissionError·IOError 포함
        return True


def _trim_grid(grid) -> list:
    """뒤쪽 빈 칸·빈 행을 떼어 낸다 — 같은 내용이 다른 모양으로 읽히는 것을 막는다.

    로더는 파일에 <c> 가 없는 꼬리를 돌려주지 않는다. 그래서 "마지막 열을 비웠다" 같은
    저장은 기대 격자엔 빈 칸이 남고 실제 파일엔 없어서, 내용이 같은데도 다르다고 잡힌다.
    비교 전에 양쪽을 같은 모양으로 맞춘다(중간의 빈 칸은 그대로 — 그건 진짜 내용이다).
    """
    out = [list(r) for r in (grid or [])]
    for row in out:
        while row and (row[-1] or "") == "":
            row.pop()
    while out and not out[-1]:
        out.pop()
    return out


def _expected_after(before, patches, insert_rows=None,
                    delete_row_nums=None, delete_col_letters=None) -> list:
    """저장 뒤 시트가 어떤 모습이어야 하는지 — **되읽었을 때 보일 값**으로 계산한다.

    _patch_sheet_xml 의 순서를 그대로 흉내 낸다: 덮어쓰기(저장 전 좌표) → 행 삭제
    (아래가 위로 당겨짐) → 열 삭제 → 신규 행 꼬리 추가.

    수식(= 로 시작)은 캐시값 없이 쓰이므로 되읽으면 빈 칸이다 — 기대값도 빈 칸.
    """
    grid = [list(r) for r in (before or [])]

    for ref, val in (patches or {}).items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        r = int(m.group(2)) - 1
        c = column_index_from_string(m.group(1)) - 1
        if r < 0 or c < 0:
            continue
        while len(grid) <= r:
            grid.append([])
        row = grid[r]
        while len(row) <= c:
            row.append("")
        row[c] = "" if (val or "").startswith("=") else val

    dead = {n - 1 for n in (delete_row_nums or ())}
    grid = [row for i, row in enumerate(grid) if i not in dead]

    for letter in (delete_col_letters or ()):
        c = column_index_from_string(letter) - 1
        for row in grid:
            if 0 <= c < len(row):
                row[c] = ""

    for cells in (insert_rows or []):
        width = max((t[0] for t in cells), default=-1) + 1
        row = [""] * width
        for t in cells:
            c, val = t[0], t[1]
            if 0 <= c < width:
                row[c] = "" if (val or "").startswith("=") else val
        grid.append(row)

    return grid


def _grid_mismatches(after, expected, limit: int = 5) -> list:
    """기대 격자와 실제 격자를 전수 대조 — 어긋난 곳을 사람이 읽을 문장으로."""
    a = _trim_grid(after)
    e = _trim_grid(expected)
    bad = []
    if len(a) != len(e):
        bad.append(f"행 수가 다릅니다 — 예상 {len(e):,}행 → 파일에는 {len(a):,}행")
    for r in range(max(len(a), len(e))):
        ar = a[r] if r < len(a) else []
        er = e[r] if r < len(e) else []
        for c in range(max(len(ar), len(er))):
            av = ar[c] if c < len(ar) else ""
            ev = er[c] if c < len(er) else ""
            if av != ev:
                bad.append(f"{_cell_ref(r, c)}: 예상 {ev!r} → 파일에는 {av!r}")
                if len(bad) >= limit:
                    return bad
    return bad


def _patch_mismatches(rows, patches: dict, delete_row_nums=None,
                      limit: int = 5) -> list:
    """덮어쓴 셀이 의도한 값으로 들어갔는지 — 좌표를 짚어 주는 정밀 검사.

    ★ 행을 지우는 저장에서는 **자리가 밀린다**. _renumber_after_delete 가 VBA .Delete
    처럼 삭제된 행 아래를 위로 당기므로, 패치 좌표(저장 전 기준)를 그대로 보면 한 칸
    아래 행의 값과 비교하게 된다 — 멀쩡한 저장이 '검증 실패'로 막혔다(실사용 신고).
    그래서 삭제된 행 수만큼 행 번호를 당겨서 본다. 지워진 행 자체의 패치는 건너뛴다.

    수식(= 로 시작)은 캐시값이 없어 되읽으면 빈 칸이므로 건너뛴다.
    """
    targets = {ref: v for ref, v in (patches or {}).items()
               if not (v or "").startswith("=")}
    if not targets:
        return []
    deleted = sorted(delete_row_nums or ())
    bad = []
    for ref, expect in targets.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        col_letters, row_num = m.group(1), int(m.group(2))
        if row_num in (delete_row_nums or ()):
            continue                       # 통째로 지워진 행 — 확인할 자리가 없다
        shift = sum(1 for d in deleted if d < row_num)
        r, c = row_num - 1 - shift, column_index_from_string(col_letters) - 1
        got = rows[r][c] if (0 <= r < len(rows) and 0 <= c < len(rows[r])) else ""
        if got != expect:
            bad.append(f"{ref}: 쓰려던 값 {expect!r} → 파일에는 {got!r}")
        if len(bad) >= limit:
            break
    return bad


def _rel_target_part(rels_path: str, rel) -> str | None:
    """그 관계가 가리키는 패키지 안 부품 경로. 바깥을 가리키면 None.

    Target 은 그 .rels 가 지키는 폴더 기준의 상대 경로다
    (xl/_rels/workbook.xml.rels 의 "calcChain.xml" → xl/calcChain.xml).
    """
    if rel.get("TargetMode") == "External":
        return None
    tgt = (rel.get("Target") or "").strip()
    if not tgt or "://" in tgt:
        return None
    if tgt.startswith("/"):
        return tgt.lstrip("/")
    base = posixpath.dirname(posixpath.dirname(rels_path))   # xl/_rels/a.rels → xl
    return posixpath.normpath(posixpath.join(base, tgt)).replace("\\", "/")


def _child_elements(root) -> list:
    """주석·처리 지시를 뺀 자식 요소만 — lxml 은 주석도 자식으로 돌려준다."""
    return [el for el in root if isinstance(el.tag, str)]


def _strip_dangling_refs(name: str, data: bytes, dropped: set) -> bytes:
    """뺀 부품을 가리키는 선언([Content_Types].xml)·관계(*.rels)를 걷어낸다.

    calcChain.xml 을 빼는 것 자체는 맞다 — 값을 고치면 계산 순서 캐시는 무효가 되고,
    Excel 이 다음에 열 때 다시 만든다. 빠져 있던 건 **참조 정리** 하나뿐이었다.
    """
    if not dropped:
        return data
    if name == "[Content_Types].xml":
        want = "Override"

        def hit(el):
            return (el.get("PartName") or "").lstrip("/") in dropped
    elif name.endswith(".rels"):
        want = "Relationship"

        def hit(el):
            return _rel_target_part(name, el) in dropped
    else:
        return data

    root = etree.fromstring(data)
    gone = [el for el in _child_elements(root)
            if etree.QName(el).localname == want and hit(el)]
    if not gone:
        return data
    for el in gone:
        root.remove(el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                          standalone=True)


def _package_mismatches(path: str, limit: int = 5) -> list:
    """패키지 안에 **없는 부품**을 가리키는 선언·관계를 찾는다.

    값 대조로는 이걸 못 잡는다. 셀은 전부 맞는데 열 때마다 Excel 이 복구 창을 띄우는
    파일이 실제로 사용자에게 나갔다 — 저장이 calcChain.xml 을 빼면서 그것을 가리키는
    Content_Types Override 와 workbook 관계를 남겼기 때문이다. 값만 보는 검증은 그
    파일을 그대로 통과시켰다. 그래서 여기서 **구조**를 본다.
    """
    bad = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        try:
            ct = zf.read("[Content_Types].xml")
        except KeyError:
            return ["[Content_Types].xml 이 없습니다 — 패키지가 깨졌습니다."]
        for el in _child_elements(etree.fromstring(ct)):
            if etree.QName(el).localname != "Override":
                continue
            part = (el.get("PartName") or "").lstrip("/")
            if part and part not in names:
                bad.append(f"[Content_Types].xml 이 없는 부품을 선언합니다: {part}")
        for rels in sorted(n for n in names if n.endswith(".rels")):
            for el in _child_elements(etree.fromstring(zf.read(rels))):
                if etree.QName(el).localname != "Relationship":
                    continue
                part = _rel_target_part(rels, el)
                if part and part not in names:
                    bad.append(
                        f"{rels} 의 {el.get('Id')} 가 없는 부품을 가리킵니다: {part}")
    return bad[:limit]


def _formula_mismatches(path: str, limit: int = 5) -> list:
    """저장한 파일의 수식이 성립하는가 — 공유 수식 그룹과 수식 ref 범위.

    값 대조로는 못 잡는다. 셀 값은 <v> 에 그대로 있어 전부 일치하는데, Excel 은
    열면서 "내용에 문제가 있습니다" 복구 창을 띄운다. 실제로 그런 파일 둘이 사용자에게
    나갔다(Data_MiniGameRoulette_CS.xlsx, Data_TextUITable_CS.xlsm).
    """
    bad = []
    with zipfile.ZipFile(path) as zf:
        sheets = sorted(n for n in zf.namelist()
                        if n.startswith("xl/worksheets/") and n.endswith(".xml"))
        for sp in sheets:
            root = etree.fromstring(zf.read(sp))
            sd = root.find(_TAG_SHEETDATA)
            if sd is None:
                continue
            masters, members = {}, {}
            for c_el in sd.iter(_TAG_C):
                f_el = c_el.find(_TAG_F)
                if f_el is None:
                    continue
                own = c_el.get("r", "")
                # 공유가 아닌 수식(배열 등)의 ref 는 **자기 셀을 담아야** 한다.
                if f_el.get("t") != "shared":
                    rng = f_el.get("ref")
                    if rng and not _ref_in(own, rng):
                        bad.append(f"{sp}: {own} 의 "
                                   f"{f_el.get('t') or '수식'} ref={rng} 가 "
                                   f"자기 셀을 담지 못합니다")
                        if len(bad) >= limit:
                            return bad[:limit]
                    continue
                si = f_el.get("si")
                members.setdefault(si, []).append(own)
                if f_el.get("ref"):
                    if si in masters:
                        bad.append(f"{sp}: 공유수식 si={si} 의 주인이 둘입니다"
                                   f" ({masters[si][0]}, {own})")
                    masters[si] = (own, f_el.get("ref"), (f_el.text or "").strip())
            for si, refs in members.items():
                if si not in masters:
                    bad.append(f"{sp}: 공유수식 si={si} 에 주인이 없습니다"
                               f" (추종자 {len(refs)}개, 예: {refs[0]})")
                    continue
                own, rng, body = masters[si]
                if not body:
                    bad.append(f"{sp}: 공유수식 si={si} 주인 {own} 의 수식이 비었습니다")
                outside = [r for r in refs if not _ref_in(r, rng)]
                if outside:
                    who = "주인" if own in outside else "추종자"
                    bad.append(f"{sp}: 공유수식 si={si} 의 ref={rng} 가 "
                               f"{who} {outside[0]} 을(를) 담지 못합니다")
                if len(bad) >= limit:
                    return bad[:limit]
    return bad[:limit]


def _self_referencing(path: str, self_sheet=None) -> set:
    """자기 셀을 가리키는 수식이 든 셀들 — 순환 참조.

    본문이 있는 수식(주인·일반)만 본다. 추종자는 본문이 없어 주인에서 유도해야 하는데,
    그 유도까지 돌리면 큰 시트에서 저장이 눈에 띄게 느려진다. 신고된 사고
    ('B3: =B2+1' 이 행 삭제로 B2 로 당겨져 '=B2+1' 순환)는 일반 수식이라 여기에 걸린다.

    판정은 **자기 셀**을 품는 참조가 있는가다. 자기 '행' 으로 보면 안 된다 — 같은 행의
    옆 칸을 쓰는 수식('UITable.'&A36 이 D36 에 있는 꼴)이 실제 데이터에 흔해서, 그렇게
    보면 멀쩡한 저장이 줄줄이 막힌다(실측으로 걸렀다).
    """
    out = set()
    with zipfile.ZipFile(path) as zf:
        for sp in sorted(n for n in zf.namelist()
                         if n.startswith("xl/worksheets/") and n.endswith(".xml")):
            sd = etree.fromstring(zf.read(sp)).find(_TAG_SHEETDATA)
            if sd is None:
                continue
            for c_el in sd.iter(_TAG_C):
                f_el = c_el.find(_TAG_F)
                if f_el is None:
                    continue
                body = (f_el.text or "").strip()
                own = c_el.get("r", "")
                if not body or _ref_parts(own) is None:
                    continue
                if any(tok[0] == "ref"
                       and _prefix_targets_self(tok[1], self_sheet)
                       and _ref_covers(tok[2], own)
                       for tok in _formula_tokens(body)):
                    out.add(f"{sp}!{own}")
    return out


def _new_circular_formulas(before_path, after_path, self_sheet, limit: int = 5) -> list:
    """저장 전엔 없던 순환 참조가 생겼나.

    행을 지우면 엑셀은 수식도 함께 고쳐 준다. 우리가 그걸 빠뜨리면 'B3: =B2+1' 이
    B2 로 당겨지며 자기 자신을 가리키게 된다 — 값 대조로는 안 잡힌다(캐시 값은 그대로다).

    **이미 있던** 순환은 건드리지 않는다. 성한 파일만 들어온다고 가정하면 안 된다 —
    실제 데이터에도 과거 편집이 남긴 #REF! 가 7천 셀 넘게 들어 있다.
    """
    try:
        before = _self_referencing(before_path, self_sheet)
    except Exception:
        log.warning("저장 전 순환 참조를 읽지 못해 검사를 건너뜁니다", exc_info=True)
        return []
    fresh = sorted(_self_referencing(after_path, self_sheet) - before)
    return [f"{ref} 의 수식이 자기 셀을 가리킵니다(순환 참조)" for ref in fresh[:limit]]


def _sheet_order_mismatches(path: str, limit: int = 5) -> list:
    """행·셀이 번호 순서대로, 겹치지 않게 들어 있는가.

    엑셀은 sheetData 안의 <row> 가 오름차순이고 번호가 겹치지 않기를 요구한다. 행 안의
    <c> 도 마찬가지다. 어기면 "내용에 문제가 있습니다" 복구 창이 뜬다.

    값 대조로는 못 잡는다 — 로더는 번호로 격자를 채우므로 순서가 어긋나도 같은 값이
    나온다. 실제로 행 삽입이 빈 <row> 와 번호가 겹쳐 이 상태로 저장된 적이 있다.
    """
    bad = []
    with zipfile.ZipFile(path) as zf:
        for sp in sorted(n for n in zf.namelist()
                         if n.startswith("xl/worksheets/") and n.endswith(".xml")):
            sd = etree.fromstring(zf.read(sp)).find(_TAG_SHEETDATA)
            if sd is None:
                continue
            prev, seen = 0, set()
            for row_el in sd:
                try:
                    rn = int(row_el.get("r", 0))
                except (TypeError, ValueError):
                    continue
                if rn in seen:
                    bad.append(f"{sp}: {rn}행이 두 번 들어 있습니다")
                elif rn <= prev:
                    bad.append(f"{sp}: 행 순서가 뒤집혔습니다({prev} → {rn})")
                seen.add(rn)
                prev = rn
                pc, cols = 0, set()
                for c_el in row_el:
                    m = _COL_RE.match(c_el.get("r", ""))
                    if not m:
                        continue
                    ci = column_index_from_string(m.group(1))
                    if int(m.group(2)) != rn:
                        bad.append(f"{sp}: {c_el.get('r')} 가 {rn}행 안에 있습니다")
                    elif ci in cols:
                        bad.append(f"{sp}: {c_el.get('r')} 가 두 번 들어 있습니다")
                    elif ci <= pc:
                        bad.append(f"{sp}: {rn}행의 셀 순서가 뒤집혔습니다"
                                   f"({c_el.get('r')})")
                    cols.add(ci)
                    pc = ci
                if len(bad) >= limit:
                    return bad[:limit]
    return bad[:limit]


def _verify_saved(tmp_path: str, before, patches: dict, insert_rows=None,
                  delete_row_nums=None, delete_col_letters=None,
                  sheet_name=None, base_path=None, self_sheet=None) -> None:
    """갓 쓴 임시 파일을 되읽어 **화면이 약속한 결과 전체**와 대조. 다르면 예외.

    '쓰고 나서 확인'이 아니라 **'확인하고 나서 바꾼다'** — 원본을 교체하기 전에 본다.
    읽기는 비교에 쓰는 로더 그대로라, 사용자가 다음에 열었을 때 보게 될 글자를 그대로
    본다(예전 _is_numeric 버그처럼 XML 은 멀쩡한데 값이 달라 보이는 경우까지 잡힌다).

    세 단계로 본다.
      0. **파일 구조** — 없는 부품을 가리키는 선언·관계(_package_mismatches),
         그리고 성립하지 않는 수식(_formula_mismatches).
      1. 덮어쓴 셀 — 좌표를 짚어 주는 정밀 검사(_patch_mismatches).
      2. **시트 전체** — 기대 격자와 전수 대조(_grid_mismatches).

    0 이 필요한 이유: 1·2 는 **값만** 본다. 그래서 셀은 전부 맞는데 Excel 이 열 때마다
    복구 창을 띄우는 파일을 그대로 통과시켰다(실사용 신고 — calcChain.xml 을 빼면서
    참조를 남겼다). 값이 맞다고 파일이 멀쩡한 건 아니다.

    2 가 필요한 이유: 1 은 **자기가 쓴 셀만** 본다. 그래서 '지웠어야 할 행이 빈 행으로
    살아남은' 경우(실사용 신고, v206)나 건드리지 않기로 한 셀이 바뀐 경우를 못 잡는다.
    앞으로 열 매칭이 들어오면 '엉뚱한 열에 썼다'도 1 로는 잡히지 않는다 — 의도한 자리에
    의도한 값을 썼으니 통과한다. 그 구멍을 2 가 막는다.

    before 가 None 이면(원본을 못 읽었다) 2 는 건너뛴다 — 확인 수단이 없다고 멀쩡한
    저장을 막지는 않는다. 1 은 그대로 돈다.
    """
    bad = (_package_mismatches(tmp_path) or _sheet_order_mismatches(tmp_path)
           or _formula_mismatches(tmp_path))
    if not bad and delete_row_nums and base_path:
        bad = _new_circular_formulas(base_path, tmp_path, self_sheet)
    if not bad:
        from .loaders import load_values_any   # 순환 import 방지 — 저장 시점에만 필요
        after = load_values_any(tmp_path, sheet_name=sheet_name)
        bad = _patch_mismatches(after, patches, delete_row_nums)
        if not bad and before is not None:
            bad = _grid_mismatches(
                after,
                _expected_after(before, patches, insert_rows,
                                delete_row_nums, delete_col_letters))
    if bad:
        raise ValueError(
            "저장 검증 실패 — 원본 파일을 그대로 두었습니다." + \
            chr(10) + chr(10).join(bad))


def _before_values(path: str, sheet_name):
    """저장 전 원본 값 격자 — 전수 대조의 기준. 못 읽으면 None(대조를 건너뛴다).

    원본은 os.replace 전까지 그대로이므로 언제 읽어도 같다. 보통은 방금 비교하며 읽은
    것이 값 캐시에 남아 있어 공짜에 가깝다(캐시 열쇠 = 경로·mtime·시트).
    """
    try:
        from .loaders import load_values_any
        return load_values_any(path, sheet_name=sheet_name)
    except Exception:
        log.warning("저장 전 원본을 읽지 못해 전수 대조를 건너뜁니다: %s",
                    path, exc_info=True)
        return None


def _write_patches_to_file(
    path_base: str,
    patches: dict[str, str],
    insert_rows: list[list[tuple]] | None = None,
    delete_row_nums: set[int] | None = None,
    delete_col_letters: set[str] | None = None,
    sheet_name=None,
    src_path=None,
    src_sheet_name=None,
    patch_style_src=None,
) -> None:
    """
    patches            : {cell_ref: value}          — 기존 셀 덮어쓰기
    insert_rows        : [[(col_idx, value[, src_ref]), ...]]  — 파일 끝에 새 행 추가
    delete_row_nums    : {1-based row number}        — 해당 <row> 요소 자체 삭제
    delete_col_letters : {'A', 'B', ...}             — 해당 열의 모든 <c> 삭제
    src_path           : 서식을 읽어올 소스 xlsx 경로 (None이면 서식 병합 안 함)
    src_sheet_name     : 소스 시트 이름
    patch_style_src    : {target_ref: source_ref}    — 덮어쓰기 셀의 소스 서식 좌표
    """
    if _is_file_locked(path_base):
        # 이유를 단정하지 않는다 — 열려 있을 수도, 읽기 전용일 수도 있다.
        raise PermissionError(f"파일을 저장할 수 없습니다.\n{path_base}")

    tmp = path_base + ".tmp_merge"
    try:
        with zipfile.ZipFile(path_base, "r") as zin, \
             zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            sheet_path = _resolve_sheet_path(zin, sheet_name)
            # 시트를 해결하지 못하면 패치가 조용히 유실된다 — 데이터 손실 대신 오류로 노출.
            if sheet_path not in zin.namelist():
                raise ValueError(f"워크시트를 찾을 수 없습니다: {sheet_path}")

            # 시트 이름은 **루프 밖에서** 한 번만 읽는다. 아래 루프 안에서 읽으면
            # zipfile 이 'Bad magic number for file header' 로 깨진다(멤버를 훑는
            # 도중에 다른 멤버를 읽는 탓). 조용히 None 이 되어 자기 시트를 이름으로
            # 가리키는 수식이 통째로 안 고쳐졌다 — 정답지 대조로 잡았다.
            self_sheet = _sheet_name_for_path(zin, sheet_path)

            # 서식 병합 pre-pass — 실패해도 값 병합은 그대로 진행 (내부 try/except).
            patch_styles, new_styles_bytes, insert_rows = _prepare_style_merge(
                zin, sheet_path, src_path, src_sheet_name,
                patch_style_src, insert_rows)

            # 뺄 부품을 **먼저** 정한다 — 참조를 걷어내려면 쓰기 전에 알아야 한다.
            # ([Content_Types].xml 은 보통 zip 맨 앞이라 calcChain 보다 먼저 지나간다.)
            dropped = {n for n in zin.namelist() if n in _DROP_PARTS}

            for item in zin.infolist():
                # calcChain.xml 항상 제거 — 수식 패치/삽입/삭제 모두 계산 체인을 무효화함
                if item.filename in dropped:
                    continue
                data = zin.read(item.filename)
                if item.filename == "xl/styles.xml" and new_styles_bytes is not None:
                    data = new_styles_bytes
                elif item.filename == sheet_path:
                    data = _patch_sheet_xml(
                        data, patches,
                        insert_rows or [], delete_row_nums or set(),
                        delete_col_letters or set(),
                        patch_styles,
                        self_sheet,
                    )
                # 뺀 부품을 가리키는 선언·관계를 걷어낸다 — 남기면 Excel 복구 창이 뜬다.
                data = _strip_dangling_refs(item.filename, data, dropped)
                zout.writestr(item, data)
        # 바꾸기 **전에** 확인한다 — 임시 파일을 비교와 같은 로더로 되읽어, 화면이
        # 약속한 결과 전체와 대조한다. 백업(.bak)은 만들지 않으므로(사용자 요청) 잘못 쓴
        # 파일로 원본을 덮으면 되돌릴 수단이 없다. 검증이 실패하면 임시 파일만 버리고
        # 원본은 손대지 않는다.
        _verify_saved(tmp, _before_values(path_base, sheet_name), patches,
                      insert_rows, delete_row_nums, delete_col_letters,
                      sheet_name, path_base, self_sheet)
        # 임시 파일을 원자적으로 교체(백업 .bak 은 만들지 않음 — 사용자 요청으로 제거).
        os.replace(tmp, path_base)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def _find_sheet_path_by_name(zin: zipfile.ZipFile, sheet_name: str):
    """워크북에서 sheet_name에 해당하는 워크시트 XML 경로를 해석. 못 찾으면 None.

    해석은 ooxml 모듈(단일 출처)에 위임한다 — 읽기(loaders)와 쓰기가 서로 다른 구현을
    갖고 있으면 '읽은 시트'와 '쓴 시트'가 갈라져 엉뚱한 시트에 저장될 수 있다.
    쓰기 경로는 worksheet 타입 rel 만 인정한다(엄격).
    """
    return ooxml.sheet_path_by_name(zin, sheet_name, require_worksheet_type=True)


def _resolve_sheet_path(zin: zipfile.ZipFile, sheet_name) -> str:
    """저장/서식병합 대상 시트 XML 경로 결정.
    이름이 지정됐는데 해당 워크북에 그 시트가 없으면 **activeTab로 폴백하지 않고 raise** —
    폴백하면 엉뚱한 시트에 패치가 쓰여 데이터가 손상되기 때문(예: A에만 있는 시트를 B에 저장).
    이름 미지정(단일 시트 등)일 때만 activeTab를 사용한다."""
    if sheet_name:
        p = _find_sheet_path_by_name(zin, sheet_name)
        if p is None:
            raise ValueError(f"이 파일에는 '{sheet_name}' 시트가 없어 저장할 수 없습니다.")
        return p
    return _find_active_sheet_path(zin)


def _find_active_sheet_path(zin: zipfile.ZipFile) -> str:
    """시트 이름 미지정(단일 시트 등)일 때 쓸 activeTab 시트 경로. 실패 시 관례적 첫 시트."""
    return ooxml.active_sheet_path(zin) or ooxml.DEFAULT_SHEET_PATH


# ── 서식(cell style) 크로스-파일 병합 ────────────────────────────────────────
# styleSheet 자식 스키마 순서 (신설 컨테이너 삽입 위치 결정용)
_STYLE_CHILD_ORDER = [
    "numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
    "cellStyles", "dxfs", "tableStyles", "colors", "extLst",
]


def _canon(el) -> bytes:
    """요소를 정규화(c14n)해 dedup 키로 사용 — 속성 순서/공백 차이에도 안정."""
    return etree.tostring(el, method="c14n")


def _child_container(root, name):
    """styleSheet 직계 자식 컨테이너를 로컬명으로 조회 (없으면 None)."""
    if root is None:
        return None
    return root.find(f"{{{_NS}}}{name}")


def _ensure_container(root, name):
    """컨테이너를 조회하고 없으면 스키마 순서를 지켜 생성해 반환."""
    el = _child_container(root, name)
    if el is not None:
        return el
    el = etree.Element(f"{{{_NS}}}{name}")
    order = _STYLE_CHILD_ORDER
    my_pos = order.index(name) if name in order else len(order)
    # 나보다 스키마상 뒤에 오는 첫 형제 앞에 삽입
    insert_before = None
    for child in root:
        tag = etree.QName(child).localname
        pos = order.index(tag) if tag in order else len(order)
        if pos > my_pos:
            insert_before = child
            break
    if insert_before is not None:
        insert_before.addprevious(el)
    else:
        root.append(el)
    return el


def _read_source_cell_styles(src_zip, sheet_name) -> dict:
    """소스 시트의 <c r= s=> 만 스캔해 {cell_ref: style_index} 반환."""
    sheet_path = _resolve_sheet_path(src_zip, sheet_name)
    if sheet_path not in src_zip.namelist():
        return {}
    root = etree.fromstring(src_zip.read(sheet_path))
    sd = root.find(f"{{{_NS}}}sheetData")
    out: dict[str, int] = {}
    if sd is not None:
        for row_el in sd:
            for c_el in row_el:
                ref = c_el.get("r")
                s = c_el.get("s")
                if ref and s is not None:
                    try:
                        out[ref] = int(s)
                    except ValueError:
                        pass
    return out


class _StyleMerger:
    """소스 styles.xml의 셀 스타일을 대상 styles.xml에 병합하고
    소스 s 인덱스 → 대상 s 인덱스 매핑을 제공한다.
    동일 요소는 c14n 동등성으로 dedup, 소스 인덱스별 캐시로 재작업 방지.
    theme/indexed 색상은 원문 그대로 복사(재매핑하지 않음 — 알려진 한계)."""

    def __init__(self, src_root, dst_root):
        self.modified = False
        self.src_fonts = _child_container(src_root, "fonts")
        self.src_fills = _child_container(src_root, "fills")
        self.src_borders = _child_container(src_root, "borders")
        self.src_cellxfs = _child_container(src_root, "cellXfs")
        self.src_csxfs = _child_container(src_root, "cellStyleXfs")

        self.dst_root = dst_root
        self.dst_fonts = _ensure_container(dst_root, "fonts")
        self.dst_fills = _ensure_container(dst_root, "fills")
        self.dst_borders = _ensure_container(dst_root, "borders")
        self.dst_cellxfs = _ensure_container(dst_root, "cellXfs")
        self.dst_csxfs = _ensure_container(dst_root, "cellStyleXfs")

        # dst dedup 인덱스 (canon → index)
        self._idx_fonts = self._build_idx(self.dst_fonts)
        self._idx_fills = self._build_idx(self.dst_fills)
        self._idx_borders = self._build_idx(self.dst_borders)
        self._idx_cellxfs = self._build_idx(self.dst_cellxfs)
        self._idx_csxfs = self._build_idx(self.dst_csxfs)

        # numFmt: 소스 id→code, 대상 code→id + 사용중 id 집합
        self.src_numfmt_code = self._read_numfmts(src_root)
        self.dst_numfmt_by_code, self.dst_numfmt_ids = self._read_dst_numfmts(dst_root)

        # 소스 인덱스별 캐시
        self._cache_xf: dict = {}
        self._cache_font: dict = {}
        self._cache_fill: dict = {}
        self._cache_border: dict = {}
        self._cache_csxf: dict = {}
        self._cache_numfmt: dict = {}

    @staticmethod
    def _build_idx(container) -> dict:
        idx: dict[bytes, int] = {}
        if container is not None:
            for i, el in enumerate(container):
                idx.setdefault(_canon(el), i)
        return idx

    @staticmethod
    def _read_numfmts(root) -> dict:
        out: dict[int, str] = {}
        nf = _child_container(root, "numFmts")
        if nf is not None:
            for el in nf:
                try:
                    out[int(el.get("numFmtId"))] = el.get("formatCode", "")
                except (TypeError, ValueError):
                    pass
        return out

    @staticmethod
    def _read_dst_numfmts(root):
        by_code: dict[str, int] = {}
        ids: set[int] = set()
        nf = _child_container(root, "numFmts")
        if nf is not None:
            for el in nf:
                try:
                    fid = int(el.get("numFmtId"))
                except (TypeError, ValueError):
                    continue
                ids.add(fid)
                by_code.setdefault(el.get("formatCode", ""), fid)
        return by_code, ids

    def _set_count(self, container):
        container.set("count", str(len(container)))

    def _dedup_append(self, el, dst_container, idx_map, cache, cache_key):
        """el 을 dst_container 에 c14n-dedup 삽입하고, 그 인덱스를 cache[cache_key]에
        기록·반환한다. 세 remap 메서드(_map_container/_map_style_xf/map_index)의 공통 꼬리."""
        key = _canon(el)
        i = idx_map.get(key)
        if i is None:
            i = len(dst_container)
            dst_container.append(el)
            idx_map[key] = i
            self._set_count(dst_container)
            self.modified = True
        cache[cache_key] = i
        return i

    def _map_container(self, src_container, dst_container, idx_map, cache, src_idx):
        """fonts/fills/borders 공통 remap — dst 인덱스 반환."""
        if src_idx in cache:
            return cache[src_idx]
        if src_container is None or src_idx < 0 or src_idx >= len(src_container):
            cache[src_idx] = 0
            return 0
        el = deepcopy(src_container[src_idx])
        return self._dedup_append(el, dst_container, idx_map, cache, src_idx)

    def _map_numfmt(self, src_id: int) -> int:
        if src_id in self._cache_numfmt:
            return self._cache_numfmt[src_id]
        code = self.src_numfmt_code.get(src_id)
        if code is None:
            self._cache_numfmt[src_id] = 0
            return 0
        nid = self.dst_numfmt_by_code.get(code)
        if nid is None:
            nid = max(163, *self.dst_numfmt_ids) + 1 if self.dst_numfmt_ids else 164
            nf = _ensure_container(self.dst_root, "numFmts")
            el = etree.SubElement(nf, f"{{{_NS}}}numFmt")
            el.set("numFmtId", str(nid))
            el.set("formatCode", code)
            self._set_count(nf)
            self.dst_numfmt_by_code[code] = nid
            self.dst_numfmt_ids.add(nid)
            self.modified = True
        self._cache_numfmt[src_id] = nid
        return nid

    def _remap_xf_ids(self, xf):
        """xf 요소(deepcopy본)의 numFmtId/fontId/fillId/borderId를 대상 인덱스로 remap."""
        nid = xf.get("numFmtId")
        if nid is not None:
            try:
                if int(nid) >= 164:
                    xf.set("numFmtId", str(self._map_numfmt(int(nid))))
            except ValueError:
                pass
        for attr, src_c, dst_c, idx_map, cache in (
            ("fontId", self.src_fonts, self.dst_fonts, self._idx_fonts, self._cache_font),
            ("fillId", self.src_fills, self.dst_fills, self._idx_fills, self._cache_fill),
            ("borderId", self.src_borders, self.dst_borders, self._idx_borders, self._cache_border),
        ):
            v = xf.get(attr)
            if v is not None:
                try:
                    xf.set(attr, str(self._map_container(src_c, dst_c, idx_map, cache, int(v))))
                except ValueError:
                    pass

    def _map_style_xf(self, src_idx: int) -> int:
        """cellStyleXfs 항목 remap — dst cellStyleXfs 인덱스 반환."""
        if src_idx in self._cache_csxf:
            return self._cache_csxf[src_idx]
        if self.src_csxfs is None or src_idx < 0 or src_idx >= len(self.src_csxfs):
            self._cache_csxf[src_idx] = 0
            return 0
        xf = deepcopy(self.src_csxfs[src_idx])
        self._remap_xf_ids(xf)
        return self._dedup_append(
            xf, self.dst_csxfs, self._idx_csxfs, self._cache_csxf, src_idx)

    def map_index(self, src_s):
        """소스 <c s> 인덱스 → 대상 cellXfs 인덱스. 해결 불가 시 None."""
        if src_s is None:
            return None
        if src_s in self._cache_xf:
            return self._cache_xf[src_s]
        if self.src_cellxfs is None or src_s < 0 or src_s >= len(self.src_cellxfs):
            return None
        xf = deepcopy(self.src_cellxfs[src_s])
        self._remap_xf_ids(xf)
        xfid = xf.get("xfId")
        if xfid is not None:
            try:
                xf.set("xfId", str(self._map_style_xf(int(xfid))))
            except ValueError:
                xf.set("xfId", "0")
        return self._dedup_append(
            xf, self.dst_cellxfs, self._idx_cellxfs, self._cache_xf, src_s)


def _prepare_style_merge(zin, sheet_path, src_path, src_sheet_name,
                         patch_style_src, insert_rows):
    """저장 전 스타일 pre-pass. 반환: (patch_styles|None, new_styles_bytes|None, insert_rows).
    insert_rows는 (col, val, src_ref) 3원소를 (col, val, dst_s|None)로 변환한 결과.
    실패/미적용 시 patch_styles/new_styles_bytes=None, insert는 (col, val) 2원소로 스트립."""
    def _strip(rows):
        if not rows:
            return rows
        return [[(c[0], c[1]) for c in cells] for cells in rows]

    try:
        has_insert_src = bool(insert_rows) and any(
            len(c) > 2 and c[2] is not None for cells in insert_rows for c in cells)
        if not src_path or not (patch_style_src or has_insert_src):
            return None, None, _strip(insert_rows)
        if "xl/styles.xml" not in zin.namelist():
            return None, None, _strip(insert_rows)

        dst_styles_root = etree.fromstring(zin.read("xl/styles.xml"))
        with zipfile.ZipFile(src_path, "r") as src_zin:
            if "xl/styles.xml" not in src_zin.namelist():
                return None, None, _strip(insert_rows)
            src_styles_root = etree.fromstring(src_zin.read("xl/styles.xml"))
            src_cell_s = _read_source_cell_styles(src_zin, src_sheet_name)

        merger = _StyleMerger(src_styles_root, dst_styles_root)

        patch_styles: dict[str, int] = {}
        for tref, sref in (patch_style_src or {}).items():
            s = src_cell_s.get(sref)
            if s is not None:
                di = merger.map_index(s)
                if di is not None:
                    patch_styles[tref] = di

        new_inserts = []
        for cells in (insert_rows or []):
            row2 = []
            for cell in cells:
                col_idx, val = cell[0], cell[1]
                sref = cell[2] if len(cell) > 2 else None
                dst_s = None
                if sref is not None:
                    s = src_cell_s.get(sref)
                    if s is not None:
                        dst_s = merger.map_index(s)
                row2.append((col_idx, val, dst_s))
            new_inserts.append(row2)

        new_styles_bytes = None
        if merger.modified:
            new_styles_bytes = etree.tostring(
                dst_styles_root, xml_declaration=True, encoding="UTF-8", standalone=True)

        return (patch_styles or None), new_styles_bytes, (
            new_inserts if insert_rows else insert_rows)
    except Exception:
        # 서식 병합 실패 → 값만 병합(서식 없이). 사용자에겐 조용하지만 진단 로그로 남긴다.
        log.warning("서식 병합 pre-pass 실패(값만 병합): src=%s", src_path, exc_info=True)
        return None, None, _strip(insert_rows)


def _is_numeric(val: str) -> bool:
    """숫자 셀로 써도 **보이는 글자가 한 자도 바뀌지 않는** 값만 True.

    예전에는 float() 로 파싱만 되면 숫자로 썼다. 그래서 '01'·'007'·'210000\\t'·'nan'·
    '1_0' 같은 값이 숫자 셀이 돼, 저장하면 값이 조용히 바뀌었다 —
    실측: 실제 빌드 파일 사본에 '01' 을 그대로 저장하면 파일에는 '1' 이 남았다.

    그래서 '숫자로 썼다가 우리 로더로 되읽으면 원문 그대로인가'를 기준으로 삼는다.
    되읽기 표기는 loaders._cell_to_str 과 같은 규칙(정수형 float 은 정수로)이며,
    두 규칙이 어긋나지 않도록 테스트로 묶어 두었다.

    (실측 영향 범위: 실제 빌드 데이터 60개 파일의 숫자형 셀 193,772개 중 이 규칙으로
     문자 셀이 되는 것은 27개뿐이다 — 앞자리 0 23개, 끝에 탭이 붙은 4개. 둘 다 예전엔
     저장하면서 값이 뭉개지던 셀이다.)
    """
    if not isinstance(val, str):
        return False
    try:
        f = float(val)
    except (ValueError, TypeError, OverflowError):
        return False
    try:
        back = str(int(f)) if f == int(f) else str(f)
    except (ValueError, OverflowError):     # nan / inf — 숫자로 쓰면 값이 사라진다
        return False
    return back == val


def _set_cell_value(c_el, new_val: str):
    """셀 <c> 요소의 값을 설정 — 기존 자식 제거 후 수식(<f>)/숫자·문자(<v>)로 재작성."""
    for child in list(c_el):
        c_el.remove(child)
    c_el.attrib.pop("t", None)
    if new_val == "":
        return
    if new_val.startswith("="):
        f_el = etree.SubElement(c_el, _TAG_F)
        f_el.text = new_val[1:]   # '=' 제외한 수식 본문
    elif _is_numeric(new_val):
        v_el = etree.SubElement(c_el, _TAG_V)
        v_el.text = new_val
    else:
        # t="str": sharedStrings.xml 변경 없이 Excel이 안전하게 수용하는 문자열 타입
        c_el.set("t", "str")
        v_el = etree.SubElement(c_el, _TAG_V)
        v_el.text = new_val


# 수식 안의 셀 참조를 알아보기 위한 조각들.
#   접두  : [1]Sheet!  /  'My Sheet'!  /  Sheet1!     (외부 통합문서·시트 지정)
#   A1    : B9  $B$9  B$9  $B9
#   행참조: 9:9  $9:$9                                 (행 전체)
# 시트 이름은 ASCII 가 아닐 수 있다 — [^\W\d] 는 '숫자가 아닌 낱말 글자'(유니코드
# 글자·밑줄)다. 이게 [A-Za-z_] 였을 때 [7]메뉴얼!$L$8:$M$15 를 외부 참조로 못 알아보고
# 그 범위를 조정해 버렸다(정답지 대조에서 잡힘).
_F_PREFIX = r"(?:\[[^\]]*\])?(?:'(?:[^']|'')*'|[^\W\d][\w.]*)[ ]*!"
_F_A1 = r"\$?[A-Za-z]{1,3}\$?\d+"
_F_ROW = r"\$?\d+"
_F_TOKEN = re.compile(
    rf"(?P<prefix>{_F_PREFIX})?"
    rf"(?P<ref>{_F_A1}:{_F_A1}|{_F_A1}|{_F_ROW}:{_F_ROW})")
_F_A1_PARTS = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_F_ROW_PARTS = re.compile(r"^(\$?)(\d+)$")
REF_ERROR = "#REF!"


def _sheet_name_for_path(zin, sheet_path: str):
    """그 워크시트 XML 경로의 **시트 이름**. 못 찾으면 None.

    수식이 자기 시트를 이름으로 가리키는 경우가 있다(실측: 18개 파일 85,429건,
    `VLOOKUP(E2,TextUITable!$D:$E,2,0)` 같은 꼴). 그런 참조도 행 삭제 때 조정돼야
    하므로, 지금 고치는 시트가 어떤 이름인지 알아야 한다.
    """
    try:
        # Target 은 상대('worksheets/sheet1.xml')일 수도 절대('/xl/worksheets/...')일
        # 수도 있다 — 엑셀은 상대, openpyxl 은 절대로 쓴다. 둘 다 풀어 주는 기존
        # 헬퍼를 그대로 쓴다.
        rp = "xl/_rels/workbook.xml.rels"
        rels = {r.get("Id"): _rel_target_part(rp, r)
                for r in etree.fromstring(zin.read(rp))
                if isinstance(r.tag, str) and r.tag.endswith("Relationship")}
        for sh in etree.fromstring(zin.read("xl/workbook.xml")).iter(f"{{{_NS}}}sheet"):
            rid = next((v for k, v in sh.attrib.items() if k.endswith("}id")), None)
            if rels.get(rid) == sheet_path:
                return sh.get("name")
    except Exception:
        log.warning("시트 이름을 찾지 못했습니다: %s", sheet_path, exc_info=True)
    return None


def _shift_row_num(row: int, deleted_sorted: list, deleted_set: set):
    """삭제 뒤 그 행의 새 번호. 그 행이 지워졌으면 None."""
    if row in deleted_set:
        return None
    return row - bisect.bisect_left(deleted_sorted, row)


def _shift_span(r1: int, r2: int, deleted_sorted: list, deleted_set: set):
    """범위 [r1, r2] 의 삭제 뒤 범위. 하나도 안 남으면 None.

    **끝점에 규칙을 적용하는 게 아니다.** 양 끝이 모두 지워져도 가운데가 남으면 범위는
    살아 있다(실측: D={5,8} 일 때 B5:B8 → B5:B6 — 남은 6,7 이 새 5,6 이 된다).
    """
    lo, hi = (r1, r2) if r1 <= r2 else (r2, r1)
    while lo <= hi and lo in deleted_set:
        lo += 1
    while hi >= lo and hi in deleted_set:
        hi -= 1
    if lo > hi:
        return None
    return (_shift_row_num(lo, deleted_sorted, deleted_set),
            _shift_row_num(hi, deleted_sorted, deleted_set))


def _prefix_targets_self(prefix: str, self_sheet) -> bool:
    """그 접두가 **지금 고치는 시트**를 가리키나.

    외부 통합문서([1]X! 꼴)는 행을 지워도 바뀌지 않는다 — 다른 파일이니까. 같은 파일의
    다른 시트도 마찬가지다. 자기 시트를 이름으로 쓴 것만 조정 대상이다.
    """
    if prefix is None:
        return True                         # 이름 없음 = 자기 시트
    if "[" in prefix:
        return False                        # 외부 통합문서
    name = prefix.rstrip().rstrip("!").rstrip()
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    return bool(self_sheet) and name == self_sheet


def _rewrite_one_ref(ref: str, deleted_sorted: list, deleted_set: set):
    """A1/범위/행참조 하나를 다시 쓴다. 통째로 사라지면 None."""
    if ":" in ref:
        a, b = ref.split(":", 1)
        ma, mb = _F_A1_PARTS.match(a), _F_A1_PARTS.match(b)
        if ma and mb:
            span = _shift_span(int(ma.group(4)), int(mb.group(4)),
                               deleted_sorted, deleted_set)
            if span is None:
                return None
            return (f"{ma.group(1)}{ma.group(2)}{ma.group(3)}{span[0]}:"
                    f"{mb.group(1)}{mb.group(2)}{mb.group(3)}{span[1]}")
        ra, rb = _F_ROW_PARTS.match(a), _F_ROW_PARTS.match(b)
        if ra and rb:
            span = _shift_span(int(ra.group(2)), int(rb.group(2)),
                               deleted_sorted, deleted_set)
            if span is None:
                return None
            return f"{ra.group(1)}{span[0]}:{rb.group(1)}{span[1]}"
        return ref
    m = _F_A1_PARTS.match(ref)
    if not m:
        return ref
    new = _shift_row_num(int(m.group(4)), deleted_sorted, deleted_set)
    if new is None:
        return None
    return f"{m.group(1)}{m.group(2)}{m.group(3)}{new}"


def _rewrite_formula_rows(formula: str, deleted, self_sheet=None) -> str:
    """행을 지운 뒤의 수식 본문. 엑셀이 행 삭제 때 하는 일을 그대로 한다.

    규칙은 추측이 아니라 **엑셀에게 물어서** 정했다(탐침 시트 52개 사례).

      · 지워진 행을 가리키면 그 **참조 토큰 전체**가 #REF! 가 된다. 시트 접두는 남는다
        (Sheet1!B5 → Sheet1!#REF!).
      · 그 외에는 자기보다 위에서 지워진 행 수만큼 당긴다. 절대 표시($)는 유지되지만
        행 번호는 **절대여도 조정된다**($B$9 → $B$7).
      · 범위는 끝점이 아니라 **살아남은 행**으로 다시 잡는다(_shift_span 참조).
      · 문자열 리터럴 안의 좌표, 다른 시트·외부 통합문서 참조, 열 전체 범위($A:$W),
        ROW() 류는 건드리지 않는다.
    """
    if not formula or not deleted:
        return formula
    deleted_set = set(deleted)
    deleted_sorted = sorted(deleted_set)
    out = []
    for tok in _formula_tokens(formula):
        if tok[0] == "text":
            out.append(tok[1])
            continue
        _k, prefix, ref = tok
        if _prefix_targets_self(prefix, self_sheet):
            new = _rewrite_one_ref(ref, deleted_sorted, deleted_set)
            ref = REF_ERROR if new is None else new
        out.append((prefix or "") + ref)
    return "".join(out)


def _formula_tokens(formula: str):
    """수식을 (종류, 조각) 으로 훑는다 — ('text', 글자들) 또는 ('ref', 접두, 참조).

    재작성과 순환 검사가 **같은 눈**으로 봐야 한다. 둘이 따로 훑으면 한쪽만 고쳐져
    어긋난다(실제로 순환 검사를 따로 짰다가 '자기 행'과 '자기 셀'을 혼동했다).
    문자열 리터럴은 통째로 'text' 로 넘긴다 — 그 안의 좌표는 참조가 아니다.
    """
    i, n = 0, len(formula or "")
    buf = []
    while i < n:
        ch = formula[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':
                        j += 2
                        continue
                    j += 1
                    break
                j += 1
            buf.append(formula[i:j])
            i = j
            continue
        if i == 0 or not (formula[i - 1].isalnum() or formula[i - 1] in "_.$!"):
            m = _F_TOKEN.match(formula, i)
            if m:
                if buf:
                    yield ("text", "".join(buf))
                    buf = []
                yield ("ref", m.group("prefix"), m.group("ref"))
                i = m.end()
                continue
        buf.append(ch)
        i += 1
    if buf:
        yield ("text", "".join(buf))


def _ref_covers(ref: str, cell: str) -> bool:
    """그 참조가 cell 을 품는가 — 단일 셀·범위·행 전체 모두."""
    p = _ref_parts(cell)
    if p is None:
        return False
    col, row = column_index_from_string(p[0]), p[1]
    parts = ref.split(":")
    a, b = _F_A1_PARTS.match(parts[0]), _F_A1_PARTS.match(parts[-1])
    if a and b:
        c1, c2 = (column_index_from_string(a.group(2)),
                  column_index_from_string(b.group(2)))
        r1, r2 = int(a.group(4)), int(b.group(4))
        return (min(c1, c2) <= col <= max(c1, c2)
                and min(r1, r2) <= row <= max(r1, r2))
    ra, rb = _F_ROW_PARTS.match(parts[0]), _F_ROW_PARTS.match(parts[-1])
    if ra and rb:                            # 행 전체 참조 — 열은 가리지 않는다
        r1, r2 = int(ra.group(2)), int(rb.group(2))
        return min(r1, r2) <= row <= max(r1, r2)
    return False


def _has_self_row_ref(formula: str, self_sheet=None) -> bool:
    """그 수식이 **자기 시트의 행**을 가리키나 — 행 삭제에 영향을 받는지의 기준.

    ROW()-1, 외부 통합문서 참조([1]X!$D:$E), 열 전체 범위($A:$W) 뿐인 수식은 행을
    지워도 그대로다. 실제 데이터의 큰 그룹들이 여기 해당해서(TextUITable 의 ROW()-1 은
    추종자가 5천 개가 넘는다) 먼저 걸러 내면 쓸데없는 일을 크게 줄인다.
    """
    if not formula:
        return False
    probe = {1, 2}
    return _rewrite_formula_rows(formula, probe, self_sheet) != formula


def _translate(formula: str, origin: str, dest: str) -> str:
    """공유 수식 본문을 origin 자리에서 dest 자리로 옮겨 쓴다. 실패하면 None."""
    try:
        from openpyxl.formula.translate import Translator
        out = Translator("=" + formula, origin=origin).translate_formula(dest)
    except Exception:
        return None
    return out[1:] if out.startswith("=") else out


def _rewrite_rows_in_formulas(sheetdata, deleted, self_sheet, orig_of) -> None:
    """행을 지운 뒤 수식 본문의 셀 참조를 다시 쓴다. _renumber_after_delete **뒤에**.

    엑셀은 행을 지우면 수식도 함께 고쳐 준다. 우리는 좌표(<c r>)만 당기고 본문은
    그대로 뒀다 — 그래서 'B3: =B2+1' 이 B2 로 당겨지면 '=B2+1' 그대로라 **순환 참조**가
    됐다. 실측으로 범위를 재 보니 367쌍 중 6개 파일·935셀이 이렇게 틀어진다.

    공유 수식은 엑셀이 하는 대로 한다 — **그룹은 유지하되 깨진 셀만 떼어낸다.**
    추종자는 주인에서 자기 자리에 맞게 유도되므로 대부분 저절로 맞는다. 유도 결과가
    정답과 다른 칸(보통 지워진 행을 직접 가리키던 한 칸)만 일반 수식으로 분리한다.
    """
    if not deleted:
        return
    new_of = {}
    groups = {}
    for row_el in sheetdata:
        for c_el in row_el:
            f_el = c_el.find(_TAG_F)
            if f_el is None:
                continue
            new_of[c_el] = c_el.get("r", "")
            if f_el.get("t") == "shared" and f_el.get("si"):
                groups.setdefault(f_el.get("si"), []).append(c_el)

    done = set()
    for si, cells in groups.items():
        master = next((c for c in cells if c.find(_TAG_F).get("ref")), None)
        if master is None:
            continue                         # _repair_shared_formulas 가 처리한다
        body = (master.find(_TAG_F).text or "").strip()
        m_orig, m_new = orig_of.get(master, new_of[master]), new_of[master]
        if not _has_self_row_ref(body, self_sheet):
            done.update(cells)               # 자리에 무관한 식(ROW() 등) — 손댈 것 없다
            continue
        new_body = _rewrite_formula_rows(body, deleted, self_sheet)
        master.find(_TAG_F).text = new_body
        done.add(master)
        for c_el in cells:
            if c_el is master:
                continue
            f_orig = _translate(body, m_orig, orig_of.get(c_el, new_of[c_el]))
            want = (_rewrite_formula_rows(f_orig, deleted, self_sheet)
                    if f_orig is not None else None)
            derived = _translate(new_body, m_new, new_of[c_el])
            if want is not None and derived is not None and want != derived:
                f_el = c_el.find(_TAG_F)     # 유도로는 못 맞춘다 — 떼어낸다
                f_el.attrib.pop("t", None)
                f_el.attrib.pop("si", None)
                f_el.attrib.pop("ref", None)
                f_el.text = want
            done.add(c_el)

    for c_el, _ref in new_of.items():
        if c_el in done:
            continue
        f_el = c_el.find(_TAG_F)
        if f_el is None:
            continue
        if f_el.get("t") == "shared" and not f_el.get("ref"):
            continue                         # 추종자는 본문이 없다
        body = (f_el.text or "").strip()
        if body:
            f_el.text = _rewrite_formula_rows(body, deleted, self_sheet)


def _ref_parts(ref: str):
    """'B12' → ('B', 12). 좌표 형식이 아니면 None."""
    m = _COL_RE.match(ref or "")
    return (m.group(1), int(m.group(2))) if m else None


def _bounding_ref(refs) -> str:
    """좌표들을 모두 담는 최소 직사각형 — 'A3' 또는 'A3:A66'."""
    cols, rows = [], []
    for r in refs:
        p = _ref_parts(r)
        if p:
            cols.append(column_index_from_string(p[0]))
            rows.append(p[1])
    if not cols:
        return ""
    first = f"{get_column_letter(min(cols))}{min(rows)}"
    last = f"{get_column_letter(max(cols))}{max(rows)}"
    return first if first == last else f"{first}:{last}"


def _ref_in(ref: str, rng: str) -> bool:
    """ref 가 'A3:A66' 같은 범위 안에 드는가. 판단할 수 없으면 True(문제 삼지 않는다)."""
    parts = (rng or "").split(":")
    a, b, p = _ref_parts(parts[0]), _ref_parts(parts[-1]), _ref_parts(ref)
    if not (a and b and p):
        return True
    ca, cb = column_index_from_string(a[0]), column_index_from_string(b[0])
    cp = column_index_from_string(p[0])
    return (min(ca, cb) <= cp <= max(ca, cb)
            and min(a[1], b[1]) <= p[1] <= max(a[1], b[1]))


def _shift_ref_rows(ref: str, sorted_deleted) -> str:
    """'H117' / 'H117:H120' 의 행 번호를 삭제분만큼 당긴다 — 셀 좌표와 같은 규칙.

    수식의 ref 는 **그 수식이 덮는 범위**다. 배열 수식은 자기 셀을, 공유 수식은 주인과
    추종자를 담는다. 행을 지우면 <c r> 은 당겨지는데 이 속성이 그대로 남아, 수식이
    자기 자리를 벗어난 범위를 가리키게 된다 — Excel 은 그런 파일을 거부한다.
    """
    out = []
    for part in (ref or "").split(":"):
        p = _ref_parts(part)
        if p is None:
            return ref                     # 모르는 모양 — 건드리지 않는다
        col, rn = p
        out.append(f"{col}{rn - sum(1 for d in sorted_deleted if d < rn)}")
    return ":".join(out)


def _cell_sort_key(c_el):
    p = _ref_parts(c_el.get("r", ""))
    return (p[1], column_index_from_string(p[0])) if p else (0, 0)


def _shared_groups(sheetdata) -> dict:
    """바꾸기 **전** 공유 수식 그룹 — si → {주인 좌표, 수식 본문, (셀, 원래 좌표)들}.

    엑셀은 같은 모양의 수식을 한 번만 적는다.

        <c r="B4"><f t="shared" ref="B4:B16" si="0">ROW()-1</f><v>3</v></c>   주인
        <c r="B5"><f t="shared" si="0"/><v>4</v></c>                          추종자

    추종자는 본문이 없다. 주인을 보고 자기 자리에 맞게 옮겨 쓴다. 그래서 **주인이
    사라지면 추종자는 읽을 수 없는 셀**이 되고 Excel 이 파일을 거부한다.

    바꾸기 **전에** 잡아 두는 이유: 주인 셀을 덮어쓰거나 그 행을 지우고 나면 수식
    본문을 되살릴 방법이 없다. 원래 좌표도 함께 남긴다 — 행을 지우면 좌표가 당겨지는데,
    수식을 옮겨 쓸 때는 **당기기 전 좌표**를 기준으로 해야 본문이 그대로 나온다.
    """
    groups: dict = {}
    for row_el in sheetdata:
        for c_el in row_el:
            f_el = c_el.find(_TAG_F)
            if f_el is None or f_el.get("t") != "shared":
                continue
            si = f_el.get("si")
            if si is None:
                continue
            g = groups.setdefault(si, {"origin": None, "formula": None, "cells": []})
            g["cells"].append((c_el, c_el.get("r", "")))
            if f_el.get("ref"):
                g["origin"] = c_el.get("r", "")
                g["formula"] = (f_el.text or "").strip()
    return groups


def _drop_shared_formula(c_el) -> None:
    """<f> 를 떼고 캐시 값(<v>)만 남긴다 — 수식은 잃지만 셀 값과 파일은 지킨다."""
    f_el = c_el.find(_TAG_F)
    if f_el is not None:
        c_el.remove(f_el)


def _make_plain_formula(c_el, body: str) -> None:
    """공유 수식을 평범한 수식으로 되돌린다(혼자 남았을 때)."""
    f_el = c_el.find(_TAG_F)
    if f_el is None:
        return
    f_el.attrib.pop("t", None)
    f_el.attrib.pop("si", None)
    f_el.attrib.pop("ref", None)
    f_el.text = body


def _promote_shared_master(cells, group, orig_of):
    """남은 셀 하나를 새 주인으로 올린다. 못 하면 None.

    수식 본문은 주인 자리 기준으로 적혀 있으므로, 새 주인 자리로 **옮겨 써야** 한다
    (B2 를 가리키던 상대 참조는 다섯 칸 아래에서 B7 이 된다). 좌표는 둘 다 '당기기 전'
    것을 쓴다 — 그래야 원래 파일에서 그 셀이 가졌을 수식과 같아진다.
    """
    if not group.get("formula") or not group.get("origin"):
        return None
    target = min(cells, key=_cell_sort_key)
    dest = orig_of.get(target) or target.get("r", "")
    try:
        from openpyxl.formula.translate import Translator
        text = Translator("=" + group["formula"],
                          origin=group["origin"]).translate_formula(dest)
    except Exception:
        log.warning("공유 수식을 새 주인 자리로 옮겨 쓰지 못했습니다: %s → %s",
                    group.get("origin"), dest, exc_info=True)
        return None
    f_el = target.find(_TAG_F)
    f_el.text = text[1:] if text.startswith("=") else text
    f_el.set("ref", target.get("r", ""))
    return target


def _repair_shared_formulas(sheetdata, groups) -> None:
    """공유 수식을 Excel 이 받아들이는 상태로 되돌린다. 바꾼 **뒤에** 부른다.

    실사용에서 둘 다 터졌다(40_Build 두 파일이 열리지 않았다).

      1. **주인이 사라졌다.** 병합이 주인 셀을 값으로 덮으면 추종자가 고아가 된다
         (Data_MiniGameRoulette_CS.xlsx: si=2 추종자 3개). 남은 셀 하나를 새 주인으로
         올린다. 옮겨 쓸 수 없으면 <f> 를 떼어 **값만 남긴다** — 수식은 잃어도 파일은
         열린다.

      2. **주인의 ref 가 자기를 안 담는다.** 행을 지우면 <c r> 은 당겨지는데 ref 는
         그대로 남는다(Data_TextUITable_CS.xlsm: 주인 A5250 의 ref 가 A5251:A5314).
         ref 는 주인 자신과 살아남은 추종자를 모두 담아야 한다.

    값은 건드리지 않는다 — <v> 의 캐시 값은 그대로다.
    """
    if not groups:
        return
    orig_of = {c_el: orig for g in groups.values() for c_el, orig in g["cells"]}

    live: dict = {}
    for row_el in sheetdata:
        for c_el in row_el:
            f_el = c_el.find(_TAG_F)
            if f_el is not None and f_el.get("t") == "shared" and f_el.get("si"):
                live.setdefault(f_el.get("si"), []).append(c_el)

    for si, cells in live.items():
        group = groups.get(si, {"origin": None, "formula": None, "cells": []})
        master = next((c for c in cells if c.find(_TAG_F).get("ref")), None)
        if master is None:
            master = _promote_shared_master(cells, group, orig_of)
            if master is None:
                for c_el in cells:
                    _drop_shared_formula(c_el)
                continue
        if len(cells) == 1:
            body = (master.find(_TAG_F).text or "").strip()
            if body:
                _make_plain_formula(master, body)   # 혼자면 공유할 이유가 없다
            else:
                _drop_shared_formula(master)
            continue
        master.find(_TAG_F).set(
            "ref", _bounding_ref([c.get("r", "") for c in cells]))


def _index_sheet(sheetdata):
    """sheetData 를 (ref→<c>, 1-based row번호→<row>) 두 인덱스로 스캔."""
    existing: dict[str, etree._Element] = {}
    row_map: dict[int, etree._Element] = {}
    for row_el in sheetdata:
        row_map[int(row_el.get("r", 0))] = row_el
        for c_el in row_el:
            ref = c_el.get("r", "")
            if ref:
                existing[ref] = c_el
    return existing, row_map


def _delete_rows(sheetdata, row_map, delete_row_nums) -> set:
    """지정 행(<row>)을 제거하고 실제 삭제된 행 번호 집합을 반환."""
    deleted: set[int] = set()
    for row_num in delete_row_nums:
        row_el = row_map.get(row_num)
        if row_el is not None:
            sheetdata.remove(row_el)
            deleted.add(row_num)
    return deleted


def _apply_patches(sheetdata, existing, row_map, patches, patch_styles):
    """기존 셀 덮어쓰기(빈값이면 <c> 제거) 또는 없는 셀/행 신규 생성."""
    for ref, new_val in patches.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        row_num = int(m.group(2))
        if ref in existing:
            c_el = existing[ref]
            if new_val == "":
                parent = c_el.getparent()   # 빈값 패치 → <c> 요소 자체 제거
                if parent is not None:
                    parent.remove(c_el)
            else:
                _set_cell_value(c_el, new_val)
                if patch_styles and ref in patch_styles:
                    c_el.set("s", str(patch_styles[ref]))
            continue
        # 대상 셀 없음 — 필요하면 행부터 만들고 셀 추가(열 순서 유지)
        row_el = row_map.get(row_num)
        if row_el is None:
            row_el = etree.SubElement(sheetdata, _TAG_ROW)
            row_el.set("r", str(row_num))
            row_map[row_num] = row_el
            sheetdata[:] = sorted(sheetdata, key=lambda e: int(e.get("r", 0)))
        if new_val != "":
            c_el = etree.SubElement(row_el, _TAG_C)
            c_el.set("r", ref)
            _set_cell_value(c_el, new_val)
            if patch_styles and ref in patch_styles:
                c_el.set("s", str(patch_styles[ref]))
            row_el[:] = sorted(row_el, key=lambda e: column_index_from_string(
                _COL_RE.match(e.get("r", "A1")).group(1)))


def _row_has_value(row_el) -> bool:
    """그 행에 **보이는 값**이 있나 — <c> 가 있는 것만으로는 부족하다.

    서식만 주고 값은 없는 셀이 꼬리에 남아 있는 파일이 흔하다(실측:
    Data_TextPCSkillTable_CS.xlsm 은 값 있는 마지막 행이 4166인데 <c> 는 4560행까지
    있다). 그걸 데이터로 세면 새 행이 빈 행 394개 **아래**에 붙는다 — 비교에 쓰는
    로더가 보는 마지막 행과 같은 기준으로 세야 한다.
    """
    for c_el in row_el:
        if c_el.find(_TAG_F) is not None:
            return True
        v_el = c_el.find(_TAG_V)
        if v_el is not None and (v_el.text or "").strip():
            return True
        if c_el.find(f"{{{_NS}}}is") is not None:
            return True
    return False


def _append_rows(sheetdata, insert_rows):
    """실제 셀이 있는 마지막 행 다음부터 새 행을 추가.

    새 행은 **셀이 있는** 마지막 행 다음에 놓는다. 그런데 그 자리에 빈 <row> 요소가
    이미 있을 수 있다 — 엑셀이 서식만 준 행을 그렇게 남긴다(실측:
    Data_RewardGroup_CS.xlsx 는 셀이 있는 행이 423개인데 <row> 는 1207개다).
    예전엔 빈 행을 무시하고 새로 만들어 **같은 번호의 <row> 가 둘** 생겼고, 순서도
    뒤집혔다(…1207, 424, 425…). 엑셀은 그런 파일을 거부한다.
    그래서 그 번호에 행이 이미 있으면 **그 행에 채워 넣는다.**
    """
    if not insert_rows:
        return
    by_num = {}
    for row_el in sheetdata:
        try:
            by_num[int(row_el.get("r", 0))] = row_el
        except (TypeError, ValueError):
            continue
    last_data_row = max((n for n, el in by_num.items() if _row_has_value(el)),
                        default=0)
    next_row = last_data_row + 1 if last_data_row > 0 else 1
    for cells in insert_rows:
        row_el = by_num.get(next_row)
        if row_el is None:
            row_el = etree.SubElement(sheetdata, _TAG_ROW)
            row_el.set("r", str(next_row))
            by_num[next_row] = row_el
        # 그 행에 이미 있는 <c>(서식만 준 빈 칸)를 다시 쓴다 — 새로 만들면 같은
        # 좌표의 셀이 둘 생기고 엑셀이 파일을 거부한다.
        by_col = {}
        for c_el in row_el:
            m = _COL_RE.match(c_el.get("r", ""))
            if m:
                by_col[column_index_from_string(m.group(1)) - 1] = c_el
        for cell in sorted(cells, key=lambda x: x[0]):
            col_idx, val = cell[0], cell[1]
            dst_s = cell[2] if len(cell) > 2 else None
            if val == "":
                continue
            c_el = by_col.get(col_idx)
            if c_el is None:
                c_el = etree.SubElement(row_el, _TAG_C)
                by_col[col_idx] = c_el
            c_el.set("r", _cell_ref(next_row - 1, col_idx))
            _set_cell_value(c_el, val)
            if dst_s is not None:
                c_el.set("s", str(dst_s))
        row_el[:] = sorted(row_el, key=lambda e: column_index_from_string(
            _COL_RE.match(e.get("r", "A1")).group(1)))
        next_row += 1
    sheetdata[:] = sorted(sheetdata, key=lambda e: int(e.get("r", 0)))


def _delete_columns(sheetdata, delete_col_letters):
    """지정 열 문자에 속하는 <c> 요소를 모든 행에서 제거."""
    if not delete_col_letters:
        return
    for row_el in sheetdata:
        to_remove = [
            c_el for c_el in row_el
            if (m := _COL_RE.match(c_el.get("r", ""))) and m.group(1) in delete_col_letters
        ]
        for c_el in to_remove:
            row_el.remove(c_el)


def _renumber_after_delete(sheetdata, deleted):
    """VBA .Delete처럼, 삭제된 행 번호만큼 아래 행들을 위로 당겨 빈 행 번호를 없앤다."""
    if not deleted:
        return
    sorted_deleted = sorted(deleted)
    for row_el in sheetdata:
        rn = int(row_el.get("r", 0))
        offset = sum(1 for d in sorted_deleted if d < rn)
        if offset > 0:
            new_rn = rn - offset
            row_el.set("r", str(new_rn))
            for c_el in row_el:
                m = _COL_RE.match(c_el.get("r", ""))
                if m:
                    c_el.set("r", f"{m.group(1)}{new_rn}")
    # 수식의 ref 도 같이 당긴다. 셀만 당기고 두면 배열 수식이 자기 자리를 벗어난
    # 범위를 가리켜 Excel 이 파일을 거부한다(실측: Data_Illustration_C.xlsx 의
    # <c r="H116"><f t="array" ref="H117">). 범위 끝이 어디에 있든 같은 규칙으로 센다.
    # 공유 수식 ref 는 _repair_shared_formulas 가 실제 그룹 범위로 다시 잡는다.
    for row_el in sheetdata:
        for c_el in row_el:
            f_el = c_el.find(_TAG_F)
            if f_el is not None and f_el.get("ref"):
                f_el.set("ref", _shift_ref_rows(f_el.get("ref"), sorted_deleted))


def _patch_sheet_xml(
    data: bytes,
    patches: dict[str, str],
    insert_rows: list[list[tuple]] | None = None,
    delete_row_nums: set[int] | None = None,
    delete_col_letters: set[str] | None = None,
    patch_styles: dict[str, int] | None = None,
    self_sheet: str | None = None,
) -> bytes:
    """sheet XML 을 직접 패치. 5개 단계를 순서대로 위임한다(각 단계는 별도 헬퍼).

    patches            : {cell_ref: value}          기존 셀 덮어쓰기
    insert_rows        : [[(col_idx, value[, dst_s]), ...]]  파일 끝에 새 행 추가
    delete_row_nums    : {1-based row number}        해당 <row> 요소 자체 삭제
    delete_col_letters : {'A', 'B', ...}             해당 열의 모든 <c> 삭제
    patch_styles       : {cell_ref: style_index}     병합할 소스 서식 인덱스 (덮어쓰기 셀)
    self_sheet         : 이 시트의 이름 — 수식이 자기 시트를 이름으로 가리킬 때 필요
    """
    insert_rows = insert_rows or []
    delete_row_nums = delete_row_nums or set()
    delete_col_letters = delete_col_letters or set()

    tree = etree.fromstring(data)
    sheetdata = tree.find(_TAG_SHEETDATA)
    if sheetdata is None:
        return data

    existing, row_map = _index_sheet(sheetdata)
    # 공유 수식 그룹은 **건드리기 전에** 잡아 둔다 — 주인을 덮어쓰거나 그 행을 지우고
    # 나면 수식 본문을 되살릴 방법이 없다.
    shared = _shared_groups(sheetdata)
    deleted = _delete_rows(sheetdata, row_map, delete_row_nums)
    _apply_patches(sheetdata, existing, row_map, patches, patch_styles)
    _append_rows(sheetdata, insert_rows)
    _delete_columns(sheetdata, delete_col_letters)
    _renumber_after_delete(sheetdata, deleted)
    # 좌표가 다 정해진 뒤에 고친다 — ref 는 **당긴 뒤** 좌표로 써야 한다.
    # 구조를 먼저 되돌린다 — 주인이 지워진 그룹은 여기서 본문이 되살아나고, 그
    # 되살린 본문도 행 삭제에 맞춰 다시 써야 한다(반대 순서면 그 그룹만 옛 참조로 남는다).
    _repair_shared_formulas(sheetdata, shared)
    _rewrite_rows_in_formulas(
        sheetdata, deleted, self_sheet,
        {c: o for g in shared.values() for c, o in g["cells"]})

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
