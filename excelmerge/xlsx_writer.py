"""xlsx 저장 — sheet XML 직접 패치로 수식 보존 기록 (excel_diff_merge.py에서 분리)."""
import os
import re
import zipfile
from collections import defaultdict

from lxml import etree
from openpyxl.utils import get_column_letter, column_index_from_string


_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_COL_RE = re.compile(r"([A-Z]+)(\d+)")


def _cell_ref(r: int, c: int) -> str:
    return f"{get_column_letter(c + 1)}{r + 1}"


def _promote_empty_cols_to_delete(
    patches: dict[str, str],
    delete_row_nums: set[int],
    path: str | None,
) -> tuple[dict[str, str], set[int], set[str]]:
    """
    1단계: 패치 적용 후 모든 셀이 빈값인 행 → delete_row_nums로 승격 (patches에서 제거)
    2단계: 행 삭제 반영 후 모든 셀이 빈값인 열 → delete_col_letters 반환
    """
    if not path:
        return patches, delete_row_nums, set()

    try:
        with zipfile.ZipFile(path, "r") as zin:
            sheet_name = _find_active_sheet_path(zin)
            xml_data = zin.read(sheet_name)
        tree = etree.fromstring(xml_data)
        ns = _NS
        sheetdata = tree.find(f"{{{ns}}}sheetData")
        file_cells: dict[str, str] = {}
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
                    if f_el is not None and f_el.text:
                        val = "=" + f_el.text
                    elif v_el is not None and v_el.text:
                        val = v_el.text
                    elif is_el is not None:
                        t_el = is_el.find(f"{{{ns}}}t")
                        val  = t_el.text if (t_el is not None and t_el.text) else ""
                    else:
                        val = ""
                    file_cells[ref] = val
                    file_row_refs[rn].add(ref)
    except Exception:
        return patches, delete_row_nums, set()

    merged: dict[str, str] = {**file_cells, **patches}

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
        if any(merged.get(ref, "") != "" for ref in non_patched_refs):
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
    """파일이 다른 프로세스에 의해 열려 있는지 확인한다."""
    try:
        with open(path, "a"):
            return False
    except (IOError, PermissionError):
        return True


def _write_patches_to_file(
    path_base: str,
    patches: dict[str, str],
    insert_rows: list[list[tuple[int, str]]] | None = None,
    delete_row_nums: set[int] | None = None,
    delete_col_letters: set[str] | None = None,
) -> None:
    """
    patches            : {cell_ref: value}          — 기존 셀 덮어쓰기
    insert_rows        : [[(col_idx, value), ...]]   — 파일 끝에 새 행 추가
    delete_row_nums    : {1-based row number}        — 해당 <row> 요소 자체 삭제
    delete_col_letters : {'A', 'B', ...}             — 해당 열의 모든 <c> 삭제
    """
    if _is_file_locked(path_base):
        raise PermissionError(f"파일이 열려 있어 저장할 수 없습니다:\n{path_base}")

    tmp = path_base + ".tmp_merge"
    try:
        with zipfile.ZipFile(path_base, "r") as zin, \
             zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            sheet_name = _find_active_sheet_path(zin)
            # 시트를 해결하지 못하면 패치가 조용히 유실된다 — 데이터 손실 대신 오류로 노출.
            if sheet_name not in zin.namelist():
                raise ValueError(f"워크시트를 찾을 수 없습니다: {sheet_name}")
            for item in zin.infolist():
                # calcChain.xml 항상 제거 — 수식 패치/삽입/삭제 모두 계산 체인을 무효화함
                if item.filename == "xl/calcChain.xml":
                    continue
                data = zin.read(item.filename)
                if item.filename == sheet_name:
                    data = _patch_sheet_xml(
                        data, patches,
                        insert_rows or [], delete_row_nums or set(),
                        delete_col_letters or set(),
                    )
                zout.writestr(item, data)
        os.replace(tmp, path_base)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def _find_active_sheet_path(zin: zipfile.ZipFile) -> str:
    try:
        ns_wb = _NS
        ns_r  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        wb_xml  = zin.read("xl/workbook.xml")
        wb_root = etree.fromstring(wb_xml)
        active_tab = 0
        for bv in wb_root.iter(f"{{{ns_wb}}}bookView"):
            active_tab = int(bv.get("activeTab", 0))
            break
        rids = [sh.get(f"{{{ns_r}}}id")
                for sh in wb_root.iter(f"{{{ns_wb}}}sheet")]
        if not rids:
            return "xl/worksheets/sheet1.xml"
        rid = rids[min(active_tab, len(rids) - 1)]
        rels_xml  = zin.read("xl/_rels/workbook.xml.rels")
        rels_root = etree.fromstring(rels_xml)
        for rel in rels_root:
            if rel.get("Id") == rid and rel.get("Type", "").endswith("/worksheet"):
                target = rel.get("Target", "")
                if target.startswith("/"):
                    # 절대경로(패키지 루트 기준, openpyxl 등) → 선행 슬래시 제거
                    target = target[1:]
                elif not target.startswith("xl/"):
                    # 상대경로(워크북 기준) → xl/ 접두
                    target = "xl/" + target
                return target
    except Exception:
        pass
    return "xl/worksheets/sheet1.xml"


def _is_numeric(val: str) -> bool:
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def _set_cell_value(c_el, tag_v, tag_is, tag_t, tag_f, new_val: str):
    for child in list(c_el):
        c_el.remove(child)
    c_el.attrib.pop("t", None)
    if new_val == "":
        return
    if new_val.startswith("="):
        f_el = etree.SubElement(c_el, tag_f)
        f_el.text = new_val[1:]   # '=' 제외한 수식 본문
    elif _is_numeric(new_val):
        v_el = etree.SubElement(c_el, tag_v)
        v_el.text = new_val
    else:
        # t="str": sharedStrings.xml 변경 없이 Excel이 안전하게 수용하는 문자열 타입
        c_el.set("t", "str")
        v_el = etree.SubElement(c_el, tag_v)
        v_el.text = new_val


def _patch_sheet_xml(
    data: bytes,
    patches: dict[str, str],
    insert_rows: list[list[tuple[int, str]]] = [],
    delete_row_nums: set[int] = set(),
    delete_col_letters: set[str] = set(),
) -> bytes:
    """
    patches            : {cell_ref: value}          기존 셀 덮어쓰기
    insert_rows        : [[(col_idx, value), ...]]   파일 끝에 새 행 추가
    delete_row_nums    : {1-based row number}        해당 <row> 요소 자체 삭제
    delete_col_letters : {'A', 'B', ...}             해당 열의 모든 <c> 삭제
    """
    tree = etree.fromstring(data)
    ns    = _NS
    tag_c  = f"{{{ns}}}c"
    tag_f  = f"{{{ns}}}f"
    tag_is = f"{{{ns}}}is"
    tag_t  = f"{{{ns}}}t"
    tag_v  = f"{{{ns}}}v"

    existing: dict[str, etree._Element] = {}
    row_map:  dict[int, etree._Element] = {}

    sheetdata = tree.find(f"{{{ns}}}sheetData")
    if sheetdata is None:
        return data

    for row_el in sheetdata:
        r_idx = int(row_el.get("r", 0))
        row_map[r_idx] = row_el
        for c_el in row_el:
            ref = c_el.get("r", "")
            if ref:
                existing[ref] = c_el

    # 행 삭제 (실제 삭제된 행 번호를 추적)
    actually_deleted: set[int] = set()
    for row_num in delete_row_nums:
        row_el = row_map.get(row_num)
        if row_el is not None:
            sheetdata.remove(row_el)
            actually_deleted.add(row_num)

    # 덮어쓰기 패치
    for ref, new_val in patches.items():
        m = _COL_RE.match(ref)
        if not m:
            continue
        row_num = int(m.group(2))

        if ref in existing:
            c_el = existing[ref]
            if new_val == "":
                # 빈값 패치 시 <c> 요소 자체를 부모 행에서 제거
                parent = c_el.getparent()
                if parent is not None:
                    parent.remove(c_el)
            else:
                _set_cell_value(c_el, tag_v, tag_is, tag_t, tag_f, new_val)
        else:
            if row_num not in row_map:
                row_el = etree.SubElement(sheetdata, f"{{{ns}}}row")
                row_el.set("r", str(row_num))
                row_map[row_num] = row_el
                sheetdata[:] = sorted(sheetdata, key=lambda e: int(e.get("r", 0)))
            else:
                row_el = row_map[row_num]
            if new_val != "":
                c_el = etree.SubElement(row_el, tag_c)
                c_el.set("r", ref)
                _set_cell_value(c_el, tag_v, tag_is, tag_t, tag_f, new_val)
                row_el[:] = sorted(row_el, key=lambda e: column_index_from_string(
                    _COL_RE.match(e.get("r", "A1")).group(1)
                ))

    # 새 행 삽입 (실제 셀이 있는 마지막 행 바로 다음 행 번호부터 — 빈 <row> 요소 무시)
    if insert_rows:
        last_data_row = max(
            (int(row_el.get("r", 0)) for row_el in sheetdata if list(row_el)),
            default=0,
        )
        next_row = last_data_row + 1 if last_data_row > 0 else 1
        for cells in insert_rows:
            row_el = etree.SubElement(sheetdata, f"{{{ns}}}row")
            row_el.set("r", str(next_row))
            for col_idx, val in sorted(cells, key=lambda x: x[0]):
                if val == "":
                    continue
                ref = _cell_ref(next_row - 1, col_idx)
                c_el = etree.SubElement(row_el, tag_c)
                c_el.set("r", ref)
                _set_cell_value(c_el, tag_v, tag_is, tag_t, tag_f, val)
            next_row += 1

    # 빈 열 삭제: 지정된 열 문자에 속하는 <c> 요소를 모든 행에서 제거
    if delete_col_letters:
        for row_el in sheetdata:
            to_remove = [
                c_el for c_el in row_el
                if (m := _COL_RE.match(c_el.get("r", ""))) and m.group(1) in delete_col_letters
            ]
            for c_el in to_remove:
                row_el.remove(c_el)

    # VBA .Delete처럼 삭제된 행 번호를 위로 당겨 빈 행 번호 없애기
    if actually_deleted:
        sorted_deleted = sorted(actually_deleted)
        for row_el in sheetdata:
            rn = int(row_el.get("r", 0))
            offset = sum(1 for d in sorted_deleted if d < rn)
            if offset > 0:
                new_rn = rn - offset
                row_el.set("r", str(new_rn))
                for c_el in row_el:
                    ref = c_el.get("r", "")
                    m = _COL_RE.match(ref)
                    if m:
                        c_el.set("r", f"{m.group(1)}{new_rn}")

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
