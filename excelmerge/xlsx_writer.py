"""xlsx 저장 — sheet XML 직접 패치로 수식 보존 기록 (excel_diff_merge.py에서 분리)."""
import os
import re
import shutil
import zipfile
from copy import deepcopy
from collections import defaultdict

from lxml import etree
from openpyxl.utils import get_column_letter, column_index_from_string

from .logutil import log


_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_COL_RE = re.compile(r"([A-Z]+)(\d+)")

# 자주 쓰는 네임스페이스 태그(Clark 표기) — 반복 f-string 생성 제거.
_TAG_SHEETDATA = f"{{{_NS}}}sheetData"
_TAG_ROW = f"{{{_NS}}}row"
_TAG_C = f"{{{_NS}}}c"
_TAG_F = f"{{{_NS}}}f"
_TAG_IS = f"{{{_NS}}}is"
_TAG_T = f"{{{_NS}}}t"
_TAG_V = f"{{{_NS}}}v"


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
        # 빈 열/행 감지 실패 → 승격 없이 진행(안전). 원인은 진단 로그로만 남긴다.
        log.warning("빈 열/행 감지 실패(승격 건너뜀): %s", path, exc_info=True)
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
        raise PermissionError(f"파일이 열려 있어 저장할 수 없습니다:\n{path_base}")

    tmp = path_base + ".tmp_merge"
    try:
        with zipfile.ZipFile(path_base, "r") as zin, \
             zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            sheet_path = _resolve_sheet_path(zin, sheet_name)
            # 시트를 해결하지 못하면 패치가 조용히 유실된다 — 데이터 손실 대신 오류로 노출.
            if sheet_path not in zin.namelist():
                raise ValueError(f"워크시트를 찾을 수 없습니다: {sheet_path}")

            # 서식 병합 pre-pass — 실패해도 값 병합은 그대로 진행 (내부 try/except).
            patch_styles, new_styles_bytes, insert_rows = _prepare_style_merge(
                zin, sheet_path, src_path, src_sheet_name,
                patch_style_src, insert_rows)

            for item in zin.infolist():
                # calcChain.xml 항상 제거 — 수식 패치/삽입/삭제 모두 계산 체인을 무효화함
                if item.filename == "xl/calcChain.xml":
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
                    )
                zout.writestr(item, data)
        # 덮어쓰기 전 원본을 .bak로 백업 — 잘못 저장 시 복구 수단(원자적 교체와 별개).
        try:
            if os.path.exists(path_base):
                shutil.copy2(path_base, path_base + ".bak")
        except OSError:
            pass   # 백업 실패는 저장 자체를 막지 않는다(권한/디스크 등)
        os.replace(tmp, path_base)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def _rel_target_to_path(target: str) -> str:
    """workbook.xml.rels의 worksheet Target을 zip 내부 경로로 정규화."""
    if target.startswith("/"):
        # 절대경로(패키지 루트 기준, openpyxl 등) → 선행 슬래시 제거
        return target[1:]
    if not target.startswith("xl/"):
        # 상대경로(워크북 기준) → xl/ 접두
        return "xl/" + target
    return target


def _find_sheet_path_by_name(zin: zipfile.ZipFile, sheet_name: str):
    """워크북에서 sheet_name에 해당하는 워크시트 XML 경로를 해석. 못 찾으면 None.
    workbook.xml의 <sheet name= r:id=>에서 이름을 매칭하고 rels에서 target을 찾는다."""
    if not sheet_name:
        return None
    try:
        ns_wb = _NS
        ns_r  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        wb_root = etree.fromstring(zin.read("xl/workbook.xml"))
        rid = None
        for sh in wb_root.iter(f"{{{ns_wb}}}sheet"):
            if sh.get("name") == sheet_name:
                rid = sh.get(f"{{{ns_r}}}id")
                break
        if rid is None:
            return None
        rels_root = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        for rel in rels_root:
            if rel.get("Id") == rid and rel.get("Type", "").endswith("/worksheet"):
                return _rel_target_to_path(rel.get("Target", ""))
    except Exception:
        pass
    return None


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
                return _rel_target_to_path(rel.get("Target", ""))
    except Exception:
        pass
    return "xl/worksheets/sheet1.xml"


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
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


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


def _append_rows(sheetdata, insert_rows):
    """실제 셀이 있는 마지막 행 다음부터 새 행을 추가(빈 <row> 요소는 무시)."""
    if not insert_rows:
        return
    last_data_row = max(
        (int(row_el.get("r", 0)) for row_el in sheetdata if list(row_el)), default=0)
    next_row = last_data_row + 1 if last_data_row > 0 else 1
    for cells in insert_rows:
        row_el = etree.SubElement(sheetdata, _TAG_ROW)
        row_el.set("r", str(next_row))
        for cell in sorted(cells, key=lambda x: x[0]):
            col_idx, val = cell[0], cell[1]
            dst_s = cell[2] if len(cell) > 2 else None
            if val == "":
                continue
            c_el = etree.SubElement(row_el, _TAG_C)
            c_el.set("r", _cell_ref(next_row - 1, col_idx))
            _set_cell_value(c_el, val)
            if dst_s is not None:
                c_el.set("s", str(dst_s))
        next_row += 1


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


def _patch_sheet_xml(
    data: bytes,
    patches: dict[str, str],
    insert_rows: list[list[tuple]] | None = None,
    delete_row_nums: set[int] | None = None,
    delete_col_letters: set[str] | None = None,
    patch_styles: dict[str, int] | None = None,
) -> bytes:
    """sheet XML 을 직접 패치. 5개 단계를 순서대로 위임한다(각 단계는 별도 헬퍼).

    patches            : {cell_ref: value}          기존 셀 덮어쓰기
    insert_rows        : [[(col_idx, value[, dst_s]), ...]]  파일 끝에 새 행 추가
    delete_row_nums    : {1-based row number}        해당 <row> 요소 자체 삭제
    delete_col_letters : {'A', 'B', ...}             해당 열의 모든 <c> 삭제
    patch_styles       : {cell_ref: style_index}     병합할 소스 서식 인덱스 (덮어쓰기 셀)
    """
    insert_rows = insert_rows or []
    delete_row_nums = delete_row_nums or set()
    delete_col_letters = delete_col_letters or set()

    tree = etree.fromstring(data)
    sheetdata = tree.find(_TAG_SHEETDATA)
    if sheetdata is None:
        return data

    existing, row_map = _index_sheet(sheetdata)
    deleted = _delete_rows(sheetdata, row_map, delete_row_nums)
    _apply_patches(sheetdata, existing, row_map, patches, patch_styles)
    _append_rows(sheetdata, insert_rows)
    _delete_columns(sheetdata, delete_col_letters)
    _renumber_after_delete(sheetdata, deleted)

    return etree.tostring(tree, xml_declaration=True, encoding="UTF-8", standalone=True)
