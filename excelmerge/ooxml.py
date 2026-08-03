"""OOXML(xlsx) 워크북 구조 해석의 단일 출처 — 시트 이름/관계(rel) → zip 내부 경로.

**왜 이 모듈이 필요한가.** 같은 해석 로직이 loaders(읽기)와 xlsx_writer(쓰기)에 각각
구현돼 있었다. 둘이 갈라지면 "읽은 시트"와 "쓴 시트"가 달라져 **엉뚱한 시트에 병합 결과가
저장되는 데이터 손상**으로 직결된다(예: Target 정규화를 한쪽만 고치는 경우).
그래서 기계적인 해석은 여기 한 곳에만 둔다.

**정책은 호출부에 남긴다.** 읽기와 쓰기는 '시트를 못 찾았을 때' 의도적으로 다르게 동작한다:
- 읽기(loaders): 첫 시트로 관대하게 폴백 — 미리보기/비교는 실패보다 표시가 낫다.
- 쓰기(xlsx_writer): **raise** — 폴백하면 엉뚱한 시트를 덮어써 데이터가 손상된다.
이 차이는 버그가 아니라 설계이므로 통합하지 않는다. 이 모듈은 '찾기'만 하고 폴백/예외는
호출부가 결정한다.
"""
import xml.etree.ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

WORKBOOK_XML = "xl/workbook.xml"
WORKBOOK_RELS = "xl/_rels/workbook.xml.rels"
# 워크북/관계 해석이 모두 실패했을 때의 관례적 첫 시트 경로.
DEFAULT_SHEET_PATH = "xl/worksheets/sheet1.xml"


def rel_target_to_path(target: str) -> str:
    """workbook.xml.rels 의 Target 을 zip 내부 경로로 정규화.

    - "/xl/worksheets/sheet1.xml" (패키지 루트 절대경로, openpyxl 등) → 선행 슬래시 제거
    - "worksheets/sheet1.xml"    (워크북 기준 상대경로)              → "xl/" 접두
    - "xl/worksheets/sheet1.xml" (이미 정규형)                       → 그대로
    """
    if target.startswith("/"):
        return target[1:]
    if not target.startswith("xl/"):
        return "xl/" + target
    return target


def sheet_entries(z) -> list:
    """workbook.xml 의 시트 목록을 **문서 순서**로 [(name, rid), ...] 반환. 실패 시 [].

    문서 순서는 activeTab 인덱싱의 기준이므로 정렬하지 말 것.

    스펙상 <sheet> 는 <sheets> 의 자식이어야 하므로 **컨테이너 범위로 먼저 찾는다**.
    통합 전 쓰기 경로는 문서 전체를 iter() 로 훑어 <sheets> 밖의 떠돌이 <sheet> 까지
    주워, activeTab 이 엉뚱한 파트를 가리켜 **그 시트에 저장**될 수 있었다(데이터 손상).
    컨테이너에서 하나도 못 찾은 경우에만 iter() 로 완화 탐색해 기존 관용성은 유지한다.
    """
    try:
        root = ET.fromstring(z.read(WORKBOOK_XML))
    except Exception:
        return []
    found = root.findall(f"{{{MAIN_NS}}}sheets/{{{MAIN_NS}}}sheet")
    if not found:
        found = list(root.iter(f"{{{MAIN_NS}}}sheet"))
    return [(sh.get("name"), sh.get(f"{{{REL_NS}}}id")) for sh in found]


def active_tab(z) -> int:
    """첫 <bookView> 의 activeTab(기본 0). 실패 시 0."""
    try:
        root = ET.fromstring(z.read(WORKBOOK_XML))
        for bv in root.iter(f"{{{MAIN_NS}}}bookView"):
            return int(bv.get("activeTab", 0))
    except Exception:
        pass
    return 0


def path_for_rid(z, rid: str, require_worksheet_type: bool) -> str:
    """rel Id(rid) → 워크시트 XML 의 zip 내부 경로. 못 찾으면 None.

    require_worksheet_type: Type 이 ".../worksheet" 로 끝나는 rel 만 인정할지.
      쓰기 경로는 True(엄격 — 엉뚱한 파트에 쓰는 것을 막는다), 읽기 경로는 기존 동작
      보존을 위해 False. 두 경로의 현재 동작을 그대로 유지하기 위한 파라미터다.
    """
    if not rid:
        return None
    try:
        rels = ET.fromstring(z.read(WORKBOOK_RELS))
    except Exception:
        return None
    for rel in rels:
        if rel.get("Id") != rid:
            continue
        if require_worksheet_type and not rel.get("Type", "").endswith("/worksheet"):
            continue
        return rel_target_to_path(rel.get("Target", ""))
    return None


def sheet_path_by_name(z, sheet_name, require_worksheet_type: bool = False) -> str:
    """이름으로 워크시트 경로 해석. 이름이 없거나 못 찾으면 None(폴백하지 않음 — 정책은 호출부)."""
    if not sheet_name:
        return None
    for name, rid in sheet_entries(z):
        if name == sheet_name:
            return path_for_rid(z, rid, require_worksheet_type)
    return None


def first_sheet_path(z, require_worksheet_type: bool = False) -> str:
    """문서 순서상 첫 시트의 경로. 못 찾으면 None."""
    entries = sheet_entries(z)
    if not entries:
        return None
    return path_for_rid(z, entries[0][1], require_worksheet_type)


def active_sheet_path(z, require_worksheet_type: bool = True) -> str:
    """activeTab 이 가리키는 시트의 경로. 범위를 넘으면 마지막 시트로 clamp. 못 찾으면 None."""
    entries = sheet_entries(z)
    if not entries:
        return None
    idx = min(active_tab(z), len(entries) - 1)
    return path_for_rid(z, entries[idx][1], require_worksheet_type)
