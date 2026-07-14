# Changelog

형식은 [Keep a Changelog](https://keepachangelog.com/)를 느슨히 따른다. 버전은
`excelmerge/__init__.py` 의 정수 `__version__`.

과거 상세 이력(~v181)은 git 커밋 메시지와 GitHub Releases 본문에 있다. 이 파일은 v182부터
사용자 관점 변경점을 요약한다.

## [182] - 2026-07-14

### 추가 / 변경 (사용자 체감)
- 상태바 우측에 진행률 바 추가 — 폴더 스캔·파일 로딩은 확정 진행률, 비교·저장은 바쁨 표시.
- 폴더 병합을 백그라운드로 수행 — 대용량 복사 시 UI가 얼지 않음.
- 폴더 비교 패널 어디에 폴더를 떨궈도 등록됨(이전엔 좁은 경로칸만 인식).
- 변경점/찾기 이동이 항상 깨끗한 단일 선택으로 착지(결정론적).
- 행 헤더 우클릭: '키 행 초기화' 항목 제거 — 키 행이 아닌 행에서만 '키 행으로 설정' 노출.

### 변경 (인프라 — 사용자 체감 없음)
- 빌드 체계 단일화: 버전마다 복제하던 `ExcelMerge_v{N}.spec`/`build_v{N}.bat` 을 버전과 무관한
  단일 `ExcelMerge.spec`(버전은 `__version__` 에서 읽음) + `build.bat` 으로 통합. 스테일 빌드 파일
  일괄 삭제.
- 의존성 정리: 미사용 `formulas`·`numpy` 제거, 상한 명시(`requirements.txt`), 재현 빌드용
  `requirements.lock` 추가.
- 테스트: `xlsx_writer` 서식 병합·`uasset_parser` 폴백·`loaders` calamine→openpyxl 폴백 등 고위험
  경로 테스트 추가(53→67개). `pyproject.toml`/`tests/conftest.py` 로 하네스 통일, GitHub Actions CI 추가.
- 안정성: `crash.log` 1MB 초과 시 1세대 로테이션(무한 증가 방지).
- 릴리스: `make_release.py --publish` 로 `gh release create` 실제 게시 지원.
- 문서: `README.md`·`CHANGELOG.md`·`RELEASE.md` 신설.
