# ExcelMerge

Excel(`xlsx/xls/xlsm/xlsb`)·JSON·Unreal `uasset` 파일을 **A/B로 비교하고 셀 단위로 병합**하는
PyQt5 Windows 데스크톱 툴. Beyond Compare 스타일의 파일/폴더 비교, 키 열 기반 행 매칭, 서식 보존
병합, 자동 업데이트를 제공한다.

기능 사양 전문은 [`ExcelMerge.md`](ExcelMerge.md) 참고.

## 설치 · 실행 (개발)

```bat
pip install -r requirements.txt      REM 또는 재현 빌드: pip install -r requirements.lock
run.bat                              REM = python excel_diff_merge.py
```

P4V 등 외부 diff 툴 연동:

```bat
python excel_diff_merge.py -s <A경로> -d <B경로>
```

파일이면 셀 비교, 폴더면 폴더 비교로 진입한다.

## 테스트

```bat
pip install pytest
pytest
```

Qt 위젯 테스트는 헤드리스에서 돌도록 `QT_QPA_PLATFORM=offscreen` 로 실행된다(설정은
`tests/conftest.py`). CI(`.github/workflows/ci.yml`)가 push/PR마다 Windows 러너에서 동일하게 돈다.

## 빌드 (exe)

버전은 `excelmerge/__init__.py` 의 `__version__` 단일 출처에서 온다. 빌드 스크립트는 버전과 무관한
단일 파일이다:

```bat
build.bat            REM PyInstaller ExcelMerge.spec → dist/ExcelMerge_v<버전>.exe → (설정 시) 코드 서명
```

`build.bat` 은 빌드 후 `sign.py`(코드 서명)를 호출한다 — 인증서 환경변수가 설정돼 있으면 서명하고,
없으면 조용히 건너뛴다. 재현성·서명·상세 절차는 [`RELEASE.md`](RELEASE.md) 참고.

> **빌드 전제(중요)**
> - **Windows SDK(UCRT 재배포)** 필요 — UCRT DLL을 번들에 포함해 UCRT 미설치 PC의 "Failed to load
>   Python DLL" 오류를 막는다(SDK 없으면 `ExcelMerge.spec` 이 빌드 중단 + 안내).
> - **빌드 Python 버전이 사용자 최소 Windows 버전을 결정한다.** Python 3.14 빌드는 **Windows 10+ 전용**
>   (구형 Windows에서 DLL 로드 실패). 구형 OS 지원이 필요하면 낮은 Python으로 빌드할 것.
> - 재현 빌드: 클린 venv + `pip install -r requirements.lock`.

## 릴리스 · 자동 업데이트

- 배포/게시 절차는 [`RELEASE.md`](RELEASE.md), 버전별 변경점은 [`CHANGELOG.md`](CHANGELOG.md).
- 앱은 실행 시 백그라운드로 GitHub `releases/latest`(`gywnsdlqkd-Nexus/ExcelMerge`)를 조회해 새 버전을
  감지·자기 교체한다(frozen exe에서만). 상세는 `ExcelMerge.md` "자동 업데이트" 절.

## 구조

```
excel_diff_merge.py        진입점(인자 파싱 · 크래시로깅 설치 · MainWindow 기동)
excelmerge/
  diff_engine.py           순수 diff 계산(키/행 매칭)
  diff_model.py            Qt 뷰-모델
  loaders.py               값 로딩(calamine → openpyxl 폴백) · 시트명 · 수식 플래그
  xlsx_writer.py           sheet XML 직접 패치 저장(수식/서식 보존)
  folder_compare.py        폴더 비교 백엔드(바이트→내용 비교, 스레드풀)
  uasset_parser.py         UE5 DataTable .uasset 파서
  workers.py               QThread 워커(로드/비교/병합/스캔)
  main_window.py · diff_view.py · folder_view.py · panels.py · widgets.py   UI
  theme.py · prefs.py · updater.py · crashlog.py                            지원
tests/                     pytest 스위트
```
