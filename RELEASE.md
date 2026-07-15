# 릴리스 절차

앱은 GitHub `releases/latest` (`gywnsdlqkd-Nexus/ExcelMerge`)를 조회해 자동 업데이트한다.
따라서 새 버전은 **`v<버전>` 태그 + `.exe` 에셋을 붙인 GitHub Release** 로 게시되어야 한다.

## 순서

1. **버전 bump** — `excelmerge/__init__.py` 의 `__version__` 을 올린다. **버전 변경만 담은 단독 커밋**을
   권장(기능 커밋에 섞지 말 것).
   ```
   __version__ = "182"
   ```

2. **테스트** — 릴리스 전 게이트.
   ```bat
   pytest
   ```

3. **빌드** — 버전 무관 단일 스크립트. 결과물은 `dist/ExcelMerge_v<버전>.exe`.
   빌드가 끝나면 `build.bat` 이 이어서 `sign.py`(코드 서명, 인증서 미설정이면 자동 건너뜀)를 호출한다.
   ```bat
   build.bat
   ```
   > 재현성을 위해 **클린 venv + `pip install -r requirements.lock`** 후 빌드하는 것을 권장.
   > 빌드 환경 전제는 아래 "빌드 환경 요구사항" 참고.

4. **exe 스모크(필수)** — 결과 exe를 실제로 실행해 확인. 파일 생성만 확인하지 말 것.
   - 직접 실행: `dist/ExcelMerge_v<버전>.exe` 더블클릭 → 파일 비교·폴더 비교 동작.
   - **자동 업데이트 경로도 확인 권장**: 이전 버전 exe에서 이 릴리스로 업데이트 → 재실행까지 정상인지.
     (과거 이 경로에서 부트로더 환경변수 상속으로 재실행 실패한 사례가 있었음 — v183에서 수정.)

5. **게시** — `make_release.py` 가 sha256 계산 후 GitHub Release 를 만든다.
   ```bat
   REM 명령을 출력만(수동 실행용):
   python make_release.py --notes "이번 변경점 요약"

   REM gh CLI 로 실제 태그+릴리스 생성:
   python make_release.py --publish --notes "이번 변경점 요약"
   ```
   `--publish` 는 `gh release create v<버전> dist/ExcelMerge_v<버전>.exe -R <repo> -t v<버전> -n <notes>`
   를 실행한다(요구: `gh` CLI 로그인 상태).

6. **CHANGELOG** — `CHANGELOG.md` 의 `[Unreleased]` 항목을 새 버전 절로 옮긴다.

## 빌드 환경 요구사항 (재현성)

- **Python** — 현재 `requirements.lock` 기준 **Python 3.14 / Windows x64**. 정확 재현은 클린 venv +
  `pip install -r requirements.lock`.
- **⚠️ 빌드 Python 버전이 = 사용자 최소 Windows 버전을 결정한다.**
  - Python 3.14로 빌드하면 `python314.dll` 이 Windows 10+ 전용 API(`api-ms-win-core-path` 등)에
    의존 → **Windows 8.1 이하에서는 실행 불가**("Failed to load Python DLL … 지정된 모듈을 찾을 수 없습니다").
  - 구형 Windows(예: 7/8.1) 지원이 필요하면 그 OS를 지원하는 **낮은 Python으로 빌드**해야 한다
    (예: 3.8=Win7, 3.11/3.12=Win8.1). 지원 목표 OS를 먼저 정하고 그에 맞는 Python으로 빌드할 것.
- **Windows SDK(UCRT 재배포)** — `ExcelMerge.spec` 이 `C:\Program Files (x86)\Windows Kits\10\Redist\
  <ver>\ucrt\DLLs\x64` 의 UCRT DLL을 번들에 포함한다(UCRT 없는 PC에서 로드 실패 방지). SDK 미설치 시
  빌드가 중단되며 설치 안내가 뜬다.
- Rust 확장(`python-calamine`, `orjson`)은 `collect_all` 로 수집된다(스펙에 반영됨).

## 코드 서명 (권장)

미서명 exe는 SmartScreen 경고 + 백신 오탐/추출 차단 위험이 있다. `build.bat` 은 인증서가 환경변수로
설정돼 있으면 자동으로 `sign.py` 로 SHA-256 Authenticode 서명한다(없으면 조용히 건너뜀).

1. **인증서 준비** — 코드서명 인증서(OV/EV, 또는 사내 CA 발급)를 확보한다. 저장소 방식이 가장 편하다:
   Windows 인증서 저장소(`certmgr.msc` → 개인)에 설치 후 지문(Thumbprint) 확인.
2. **환경변수 설정**(둘 중 하나):
   - `EXCELMERGE_SIGN_THUMBPRINT` = 저장소 인증서 지문(SHA-1). **비밀번호 불필요 — 권장.**
   - `EXCELMERGE_SIGN_PFX` = .pfx 경로 (+ 필요 시 `EXCELMERGE_SIGN_PFX_PASSWORD`).
   - (선택) `EXCELMERGE_SIGN_TIMESTAMP` = RFC3161 서버(기본 DigiCert).
3. **빌드/서명** — `build.bat` 실행(자동 서명) 또는 이미 빌드된 exe에 `python sign.py [exe경로]`.
   서명 후 `signtool verify /pa` 로 검증 결과를 출력한다.
> 비밀번호는 스크립트에 하드코딩하지 말 것(저장소 지문 방식 권장). CI에서 서명하려면 시크릿으로 주입.

## 주의

- 자동 업데이트의 미인증 다운로드는 **public repo** 에서만 된다. private 면 토큰이 필요해 현재 미지원.
- 태그명은 `v<정수>` 형식(`v` 접두 허용). 업데이터의 `is_newer` 는 두 값이 모두 숫자면 정수로 비교한다.
- 매니페스트(`latest.json`) 방식은 GitHub 대신 임의 URL 로 배포할 때만 쓴다(`--base-url`/`--gdrive-exe`).
  GitHub Releases 를 쓰면 `latest.json` 불필요.
