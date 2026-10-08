# 릴리스 절차

## 한 명령으로 (권장)

```
python release.py <버전>              # 예: python release.py 203
python release.py <버전> --dry-run    # 무엇을 할지만 보여주고 아무것도 바꾸지 않음
python release.py <버전> --no-publish # 로컬 커밋·빌드·스모크까지만
python release.py <버전> --skip-merge-check   # 전수 머지 검사만 건너뜀
```

아래 "순서"를 그대로 돌린다: 사전 점검 → 버전 bump → CHANGELOG 승격 → **테스트** →
**전수 머지 검사** → 커밋 → 빌드 → 서명 → **exe 실행 스모크** → 푸시 → 게시. 릴리스 노트는 CHANGELOG 의
해당 절을 그대로 쓰므로, **배포 전에 `## [미배포]` 에 변경점을 적어 두면 된다**.

스크립트가 막는 것:
 - `[미배포]` 가 비어 있으면 시작하지 않는다(본문 빈 릴리스 방지).
 - 버전이 현재보다 낮거나 같으면 멈춘다.
 - 버전/CHANGELOG 외의 커밋되지 않은 변경이 있으면 멈춘다(먼저 커밋할 것).
 - **테스트는 버전을 올린 뒤에** 돌린다 — 새 버전 번호를 하드코딩한 테스트가 있으면
   옛 순서(bump 전 테스트)로는 못 잡는다(v200 때 실제로 그렇게 깨졌다).
 - 푸시·게시는 빌드와 exe 실행 스모크를 통과한 뒤에만 한다.
 - **전수 머지 검사는 커밋 앞에서** 돌린다 — 실패해도 되돌릴 커밋이 없다.
 - 그 검사를 **건너뛸 때는 반드시 그렇게 말한다.** 조용히 넘어가면 "돌렸겠거니" 가
   되고, 실제로 v219~v221 이 그렇게 나갔다(v222 의 버그는 그 셋보다 앞서 있었다).

## 전수 머지 검사 — merge_check.py

**릴리스가 알아서 돌린다.** 두 빌드 폴더를 환경변수로 가리켜 두면 된다(경로는
머신마다 달라 저장소에 적을 수 없다 — sign.py 가 인증서를 받는 방식과 같다).

```
setx EXCELMERGE_CHECK_A "D:/HyoJun-Park-38_Build/GameData/ExcelData"
setx EXCELMERGE_CHECK_B "D:/HyoJun-Park-40_Build/GameData/ExcelData"
```

설정이 없으면 건너뛰되 **그 사실과 켜는 법을 출력한다.** 둘 중 하나만 설정하면
설정 실수로 보고 멈춘다. 손으로 돌리려면:

```
python merge_check.py <A폴더> <B폴더> <출력폴더> --excel
```

실제 빌드 데이터로 머지를 **전부** 돌려 보고, 결과를 엑셀로 열어 본다. 원본은 읽기만
하고 B 를 출력 폴더로 복사해 거기에만 쓴다. 하나라도 어긋나면 종료 코드 1.

`pytest` 로는 부족하다. 저장 쪽을 고칠 때마다 합성 조작("셀 하나 고치기", "행 몇 개
지우기")으로 확인했는데, 실제 머지는 그보다 많은 일을 한다 — 열 매칭, 빈 열·행 삭제
승격, **행 삽입**, 서식 병합. v215 에서 터진 셋은 전부 행 삽입 경로였고, 합성 조작이
그 경로를 안 밟아 v212~v214 를 그대로 통과했다. v222 의 "끝 행을 지우면서 행을
넣으면 번호가 겹친다" 도 여기서만 나왔다 — 삽입·삭제를 따로 시험한 테스트로는
안 보이고, 둘이 **같은 저장에서 끝쪽으로 겹쳐야** 드러난다.

`--excel` 이 최종 판정이다. 구조 검사가 놓치는 것을 엑셀은 잡는다 — 실제로 원인 여덟
중 넷은 '엑셀이 열어 주는가' 로 처음 드러났다.

커밋 이후 단계에서 실패하면 원인을 고치고 다시 돌리거나 `git reset --hard HEAD~1` 로
되돌린다. 아래는 각 단계를 손으로 할 때의 설명이다.


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

3. **빌드** — 버전 무관 단일 스크립트. 결과물은 **폴더** `dist/ExcelMerge/`
   (v217~ onedir — 켤 때마다 번들을 %TEMP% 에 푸는 비용을 없앴다. 실측 4.6초 → 2.7초).
   빌드가 끝나면 `build.bat` 이 이어서 `sign.py`(코드 서명, 인증서 미설정이면 자동 건너뜀)를 호출한다.
   ```bat
   build.bat
   ```
   > 재현성을 위해 **클린 venv + `pip install -r requirements.lock`** 후 빌드하는 것을 권장.
   > 빌드 환경 전제는 아래 "빌드 환경 요구사항" 참고.

4. **설치 파일 만들기** — `release.py` 가 ISCC(Inno Setup)로 `installer.iss` 를 컴파일해
   `dist/ExcelMerge_Setup.exe` 를 만든다(이름에 버전을 붙이지 않는다 — 릴리스마다
   태그가 달라 겹치지 않고, 받는 사람이 늘 같은 이름을 본다). 앱 exe 를 **먼저** 서명한 뒤 묶는다
   (순서가 바뀌면 설치 파일 안의 exe 가 미서명으로 들어간다).
   > ISCC 가 없으면 릴리스가 멈춘다. `winget install --id JRSoftware.InnoSetup`

5. **exe 스모크(필수)** — 결과를 실제로 실행해 확인. 파일 생성만 확인하지 말 것.
   - 직접 실행: `dist/ExcelMerge/ExcelMerge.exe` → 파일 비교·폴더 비교 동작.
   - **자동 업데이트 경로도 확인 권장**: 이전 버전 exe에서 이 릴리스로 업데이트 → 재실행까지 정상인지.
     (과거 이 경로에서 부트로더 환경변수 상속으로 재실행 실패한 사례가 있었음 — v183에서 수정.)

5. **게시** — `make_release.py` 가 sha256 계산 후 GitHub Release 를 만든다.
   ```bat
   REM 명령을 출력만(수동 실행용):
   python make_release.py --notes "이번 변경점 요약"

   REM gh CLI 로 실제 태그+릴리스 생성:
   python make_release.py --publish --notes "이번 변경점 요약"
   ```
   `--publish` 는 `gh release create v<버전> dist/ExcelMerge_Setup.exe -R <repo> ...`
   를 실행한다(요구: `gh` CLI 로그인 상태). 자동 업데이트도 이 설치 파일을 받아
   `/SILENT` 로 돌린다 — 릴리스에 설치 파일 말고 다른 exe 를 같이 올리지 말 것.

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
