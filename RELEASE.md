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
   ```bat
   build.bat
   ```
   > 빌드 머신에 Windows SDK(UCRT 재배포)가 있어야 한다. 재현성이 중요하면 클린 venv +
   > `pip install -r requirements.lock` 후 빌드.

4. **exe 스모크** — `dist/ExcelMerge_v<버전>.exe` 를 직접 실행해 파일 비교·폴더 비교가 되는지 확인.

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

## 주의

- 자동 업데이트의 미인증 다운로드는 **public repo** 에서만 된다. private 면 토큰이 필요해 현재 미지원.
- 태그명은 `v<정수>` 형식(`v` 접두 허용). 업데이터의 `is_newer` 는 두 값이 모두 숫자면 정수로 비교한다.
- 매니페스트(`latest.json`) 방식은 GitHub 대신 임의 URL 로 배포할 때만 쓴다(`--base-url`/`--gdrive-exe`).
  GitHub Releases 를 쓰면 `latest.json` 불필요.
