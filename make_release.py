"""릴리스 헬퍼 — 빌드된 exe로 latest.json(자동 업데이트 매니페스트)을 생성한다.

사용법:
    python make_release.py [--base-url https://.../ ] [--notes "이번 변경점"]

동작:
  - excelmerge.__version__ 로 대상 exe(dist/ExcelMerge_v<ver>.exe) 확인
  - sha256 계산
  - dist/latest.json 작성 { version, url, sha256, notes }
    · --base-url 을 주면 url = base_url + "ExcelMerge_v<ver>.exe"
    · 안 주면 url 은 자리표시자("") — 업로드 후 실제 직링크로 채워 넣으면 됨

이후: dist/ExcelMerge_v<ver>.exe 와 dist/latest.json 을 배포 위치(클라우드/사내 웹)에 올리면 끝.
"""
import argparse
import hashlib
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from excelmerge import __version__


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--github-repo", default="",
                    help="GitHub 릴리스로 배포할 공개 repo(owner/name). 미지정 시 updater.GITHUB_REPO")
    ap.add_argument("--base-url", default="",
                    help="exe가 올라갈 위치의 베이스 URL(끝에 / 포함). 예: https://host/dl/")
    ap.add_argument("--gdrive-exe", default="",
                    help="업로드한 exe의 Google Drive 공유 링크 또는 파일 ID → 직접 다운로드 URL로 변환")
    ap.add_argument("--notes", default="", help="이번 버전 변경점(사용자에게 표시)")
    ap.add_argument("--publish", action="store_true",
                    help="GitHub Releases 모드에서 명령을 '출력만' 하지 않고 실제로 gh 로 태그+릴리스 생성")
    args = ap.parse_args()

    exe_name = f"ExcelMerge_v{__version__}.exe"
    exe_path = os.path.join(HERE, "dist", exe_name)
    if not os.path.isfile(exe_path):
        print(f"[오류] 빌드된 exe가 없습니다: {exe_path}\n먼저 PyInstaller로 빌드하세요.")
        sys.exit(1)

    # ── GitHub Releases 모드 ──
    from excelmerge import updater as _upd
    repo = (args.github_repo or _upd.GITHUB_REPO).strip().strip("/")
    if repo:
        sha = _sha256(exe_path)
        tag = f"v{__version__}"
        print(f"[GitHub Releases] repo={repo}  tag={tag}  sha256={sha}")
        import shutil as _sh
        gh = _sh.which("gh")
        cmd = ["gh", "release", "create", tag, exe_path, "-R", repo,
               "-t", tag, "-n", (args.notes or tag)]
        printable = f'  gh release create {tag} "{exe_path}" -R {repo} -t "{tag}" -n "{args.notes or tag}"'

        if args.publish:
            if not gh:
                print("[오류] --publish 를 쓰려면 gh CLI 가 필요합니다. 설치 후 다시 시도하거나 아래 명령을 수동 실행:")
                print(printable)
                sys.exit(1)
            import subprocess
            print(f"[게시] {tag} 릴리스를 생성합니다 (gh)...")
            r = subprocess.run(cmd)
            if r.returncode != 0:
                print(f"[오류] gh release create 실패(exit={r.returncode}). 위 로그 확인.")
                sys.exit(r.returncode)
            print(f"[완료] {tag} 게시됨. 앱은 releases/latest 를 조회하므로 별도 latest.json 불필요.")
            return

        if gh:
            print("\ngh CLI 감지됨 — 실제 게시는 --publish, 또는 아래 명령을 수동 실행:")
        else:
            print("\ngh CLI 미설치 — 웹 UI로 올리거나 gh 설치 후 아래 명령 사용:")
        print(printable)
        print("\n또는 웹: repo → Releases → Draft new release →")
        print(f"  Tag: {tag}  /  첨부: dist/{exe_name}  /  Publish")
        print("\n앱은 releases/latest를 조회하므로 별도 latest.json 불필요.")
        return

    if args.gdrive_exe:
        from excelmerge.updater import gdrive_id_from_url, gdrive_direct_url
        fid = gdrive_id_from_url(args.gdrive_exe) or args.gdrive_exe.strip()
        url = gdrive_direct_url(fid)
    elif args.base_url:
        url = args.base_url.rstrip("/") + "/" + exe_name
    else:
        url = ""
    manifest = {
        "version": __version__,
        "url": url,
        "sha256": _sha256(exe_path),
        "notes": args.notes,
    }
    out = os.path.join(HERE, "dist", "latest.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(f"[완료] {out}")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    if not url:
        print("\n※ url이 비어 있습니다 — 업로드 후 latest.json의 url을 실제 직링크로 채워 넣으세요.")
    print("\n다음: dist/{0} 와 dist/latest.json 을 배포 위치에 업로드하세요.".format(exe_name))


if __name__ == "__main__":
    main()
