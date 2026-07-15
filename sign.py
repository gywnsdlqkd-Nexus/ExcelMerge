"""코드 서명 헬퍼 - 빌드된 exe를 Authenticode(SHA-256)로 서명한다.

인증서는 코드에 담지 않고 환경변수로 주입한다(둘 중 하나):
  - EXCELMERGE_SIGN_THUMBPRINT : Windows 인증서 저장소(CurrentUser\\My)의 코드서명 인증서
                                 지문(SHA-1). 비밀번호 불필요 - **권장**.
  - EXCELMERGE_SIGN_PFX         : .pfx 파일 경로. 필요 시 EXCELMERGE_SIGN_PFX_PASSWORD 로 암호.
선택:
  - EXCELMERGE_SIGN_TIMESTAMP   : RFC3161 타임스탬프 서버(기본 DigiCert).

둘 다 없으면 **서명을 건너뛰고 정상 종료(exit 0)** - 미설정 환경에서도 빌드가 막히지 않게.
(서명 안 된 exe는 SmartScreen 경고·백신 오탐 위험이 있으니 배포 전 서명 권장.)

사용:
    python sign.py [exe경로]     # 생략 시 dist/ExcelMerge_v<버전>.exe
"""
import glob
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from excelmerge import __version__

# 콘솔 인코딩(예: cp949)이 일부 문자를 못 써서 print가 죽는 것을 방지.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:
        pass

_DEFAULT_TS = "http://timestamp.digicert.com"


def find_signtool() -> str:
    """Windows SDK에서 가장 최신 signtool.exe(x64)를 찾는다. 없으면 ""."""
    pats = [
        r"C:\Program Files (x86)\Windows Kits\10\bin\*\x64\signtool.exe",
        r"C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe",
        r"C:\Program Files (x86)\Windows Kits\10\App Certification Kit\signtool.exe",
    ]
    hits = []
    for p in pats:
        hits += glob.glob(p)
    return sorted(hits, reverse=True)[0] if hits else ""


def build_sign_cmd(signtool: str, exe: str) -> list | None:
    """환경변수 기반 서명 명령을 만든다. 인증서 미설정이면 None."""
    ts = os.environ.get("EXCELMERGE_SIGN_TIMESTAMP", _DEFAULT_TS)
    thumb = os.environ.get("EXCELMERGE_SIGN_THUMBPRINT", "").strip()
    pfx = os.environ.get("EXCELMERGE_SIGN_PFX", "").strip()
    common = ["sign", "/fd", "sha256", "/tr", ts, "/td", "sha256"]
    if thumb:
        return [signtool] + common + ["/sha1", thumb, exe]
    if pfx:
        pw = os.environ.get("EXCELMERGE_SIGN_PFX_PASSWORD", "")
        cmd = [signtool] + common + ["/f", pfx]
        if pw:
            cmd += ["/p", pw]
        return cmd + [exe]
    return None


def main():
    exe = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        HERE, "dist", f"ExcelMerge_v{__version__}.exe")
    if not os.path.isfile(exe):
        print(f"[서명] 대상 exe가 없습니다: {exe}")
        sys.exit(1)

    signtool = find_signtool()
    if not signtool:
        print("[서명 건너뜀] signtool.exe 를 찾지 못했습니다(Windows SDK 필요). 미서명 상태 유지.")
        return

    cmd = build_sign_cmd(signtool, exe)
    if cmd is None:
        print("[서명 건너뜀] 인증서 미설정 - EXCELMERGE_SIGN_THUMBPRINT 또는 "
              "EXCELMERGE_SIGN_PFX 를 설정하세요. 미서명 상태로 진행합니다.")
        return

    # 비밀번호(/p 뒤 값)는 로그에 남기지 않도록 마스킹해 출력.
    shown = list(cmd)
    if "/p" in shown:
        shown[shown.index("/p") + 1] = "****"
    print("[서명] " + " ".join(f'"{a}"' if " " in a else a for a in shown))
    r = subprocess.run(cmd)
    if r.returncode != 0:
        print(f"[서명 실패] signtool exit={r.returncode}")
        sys.exit(r.returncode)

    # 서명 검증(/pa: 기본 정책). 실패해도 서명 자체는 됐을 수 있어 경고만.
    v = subprocess.run([signtool, "verify", "/pa", "/v", exe])
    print(f"[서명 완료] 검증 exit={v.returncode}  →  {exe}")


if __name__ == "__main__":
    main()
