# -*- coding: utf-8 -*-
"""릴리스 한 번에 — 게이트 → 버전 bump → CHANGELOG → 커밋 → 빌드 → exe 스모크 → 게시.

RELEASE.md 의 순서를 그대로 따르되, **빠뜨릴 수 없게** 묶은 것이다. 손으로 돌리면
일곱 단계를 매번 같은 순서로 반복해야 하고, 실제로 이런 사고가 났다:

  · 버전을 올린 뒤 테스트를 안 돌려, 새 버전 번호를 하드코딩한 테스트가 깨진 채 빌드됐다.
    → 그래서 이 스크립트는 **bump 다음에** 테스트를 돌린다(옛 순서로는 못 잡는 실패다).
  · CHANGELOG [미배포] 가 비어 있는데 릴리스를 만들어 본문이 빈 적이 있다.
    → 비어 있으면 시작조차 하지 않는다.

사용:
    python release.py 203                 # 전 과정(게시까지)
    python release.py 203 --dry-run       # 무엇을 할지만 보여주고 아무것도 바꾸지 않음
    python release.py 203 --no-publish    # 로컬 커밋·빌드·스모크까지만(푸시/게시 안 함)

중간에 실패하면 그 자리에서 멈춘다. 커밋 이후 단계에서 실패했다면 원인을 고치고 다시
돌리거나, 커밋을 되돌린다(`git reset --hard HEAD~1`) — 푸시·게시는 스모크까지 통과해야
일어나므로, 실패한 빌드가 사용자에게 나가는 일은 없다.
"""
import argparse
import datetime
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
INIT_PY = os.path.join(HERE, "excelmerge", "__init__.py")
CHANGELOG = os.path.join(HERE, "CHANGELOG.md")
UNRELEASED = "## [미배포]"


# ── 순수 함수(테스트 대상) ───────────────────────────────────────────────────

def bump_version(text: str, new_version: str) -> str:
    """__init__.py 본문의 __version__ 을 new_version 으로. 없으면 예외."""
    pat = re.compile(r'^__version__\s*=\s*"(\d+)"', re.M)
    m = pat.search(text)
    if not m:
        raise ValueError("__version__ 을 찾지 못했습니다.")
    if int(new_version) <= int(m.group(1)):
        raise ValueError(f"버전이 뒤로 갑니다: 현재 v{m.group(1)} → v{new_version}")
    return pat.sub(f'__version__ = "{new_version}"', text, count=1)


def current_version(text: str) -> str:
    m = re.search(r'^__version__\s*=\s*"(\d+)"', text, re.M)
    if not m:
        raise ValueError("__version__ 을 찾지 못했습니다.")
    return m.group(1)


def unreleased_body(text: str) -> str:
    """CHANGELOG 의 [미배포] 절 본문(다음 절 직전까지). 없거나 비면 빈 문자열."""
    i = text.find(UNRELEASED)
    if i < 0:
        return ""
    rest = text[i + len(UNRELEASED):]
    m = re.search(r"^## \[", rest, re.M)
    return (rest[:m.start()] if m else rest).strip()


def promote_changelog(text: str, version: str, date: str) -> str:
    """[미배포] 를 [<version>] - <date> 절로 승격하고, 빈 [미배포] 를 새로 얹는다."""
    if UNRELEASED not in text:
        raise ValueError("CHANGELOG 에 [미배포] 절이 없습니다.")
    if not unreleased_body(text):
        raise ValueError("[미배포] 절이 비어 있습니다 — 릴리스 노트가 빈 채로 나갑니다.")
    return text.replace(UNRELEASED, f"{UNRELEASED}\n\n## [{version}] - {date}", 1)


def notes_for(text: str, version: str) -> str:
    """게시용 릴리스 노트 = 해당 버전 절의 본문 그대로."""
    head = f"## [{version}] - "
    i = text.find(head)
    if i < 0:
        raise ValueError(f"CHANGELOG 에 [{version}] 절이 없습니다.")
    rest = text[i:]
    nl = rest.find("\n")
    rest = rest[nl + 1:]
    m = re.search(r"^## \[", rest, re.M)
    return (rest[:m.start()] if m else rest).strip()


def blocking_changes(porcelain: str, allowed: set) -> list:
    """`git status --porcelain` 에서 허용 목록 밖의 변경만 골라 낸다.

    porcelain 한 줄 = 상태 2칸 + 공백 + 경로(' M path', '?? path', 'R  old -> new').
    경로만 떼어 비교한다 — 이름 변경은 도착 경로로 본다.
    """
    out_lines = []
    for line in porcelain.splitlines():
        if not line.strip():
            continue
        path = line[3:].strip().strip('"')
        if " -> " in path:                      # 이름 변경(R) — 새 경로 기준
            path = path.split(" -> ", 1)[1].strip().strip('"')
        if path not in allowed:
            out_lines.append(line)
    return out_lines


# ── 실행 도우미 ──────────────────────────────────────────────────────────────

def run(cmd, **kw):
    print(f"  $ {' '.join(cmd) if isinstance(cmd, list) else cmd}")
    r = subprocess.run(cmd, cwd=HERE, **kw)
    if r.returncode != 0:
        fail(f"명령 실패(exit={r.returncode}): {cmd}")
    return r


def out(cmd) -> str:
    """표준출력 그대로 — **strip 하지 않는다**.

    `git status --porcelain` 은 줄 앞 두 칸이 상태 코드다(' M path' 처럼 앞이 공백일 수
    있다). 통째로 strip 하면 첫 줄의 그 공백이 사라져 경로 파싱이 한 칸 밀린다 —
    실제로 CHANGELOG.md 가 '허용 목록에 있는데도 막힘'으로 잡혔다.
    """
    return subprocess.run(cmd, cwd=HERE, capture_output=True, text=True,
                          encoding="utf-8", errors="replace").stdout


def step(msg):
    print(f"\n▶ {msg}")


def fail(msg):
    print(f"\n✖ {msg}")
    sys.exit(1)


def read(path):
    with open(path, encoding="utf-8", newline="") as f:
        return f.read()


def write(path, text):
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)


def exe_smoke(version: str, timeout: int = 90) -> None:
    """빌드된 exe 를 띄워 '창이 뜨고 살아 있는지' 확인하고 닫는다(Windows 전용).

    파일이 만들어졌다는 것만으로는 부족하다 — 과거에 부트로더/DLL 문제로 '빌드는 됐는데
    실행이 안 되는' 사고가 있었다(RELEASE.md 참고).
    """
    exe = os.path.join(HERE, "dist", f"ExcelMerge_v{version}.exe")
    if not os.path.isfile(exe):
        fail(f"빌드 결과가 없습니다: {exe}")
    if os.name != "nt":
        print("  (Windows 아님 — 실행 스모크 건너뜀)")
        return
    import ctypes
    import time
    proc = subprocess.Popen([exe], cwd=HERE)
    try:
        user32 = ctypes.windll.user32
        want = f"ExcelMerge v{version}"
        found = None
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline and found is None:
            time.sleep(1.0)
            if proc.poll() is not None:
                fail(f"exe 가 조기 종료했습니다(exit={proc.returncode}).")

            def cb(hwnd, _lp):
                nonlocal found
                buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, buf, 256)
                if buf.value.strip() == want and user32.IsWindowVisible(hwnd):
                    found = hwnd
                return True

            proto = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
            user32.EnumWindows(proto(cb), 0)
        if found is None:
            fail(f"'{want}' 창이 {timeout}초 안에 뜨지 않았습니다.")
        print(f"  창 확인: {want}")
    finally:
        try:
            proc.terminate()
            proc.wait(timeout=10)
        except Exception:
            proc.kill()


# ── 본체 ─────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="ExcelMerge 릴리스 한 번에")
    ap.add_argument("version", help="새 버전(정수). 예: 203")
    ap.add_argument("--dry-run", action="store_true",
                    help="검사만 하고 아무것도 바꾸지 않는다")
    ap.add_argument("--no-publish", action="store_true",
                    help="로컬 커밋·빌드·스모크까지만 — 푸시/게시는 하지 않는다")
    args = ap.parse_args()

    if not args.version.isdigit():
        fail("버전은 정수여야 합니다(예: 203).")

    step("사전 점검")
    init_text = read(INIT_PY)
    cur = current_version(init_text)
    print(f"  현재 v{cur} → 새 v{args.version}")
    bump_version(init_text, args.version)          # 역행이면 여기서 멈춘다

    chg = read(CHANGELOG)
    body = unreleased_body(chg)
    if not body:
        fail("CHANGELOG [미배포] 절이 비어 있습니다 — 변경점을 먼저 적으세요.")
    print(f"  [미배포] 항목 {len(body.splitlines())}줄 확인")

    allowed = {"excelmerge/__init__.py", "CHANGELOG.md"}
    blocking = blocking_changes(out(["git", "status", "--porcelain"]), allowed)
    if blocking:
        fail("커밋되지 않은 변경이 있습니다(먼저 커밋하세요):\n    "
             + "\n    ".join(blocking))
    print("  작업 트리 깨끗")

    if args.dry_run:
        date = datetime.date.today().isoformat()
        print("\n[dry-run] 다음을 수행합니다:")
        print(f"  1) __version__ = \"{args.version}\"")
        print(f"  2) CHANGELOG: [미배포] → [{args.version}] - {date}")
        print("  3) pytest (버전 올린 뒤에 돌린다)")
        print("  4) 커밋: chore(release): v{0} — 버전 bump + CHANGELOG".format(args.version))
        print("  5) PyInstaller 빌드 + 서명 시도")
        print("  6) exe 실행 스모크(창 확인)")
        if not args.no_publish:
            print("  7) 푸시 + GitHub Release 게시(노트는 CHANGELOG 절 그대로)")
        print("\n릴리스 노트 미리보기 —")
        print("\n".join("    " + l for l in body.splitlines()[:12]))
        return

    step("버전 bump + CHANGELOG 승격")
    date = datetime.date.today().isoformat()
    write(INIT_PY, bump_version(init_text, args.version))
    write(CHANGELOG, promote_changelog(chg, args.version, date))
    print(f"  v{args.version} / {date}")

    step("테스트 게이트 (버전 올린 뒤 — 버전에 묶인 테스트까지 잡는다)")
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    run([sys.executable, "-m", "pytest", "tests/", "-q"], env=env)
    run([sys.executable, "tests/smoke_test.py"], env=env)

    step("커밋")
    run(["git", "add", "excelmerge/__init__.py", "CHANGELOG.md"])
    run(["git", "commit", "-m",
         f"chore(release): v{args.version} — 버전 bump + CHANGELOG"])

    step("빌드")
    run([sys.executable, "-m", "PyInstaller", "ExcelMerge.spec"], env=env)
    run([sys.executable, "sign.py"], env=env)

    step("exe 실행 스모크")
    exe_smoke(args.version)

    if args.no_publish:
        print(f"\n✔ v{args.version} 로컬 준비 완료 — 푸시/게시는 하지 않았습니다.")
        print("   게시하려면: python release.py {0} 을 --no-publish 없이 다시 "
              "돌리거나,\n   git push 후 python make_release.py --publish --notes ..."
              .format(args.version))
        return

    step("푸시")
    run(["git", "push", "origin", "HEAD:main"])

    step("게시")
    notes = notes_for(read(CHANGELOG), args.version)
    run([sys.executable, "make_release.py", "--publish", "--notes", notes], env=env)

    print(f"\n✔ v{args.version} 배포 완료")


if __name__ == "__main__":
    main()
