"""앱 자동 업데이트 — HTTPS 매니페스트(latest.json) 기반.

동작: 실행 시 지정 URL의 latest.json을 백그라운드로 읽어 현재 버전보다 높으면 알림 →
원클릭으로 새 exe를 내려받아 자기 자신을 교체·재시작한다. URL 미설정이면 완전 비활성.

새 의존성 없이 표준 라이브러리(urllib)만 사용한다.

배포자 준비물(한 번): 클라우드/사내 웹의 직접 다운로드 위치에
  - ExcelMerge_v<N>.exe
  - latest.json  {"version","url","sha256"(선택),"notes"(선택)}
을 올리고, 그 latest.json URL을 아래 MANIFEST_URL(또는 %APPDATA%/ExcelMerge/update.json)에 설정.
"""
import hashlib
import html as _html
import http.cookiejar
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.request
from urllib.parse import urlparse, urlencode

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMessageBox, QProgressDialog

from . import __version__

# 업데이트 소스 — 둘 중 하나를 설정(둘 다 비면 자동 업데이트 비활성 = 현재 수동 배포 그대로).
#  1) GITHUB_REPO: 공개 repo "owner/name" — releases/latest API를 매니페스트로 사용(권장).
#  2) MANIFEST_URL: 범용 latest.json HTTPS URL.
# 각각 %APPDATA%/ExcelMerge/update.json 의 "github_repo" / "manifest_url" 로도 덮어쓸 수 있다.
GITHUB_REPO = "gywnsdlqkd-Nexus/ExcelMerge"
MANIFEST_URL = ""

_HTTP_TIMEOUT = 5      # 매니페스트 조회 타임아웃(초)
_DL_TIMEOUT = 30       # 다운로드 연결 타임아웃(초)
_DL_CHUNK = 1 << 16    # 64 KiB
_GH_HEADERS = {"User-Agent": "ExcelMerge-Updater",
               "Accept": "application/vnd.github+json"}


# ── 소스 결정 ────────────────────────────────────────────────────────────────
def _config_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "ExcelMerge", "update.json")


def _config() -> dict:
    try:
        p = _config_path()
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8-sig") as f:
                c = json.load(f)
                return c if isinstance(c, dict) else {}
    except Exception:
        pass
    return {}


def _source():
    """업데이트 소스 결정. 반환: (kind, url) 또는 None.
    kind='github' → GitHub Releases API, kind='manifest' → 범용 latest.json.
    설정 파일(github_repo/manifest_url)이 상수보다 우선."""
    cfg = _config()
    repo = (cfg.get("github_repo") or GITHUB_REPO or "").strip().strip("/")
    if repo:
        return ("github", f"https://api.github.com/repos/{repo}/releases/latest")
    url = (cfg.get("manifest_url") or MANIFEST_URL or "").strip()
    if url:
        return ("manifest", url)
    return None


# ── 순수 헬퍼 ────────────────────────────────────────────────────────────────
def _parse_manifest(data: bytes) -> dict:
    """latest.json 바이트를 dict로. version은 필수, 나머지는 선택."""
    m = json.loads(data.decode("utf-8-sig"))
    if not isinstance(m, dict) or not str(m.get("version", "")).strip():
        raise ValueError("매니페스트에 version이 없습니다.")
    return {
        "version": str(m["version"]).strip(),
        "url": (m.get("url") or "").strip(),
        "sha256": (m.get("sha256") or "").strip().lower(),
        "notes": (m.get("notes") or "").strip(),
    }


def _parse_github_release(data: bytes) -> dict:
    """GitHub releases/latest JSON을 매니페스트 형태로 변환.
    version=tag_name(앞 v 제거), notes=body, url=.exe asset의 browser_download_url,
    sha256=asset digest('sha256:...')가 있으면 사용."""
    j = json.loads(data.decode("utf-8"))
    tag = str(j.get("tag_name", "")).strip()
    if not tag:
        raise ValueError("릴리스에 tag_name이 없습니다.")
    version = tag[1:] if tag[:1] in ("v", "V") else tag
    url, sha = "", ""
    for a in (j.get("assets") or []):
        if str(a.get("name", "")).lower().endswith(".exe"):
            url = (a.get("browser_download_url") or "").strip()
            digest = (a.get("digest") or "")
            if digest.lower().startswith("sha256:"):
                sha = digest.split(":", 1)[1].strip().lower()
            break
    return {"version": version, "url": url, "sha256": sha,
            "notes": (j.get("body") or "").strip()}


def is_newer(remote: str, local: str) -> bool:
    """remote가 local보다 새 버전인지. 둘 다 정수면 정수 비교, 아니면 문자열 비교."""
    r, l = str(remote).strip(), str(local).strip()
    if r.isdigit() and l.isdigit():
        return int(r) > int(l)
    return r > l


def build_update_bat(new_exe: str, target_exe: str) -> str:
    """실행 중 exe가 잠겨 있으므로, 잠금이 풀릴 때까지 move 재시도 후 교체·재실행하는 배치.
    경로에 공백이 있어도 안전하도록 모두 따옴표로 감싼다."""
    n = new_exe.replace('"', '')
    t = target_exe.replace('"', '')
    return (
        "@echo off\r\n"
        "setlocal\r\n"
        "set /a tries=0\r\n"
        ":loop\r\n"
        "set /a tries+=1\r\n"
        "ping -n 2 127.0.0.1 >nul\r\n"                 # ~1초 대기(timeout 대체, 무인환경 안전)
        f'move /y "{n}" "{t}" >nul 2>&1\r\n'
        f'if exist "{n}" (\r\n'
        "  if %tries% lss 60 goto loop\r\n"            # 최대 ~60초 재시도
        "  exit /b 1\r\n"
        ")\r\n"
        # PyInstaller 부트로더 변수를 비우고 재실행 — 상속된 값이 있으면 재실행 exe가
        # 추출을 건너뛰어 python3xx.dll 로드에 실패한다. setlocal 범위라 부작용 없음.
        'set "_MEIPASS2="\r\n'
        'set "_MEIPASS="\r\n'
        'set "_PYI_ARCHIVE_FILE="\r\n'
        'set "_PYI_APPLICATION_HOME_DIR="\r\n'
        'set "_PYI_PARENT_PROCESS_LEVEL="\r\n'
        f'start "" "{t}"\r\n'
        'del "%~f0" >nul 2>&1\r\n'
    )


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(_DL_CHUNK), b""):
            h.update(chunk)
    return h.hexdigest()


# ── Google Drive 지원 ────────────────────────────────────────────────────────
# 공유 링크 → 파일 ID → 직접 다운로드 URL 변환 + 대용량(바이러스 검사 초과) 확인 페이지 처리.
_GD_ID_RE = re.compile(r"/d/([A-Za-z0-9_-]+)|[?&]id=([A-Za-z0-9_-]+)")
_GD_FORM_ACTION_RE = re.compile(rb'id="download-form"\s+action="([^"]+)"')
_GD_HIDDEN_RE = re.compile(rb'<input\s+type="hidden"\s+name="([^"]+)"\s+value="([^"]*)"')
_GD_CONFIRM_RE = re.compile(rb'confirm=([0-9A-Za-z_\-]+)')


def gdrive_id_from_url(url: str) -> str:
    """Drive 공유 링크에서 파일 ID 추출(/file/d/ID/ 또는 ?id=ID). 못 찾으면 ""."""
    m = _GD_ID_RE.search(url or "")
    return (m.group(1) or m.group(2)) if m else ""


def gdrive_direct_url(file_id: str) -> str:
    """파일 ID → 직접 다운로드 URL."""
    return f"https://drive.google.com/uc?export=download&id={file_id}"


def _is_gdrive(url: str) -> bool:
    host = urlparse(url).netloc
    return host.endswith("drive.google.com") or host.endswith("drive.usercontent.google.com")


def _open(url: str, timeout: int, headers=None):
    """URL 스트림 열기(리다이렉트·쿠키 추적). headers를 주면 요청에 부착(GitHub API의 UA 등).
    Google Drive 대용량 파일의 '바이러스 검사 확인' HTML 페이지가 오면 폼/토큰을 해석해 재요청."""
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()))
    req = urllib.request.Request(url, headers=headers or {})
    resp = opener.open(req, timeout=timeout)
    if not _is_gdrive(url) or "text/html" not in resp.headers.get("Content-Type", ""):
        return resp
    body = resp.read()
    # 신형: <form id="download-form" action="..."> + hidden inputs
    m = _GD_FORM_ACTION_RE.search(body)
    if m:
        action = _html.unescape(m.group(1).decode())
        params = {k.decode(): _html.unescape(v.decode())
                  for k, v in _GD_HIDDEN_RE.findall(body)}
        sep = "&" if "?" in action else "?"
        return opener.open(action + sep + urlencode(params), timeout=timeout)
    # 구형: confirm 토큰 파라미터
    m = _GD_CONFIRM_RE.search(body)
    if m:
        sep = "&" if "?" in url else "?"
        return opener.open(url + f"{sep}confirm={m.group(1).decode()}", timeout=timeout)
    raise RuntimeError("Google Drive 다운로드 확인 페이지를 처리하지 못했습니다.")


# ── 워커 ─────────────────────────────────────────────────────────────────────
class UpdateCheckWorker(QThread):
    """업데이트 소스를 백그라운드로 조회. 성공 시 done(dict), 실패는 조용히 무시(앱 시작 방해 금지).
    kind='github'이면 릴리스 API 파서, 그 외엔 latest.json 파서를 쓴다."""
    done = pyqtSignal(object)

    def __init__(self, url: str, kind: str = "manifest"):
        super().__init__()
        self.url = url
        self.kind = kind

    def run(self):
        try:
            headers = _GH_HEADERS if self.kind == "github" else None
            with _open(self.url, _HTTP_TIMEOUT, headers) as r:
                data = r.read()
            parse = _parse_github_release if self.kind == "github" else _parse_manifest
            self.done.emit(parse(data))
        except Exception:
            self.done.emit(None)


class UpdateDownloadWorker(QThread):
    """새 exe를 임시 파일로 스트리밍 다운로드(진행률) 후 sha256 검증."""
    progress = pyqtSignal(int, int)   # (받은 바이트, 전체 바이트 — 미상이면 0)
    done = pyqtSignal(str)            # 저장된 임시 exe 경로
    error = pyqtSignal(str)

    def __init__(self, url: str, sha256: str, version: str):
        super().__init__()
        self.url = url
        self.sha256 = (sha256 or "").lower()
        self.version = version
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        tmp = os.path.join(tempfile.gettempdir(),
                           f"ExcelMerge_update_v{self.version}.exe")
        try:
            with _open(self.url, _DL_TIMEOUT) as r:
                total = int(r.headers.get("Content-Length") or 0)
                got = 0
                with open(tmp, "wb") as f:
                    while True:
                        if self._cancel:
                            raise RuntimeError("취소됨")
                        chunk = r.read(_DL_CHUNK)
                        if not chunk:
                            break
                        f.write(chunk)
                        got += len(chunk)
                        self.progress.emit(got, total)
            if self.sha256 and _sha256(tmp) != self.sha256:
                raise ValueError("체크섬 불일치 — 파일이 손상됐거나 변조됐을 수 있습니다.")
            self.done.emit(tmp)
        except Exception as e:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
            self.error.emit(str(e))


# ── 적용 ─────────────────────────────────────────────────────────────────────
# PyInstaller onefile 부트로더가 '이미 추출된 자식'을 알리려고 쓰는 환경변수들.
# 이걸 재실행 exe가 상속하면 압축 해제를 건너뛰고 (종료된 이전 프로세스의) 사라진 임시
# 폴더에서 python3xx.dll을 찾다 실패한다 → 'Failed to load Python DLL … 지정된 모듈을
# 찾을 수 없습니다'. 업데이트 재실행 시 반드시 제거해야 한다(=깨끗한 첫 실행처럼).
_PYI_BOOT_ENV = ("_MEIPASS2", "_MEIPASS")


def _clean_child_env() -> dict:
    """PyInstaller 부트로더 변수를 제거한 자식 프로세스용 환경."""
    return {k: v for k, v in os.environ.items()
            if k not in _PYI_BOOT_ENV and not k.startswith("_PYI_")}


def apply_update(new_exe: str) -> bool:
    """새 exe로 자기 자신을 교체하고 재시작한다(frozen에서만). 성공 시 True(호출측이 앱 종료)."""
    if not getattr(sys, "frozen", False):
        return False   # dev(python 실행)에서는 자기 교체 불가
    target = sys.executable
    bat = os.path.join(tempfile.gettempdir(), "ExcelMerge_apply_update.bat")
    try:
        with open(bat, "w", encoding="mbcs", errors="replace") as f:
            f.write(build_update_bat(new_exe, target))
        # cmd 콘솔 창이 잠깐 떴다 사라지는 것 방지.
        # CREATE_NO_WINDOW 만 사용한다 — DETACHED_PROCESS 와는 상호 배타(동시 지정 시
        # Windows가 콘솔을 붙여 창이 번쩍인다). 추가로 STARTUPINFO 로 창을 숨긴다(벨트+멜빵).
        CREATE_NO_WINDOW = 0x08000000
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = subprocess.SW_HIDE
        # 부트로더 변수를 제거한 환경으로 배치를 띄운다(재실행 exe가 정상 추출하도록).
        subprocess.Popen(["cmd", "/c", bat], close_fds=True,
                         creationflags=CREATE_NO_WINDOW, startupinfo=si,
                         env=_clean_child_env())
        return True
    except Exception:
        return False


# ── 오케스트레이션 ────────────────────────────────────────────────────────────
def check_for_updates(win, silent: bool = True):
    """win(메인 윈도우) 기준 자동/수동 업데이트 확인.
    silent=True(자동): 최신이거나 실패면 조용히. silent=False(수동): 결과를 메시지로 안내."""
    src = _source()
    if not src:
        if not silent:
            QMessageBox.information(
                win, "업데이트",
                "업데이트 소스가 설정되지 않았습니다.\n"
                "updater.GITHUB_REPO(공개 repo) 또는 %APPDATA%\\ExcelMerge\\update.json 을 설정하세요.")
        return

    kind, url = src
    worker = UpdateCheckWorker(url, kind)
    # win에 참조를 보관해 GC/조기 파괴 방지.
    win._update_check_worker = worker
    worker.done.connect(lambda m: _on_manifest(win, m, silent))
    worker.finished.connect(worker.deleteLater)
    worker.start()


def _on_manifest(win, manifest, silent: bool):
    if not manifest:
        if not silent:
            QMessageBox.warning(win, "업데이트", "업데이트 정보를 가져오지 못했습니다(네트워크/URL 확인).")
        return
    if not is_newer(manifest["version"], __version__):
        if not silent:
            QMessageBox.information(win, "업데이트", f"이미 최신 버전입니다 (v{__version__}).")
        return
    if not manifest["url"]:
        if not silent:
            QMessageBox.warning(win, "업데이트", "새 버전은 있으나 다운로드 링크가 없습니다.")
        return

    # 새 버전이 있으면 확인 없이 무조건 업데이트한다(다운로드 후 자동 재시작).
    # 정책: 최신 버전으로만 이용 가능 → '지금 업데이트하시겠습니까?' 확인 창을 두지 않는다.
    if not getattr(sys, "frozen", False):
        # 개발(비패키지) 실행은 자기 교체 불가 — 수동 확인일 때만 안내(자동 시작 시 조용히 무시).
        if not silent:
            QMessageBox.information(
                win, "업데이트",
                "개발(비패키지) 실행 상태에서는 자동 교체를 할 수 없습니다.\n"
                f"수동으로 v{manifest['version']}을 받으세요:\n{manifest['url']}")
        return

    _download_and_apply(win, manifest)


def _download_and_apply(win, manifest):
    dlg = QProgressDialog("업데이트 다운로드 중…", "취소", 0, 100, win)
    dlg.setWindowTitle("업데이트")
    dlg.setAutoClose(False)
    dlg.setAutoReset(False)
    dlg.setMinimumDuration(0)

    dl = UpdateDownloadWorker(manifest["url"], manifest["sha256"], manifest["version"])
    win._update_dl_worker = dl

    def _prog(got, total):
        if total > 0:
            dlg.setMaximum(total)
            dlg.setValue(got)
        else:
            dlg.setMaximum(0)   # 불확정(전체 크기 모름) — 바쁨 표시

    st = {"canceling": False}

    def _err(m):
        dlg.close()
        if st["canceling"]:
            return   # 사용자 취소는 아래 _cancel_and_quit에서 종료 처리(중복 경고 방지)
        # 네트워크 등 실제 실패 — 락아웃을 피하기 위해 경고만 하고 현재 버전으로 계속 진행.
        QMessageBox.warning(win, "업데이트 실패", f"업데이트를 완료하지 못했습니다:\n{m}")

    def _ok(path):
        dlg.close()
        if apply_update(path):
            from PyQt5.QtWidgets import QApplication
            QApplication.quit()   # 배치가 교체 후 재실행
        else:
            QMessageBox.warning(win, "업데이트 실패", "교체 프로세스를 시작하지 못했습니다.")

    def _cancel_and_quit():
        # 정책: 최신 버전으로만 이용 가능 → 업데이트 취소 시 앱을 종료한다
        # (구 버전으로 계속 사용하지 못하게 함). 다시 실행하면 재다운로드된다.
        st["canceling"] = True
        dl.cancel()
        from PyQt5.QtWidgets import QApplication
        QApplication.quit()

    dl.progress.connect(_prog)
    dl.error.connect(_err)
    dl.done.connect(_ok)
    dl.finished.connect(dl.deleteLater)
    dlg.canceled.connect(_cancel_and_quit)
    dl.start()
