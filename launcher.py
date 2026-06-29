"""
launcher.py
- GitHub에서 최신 버전 확인
- 새 버전이면 파일 자동 다운로드
- gui.py 실행
"""

import os
import sys
import json
import subprocess
import urllib.request
import shutil
import traceback
from pathlib import Path

# ── 설정 ──────────────────────────────────────────────────────────
GITHUB_USER   = "lullabye416"
GITHUB_REPO   = "noim"
GITHUB_BRANCH = "main"
GITHUB_TOKEN  = "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"

APP_DIR     = Path(os.environ["LOCALAPPDATA"]) / "noim"
SCRIPTS_DIR = APP_DIR / "scripts"
LOGS_DIR    = APP_DIR / "logs"
VERSION_FILE = APP_DIR / "version.json"

# GitHub에서 받아올 파일 목록
DOWNLOAD_FILES = [
    "gui.py",
    "labor.py",
    "excel_template.py",
    "work_status.py",
    "version.json",
    "sample_노임일보.xlsx",
    "sample_작업현황.xlsx",
]

# ── 유틸 ──────────────────────────────────────────────────────────

def log(msg: str):
    """로그 파일에 기록"""
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_file = LOGS_DIR / "launcher.log"
    from datetime import datetime
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    print(msg)


def raw_url(filename: str) -> str:
    """GitHub raw 파일 URL 생성"""
    encoded = urllib.parse.quote(filename)
    return (
        f"https://raw.githubusercontent.com/"
        f"{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/{encoded}"
    )


def fetch(url: str) -> bytes:
    """URL에서 파일 다운로드"""
    req = urllib.request.Request(url)
    if GITHUB_TOKEN:
        req.add_header("Authorization", f"token {GITHUB_TOKEN}")
    with urllib.request.urlopen(req, timeout=10) as r:
        return r.read()


# ── 버전 확인 ─────────────────────────────────────────────────────

def get_local_version() -> str:
    if VERSION_FILE.exists():
        with open(VERSION_FILE, encoding="utf-8") as f:
            return json.load(f).get("version", "0.0.0")
    return "0.0.0"


def get_remote_version() -> tuple[str, str]:
    """GitHub의 version.json에서 최신 버전과 changelog 반환"""
    import urllib.parse
    data = fetch(raw_url("version.json"))
    info = json.loads(data)
    return info.get("version", "0.0.0"), info.get("changelog", "")


def version_newer(remote: str, local: str) -> bool:
    """remote가 local보다 새 버전이면 True"""
    def parse(v): return tuple(int(x) for x in v.split("."))
    return parse(remote) > parse(local)


# ── 업데이트 ──────────────────────────────────────────────────────

def download_all():
    """DOWNLOAD_FILES 전부 SCRIPTS_DIR에 다운로드"""
    import urllib.parse
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    for fname in DOWNLOAD_FILES:
        url = raw_url(fname)
        log(f"  다운로드: {fname}")
        data = fetch(url)
        dest = SCRIPTS_DIR / fname
        dest.write_bytes(data)
    log("업데이트 완료")


# ── 메인 ──────────────────────────────────────────────────────────

def main():
    import urllib.parse
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    log("=== 런처 시작 ===")

    # 1. 버전 확인
    local_ver = get_local_version()
    log(f"로컬 버전: {local_ver}")

    try:
        remote_ver, changelog = get_remote_version()
        log(f"원격 버전: {remote_ver}")

        if version_newer(remote_ver, local_ver):
            log(f"새 버전 발견: {local_ver} → {remote_ver}")

            # 업데이트 알림 팝업
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showinfo(
                "업데이트",
                f"새 버전 {remote_ver}을 다운로드합니다.\n\n{changelog}"
            )
            root.destroy()

            download_all()

            # 로컬 버전 파일 갱신
            VERSION_FILE.write_text(
                json.dumps({"version": remote_ver, "changelog": changelog},
                           ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
        else:
            log("최신 버전입니다.")

            # scripts 폴더가 비어있으면 (첫 실행) 무조건 다운로드
            if not (SCRIPTS_DIR / "gui.py").exists():
                log("scripts 폴더 없음 → 초기 다운로드")
                download_all()
                VERSION_FILE.parent.mkdir(parents=True, exist_ok=True)
                VERSION_FILE.write_text(
                    json.dumps({"version": remote_ver, "changelog": changelog},
                               ensure_ascii=False, indent=2),
                    encoding="utf-8"
                )

    except Exception as e:
        log(f"버전 확인 실패 (오프라인?): {e}")
        # 오프라인이어도 로컬 파일로 실행은 계속
        if not (SCRIPTS_DIR / "gui.py").exists():
            log("오프라인 + scripts 없음 → 실행 불가")
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk(); root.withdraw()
            messagebox.showerror("오류", "인터넷 연결이 필요합니다.\n(최초 실행 시 1회 필요)")
            root.destroy()
            return

    # 2. gui.py 실행
    gui_path = SCRIPTS_DIR / "gui.py"
    log(f"실행: {gui_path}")
    subprocess.run([sys.executable, str(gui_path)], cwd=str(SCRIPTS_DIR))
    log("=== 종료 ===")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        log(f"런처 오류:\n{err}")
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("런처 오류", err)
        root.destroy()