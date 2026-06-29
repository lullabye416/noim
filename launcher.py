"""
launcher.py — 통짜 exe의 진입점
- GitHub에서 로직 파일(labor/excel_template/work_status)만 다운로드
- scripts 폴더를 sys.path에 추가
- 내장된 gui.py를 같은 프로세스에서 실행 (subprocess 아님 → 루프 불가)
"""

import os
import sys
import json
import urllib.request
import urllib.parse
import traceback
from pathlib import Path

# ── 설정 ──────────────────────────────────────────────────────────
GITHUB_USER   = "lullabye416"
GITHUB_REPO   = "noim"
GITHUB_BRANCH = "main"

APP_DIR      = Path(os.environ["LOCALAPPDATA"]) / "noim"
SCRIPTS_DIR  = APP_DIR / "scripts"
LOGS_DIR     = APP_DIR / "logs"
VERSION_FILE = APP_DIR / "version.json"

# GitHub에서 받아올 "로직" 파일만 (gui.py는 exe에 내장이라 제외)
DOWNLOAD_FILES = [
    "labor.py",
    "excel_template.py",
    "work_status.py",
    "version.json",
    "sample_노임일보.xlsx",
    "sample_작업현황.xlsx",
]

# ── 유틸 ──────────────────────────────────────────────────────────

def log(msg: str):
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    from datetime import datetime
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    with open(LOGS_DIR / "launcher.log", "a", encoding="utf-8") as f:
        f.write(line + "\n")
    print(msg)


def fetch(url: str, retries: int = 3) -> bytes:
    import time
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": "noim-launcher"}
            )
            with urllib.request.urlopen(req, timeout=15) as r:
                return r.read()
        except Exception as e:
            last_err = e
            time.sleep(1.5 * (attempt + 1))   # 점점 길게 대기 후 재시도
    raise last_err


def raw_url(filename: str) -> str:
    return (
        f"https://raw.githubusercontent.com/"
        f"{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/"
        f"{urllib.parse.quote(filename)}"
    )


# ── 버전 ──────────────────────────────────────────────────────────

def get_local_version() -> str:
    if VERSION_FILE.exists():
        try:
            with open(VERSION_FILE, encoding="utf-8") as f:
                return json.load(f).get("version", "0.0.0")
        except Exception:
            return "0.0.0"
    return "0.0.0"


def get_remote_version():
    data = fetch(raw_url("version.json"))
    info = json.loads(data)
    return info.get("version", "0.0.0"), info.get("changelog", "")


def version_newer(remote: str, local: str) -> bool:
    def parse(v): return tuple(int(x) for x in v.split("."))
    try:
        return parse(remote) > parse(local)
    except Exception:
        return True


# ── 다운로드 ──────────────────────────────────────────────────────

def download_all():
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    for fname in DOWNLOAD_FILES:
        log(f"  다운로드: {fname}")
        data = fetch(raw_url(fname))
        (SCRIPTS_DIR / fname).write_bytes(data)
    log("다운로드 완료")


def update_logic():
    """필요 시 로직 파일 업데이트. 실패해도 기존 파일로 진행."""
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    local_ver = get_local_version()
    log(f"로컬 버전: {local_ver}")

    is_first = not (SCRIPTS_DIR / "labor.py").exists()

    try:
        remote_ver, changelog = get_remote_version()
        log(f"원격 버전: {remote_ver}")

        if is_first or version_newer(remote_ver, local_ver):
            log(f"업데이트: {local_ver} → {remote_ver}")
            download_all()
            # 다운로드 성공 후 버전 기록
            VERSION_FILE.write_text(
                json.dumps({"version": remote_ver, "changelog": changelog},
                           ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            log(f"버전 기록 완료: {remote_ver}")
        else:
            log("최신 버전입니다.")

    except Exception as e:
        log(f"업데이트 건너뜀 (오프라인?): {e}")
        if is_first:
            raise RuntimeError(
                "최초 실행에는 인터넷 연결이 필요합니다.\n"
                "인터넷 연결 후 다시 실행해 주세요."
            )


# ── 메인 ──────────────────────────────────────────────────────────

def main():
    log("=== 런처 시작 ===")
    update_logic()

    # scripts 폴더를 import 경로 최상단에 추가 → 최신 로직을 import
    sys.path.insert(0, str(SCRIPTS_DIR))
    # gui.py가 샘플 파일을 찾을 수 있도록 작업 디렉터리도 이동
    os.chdir(str(SCRIPTS_DIR))

    # exe(런처)가 설치된 폴더를 작업 폴더로 알려줌 → input/output/sample 위치 기준
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent       # C:\노임일보\
    else:
        base_dir = Path(__file__).parent             # 개발 시: 프로젝트 폴더
    os.environ["NOIM_BASE_DIR"] = str(base_dir)
    log(f"작업 폴더: {base_dir}")

    log("gui 실행")
    import gui          # exe에 내장된 gui.py
    gui.run()           # ↓ gui.py에 run() 추가 예정
    log("=== 종료 ===")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        try:
            log(f"런처 오류:\n{err}")
        except Exception:
            pass
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("실행 오류", err)
        root.destroy()