"""
launcher.py — 통짜 exe의 진입점 (운영 개선판)

변경점
------
- 자체 log() 함수 제거 → app_logging 표준 로거 사용
- app_logging.py 를 GitHub 다운로드 목록에 포함 (전 PC 로거 동기화)
- 미처리 예외를 파일 로그 + Sentry로 수집
- 네트워크/버전 실패 원인을 레벨 구분해 기록
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

# GitHub에서 받아올 "로직" 파일 (gui.py는 exe 내장이라 제외)
# app_logging.py 추가 → 로거 자체도 원격 갱신 가능
DOWNLOAD_FILES = [
    "app_logging.py",
    "labor.py",
    "excel_template.py",
    "work_status.py",
    "version.json",
    "sample_노임일보.xlsx",
    "sample_작업현황.xlsx",
]

# ── 부트스트랩 로거 ────────────────────────────────────────────────
# app_logging은 최초 실행 시 아직 SCRIPTS_DIR에 없을 수 있으므로
# import 실패 시 임시 파일 로거로 폴백한 뒤, 다운로드 후 재설정한다.

def _bootstrap_log(msg: str, level: str = "INFO"):
    """app_logging import 이전 단계용 최소 로거."""
    from datetime import datetime
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    line = f"[{datetime.now():%Y-%m-%d %H:%M:%S}] [{level}] [launcher-boot] {msg}"
    try:
        with open(LOGS_DIR / "launcher.log", "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass
    print(line)


# ── 유틸 ──────────────────────────────────────────────────────────

def fetch(url: str, retries: int = 3) -> bytes:
    import time
    last_err = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "noim-launcher"})
            with urllib.request.urlopen(req, timeout=15) as r:
                return r.read()
        except Exception as e:
            last_err = e
            _bootstrap_log(f"다운로드 재시도 {attempt+1}/{retries} 실패: {url} | {type(e).__name__}: {e}", "WARN")
            time.sleep(1.5 * (attempt + 1))
    raise last_err


def raw_url(filename: str) -> str:
    return (
        f"https://github.com/"
        f"{GITHUB_USER}/{GITHUB_REPO}/raw/{GITHUB_BRANCH}/"
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
        _bootstrap_log(f"다운로드: {fname}")
        data = fetch(raw_url(fname))
        (SCRIPTS_DIR / fname).write_bytes(data)
    _bootstrap_log("다운로드 완료")


def update_logic():
    """필요 시 로직 파일 업데이트. 실패해도 기존 파일로 진행."""
    SCRIPTS_DIR.mkdir(parents=True, exist_ok=True)
    LOGS_DIR.mkdir(parents=True, exist_ok=True)

    local_ver = get_local_version()
    _bootstrap_log(f"로컬 버전: {local_ver}")

    is_first = not (SCRIPTS_DIR / "labor.py").exists()

    try:
        remote_ver, changelog = get_remote_version()
        _bootstrap_log(f"원격 버전: {remote_ver}")

        if is_first or version_newer(remote_ver, local_ver):
            _bootstrap_log(f"업데이트: {local_ver} → {remote_ver}")
            download_all()
            VERSION_FILE.write_text(
                json.dumps({"version": remote_ver, "changelog": changelog},
                           ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            _bootstrap_log(f"버전 기록 완료: {remote_ver}")
        else:
            _bootstrap_log("최신 버전입니다.")

    except Exception as e:
        # 네트워크/프록시/방화벽 실패 원인을 명시적으로 기록
        _bootstrap_log(
            f"업데이트 건너뜀 (네트워크 실패?): {type(e).__name__}: {e}", "WARN"
        )
        if is_first:
            raise RuntimeError(
                "최초 실행에는 인터넷 연결이 필요합니다.\n"
                "회사 방화벽이 github.com 접근을 막고 있을 수 있습니다.\n"
                "인터넷/프록시 설정 확인 후 다시 실행해 주세요."
            )


# ── 메인 ──────────────────────────────────────────────────────────

def main():
    _bootstrap_log("=== 런처 시작 ===")
    update_logic()

    # scripts 폴더를 import 경로 최상단에 추가
    sys.path.insert(0, str(SCRIPTS_DIR))
    os.chdir(str(SCRIPTS_DIR))

    # 이제 표준 로거로 승격 (다운로드된 app_logging 사용)
    try:
        from app_logging import setup_logging, get_logger, install_excepthook
        setup_logging(component="launcher")
        install_excepthook()
        log = get_logger("launcher")
        log.info("표준 로깅으로 전환 완료")
    except Exception as e:
        _bootstrap_log(f"app_logging 로드 실패, 부트스트랩 로거 유지: {e}", "WARN")
        log = None

    # exe 설치 폴더를 작업 폴더로 알려줌
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys.executable).parent
    else:
        base_dir = Path(__file__).parent
    os.environ["NOIM_BASE_DIR"] = str(base_dir)

    (log.info if log else _bootstrap_log)(f"작업 폴더: {base_dir}")
    (log.info if log else _bootstrap_log)("gui 실행")

    import gui
    gui.run()

    (log.info if log else _bootstrap_log)("=== 종료 ===")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        try:
            _bootstrap_log(f"런처 오류:\n{err}", "ERROR")
            # Sentry가 로드됐다면 함께 전송
            try:
                from app_logging import capture_exception
                capture_exception(step="launcher_main")
            except Exception:
                pass
        except Exception:
            pass
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("실행 오류", err)
        root.destroy()