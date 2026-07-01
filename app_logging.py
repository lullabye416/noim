# -*- coding: utf-8 -*-
"""
app_logging.py — 운영용 표준 로깅 + Sentry 중앙 수집

설계 목표
---------
1. 모든 로그를 한 형식으로 정규화:
   [시각] [레벨] [vX.Y.Z] [머신/유저] 메시지
2. 날짜별 파일 분리 + 보관기간 제한 (무한 증식 방지)
3. launcher / gui / labor / work_status 전부 동일 로거 사용
4. Sentry 연결 시 모든 미처리 예외 + logging.ERROR 이상 자동 원격 수집
5. Sentry DSN 없거나 오프라인이어도 로컬 로깅은 항상 정상 동작

사용법
------
    from app_logging import setup_logging, get_logger, install_excepthook

    setup_logging(component="gui")   # 프로세스당 1회
    log = get_logger(__name__)
    log.info("시작")
    log.error("문제 발생", exc_info=True)

    # main 진입부:
    install_excepthook()  # 미처리 예외를 로그+Sentry로
"""

import os
import sys
import json
import socket
import getpass
import logging
import platform
import traceback
from pathlib import Path
from logging.handlers import TimedRotatingFileHandler

# ── 경로 ──────────────────────────────────────────────────────────
_APP_DIR  = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "noim"
_LOGS_DIR = _APP_DIR / "logs"

# ── Sentry DSN ────────────────────────────────────────────────────
# 우선순위: 환경변수 > version.json의 sentry_dsn 필드 > 빈 문자열(비활성)
# DSN을 코드에 하드코딩하지 않고 배포 설정으로 주입하는 것을 권장.
_SENTRY_DSN_ENV = "NOIM_SENTRY_DSN"

# 로그 보관 일수
_BACKUP_DAYS = 30

# ── 전역 상태 ─────────────────────────────────────────────────────
_INITIALIZED = False
_APP_VERSION = "0.0.0"
_MACHINE     = "unknown"
_SENTRY_ON   = False


# ── 버전/식별자 수집 ──────────────────────────────────────────────

def _read_version() -> str:
    """version.json에서 버전을 읽는다. scripts 폴더와 현재 폴더 모두 탐색."""
    candidates = [
        _APP_DIR / "scripts" / "version.json",
        _APP_DIR / "version.json",
        Path(__file__).parent / "version.json",
        Path.cwd() / "version.json",
    ]
    for p in candidates:
        try:
            if p.exists():
                with open(p, encoding="utf-8") as f:
                    return json.load(f).get("version", "0.0.0")
        except Exception:
            continue
    return "0.0.0"


def _read_sentry_dsn() -> str:
    """DSN 조회: 환경변수 우선, 없으면 version.json의 sentry_dsn."""
    dsn = os.environ.get(_SENTRY_DSN_ENV, "").strip()
    if dsn:
        return dsn
    for p in [_APP_DIR / "scripts" / "version.json",
              Path(__file__).parent / "version.json",
              Path.cwd() / "version.json"]:
        try:
            if p.exists():
                with open(p, encoding="utf-8") as f:
                    v = json.load(f).get("sentry_dsn", "").strip()
                    if v:
                        return v
        except Exception:
            continue
    return ""


def _machine_id() -> str:
    """PC/유저 식별자 — 여러 PC 로그를 구분하기 위함."""
    try:
        host = socket.gethostname()
    except Exception:
        host = "nohost"
    try:
        user = getpass.getuser()
    except Exception:
        user = "nouser"
    return f"{host}/{user}"


# ── 포매터 ────────────────────────────────────────────────────────

_LEVEL_ALIAS = {
    "DEBUG": "DBG", "INFO": "INF", "WARNING": "WRN",
    "ERROR": "ERR", "CRITICAL": "CRT",
}


class _NormalizedFormatter(logging.Formatter):
    """모든 라인을 동일 형식으로 정규화 — 레벨 3글자 고정폭."""
    def __init__(self):
        super().__init__(
            fmt=("%(asctime)s [%(levelname)s] [v{ver}] [{mac}] "
                 "[%(name)s] %(message)s").format(ver=_APP_VERSION, mac=_MACHINE),
            datefmt="%Y-%m-%d %H:%M:%S",
        )

    def format(self, record):
        record.levelname = _LEVEL_ALIAS.get(record.levelname, record.levelname[:3])
        return super().format(record)


# ── 초기화 ────────────────────────────────────────────────────────

def setup_logging(component: str = "app", level: int = logging.INFO) -> None:
    """
    프로세스당 1회 호출. 컴포넌트별 날짜 파일로 로그를 남기고
    Sentry가 설정돼 있으면 원격 수집을 활성화한다.

    component: 'launcher' | 'gui' | 'worker' 등 — 파일명 접두사
    """
    global _INITIALIZED, _APP_VERSION, _MACHINE, _SENTRY_ON

    _APP_VERSION = _read_version()
    _MACHINE     = _machine_id()

    _LOGS_DIR.mkdir(parents=True, exist_ok=True)

    root = logging.getLogger()
    root.setLevel(level)

    # 중복 핸들러 방지 (재호출 안전)
    for h in list(root.handlers):
        root.removeHandler(h)

    fmt = _NormalizedFormatter()

    # 1) 날짜별 롤링 파일 핸들러 — 자정마다 새 파일, N일 보관
    log_path = _LOGS_DIR / f"{component}.log"
    fh = TimedRotatingFileHandler(
        log_path, when="midnight", interval=1,
        backupCount=_BACKUP_DAYS, encoding="utf-8", delay=True
    )
    fh.suffix = "%Y-%m-%d"          # 롤오버 파일명: component.log.2026-07-01
    fh.setFormatter(fmt)
    fh.setLevel(level)
    root.addHandler(fh)

    # 2) 콘솔 핸들러 (개발/디버깅용)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    ch.setLevel(level)
    root.addHandler(ch)

    # 3) Sentry 연결 (선택)
    _SENTRY_ON = _init_sentry()

    _INITIALIZED = True

    log = get_logger("app_logging")
    log.info("로깅 초기화 | component=%s | version=%s | machine=%s | os=%s | sentry=%s",
             component, _APP_VERSION, _MACHINE, platform.platform(),
             "on" if _SENTRY_ON else "off")


def _init_sentry() -> bool:
    """Sentry SDK 초기화. 실패해도 로컬 로깅은 계속된다."""
    dsn = _read_sentry_dsn()
    if not dsn:
        return False
    try:
        import sentry_sdk
        from sentry_sdk.integrations.logging import LoggingIntegration

        sentry_logging = LoggingIntegration(
            level=logging.INFO,        # breadcrumb로 기록할 최소 레벨
            event_level=logging.ERROR, # 이벤트로 전송할 최소 레벨
        )
        sentry_sdk.init(
            dsn=dsn,
            release=f"noim@{_APP_VERSION}",   # 버전별 에러 집계
            environment=os.environ.get("NOIM_ENV", "production"),
            integrations=[sentry_logging],
            traces_sample_rate=0.0,           # 성능추적 off (에러만)
            send_default_pii=False,
        )
        sentry_sdk.set_tag("machine", _MACHINE)
        sentry_sdk.set_tag("app_version", _APP_VERSION)
        return True
    except Exception as e:
        # Sentry가 없거나 초기화 실패해도 앱은 정상 진행
        logging.getLogger("app_logging").warning("Sentry 초기화 실패 (무시): %s", e)
        return False


def get_logger(name: str = "noim") -> logging.Logger:
    """설정된 로거 반환. setup_logging 전이면 자동 기본 설정."""
    if not _INITIALIZED:
        setup_logging(component="app")
    return logging.getLogger(name)


# ── 미처리 예외 훅 ────────────────────────────────────────────────

def install_excepthook() -> None:
    """
    sys.excepthook을 교체해 모든 미처리 예외를
    파일 로그 + Sentry로 남긴다. (창이 닫혀도 흔적이 남음)
    """
    log = get_logger("excepthook")

    def _hook(exc_type, exc_value, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        log.error("미처리 예외:\n%s", text)
        # Sentry가 있으면 명시적으로 캡처
        if _SENTRY_ON:
            try:
                import sentry_sdk
                sentry_sdk.capture_exception((exc_type, exc_value, exc_tb))
            except Exception:
                pass

    sys.excepthook = _hook


def capture_exception(exc: BaseException = None, **context) -> None:
    """
    코드 내에서 명시적으로 예외를 기록·전송할 때 사용.
    except 블록에서 QMessageBox만 띄우던 자리에 함께 호출.
    """
    log = get_logger("capture")
    if exc is not None:
        log.error("예외 캡처 | %s | %s", context, exc, exc_info=exc)
    else:
        log.error("예외 캡처 | %s", context, exc_info=True)
    if _SENTRY_ON:
        try:
            import sentry_sdk
            with sentry_sdk.push_scope() as scope:
                for k, v in context.items():
                    scope.set_extra(k, v)
                if exc is not None:
                    sentry_sdk.capture_exception(exc)
                else:
                    sentry_sdk.capture_exception()
        except Exception:
            pass