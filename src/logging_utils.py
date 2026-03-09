from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone


@dataclass(frozen=True)
class LogEvent:
    ts_iso: str
    level: str
    message: str


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def info(message: str) -> LogEvent:
    return LogEvent(ts_iso=now_iso(), level="INFO", message=message)


def warn(message: str) -> LogEvent:
    return LogEvent(ts_iso=now_iso(), level="WARN", message=message)


def error(message: str) -> LogEvent:
    return LogEvent(ts_iso=now_iso(), level="ERROR", message=message)

