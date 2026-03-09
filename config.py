from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


@dataclass(frozen=True)
class SmtpConfig:
    host: str
    port: int
    email: str
    app_password: str
    sender_name: str | None = None


def sanitize_app_password(value: str) -> str:
    # Google "App Passwords" are often displayed with spaces (and sometimes NBSPs).
    return re.sub(r"\s+", "", (value or "").strip())


def build_smtp_config(
    *,
    email: str,
    app_password: str,
    sender_name: str | None = None,
    host: str = "smtp.gmail.com",
    port: int = 587,
) -> SmtpConfig:
    email = (email or "").strip()
    app_password = sanitize_app_password(app_password)
    sender_name = (sender_name or "").strip() or None

    if not email:
        raise ValueError("Missing SMTP email")
    if not app_password:
        raise ValueError("Missing SMTP app password")

    return SmtpConfig(host=host, port=port, email=email, app_password=app_password, sender_name=sender_name)


def load_smtp_config() -> SmtpConfig:
    """
    Load SMTP configuration from environment variables.

    Supports local development via a `.env` file (loaded if present),
    but the source of truth remains environment variables.
    """
    project_root = Path(__file__).resolve().parents[1]
    dotenv_path = project_root / ".env"
    load_dotenv(dotenv_path=dotenv_path, override=False)

    host = os.getenv("SMTP_HOST", "smtp.gmail.com").strip()
    port_raw = os.getenv("SMTP_PORT", "587").strip()
    email = (os.getenv("SMTP_EMAIL") or "").strip()
    app_password_raw = (os.getenv("SMTP_APP_PASSWORD") or "").strip()
    app_password = sanitize_app_password(app_password_raw)
    sender_name = (os.getenv("SMTP_SENDER_NAME") or "").strip() or None

    try:
        port = int(port_raw)
    except ValueError as e:
        raise ValueError("SMTP_PORT must be an integer") from e

    if dotenv_path.exists() and dotenv_path.stat().st_size == 0 and not (email or app_password_raw):
        raise ValueError("Your .env file exists but is empty. Add SMTP_EMAIL and SMTP_APP_PASSWORD.")

    if not email:
        raise ValueError("Missing SMTP_EMAIL environment variable")
    if not app_password:
        raise ValueError("Missing SMTP_APP_PASSWORD environment variable")

    return SmtpConfig(
        host=host,
        port=port,
        email=email,
        app_password=app_password,
        sender_name=sender_name,
    )

