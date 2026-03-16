from __future__ import annotations

from dataclasses import dataclass
import io
import re
from typing import Iterable

import pandas as pd
import requests


@dataclass(frozen=True)
class Recipient:
    name: str
    email: str
    fields: dict[str, str]


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def _require_columns(df: pd.DataFrame, required: Iterable[str]) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")


def load_recipients_from_upload(filename: str, content: bytes) -> list[Recipient]:
    """
    Load recipients from a CSV/XLSX/XLS upload.

    Expects columns `Name` and `Email` (case-insensitive).
    """
    lower = filename.lower()
    if lower.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(content))
    elif lower.endswith(".xlsx") or lower.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(content))
    else:
        raise ValueError("Unsupported file type. Upload a CSV or Excel file.")

    df = _normalize_columns(df)
    _require_columns(df, required=["name", "email"])

    df = df.dropna(subset=["email"])
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()
    df = df[(df["email"] != "")]
    df["name"] = df.get("name", "").astype(str).str.strip()

    recipients: list[Recipient] = []
    for _, row in df.iterrows():
        fields = {str(k).strip().lower(): ("" if pd.isna(v) else str(v).strip()) for k, v in row.items()}
        # Ensure baseline keys exist.
        fields.setdefault("name", fields.get("name", ""))
        fields.setdefault("email", fields.get("email", ""))
        recipients.append(Recipient(name=fields["name"], email=fields["email"], fields=fields))

    if not recipients:
        raise ValueError("No valid recipients found after cleaning.")

    return recipients


_SHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")
_GID_RE = re.compile(r"(?:\?|#|&|/)(?:gid=)(\d+)")


def google_sheet_to_csv_url(sheet_url: str) -> str:
    """
    Convert a Google Sheets share URL into a public CSV export URL.

    Requires the sheet to be shared as "Anyone with the link can view".
    """
    url = (sheet_url or "").strip()
    m = _SHEET_ID_RE.search(url)
    if not m:
        raise ValueError("Invalid Google Sheets link. It should contain '/spreadsheets/d/<sheetId>'.")

    sheet_id = m.group(1)
    gid_match = _GID_RE.search(url)
    gid = gid_match.group(1) if gid_match else "0"

    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"


def load_recipients_from_google_sheet(sheet_url: str) -> list[Recipient]:
    csv_url = google_sheet_to_csv_url(sheet_url)
    try:
        resp = requests.get(csv_url, timeout=20)
    except Exception as e:
        raise ValueError(f"Failed to fetch Google Sheet. Error: {e}") from e

    if resp.status_code != 200:
        raise ValueError(
            "Failed to fetch Google Sheet. "
            "Make sure the sheet is shared as 'Anyone with the link can view'."
        )

    return load_recipients_from_upload("sheet.csv", resp.content)

