from __future__ import annotations

from dataclasses import dataclass
import io
from typing import Iterable

import pandas as pd


@dataclass(frozen=True)
class Recipient:
    name: str
    email: str


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

    df = df[["name", "email"]].dropna()
    df["name"] = df["name"].astype(str).str.strip()
    df["email"] = df["email"].astype(str).str.strip()

    df = df[(df["name"] != "") & (df["email"] != "")]

    recipients: list[Recipient] = []
    for _, row in df.iterrows():
        recipients.append(Recipient(name=row["name"], email=row["email"]))

    if not recipients:
        raise ValueError("No valid recipients found after cleaning.")

    return recipients

