"""Normalization helpers for rego strings, dates, and monetary values."""

import re
from typing import Any

import pandas as pd


def normalize_rego(value: Any) -> str:
    """Normalize a registration number to uppercase letters/digits only."""
    if value is None:
        return ""
    cleaned = str(value).strip().upper()
    cleaned = re.sub(r"[^A-Z0-9]", "", cleaned)
    return cleaned


def parse_flexible_datetime(value: Any) -> pd.Timestamp | None:
    """Parse common Australian and ISO datetime formats into a Pandas Timestamp."""
    if value is None or value == "":
        return None
    if isinstance(value, pd.Timestamp):
        return value
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return None
        try:
            if re.match(r"^\d{4}[-/]", value):
                parsed = pd.to_datetime(value, errors="raise")
            else:
                parsed = pd.to_datetime(value, dayfirst=True, errors="raise")
            return pd.Timestamp(parsed)
        except Exception:
            return None
    try:
        parsed = pd.Timestamp(value)
        return parsed
    except Exception:
        return None


def normalize_amount(value: Any, make_positive: bool = False) -> float:
    """Convert a currency-like value to a 2-decimal float."""
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        amount = float(value)
    else:
        text = str(value).strip()
        if not text:
            return 0.0
        text = text.replace("$", "").replace(",", "")
        text = text.replace("(", "-").replace(")", "")
        try:
            amount = float(text)
        except ValueError:
            return 0.0
    if make_positive:
        amount = abs(amount)
    return round(amount, 2)


def generate_duplicate_key(rego: Any, dt: Any, amount: Any) -> str:
    """Create a stable duplicate key for a toll record."""
    rego_key = normalize_rego(rego)
    dt_value = parse_flexible_datetime(dt)
    if dt_value is None:
        dt_text = ""
    else:
        dt_text = dt_value.strftime("%Y-%m-%d %H:%M:%S")
    amt_value = normalize_amount(amount, make_positive=True)
    return f"{rego_key}|{dt_text}|{amt_value:.2f}"
