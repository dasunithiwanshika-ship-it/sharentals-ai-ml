"""Readers for LINKT, 365 booking, and master Excel files."""

from __future__ import annotations

import io
import os
from typing import Any, BinaryIO, Dict, Tuple, Union

import pandas as pd

from .normalizer import generate_duplicate_key, normalize_amount, normalize_rego, parse_flexible_datetime


def _coerce_file(file: Union[str, bytes, BinaryIO, os.PathLike, None]) -> Any:
    if file is None:
        return pd.DataFrame()
    if isinstance(file, (str, os.PathLike)):
        path = str(file)
        if not os.path.exists(path):
            return pd.DataFrame()
        if path.lower().endswith(".csv"):
            return pd.read_csv(path)
        return pd.read_excel(path)
    if isinstance(file, bytes):
        return pd.read_excel(io.BytesIO(file))
    if hasattr(file, "read"):
        pos = getattr(file, "tell", lambda: None)()
        data = file.read()
        if isinstance(data, bytes):
            return pd.read_excel(io.BytesIO(data))
        if hasattr(file, "seek"):
            try:
                file.seek(pos)
            except Exception:
                pass
        return pd.DataFrame(data)
    return pd.DataFrame(file)


def _read_input_frame(file: Union[str, bytes, BinaryIO, os.PathLike, None]) -> pd.DataFrame:
    df = _coerce_file(file)
    if df is None or df.empty:
        return pd.DataFrame()
    return df


def read_master_file(file: Union[str, bytes, BinaryIO, os.PathLike, None]) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Read the master dataset from a file path, bytes, or stream."""
    df = _read_input_frame(file)
    return df, {}


def read_linkt_file(file: Union[str, bytes, BinaryIO, os.PathLike, None]) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Read and normalize the LINKT import into the internal toll record schema."""
    df = _read_input_frame(file)
    if df.empty:
        return df, {}

    if "LPN" in df.columns:
        rego_col = "LPN"
    elif "Registration" in df.columns:
        rego_col = "Registration"
    elif "Vehicle Rego" in df.columns:
        rego_col = "Vehicle Rego"
    else:
        rego_col = next((c for c in df.columns if "rego" in str(c).lower()), None)

    start_col = next((c for c in df.columns if "start" in str(c).lower() and "date" in str(c).lower()), None)
    if start_col is None:
        start_col = next((c for c in df.columns if "start" in str(c).lower()), None)
    end_col = next((c for c in df.columns if "end" in str(c).lower() and "date" in str(c).lower()), None)
    if end_col is None:
        end_col = next((c for c in df.columns if "end" in str(c).lower()), None)
    amt_col = next((c for c in df.columns if "amount" in str(c).lower()), None)

    df = df.copy()
    df["raw_rego"] = df[rego_col].fillna("") if rego_col else ""
    df["norm_rego"] = df["raw_rego"].map(normalize_rego)
    df["parsed_start_dt"] = pd.to_datetime(df[start_col], dayfirst=True, errors="coerce") if start_col else pd.NaT
    df["parsed_end_dt"] = pd.to_datetime(df[end_col], dayfirst=True, errors="coerce") if end_col else pd.NaT
    df["toll_amount"] = df[amt_col].map(normalize_amount) if amt_col else 0.0
    df["type"] = df.get("Type", df.get("Type of activity", "Toll")) if "Type" in df.columns or "Type of activity" in df.columns else "Toll"
    df["details"] = df.get("Details", df.get("Description", ""))
    df["concession"] = df.get("Concession", "")
    df["dup_key"] = df.apply(lambda row: generate_duplicate_key(row["raw_rego"], row["parsed_start_dt"], row["toll_amount"]), axis=1)

    return df, {"rego_col": rego_col or "", "start_col": start_col or "", "end_col": end_col or "", "amount_col": amt_col or ""}


def read_365_booking_file(file: Union[str, bytes, BinaryIO, os.PathLike, None]) -> Tuple[pd.DataFrame, Dict[str, str]]:
    """Read the 365 bookings file and normalize booking date columns."""
    df = _read_input_frame(file)
    if df.empty:
        return df, {}

    rego_col = next((c for c in df.columns if "rego" in str(c).lower()), None)
    if rego_col is None:
        rego_col = next((c for c in df.columns if "plate" in str(c).lower() or "license" in str(c).lower()), None)

    start_col = next((c for c in df.columns if "start" in str(c).lower() and "date" in str(c).lower()), None)
    if start_col is None:
        start_col = next((c for c in df.columns if "start" in str(c).lower()), None)
    finish_col = next((c for c in df.columns if "finish" in str(c).lower() and "date" in str(c).lower()), None)
    if finish_col is None:
        finish_col = next((c for c in df.columns if "finish" in str(c).lower()), None)
    client_col = next((c for c in df.columns if "client" in str(c).lower()), None)
    booking_ref_col = next((c for c in df.columns if "booking" in str(c).lower() or "ref" in str(c).lower()), None)
    contact_col = next((c for c in df.columns if "contact" in str(c).lower() or "phone" in str(c).lower()), None)

    df = df.copy()
    if rego_col:
        df["norm_rego"] = df[rego_col].map(normalize_rego)
    else:
        df["norm_rego"] = ""
    if start_col:
        df["start_dt"] = pd.to_datetime(df[start_col], dayfirst=True, errors="coerce")
    else:
        df["start_dt"] = pd.NaT
    if finish_col:
        df["finish_dt"] = pd.to_datetime(df[finish_col], dayfirst=True, errors="coerce")
    else:
        df["finish_dt"] = pd.NaT
    df["client"] = df[client_col] if client_col else ""
    df["booking_ref"] = df[booking_ref_col] if booking_ref_col else ""
    df["contact"] = df[contact_col] if contact_col else ""

    return df, {
        "rego_col": rego_col or "",
        "start_col": start_col or "",
        "finish_col": finish_col or "",
        "client_col": client_col or "",
        "booking_ref_col": booking_ref_col or "",
        "contact_col": contact_col or "",
    }
