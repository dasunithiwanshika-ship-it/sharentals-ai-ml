"""Driver matching logic for LINKT tolls against 365 bookings."""

from __future__ import annotations

from typing import Any, Dict

import pandas as pd

from .normalizer import normalize_rego
from .schemas import STATUS_MULTIPLE_MATCH, STATUS_UNMATCHED


def match_driver_for_toll(toll_row: pd.Series, bookings_df: pd.DataFrame) -> Dict[str, Any]:
    """Match the toll to a single active booking or mark it unmatched."""
    if bookings_df is None or bookings_df.empty:
        return {
            "matched": False,
            "driver": "UNMATCHED",
            "status": STATUS_UNMATCHED,
            "reason": "No 365 booking data available",
        }

    raw_rego = toll_row.get("raw_rego", toll_row.get("REGO", ""))
    norm_rego = toll_row.get("norm_rego", normalize_rego(raw_rego))
    toll_dt = toll_row.get("parsed_start_dt")
    if toll_dt is None:
        toll_dt = toll_row.get("StartDateTime")
    if toll_dt is None:
        return {
            "matched": False,
            "driver": "UNMATCHED",
            "status": STATUS_UNMATCHED,
            "reason": "No toll date available",
        }

    df = bookings_df.copy()
    for col in ["norm_rego", "Norm Rego", "Rego No.", "rego", "REGO", "vehicle_rego"]:
        if col in df.columns:
            df["norm_rego"] = df[col].map(normalize_rego)
            break
    else:
        df["norm_rego"] = ""

    for col in ["start_dt", "Start Date", "StartDate", "Hire Date", "start_date"]:
        if col in df.columns:
            df["start_dt"] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
            break
    else:
        df["start_dt"] = pd.NaT

    for col in ["finish_dt", "Finish Date", "FinishDate", "Return Date", "end_date"]:
        if col in df.columns:
            df["finish_dt"] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
            break
    else:
        df["finish_dt"] = pd.NaT

    candidate = df[df["norm_rego"] == norm_rego]
    if candidate.empty:
        return {
            "matched": False,
            "driver": "UNMATCHED",
            "status": STATUS_UNMATCHED,
            "reason": f"No matching Rego for {raw_rego} in 365 bookings",
        }

    active = candidate[(candidate["start_dt"].notna()) & (candidate["finish_dt"].notna())]
    active = active[(toll_dt >= active["start_dt"]) & (toll_dt <= active["finish_dt"])]

    if active.empty:
        return {
            "matched": False,
            "driver": "UNMATCHED",
            "status": STATUS_UNMATCHED,
            "reason": "No active booking at toll time",
        }

    if len(active) > 1:
        return {
            "matched": False,
            "driver": "UNMATCHED",
            "status": STATUS_MULTIPLE_MATCH,
            "reason": "Multiple active bookings found for this rego at the toll time",
        }

    row = active.iloc[0]

    driver_name = row.get("client")
    if driver_name is None or str(driver_name).strip() == "":
        driver_name = row.get("Client")
    if driver_name is None or str(driver_name).strip() == "":
        driver_name = "Unknown Driver"

    booking_ref = row.get("booking_ref")
    if booking_ref is None or str(booking_ref).strip() == "":
        booking_ref = row.get("Booking")
    if booking_ref is None or str(booking_ref).strip() == "":
        booking_ref = row.get("Reference")
    if booking_ref is None or str(booking_ref).strip() == "":
        booking_ref = ""

    hire_date = row.get("start_dt")
    if hire_date is None:
        hire_date = row.get("Start Date")

    return_date = row.get("finish_dt")
    if return_date is None:
        return_date = row.get("Finish Date")

    contact = row.get("contact")
    if contact is None or str(contact).strip() == "":
        contact = row.get("Contact")

    return {
        "matched": True,
        "driver": driver_name,
        "booking_ref": booking_ref,
        "hire_date": hire_date,
        "return_date": return_date,
        "contact": contact,
        "status": "MATCHED",
        "reason": "Matched to a single valid 365 booking",
    }
