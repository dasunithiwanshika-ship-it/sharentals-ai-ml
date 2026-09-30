"""Duplicate detection and historical master cleanup."""

from __future__ import annotations

from typing import Any, Dict, List, Tuple

import pandas as pd

from .normalizer import generate_duplicate_key, normalize_rego, parse_flexible_datetime
from .schemas import STATUS_ALREADY_PAID, STATUS_DUPLICATE


class DuplicateDetector:
    """Tracks historical duplicate keys and current-batch duplicates."""

    def __init__(self, master_df: pd.DataFrame):
        self.master_df = master_df.copy() if master_df is not None else pd.DataFrame()
        self._batch_seen: set[str] = set()
        self._master_index: Dict[str, Dict[str, Any]] = {}

        if self.master_df.empty:
            return

        for _, row in self.master_df.iterrows():
            key = self._row_to_key(row)
            if key:
                self._master_index.setdefault(key, row.to_dict())

    def _row_to_key(self, row: pd.Series) -> str | None:
        rego = row.get("REGO", row.get("rego", ""))
        start_dt = row.get("StartDateTime", row.get("start_datetime", row.get("Date", "")))
        amount = row.get("Toll Amount", row.get("toll_amount", 0.0))
        if rego is None or str(rego).strip() == "":
            return None
        dt = parse_flexible_datetime(start_dt)
        if dt is None:
            return None
        return generate_duplicate_key(rego, dt, amount)

    def check_duplicate(self, dup_key: str) -> Tuple[bool, str | None, str | None]:
        """Return (is_duplicate, status, reason)."""
        if dup_key in self._batch_seen:
            return True, STATUS_DUPLICATE, "Duplicate record already found in this batch"

        if dup_key in self._master_index:
            prior = self._master_index[dup_key]
            payment_status = str(prior.get("Payment Status", prior.get("payment_status", ""))).strip().lower()
            if payment_status in {"paid", "card", "cash", "settled", "already paid"} or "paid" in payment_status or "settled" in payment_status:
                self._batch_seen.add(dup_key)
                return True, STATUS_ALREADY_PAID, "Already paid in historical master data"
            self._batch_seen.add(dup_key)
            return True, STATUS_DUPLICATE, "Duplicate toll already exists in master data"

        self._batch_seen.add(dup_key)
        return False, None, None


def clean_master_dataframe(master_df: pd.DataFrame) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
    """Remove duplicate rows from historical master data while keeping the first record."""
    if master_df is None or master_df.empty:
        return master_df.copy(), []

    cleaned_rows: List[pd.Series] = []
    purged: List[Dict[str, Any]] = []
    seen: set[str] = set()

    for _, row in master_df.iterrows():
        key = generate_duplicate_key(
            row.get("REGO", row.get("rego", "")),
            row.get("StartDateTime", row.get("start_datetime", row.get("Date", ""))),
            row.get("Toll Amount", row.get("toll_amount", 0.0)),
        )
        if key in seen:
            purged.append(row.to_dict())
            continue
        seen.add(key)
        cleaned_rows.append(row)

    if not cleaned_rows:
        return pd.DataFrame(columns=master_df.columns), purged

    return pd.DataFrame(cleaned_rows).reset_index(drop=True), purged
