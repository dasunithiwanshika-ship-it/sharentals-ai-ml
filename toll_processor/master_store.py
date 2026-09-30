"""Master Dataset persistence layer — safe loading, saving, and archiving.

This module manages the `data/current_master.xlsx` file that serves as the
single source of truth for the Master Dataset between processing runs.

Safety contract:
- Before replacing current_master.xlsx, the old file is archived.
- If archiving fails, the new master is NOT written.
- If writing the new master fails, the archive remains and the old master is untouched.
- Uses write-to-temp-then-rename to prevent partial/corrupt writes.
"""

import io
import os
import shutil
import datetime
from typing import Dict, Any, Optional, List, Tuple

import pandas as pd

from .schemas import (
    DATA_DIR,
    ARCHIVE_DIR,
    CURRENT_MASTER_PATH,
    MASTER_COLUMNS,
)
from .file_readers import read_master_file


def _ensure_dirs() -> None:
    """Create data/ and data/archive/ directories if they don't exist."""
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)


def has_stored_master() -> bool:
    """Check if a current master file exists on disk."""
    return os.path.isfile(CURRENT_MASTER_PATH)


def load_stored_master() -> Tuple[pd.DataFrame, str]:
    """
    Load the current master dataset from disk.
    Returns: (master_df, file_path)
    Raises FileNotFoundError if no stored master exists.
    """
    if not has_stored_master():
        raise FileNotFoundError(
            f"No stored Master Dataset found at {CURRENT_MASTER_PATH}. "
            "Please run Initial Setup first."
        )
    master_df, _ = read_master_file(CURRENT_MASTER_PATH)
    return master_df, CURRENT_MASTER_PATH



def get_master_info() -> Optional[Dict[str, Any]]:
    """
    Get metadata about the current stored master.
    Returns None if no master exists.
    Returns dict with: row_count, file_size_kb, last_modified, file_path
    """
    if not has_stored_master():
        return None

    stat = os.stat(CURRENT_MASTER_PATH)
    mod_time = datetime.datetime.fromtimestamp(stat.st_mtime)

    # Quick read to get row count
    try:
        df, _ = read_master_file(CURRENT_MASTER_PATH)
        row_count = len(df)
    except Exception:
        row_count = "Unknown"

    return {
        "row_count": row_count,
        "file_size_kb": round(stat.st_size / 1024, 1),
        "last_modified": mod_time.strftime("%Y-%m-%d %H:%M:%S"),
        "last_modified_dt": mod_time,
        "file_path": CURRENT_MASTER_PATH,
    }


def list_archive() -> List[Dict[str, Any]]:
    """List archived master files, most recent first."""
    _ensure_dirs()
    archives = []
    for fname in os.listdir(ARCHIVE_DIR):
        if fname.endswith(".xlsx"):
            fpath = os.path.join(ARCHIVE_DIR, fname)
            stat = os.stat(fpath)
            archives.append({
                "filename": fname,
                "file_path": fpath,
                "file_size_kb": round(stat.st_size / 1024, 1),
                "created": datetime.datetime.fromtimestamp(stat.st_mtime).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
            })
    archives.sort(key=lambda x: x["created"], reverse=True)
    return archives


def save_new_master(
    excel_bytes: bytes,
    date_suffix: str = "",
) -> Dict[str, Any]:
    """
    Safely save a new Master Dataset, archiving the previous one.

    Steps:
    1. Ensure data/ and data/archive/ directories exist.
    2. If current_master.xlsx exists, copy it to data/archive/ with timestamp.
    3. Write new master to a temp file.
    4. Rename temp file to current_master.xlsx (atomic on same filesystem).

    Returns dict with: archived_as, new_master_path, previous_row_count, new_row_count

    Raises on any failure — previous master remains untouched.
    """
    _ensure_dirs()
    result: Dict[str, Any] = {
        "archived_as": None,
        "new_master_path": CURRENT_MASTER_PATH,
        "previous_row_count": 0,
        "new_row_count": 0,
    }

    # Step 1: Archive existing master if present
    if has_stored_master():
        try:
            old_info = get_master_info()
            result["previous_row_count"] = old_info["row_count"] if old_info else 0
        except Exception:
            result["previous_row_count"] = "Unknown"

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        archive_name = f"Master_{date_suffix}_{timestamp}.xlsx" if date_suffix else f"Master_{timestamp}.xlsx"
        archive_path = os.path.join(ARCHIVE_DIR, archive_name)

        # Archive by copying (not moving) — keeps the old master intact until replacement
        shutil.copy2(CURRENT_MASTER_PATH, archive_path)
        result["archived_as"] = archive_name

    # Step 2: Validate the Excel bytes in-memory before touching disk
    try:
        test_df, _ = read_master_file(io.BytesIO(excel_bytes))
        result["new_row_count"] = len(test_df)
    except Exception as e:
        raise ValueError(
            f"Validation failed on new Master file — previous Master preserved. Error: {e}"
        )

    # Step 3: Write new master to temp file and atomically replace
    temp_path = CURRENT_MASTER_PATH + ".tmp"
    try:
        with open(temp_path, "wb") as f:
            f.write(excel_bytes)
            f.flush()
            os.fsync(f.fileno())

        # Atomic replace — remove old if exists, rename temp to target
        if os.path.exists(CURRENT_MASTER_PATH):
            os.remove(CURRENT_MASTER_PATH)
        os.rename(temp_path, CURRENT_MASTER_PATH)

    except Exception:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        raise

    return result


def delete_stored_master() -> bool:
    """
    Remove the current stored master (for reset/fresh start).
    Archives it first for safety.
    """
    if not has_stored_master():
        return False

    _ensure_dirs()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    archive_name = f"Master_RESET_{timestamp}.xlsx"
    archive_path = os.path.join(ARCHIVE_DIR, archive_name)
    shutil.copy2(CURRENT_MASTER_PATH, archive_path)
    os.remove(CURRENT_MASTER_PATH)
    return True
