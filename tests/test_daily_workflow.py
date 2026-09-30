"""Tests for Master Store persistence and Daily Workflow carry-forward."""

import os
import shutil
import tempfile
import pandas as pd
import pytest

from toll_processor.sample_generator import generate_sample_datasets
from toll_processor.processor import process_daily_run, process_toll_files
from toll_processor import master_store
from toll_processor.schemas import CURRENT_MASTER_PATH, ARCHIVE_DIR, DATA_DIR


@pytest.fixture
def clean_data_env(tmp_path, monkeypatch):
    """Isolate data directory to a temporary path for testing."""
    test_data_dir = tmp_path / "data"
    test_archive_dir = test_data_dir / "archive"
    test_master_path = test_data_dir / "current_master.xlsx"

    monkeypatch.setattr(master_store, "DATA_DIR", str(test_data_dir))
    monkeypatch.setattr(master_store, "ARCHIVE_DIR", str(test_archive_dir))
    monkeypatch.setattr(master_store, "CURRENT_MASTER_PATH", str(test_master_path))

    return {
        "data_dir": str(test_data_dir),
        "archive_dir": str(test_archive_dir),
        "master_path": str(test_master_path),
    }


def test_initial_setup_and_carry_forward(clean_data_env, tmp_path):
    # 1. Generate realistic test files
    sample_dir = tmp_path / "samples"
    linkt_file, bookings_file, master_file = generate_sample_datasets(str(sample_dir))

    # Read original master row count
    orig_master_df, _ = master_store.read_master_file(master_file)
    orig_rows = len(orig_master_df)

    assert not master_store.has_stored_master()

    # Day 1: Initial Setup
    res_day1, save_info_day1 = process_daily_run(
        linkt_file=linkt_file,
        bookings_365_file=bookings_file,
        master_file=master_file,
        date_received="2026-09-08",
        admin_fee=5.55,
        is_initial_setup=True,
    )

    assert master_store.has_stored_master()
    assert save_info_day1["archived_as"] is None
    new_tolls_day1 = len(res_day1.new_tolls_df)
    assert len(res_day1.master_df) == orig_rows + new_tolls_day1
    assert save_info_day1["new_row_count"] == orig_rows + new_tolls_day1

    # Verify info
    info = master_store.get_master_info()
    assert info is not None
    assert info["row_count"] == orig_rows + new_tolls_day1

    # Day 2: Run AGAIN with the SAME LINKT file (Simulate user re-uploading or historical overlap)
    # This must produce 0 NEW chargeable tolls because all were added to current_master
    res_day2, save_info_day2 = process_daily_run(
        linkt_file=linkt_file,
        bookings_365_file=bookings_file,
        master_file=None,  # auto-load current_master
        date_received="2026-09-09",
        admin_fee=5.55,
        is_initial_setup=False,
    )

    # 0 new chargeable tolls
    assert len(res_day2.new_tolls_df) == 0
    assert res_day2.summary_stats["new_tolls_count"] == 0
    assert res_day2.summary_stats["total_admin_fee"] == 0.0
    assert res_day2.summary_stats["total_customer_amount"] == 0.0

    # All previously added tolls are now detected as duplicates!
    assert len(res_day2.duplicates_df) >= new_tolls_day1

    # Master row count remained exactly the same
    assert len(res_day2.master_df) == orig_rows + new_tolls_day1
    assert save_info_day2["archived_as"] is not None

    # Archive folder has the Day 1 version
    archives = master_store.list_archive()
    assert len(archives) == 1
    assert archives[0]["filename"] == save_info_day2["archived_as"]


def test_unmatched_tolls_not_charged(clean_data_env, tmp_path):
    sample_dir = tmp_path / "samples"
    linkt_file, bookings_file, master_file = generate_sample_datasets(str(sample_dir))

    # Initial setup
    res, _ = process_daily_run(
        linkt_file=linkt_file,
        bookings_365_file=bookings_file,
        master_file=master_file,
        date_received="2026-09-08",
        admin_fee=5.55,
        is_initial_setup=True,
    )

    if not res.unmatched_df.empty:
        # None of the unmatched tolls should appear in new_tolls_df or customer_summary_df
        matched_drivers = set(res.customer_summary_df["Driver"].tolist())
        assert "UNMATCHED" not in matched_drivers
        assert "Unknown" not in matched_drivers
