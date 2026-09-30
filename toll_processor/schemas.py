"""Shared constants, statuses, and workbook metadata for the toll processor."""

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = str(BASE_DIR / "data")
ARCHIVE_DIR = str(Path(DATA_DIR) / "archive")
CURRENT_MASTER_PATH = str(Path(DATA_DIR) / "current_master.xlsx")

MASTER_COLUMNS = [
    "Unique Key",
    "Duplicate Count",
    "Date Received",
    "Date",
    "Time",
    "StartDateTime",
    "EndDateTime",
    "Type of activity",
    "REGO",
    "Details",
    "Concession",
    "Toll Amount",
    "Admin Fee",
    "TOTAL AMOUNT ",
    "Driver",
    "Hire Date",
    "Return Date",
    "invoice added to 365",
    "Invoice\nsent to customer",
    "Payment Method",
    "Payment Status",
    "Follow-up Date",
    "Source",
]

STATUS_NEW = "NEW"
STATUS_DUPLICATE = "DUPLICATE"
STATUS_ALREADY_PAID = "ALREADY_PAID"
STATUS_UNMATCHED = "UNMATCHED"
STATUS_MULTIPLE_MATCH = "MULTIPLE MATCH"

DEFAULT_ADMIN_FEE = 5.55
DEFAULT_DATE_RECEIVED = "2026-09-08"
DEFAULT_START_DATE_STR = "08/09/2026"

SHEET_MASTER_DATA = "Paste_eToll_Data"
SHEET_NEW_TOLLS = "New Tolls"
SHEET_CUSTOMER_SUMMARY = "Customer Summary"
SHEET_DUPLICATES = "Duplicates"
SHEET_UNMATCHED = "Unmatched"
SHEET_ALREADY_PAID = "Already Paid"
SHEET_SUMMARY = "Summary"
