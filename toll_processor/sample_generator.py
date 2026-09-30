"""Sample dataset generator for testing and user demonstrations."""

import os
from typing import Tuple
import pandas as pd
import openpyxl

from .schemas import MASTER_COLUMNS


def generate_sample_datasets(output_dir: str = "sample_data") -> Tuple[str, str, str]:
    """
    Generate sample LINKT, 365 Booking, and Master Dataset Excel files.
    Returns: (linkt_filepath, bookings_filepath, master_filepath)
    """
    os.makedirs(output_dir, exist_ok=True)

    linkt_file = os.path.join(output_dir, "sample_linkt_report.xlsx")
    bookings_file = os.path.join(output_dir, "sample_365_bookings.xlsx")
    master_file = os.path.join(output_dir, "sample_master_dataset.xlsx")

    # 1. 365 Bookings Data
    bookings_data = [
        {
            "Rego No.": "ABC123",
            "Date": "01/09/2026",
            "Booking": "BK-2026-001",
            "Reference": "REF-8801",
            "Client": "John Smith",
            "Start Date": "07/09/2026 09:00:00",
            "Finish Date": "12/09/2026 17:00:00",
            "Rent": 450.00,
            "Duration (Weeks)": 1.0,
            "Contact": "+61 412 345 678",
        },
        {
            "Rego No.": "XYZ789",
            "Date": "02/09/2026",
            "Booking": "BK-2026-002",
            "Reference": "REF-8802",
            "Client": "Sarah Connor",
            "Start Date": "08/09/2026 08:00:00",
            "Finish Date": "15/09/2026 18:00:00",
            "Rent": 600.00,
            "Duration (Weeks)": 1.0,
            "Contact": "+61 423 456 789",
        },
        {
            "Rego No.": "TOLL99",
            "Date": "05/09/2026",
            "Booking": "BK-2026-003",
            "Reference": "REF-8803",
            "Client": "Michael Brown",
            "Start Date": "08/09/2026 10:00:00",
            "Finish Date": "09/09/2026 12:00:00",
            "Rent": 200.00,
            "Duration (Weeks)": 0.5,
            "Contact": "+61 434 567 890",
        },
        {
            "Rego No.": "OVERLAP1",
            "Date": "06/09/2026",
            "Booking": "BK-2026-004A",
            "Reference": "REF-8804A",
            "Client": "David Miller",
            "Start Date": "08/09/2026 06:00:00",
            "Finish Date": "10/09/2026 18:00:00",
            "Rent": 300.00,
            "Duration (Weeks)": 0.5,
            "Contact": "+61 445 678 901",
        },
        {
            "Rego No.": "OVERLAP1",
            "Date": "06/09/2026",
            "Booking": "BK-2026-004B",
            "Reference": "REF-8804B",
            "Client": "Emma Wilson",
            "Start Date": "08/09/2026 12:00:00",
            "Finish Date": "11/09/2026 12:00:00",
            "Rent": 320.00,
            "Duration (Weeks)": 0.5,
            "Contact": "+61 456 789 012",
        },
        {
            "Rego No.": "HIST001",
            "Date": "20/08/2026",
            "Booking": "BK-2026-000",
            "Reference": "REF-8800",
            "Client": "Alice Walker",
            "Start Date": "20/08/2026 09:00:00",
            "Finish Date": "25/08/2026 17:00:00",
            "Rent": 500.00,
            "Duration (Weeks)": 1.0,
            "Contact": "+61 467 890 123",
        },
    ]
    df_bookings = pd.DataFrame(bookings_data)
    df_bookings.to_excel(bookings_file, index=False)

    # 2. LINKT Toll Data
    # Includes new valid matches, negative amounts, duplicate from historical master,
    # already-paid toll, multiple-match overlap, unmatched rego, and within-batch duplicate.
    linkt_data = [
        # Match with John Smith (ABC123) - Negative amount converted to positive
        {
            "Start Date": "08/09/2026 10:30:00",
            "End Date": "08/09/2026 10:35:00",
            "Type": "Toll",
            "Details": "M2 Hills Motorway - Macquarie Park",
            "LPN": "ABC 123",  # Spaces will be normalized
            "Tag Number": "TAG-1001",
            "Fleet ID": "FL-01",
            "Vehicle Class": "Car",
            "Amount": -7.86,
        },
        # Another trip by John Smith (ABC123)
        {
            "Start Date": "08/09/2026 16:45:00",
            "End Date": "08/09/2026 16:50:00",
            "Type": "Toll",
            "Details": "Lane Cove Tunnel - Westbound",
            "LPN": "ABC123",
            "Tag Number": "TAG-1001",
            "Fleet ID": "FL-01",
            "Vehicle Class": "Car",
            "Amount": -3.54,
        },
        # Match with Sarah Connor (XYZ789)
        {
            "Start Date": "08/09/2026 14:15:00",
            "End Date": "08/09/2026 14:20:00",
            "Type": "Toll",
            "Details": "WestConnex M4 - Parramatta",
            "LPN": "XYZ789",
            "Tag Number": "TAG-1002",
            "Fleet ID": "FL-01",
            "Vehicle Class": "Car",
            "Amount": -9.20,
        },
        # Match with Michael Brown (TOLL99)
        {
            "Start Date": "08/09/2026 11:00:00",
            "End Date": "08/09/2026 11:05:00",
            "Type": "Toll",
            "Details": "Sydney Harbour Bridge - Southbound",
            "LPN": "TOLL99",
            "Tag Number": "TAG-1003",
            "Fleet ID": "FL-02",
            "Vehicle Class": "Car",
            "Amount": -4.00,
        },
        # Overlapping bookings -> MULTIPLE MATCH (OVERLAP1 at 08/09/2026 14:00)
        {
            "Start Date": "08/09/2026 14:00:00",
            "End Date": "08/09/2026 14:10:00",
            "Type": "Toll",
            "Details": "Eastern Distributor - Northbound",
            "LPN": "OVERLAP1",
            "Tag Number": "TAG-1004",
            "Fleet ID": "FL-02",
            "Vehicle Class": "Car",
            "Amount": -8.50,
        },
        # Unknown Vehicle -> UNMATCHED (NO_MATCH_99)
        {
            "Start Date": "08/09/2026 15:30:00",
            "End Date": "08/09/2026 15:35:00",
            "Type": "Toll",
            "Details": "Cross City Tunnel",
            "LPN": "NOMATCH99",
            "Tag Number": "TAG-9999",
            "Fleet ID": "FL-99",
            "Vehicle Class": "Car",
            "Amount": -6.10,
        },
        # Historical Duplicate (Already exists in Master from August)
        {
            "Start Date": "22/08/2026 10:00:00",
            "End Date": "22/08/2026 10:05:00",
            "Type": "Toll",
            "Details": "M5 South-West Motorway",
            "LPN": "HIST001",
            "Tag Number": "TAG-0001",
            "Fleet ID": "FL-01",
            "Vehicle Class": "Car",
            "Amount": -5.12,
        },
        # Historical Already Paid Toll (Already paid in Master)
        {
            "Start Date": "23/08/2026 11:30:00",
            "End Date": "23/08/2026 11:35:00",
            "Type": "Toll",
            "Details": "M2 Hills Motorway",
            "LPN": "HIST001",
            "Tag Number": "TAG-0001",
            "Fleet ID": "FL-01",
            "Vehicle Class": "Car",
            "Amount": -7.86,
        },
        # Batch Duplicate (Same trip as the first one above in this batch)
        {
            "Start Date": "08/09/2026 10:30:00",
            "End Date": "08/09/2026 10:35:00",
            "Type": "Toll",
            "Details": "M2 Hills Motorway - Macquarie Park",
            "LPN": "ABC123",
            "Tag Number": "TAG-1001",
            "Fleet ID": "FL-01",
            "Vehicle Class": "Car",
            "Amount": -7.86,
        },
    ]
    df_linkt = pd.DataFrame(linkt_data)
    df_linkt.to_excel(linkt_file, index=False)

    # 3. Existing Master Dataset (Contains historical records before 08/09/2026)
    master_historical = [
        {
            "Unique Key": "HIST001_20260822100000",
            "Duplicate Count": 0,
            "Date Received": "25/08/2026",
            "Date": "22/08/2026",
            "Time": "10:00:00",
            "StartDateTime": "22/08/2026 10:00:00",
            "EndDateTime": "22/08/2026 10:05:00",
            "Type of activity": "Toll",
            "REGO": "HIST001",
            "Details": "M5 South-West Motorway",
            "Concession": "",
            "Toll Amount": 5.12,
            "Admin Fee": 5.00,  # Historical admin fee
            "TOTAL AMOUNT": 10.12,
            "Driver": "Alice Walker",
            "Contact": "+61 467 890 123",
            "Hire Date": "20/08/2026 09:00:00",
            "Return Date": "25/08/2026 17:00:00",
            "Payment Method": "",
            "Payment Status": "Pending",
            "Follow-up Date": "",
            "Source": "LINKT",
            "invoice added to 365": "Yes",
            "Invoice sent to customer": "Yes",
        },
        {
            "Unique Key": "HIST001_20260823113000",
            "Duplicate Count": 0,
            "Date Received": "25/08/2026",
            "Date": "23/08/2026",
            "Time": "11:30:00",
            "StartDateTime": "23/08/2026 11:30:00",
            "EndDateTime": "23/08/2026 11:35:00",
            "Type of activity": "Toll",
            "REGO": "HIST001",
            "Details": "M2 Hills Motorway",
            "Concession": "",
            "Toll Amount": 7.86,
            "Admin Fee": 5.00,
            "TOTAL AMOUNT": 12.86,
            "Driver": "Alice Walker",
            "Contact": "+61 467 890 123",
            "Hire Date": "20/08/2026 09:00:00",
            "Return Date": "25/08/2026 17:00:00",
            "Payment Method": "Card",
            "Payment Status": "Paid",
            "Follow-up Date": "",
            "Source": "LINKT",
            "invoice added to 365": "Yes",
            "Invoice sent to customer": "Yes",
        },
    ]
    df_master = pd.DataFrame(master_historical, columns=MASTER_COLUMNS)
    df_master.to_excel(master_file, sheet_name="Master Data", index=False)

    return linkt_file, bookings_file, master_file
