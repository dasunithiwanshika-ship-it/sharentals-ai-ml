"""Test calamine reading Booking-list.xls."""

import pandas as pd
from python_calamine import CalamineWorkbook

fpath = "input_data/Booking-list.xls"
wb = CalamineWorkbook.from_file(fpath)
print("Sheet names from Calamine:", wb.sheet_names)

for sname in wb.sheet_names:
    rows = wb.get_sheet_by_name(sname).to_python()
    print(f"\nSheet: '{sname}', Total Rows: {len(rows)}")
    df = pd.DataFrame(rows[1:], columns=rows[0])
    print(f"Columns ({len(df.columns)}):", list(df.columns))
    print("\nFirst 5 rows:")
    print(df.head(5))
    print("\nData Types preview:")
    print(df.dtypes)
    print("\nSample Rego No values:")
    if 'Rego No.' in df.columns:
        print(df['Rego No.'].dropna().unique()[:10])
    elif 'Rego' in df.columns:
        print(df['Rego'].dropna().unique()[:10])
