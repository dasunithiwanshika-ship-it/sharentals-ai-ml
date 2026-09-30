"""Test pandas calamine engine on Booking-list.xls."""

import pandas as pd

fpath = "input_data/Booking-list.xls"
df = pd.read_excel(fpath, engine="calamine")
print("Successfully read Booking-list.xls with calamine!")
print("Shape:", df.shape)
print("Columns (", len(df.columns), "):", list(df.columns))
print("\nFirst 10 rows:")
print(df.head(10))
print("\nDate Column types & sample values:")
for col in df.columns:
    print(f"  {col}: {df[col].dropna().iloc[:3].tolist() if not df[col].dropna().empty else 'EMPTY'}")
