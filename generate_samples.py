"""Script to pre-generate sample datasets."""

from toll_processor.sample_generator import generate_sample_datasets

if __name__ == "__main__":
    l, b, m = generate_sample_datasets("sample_data")
    print(f"Sample datasets generated in sample_data/:")
    print(f"  1. LINKT Report      : {l}")
    print(f"  2. 365 Bookings      : {b}")
    print(f"  3. Master Dataset    : {m}")
