"""Diagnose Booking-list.xls OLE2 structure and reading options."""

import olefile
import os

fpath = "input_data/Booking-list.xls"
print(f"Is OLE file: {olefile.isOleFile(fpath)}")
if olefile.isOleFile(fpath):
    ole = olefile.OleFileIO(fpath)
    print(f"Streams in OLE file: {ole.listdir()}")
    for stream in ole.listdir():
        size = ole.get_size(stream)
        print(f"  Stream: {stream}, Size: {size}")
    
    # Try reading Workbook stream
    if ole.exists('Workbook'):
        wb_stream = ole.openstream('Workbook')
        data = wb_stream.read()
        print(f"Workbook stream read successfully! Length: {len(data)} bytes")
