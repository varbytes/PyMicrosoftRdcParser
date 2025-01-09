#!/usr/bin/env python
# 
# PyMicrosoftRdcParser.py
# Version 1.0
#

import argparse
import plistlib
import sqlite3
from datetime import datetime, timedelta
from openpyxl import Workbook

def convert_mac_datetime(mac_timestamp):
    """Convert macOS timestamp to a Python datetime object."""
    mac_epoch = datetime(2001, 1, 1)
    return mac_epoch + timedelta(seconds=mac_timestamp)

def extract_plist_time(plist_blob):
    """Extract 'NS.time' value from stored plist."""
    try:
        plist = plistlib.loads(plist_blob)
        return convert_mac_datetime(plist['$objects'][1]['NS.time'])
    except Exception as e:
        print(f"Error parsing plist: {e}")
        return None

def main(db_path, output_file):
    """Main function to extract data from the database and output to Excel."""
    conn = sqlite3.connect(f'file:{db_path}?mode=ro', uri=True)
    cursor = conn.cursor()

    cursor.execute('''SELECT ZBOOKMARKENTITY.ZHOSTNAME, ZBOOKMARKENTITY.ZFRIENDLYNAME,
                             ZBOOKMARKENTITY.ZLASTCONNECTED, ZBOOKMARKENTITY.ZCONNECTIONCOUNT,
                             ZBOOKMARKENTITY.ZRDPSTRING,
                             ZCREDENTIALENTITY.ZUSERNAME, ZCREDENTIALENTITY.ZID,
                             CASE ZCREDENTIALENTITY.ZNILPASSWORD
                                  WHEN 1 THEN 'True'
                                  WHEN 0 THEN 'False'
                             END AS 'ZNILPASSWORD'
                        FROM ZBOOKMARKENTITY
                        JOIN ZCREDENTIALENTITY
                          ON ZBOOKMARKENTITY.ZCREDENTIAL = ZCREDENTIALENTITY.Z_PK''')
    
    rows = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Connection Bookmarks'
    ws.append(['Hostname', 'Friendly Name', 'Last Connected', 'Connection String', 'RDP String', 'Username', 'ID', 'NIL Password'])

    for row in rows:
        conn_details = list(row)
        
        time_value = extract_plist_time(conn_details[2])
        if time_value is not None:
            conn_details[2] = time_value.strftime('%Y-%m-%d %H:%M:%S')
        else:
            conn_details[2] = None

        ws.append(conn_details)

    wb.save(output_file)
    print(f"Data has been written to {output_file} with {len(rows)} connections processed.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Parses macOS 'Microsoft Remote Desktop' application database for RDP connection bookmarks and export to Excel.")
    parser.add_argument('--db', type=str, default='com.microsoft.rdc.application-data.sqlite', help="Path to the SQLite database file.")
    parser.add_argument('--outfile', type=str, default='connection-bookmarks.xlsx', help="Path to output Excel file.")
    args = parser.parse_args()

    main(args.db, args.outfile)
