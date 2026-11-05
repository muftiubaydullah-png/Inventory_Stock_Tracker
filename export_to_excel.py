"""
export_to_excel.py
------------------
CLI script to export all current inventory records directly from MySQL into inventory_updated.xlsx.

Requirements satisfied:
1. Fetches all current inventory records directly from MySQL (table: inventory).
2. Exports the latest data into an Excel file named inventory_updated.xlsx.
3. Includes all inventory columns and records.
4. Preserves MySQL as the primary database without relying on old Excel files.
"""

import os
import sys
import pandas as pd
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if os.path.isdir(os.path.join(BASE_DIR, "utils")):
    sys.path.insert(0, BASE_DIR)
elif os.path.isdir(os.path.join(BASE_DIR, "Inventory_Stock_TrackerFF", "utils")):
    sys.path.insert(0, os.path.join(BASE_DIR, "Inventory_Stock_TrackerFF"))

from utils.db_handler import DBHandler, DatabaseError
from utils.excel_handler import build_inventory_export


def run_export(output_filename="inventory_updated.xlsx"):
    print("=" * 65)
    print("      Inventory Stock Tracker: Export MySQL to Excel")
    print("=" * 65)

    print("[1/3] Connecting to MySQL database...")
    try:
        db = DBHandler()
        df = db.get_all_dataframe()
    except DatabaseError as err:
        print(f"[ERROR] Database connection failed: {err}")
        return False

    total_records = len(df)
    print(f"      Fetched {total_records} current records directly from MySQL.")

    if df.empty:
        print("[WARNING] MySQL inventory table is empty. Nothing to export.")
        return False

    print(f"[2/3] Building formatted Excel workbook with all 18 columns...")
    target_path = os.path.abspath(output_filename)
    root_target_path = os.path.abspath(os.path.join(BASE_DIR, "..", output_filename))
    
    file_paths = [target_path]
    if os.path.isdir(os.path.dirname(root_target_path)) and root_target_path != target_path:
        file_paths.append(root_target_path)

    buffer = build_inventory_export(df, file_paths=file_paths)

    print(f"[3/3] Successfully exported to: {target_path}")
    print("=" * 65)
    print("                     EXPORT SUMMARY")
    print("=" * 65)
    print(f"  • Source Database:       MySQL Server 8.0 ({db.db_name})")
    print(f"  • Total Records:         {total_records}")
    print(f"  • Output File:           {output_filename}")
    print(f"  • File Size:             {os.path.getsize(target_path):,} bytes")
    print("=" * 65)
    return True


if __name__ == "__main__":
    out_name = sys.argv[1] if len(sys.argv) > 1 else "inventory_updated.xlsx"
    success = run_export(out_name)
    sys.exit(0 if success else 1)

