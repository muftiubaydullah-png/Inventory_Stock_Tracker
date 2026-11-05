"""
migrate_excel_to_mysql.py
-------------------------
Migration script to import all inventory records from inventory.xlsx into MySQL database (inventory_tracker).

Features:
- Preserves all 18 Excel fields (IDs, names, categories, quantities, costs, stock values, locations, dates, etc.)
- Avoids duplicate records by checking existing Item IDs in the database
- Displays detailed progress, import count, skipped duplicates, and any failed records
- Never deletes or modifies the original Excel file
"""

import os
import sys
from datetime import datetime, date
import pandas as pd
from dotenv import load_dotenv

# Ensure project utils can be imported whether script is run from project root or inside Inventory_Stock_TrackerFF
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if os.path.isdir(os.path.join(BASE_DIR, "utils")):
    sys.path.insert(0, BASE_DIR)
elif os.path.isdir(os.path.join(BASE_DIR, "Inventory_Stock_TrackerFF", "utils")):
    sys.path.insert(0, os.path.join(BASE_DIR, "Inventory_Stock_TrackerFF"))

from utils.db_handler import DBHandler, DatabaseError
from utils.inventory_logic import (
    COLUMNS, HEADERS, HEADER_TO_SLUG,
    compute_status, compute_value
)


def find_excel_file():
    """Locate the inventory.xlsx file."""
    env_file = os.environ.get("INVENTORY_FILE")
    candidates = [
        env_file,
        os.path.join(BASE_DIR, "inventory.xlsx"),
        os.path.join(BASE_DIR, "Inventory_Stock_TrackerFF", "inventory.xlsx"),
        os.path.join(os.path.dirname(BASE_DIR), "inventory.xlsx"),
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return os.path.abspath(path)
    return None


def clean_val(val, default=""):
    """Convert pandas/numpy values to clean strings or None."""
    if pd.isna(val) or val is None:
        return default
    text = str(val).strip()
    return text if text != "" else default


def clean_int(val, default=0):
    """Safely convert to int."""
    try:
        if pd.isna(val) or val is None:
            return default
        return int(float(val))
    except (ValueError, TypeError):
        return default


def clean_float(val, default=0.0):
    """Safely convert to float."""
    try:
        if pd.isna(val) or val is None:
            return default
        return round(float(val), 2)
    except (ValueError, TypeError):
        return default


def clean_date(val):
    """Safely convert date/timestamp to YYYY-MM-DD string or None."""
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    try:
        dt = pd.to_datetime(val)
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None


def run_migration():
    print("=" * 65)
    print("      Inventory Stock Tracker: Excel to MySQL Migration")
    print("=" * 65)

    # 1. Locate Excel file
    excel_path = find_excel_file()
    if not excel_path:
        print("[ERROR] inventory.xlsx file was not found!")
        print("Please place inventory.xlsx in the project directory.")
        return False

    print(f"[1/4] Found Excel file: {excel_path}")

    # 2. Connect to MySQL & ensure DB and table exist
    print("[2/4] Connecting to MySQL and verifying database structure...")
    try:
        db = DBHandler()
        db.init_db()
        print(f"      Connected to MySQL database: '{db.db_name}' at {db.host}:{db.port}")
    except DatabaseError as err:
        print(f"[ERROR] Database connection failed:\n      {err}")
        print("\nPlease check your MySQL credentials in .env and make sure MySQL service is running.")
        return False

    # 3. Read Excel workbook
    print("[3/4] Reading records from Excel...")
    try:
        df = pd.read_excel(excel_path, engine="openpyxl")
    except Exception as exc:
        print(f"[ERROR] Failed to read Excel file: {exc}")
        return False

    total_rows = len(df)
    print(f"      Total records found in Excel: {total_rows}")

    # 4. Check existing records in MySQL to avoid duplicates
    conn = db.get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT `item_id` FROM `inventory`")
    existing_ids = {row[0].strip().lower() for row in cursor.fetchall()}

    insert_sql = """
        INSERT INTO `inventory` (
            `item_id`, `name`, `category`, `spec`, `brand`, `unit`,
            `qty`, `reorder`, `min_stock`, `max_stock`, `cost`, `value`,
            `supplier`, `warehouse`, `bin`, `status`, `received`, `remarks`
        ) VALUES (
            %s, %s, %s, %s, %s, %s,
            %s, %s, %s, %s, %s, %s,
            %s, %s, %s, %s, %s, %s
        )
    """

    imported_count = 0
    skipped_count = 0
    failed_records = []

    print("[4/4] Importing records into MySQL...")

    for index, row in df.iterrows():
        raw_id = clean_val(row.get("Item ID"))
        if not raw_id:
            failed_records.append((f"Row {index + 2}", "Missing Item ID"))
            continue

        item_id_lower = raw_id.lower()
        if item_id_lower in existing_ids:
            skipped_count += 1
            continue

        try:
            name = clean_val(row.get("Item Name"), "Unnamed Item")
            category = clean_val(row.get("Category"), "Miscellaneous")
            spec = clean_val(row.get("Specification / Rating"), None)
            brand = clean_val(row.get("Brand"), "Unknown")
            unit = clean_val(row.get("Unit"), "Nos")
            qty = clean_int(row.get("Quantity in Stock"), 0)
            reorder = clean_int(row.get("Reorder Level"), 50)
            min_stock = clean_int(row.get("Minimum Stock"), 10)
            max_stock = clean_int(row.get("Maximum Stock"), 1000)
            cost = clean_float(row.get("Unit Cost (INR)"), 0.0)
            value = clean_float(row.get("Stock Value (INR)"), compute_value(qty, cost))
            supplier = clean_val(row.get("Supplier"), "Unknown Supplier")
            warehouse = clean_val(row.get("Warehouse"), "Main Store")
            bin_loc = clean_val(row.get("Bin Location"), "Not Assigned")
            status = clean_val(row.get("Stock Status"), compute_status(qty, reorder))
            received = clean_date(row.get("Last Received Date"))
            remarks = clean_val(row.get("Remarks"), None)

            params = (
                raw_id, name, category, spec, brand, unit,
                qty, reorder, min_stock, max_stock, cost, value,
                supplier, warehouse, bin_loc, status, received, remarks
            )
            cursor.execute(insert_sql, params)
            existing_ids.add(item_id_lower)
            imported_count += 1
        except Exception as exc:
            failed_records.append((raw_id, str(exc)))

    cursor.close()
    conn.close()

    # 5. Summary Report
    print("\n" + "=" * 65)
    print("                     MIGRATION SUMMARY")
    print("=" * 65)
    print(f"  • Source Excel File:        {excel_path}")
    print(f"  • Target MySQL Database:    {db.db_name}")
    print(f"  • Total Excel Records:      {total_rows}")
    print(f"  • Successfully Imported:    {imported_count}")
    print(f"  • Skipped (Already Exist):  {skipped_count}")
    print(f"  • Failed Records:           {len(failed_records)}")

    if failed_records:
        print("\nFailed Record Details:")
        for item_ref, err in failed_records[:10]:
            print(f"  - [{item_ref}]: {err}")
        if len(failed_records) > 10:
            print(f"  ... and {len(failed_records) - 10} more.")

    print("\n[NOTE] The original Excel file was NOT modified or deleted.")
    print("=" * 65)
    return len(failed_records) == 0


if __name__ == "__main__":
    success = run_migration()
    sys.exit(0 if success else 1)

