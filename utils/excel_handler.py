"""
excel_handler.py
----------------
All reading and writing of the Excel file lives here.

* Reading uses pandas (fast, gives us a table to filter and sort).
* Writing uses openpyxl directly on the existing workbook, so your formatting,
  Excel table, colour rules for Stock Status and the Summary sheet are kept.
* Every write is saved to a temporary file first and then swapped in, so a
  crash half-way through can never leave you with a corrupted inventory file.
"""

import os
import re
import shutil
import threading
from copy import copy
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .inventory_logic import (
    COLUMNS, HEADERS, HEADER_TO_SLUG, INT_COLUMNS, SLUGS, TEXT_COLUMNS,
    STATUS_LOW, STATUS_OK, STATUS_OUT, STATUS_REORDER,
    compute_status, compute_value,
)

PREFERRED_SHEET = "Inventory"
SUMMARY_SHEET = "Summary"


class ExcelError(Exception):
    """A problem with the Excel file that we can explain to the user in plain words."""


class ExcelHandler:
    def __init__(self, path):
        self.path = path
        self.lock = threading.RLock()  # one change at a time
        self._cache = None
        self._signature = None

    # ------------------------------------------------------------------
    # Reading
    # ------------------------------------------------------------------
    def read(self):
        """
        Return the inventory as a DataFrame (a fresh copy each time).
        The file is only re-read when it has changed on disk, so edits you make
        directly in Excel show up automatically.
        """
        with self.lock:
            if not os.path.isfile(self.path):
                raise ExcelError(
                    f"The inventory file was not found at: {self.path}. "
                    "Place your Excel file next to app.py and name it inventory.xlsx."
                )
            stat = os.stat(self.path)
            signature = (stat.st_mtime_ns, stat.st_size)
            if self._cache is None or signature != self._signature:
                self._cache = self._load_dataframe()
                self._signature = signature
            return self._cache.copy()

    def _load_dataframe(self):
        try:
            with pd.ExcelFile(self.path, engine="openpyxl") as book:
                sheet = PREFERRED_SHEET if PREFERRED_SHEET in book.sheet_names else book.sheet_names[0]
                raw = pd.read_excel(book, sheet_name=sheet)
        except PermissionError:
            raise ExcelError("The inventory file cannot be opened. Close it in Excel and try again.")
        except Exception as exc:  # corrupt file, wrong format, etc.
            raise ExcelError(f"The inventory file could not be read ({exc}). "
                             "Make sure it is a valid .xlsx file.")

        raw.columns = [str(c).strip() for c in raw.columns]
        missing = [h for h in HEADERS if h not in raw.columns]
        if missing:
            raise ExcelError("The Excel file is missing these columns: " + ", ".join(missing))

        df = raw[HEADERS].rename(columns=HEADER_TO_SLUG)
        for col in TEXT_COLUMNS:
            df[col] = df[col].fillna("").astype(str).str.strip()
        for col in INT_COLUMNS:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
        df["cost"] = pd.to_numeric(df["cost"], errors="coerce").fillna(0.0)
        df["received"] = pd.to_datetime(df["received"], errors="coerce")
        df = df[df["item_id"] != ""].reset_index(drop=True)

        # Value and status are always derived from quantity, so they can never be stale
        # (for example if someone edited a quantity by hand in Excel).
        df["value"] = (df["qty"] * df["cost"]).round(2)
        df["status"] = [compute_status(q, r) for q, r in zip(df["qty"], df["reorder"])]
        return df

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------
    def add_item(self, record):
        """Append a new item as the last row of the sheet."""
        with self.lock:
            wb, ws = self._open_for_write()
            if self._find_row(ws, record["item_id"]):
                raise ExcelError(f"Item ID {record['item_id']} already exists.")
            last = self._last_data_row(ws)
            new_row = last + 1
            for col_index, slug in enumerate(SLUGS, start=1):
                cell = ws.cell(row=new_row, column=col_index)
                value = record.get(slug)
                cell.value = None if value == "" else value
                # copy number formats, fonts, borders from the row above
                source = ws.cell(row=last, column=col_index)
                if source.has_style:
                    cell._style = copy(source._style)
            ws.row_dimensions[new_row].height = ws.row_dimensions[last].height
            self._finish(wb, ws, new_row)

    def update_quantity(self, item_id, new_qty, received=None):
        """Change an item's quantity and recalculate its value and status."""
        with self.lock:
            wb, ws = self._open_for_write()
            row = self._find_row(ws, item_id)
            if not row:
                raise ExcelError(f"Item {item_id} was not found in the Excel file.")
            cost = ws.cell(row=row, column=SLUGS.index("cost") + 1).value or 0
            reorder = ws.cell(row=row, column=SLUGS.index("reorder") + 1).value or 0
            ws.cell(row=row, column=SLUGS.index("qty") + 1).value = new_qty
            ws.cell(row=row, column=SLUGS.index("value") + 1).value = compute_value(new_qty, float(cost))
            ws.cell(row=row, column=SLUGS.index("status") + 1).value = compute_status(new_qty, int(reorder))
            if received is not None:
                ws.cell(row=row, column=SLUGS.index("received") + 1).value = received
            self._finish(wb, ws, self._last_data_row(ws))

    def update_item(self, item_id, record):
        """Update an existing item's details in the Excel sheet."""
        with self.lock:
            wb, ws = self._open_for_write()
            row = self._find_row(ws, item_id)
            if not row:
                raise ExcelError(f"Item {item_id} was not found in the Excel file.")
            for col_index, slug in enumerate(SLUGS, start=1):
                if slug == "item_id":
                    continue  # Keep unique Item ID intact
                if slug in record:
                    val = record.get(slug)
                    ws.cell(row=row, column=col_index).value = None if val == "" else val
            self._finish(wb, ws, self._last_data_row(ws))

    def delete_item(self, item_id):
        """Remove an item's row from the sheet."""
        with self.lock:
            wb, ws = self._open_for_write()
            row = self._find_row(ws, item_id)
            if not row:
                raise ExcelError(f"Item {item_id} was not found in the Excel file.")
            ws.delete_rows(row)
            self._finish(wb, ws, self._last_data_row(ws))

    def ensure_backup(self):
        """Keep one untouched copy of the original file (made the first time the app runs)."""
        backup = os.path.splitext(self.path)[0] + "_original_backup.xlsx"
        if os.path.isfile(self.path) and not os.path.exists(backup):
            try:
                shutil.copy2(self.path, backup)
            except OSError:
                pass  # a backup is nice to have, never a reason to stop

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _open_for_write(self):
        if not os.path.isfile(self.path):
            raise ExcelError("The inventory file was not found, so nothing was changed.")
        try:
            wb = load_workbook(self.path)
        except PermissionError:
            raise ExcelError("The inventory file is locked. Close it in Excel and try again.")
        except Exception as exc:
            raise ExcelError(f"The inventory file could not be opened ({exc}).")
        ws = wb[PREFERRED_SHEET] if PREFERRED_SHEET in wb.sheetnames else wb.worksheets[0]
        header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]][:len(HEADERS)]
        if header != HEADERS:
            raise ExcelError("The Excel column headers do not match the expected layout, so nothing was changed.")
        return wb, ws

    @staticmethod
    def _last_data_row(ws):
        row = ws.max_row
        while row > 1 and ws.cell(row=row, column=1).value in (None, ""):
            row -= 1
        return row

    @staticmethod
    def _find_row(ws, item_id):
        target = str(item_id).strip().lower()
        for row in range(2, ws.max_row + 1):
            if str(ws.cell(row=row, column=1).value).strip().lower() == target:
                return row
        return None

    def _finish(self, wb, ws, last_row):
        """Tidy the sheet structures, refresh the Summary sheet, then save safely."""
        self._refresh_structures(ws, last_row)
        try:
            self._refresh_summary(wb, ws)
        except Exception:
            pass  # the Summary sheet is a bonus; never block a stock change because of it
        self._save(wb)

    @staticmethod
    def _refresh_structures(ws, last_row):
        """Make the Excel table, filter and colour rules cover exactly the data rows."""
        last_row = max(last_row, 2)
        ref = f"A1:{get_column_letter(len(COLUMNS))}{last_row}"
        for table in ws.tables.values():
            table.ref = ref
            if table.autoFilter is not None:
                table.autoFilter.ref = ref
        if ws.auto_filter.ref:
            ws.auto_filter.ref = ref

        rebuilt = ConditionalFormattingList()
        for cf in ws.conditional_formatting:
            match = re.fullmatch(r"([A-Z]+)2:([A-Z]+)\d+", str(cf.sqref))
            target = f"{match.group(1)}2:{match.group(2)}{last_row}" if match else str(cf.sqref)
            for rule in cf.rules:
                rebuilt.add(target, rule)
        ws.conditional_formatting = rebuilt

    @staticmethod
    def _refresh_summary(wb, ws):
        """Update the numbers on the Summary sheet (if it exists) after a change."""
        if SUMMARY_SHEET not in wb.sheetnames:
            return
        col = {slug: i + 1 for i, slug in enumerate(SLUGS)}
        records = total_qty = 0
        total_value = 0.0
        counts = {STATUS_OUT: 0, STATUS_REORDER: 0, STATUS_LOW: 0, STATUS_OK: 0}
        for row in range(2, ExcelHandler._last_data_row(ws) + 1):
            qty = int(ws.cell(row=row, column=col["qty"]).value or 0)
            reorder = int(ws.cell(row=row, column=col["reorder"]).value or 0)
            cost = float(ws.cell(row=row, column=col["cost"]).value or 0)
            records += 1
            total_qty += qty
            total_value += compute_value(qty, cost)
            counts[compute_status(qty, reorder)] += 1
        values = {
            "Total inventory records": records,
            "Total quantity across all items": total_qty,
            "Total stock value (INR)": round(total_value, 2),
            "Items out of stock": counts[STATUS_OUT],
            "Items requiring reorder": counts[STATUS_REORDER],
            "Items with low stock": counts[STATUS_LOW],
            "Items in stock": counts[STATUS_OK],
        }
        summary = wb[SUMMARY_SHEET]
        for row in range(2, summary.max_row + 1):
            label = summary.cell(row=row, column=1).value
            if label in values:
                summary.cell(row=row, column=2).value = values[label]

    def _save(self, wb):
        """Save to a temporary file, then replace the real file in one step."""
        temp_path = self.path + ".tmp"
        try:
            wb.save(temp_path)
            os.replace(temp_path, self.path)
        except PermissionError:
            self._remove_quietly(temp_path)
            raise ExcelError("Could not save: the inventory file is open in Excel. "
                             "Close it and try again. Nothing was changed.")
        except OSError as exc:
            self._remove_quietly(temp_path)
            raise ExcelError(f"Could not save the inventory file ({exc}). Nothing was changed.")
        self._cache = None  # force a fresh read next time

    @staticmethod
    def _remove_quietly(path):
        try:
            os.remove(path)
        except OSError:
            pass


# ----------------------------------------------------------------------
# Report export
# ----------------------------------------------------------------------
def build_report_workbook(sheets):
    """
    Create an .xlsx file in memory from {sheet_name: DataFrame}.
    Columns with INR in the name get rupee formatting.
    """
    buffer = BytesIO()
    header_fill = PatternFill("solid", fgColor="0F6B78")
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            sheet_name = name[:31]
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"
            for idx, title in enumerate(frame.columns, start=1):
                head = ws.cell(row=1, column=idx)
                head.font = Font(bold=True, color="FFFFFF")
                head.fill = header_fill
                head.alignment = Alignment(vertical="center")
                longest = max([len(str(title))] + [len(str(v)) for v in frame.iloc[:, idx - 1].head(200)])
                ws.column_dimensions[get_column_letter(idx)].width = min(max(12, longest + 2), 45)
                if "INR" in str(title):
                    for row in range(2, len(frame) + 2):
                        ws.cell(row=row, column=idx).number_format = '"₹"#,##0.00'
                if "Date" in str(title):
                    for row in range(2, len(frame) + 2):
                        ws.cell(row=row, column=idx).number_format = "DD-MMM-YYYY"
    buffer.seek(0)
    return buffer
