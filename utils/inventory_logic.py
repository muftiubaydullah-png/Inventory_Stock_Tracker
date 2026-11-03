"""
inventory_logic.py
------------------
Business rules for the Inventory Stock Tracker.

Nothing in this file touches the Excel file directly. It works on pandas
DataFrames and plain dictionaries, which keeps the rules easy to read and test.
"""

import re
from datetime import date, datetime

import pandas as pd

# ---------------------------------------------------------------------------
# Excel structure
# ---------------------------------------------------------------------------
# (short_name, exact Excel header). The ORDER matches the columns in the sheet.
# The app uses the short names internally and the headers when talking to Excel.
COLUMNS = [
    ("item_id", "Item ID"),
    ("name", "Item Name"),
    ("category", "Category"),
    ("spec", "Specification / Rating"),
    ("brand", "Brand"),
    ("unit", "Unit"),
    ("qty", "Quantity in Stock"),
    ("reorder", "Reorder Level"),
    ("min_stock", "Minimum Stock"),
    ("max_stock", "Maximum Stock"),
    ("cost", "Unit Cost (INR)"),
    ("value", "Stock Value (INR)"),
    ("supplier", "Supplier"),
    ("warehouse", "Warehouse"),
    ("bin", "Bin Location"),
    ("status", "Stock Status"),
    ("received", "Last Received Date"),
    ("remarks", "Remarks"),
]
SLUGS = [c[0] for c in COLUMNS]
HEADERS = [c[1] for c in COLUMNS]
HEADER_TO_SLUG = {header: slug for slug, header in COLUMNS}

TEXT_COLUMNS = ["item_id", "name", "category", "spec", "brand", "unit",
                "supplier", "warehouse", "bin", "status", "remarks"]
INT_COLUMNS = ["qty", "reorder", "min_stock", "max_stock"]

ID_PREFIX = "ELC-"

# Stock statuses, from most to least urgent
STATUS_OUT = "Out of Stock"
STATUS_REORDER = "Reorder Required"
STATUS_LOW = "Low Stock"
STATUS_OK = "In Stock"
STATUS_ORDER = {STATUS_OUT: 0, STATUS_REORDER: 1, STATUS_LOW: 2, STATUS_OK: 3}


# ---------------------------------------------------------------------------
# Core calculations
# ---------------------------------------------------------------------------
def compute_status(qty, reorder):
    """
    Work out the Stock Status. These rules reproduce the statuses already in
    your Excel file for all 600 items:

      quantity = 0                     -> Out of Stock
      quantity <= reorder level        -> Reorder Required
      quantity <= 2 x reorder level    -> Low Stock  (running down, keep an eye on it)
      anything higher                  -> In Stock
    """
    if qty <= 0:
        return STATUS_OUT
    if qty <= reorder:
        return STATUS_REORDER
    if qty <= 2 * reorder:
        return STATUS_LOW
    return STATUS_OK


def compute_value(qty, cost):
    """Stock Value = quantity x unit cost, rounded to 2 decimals."""
    return round(qty * cost, 2)


def next_item_id(df):
    """Return the next free Item ID, e.g. ELC-0601, and never an existing one."""
    highest = 0
    for item_id in df["item_id"]:
        match = re.fullmatch(re.escape(ID_PREFIX) + r"(\d+)", str(item_id))
        if match:
            highest = max(highest, int(match.group(1)))
    existing = set(df["item_id"])
    number = highest + 1
    while f"{ID_PREFIX}{number:04d}" in existing:  # safety net
        number += 1
    return f"{ID_PREFIX}{number:04d}"


# ---------------------------------------------------------------------------
# Dashboard and reports
# ---------------------------------------------------------------------------
def dashboard_stats(df):
    """Numbers for the dashboard summary cards."""
    counts = df["status"].value_counts()
    return {
        "total_items": int(len(df)),
        "total_qty": int(df["qty"].sum()),
        "total_value": float(df["value"].sum()),
        "in_stock": int(counts.get(STATUS_OK, 0)),
        "low_stock": int(counts.get(STATUS_LOW, 0)),
        "reorder": int(counts.get(STATUS_REORDER, 0)),
        "out_of_stock": int(counts.get(STATUS_OUT, 0)),
    }


def status_distribution(df):
    """List of {label, value} for the status chart, in a fixed order."""
    counts = df["status"].value_counts()
    return [{"label": s, "value": int(counts.get(s, 0))}
            for s in (STATUS_OK, STATUS_LOW, STATUS_REORDER, STATUS_OUT)]


def category_summary(df):
    """One row per category: item count, total quantity, value, items needing action."""
    if df.empty:
        return pd.DataFrame(columns=["category", "items", "qty", "value", "needs_reorder"])
    work = df.assign(needs=df["qty"] <= df["reorder"])
    out = work.groupby("category").agg(
        items=("item_id", "count"),
        qty=("qty", "sum"),
        value=("value", "sum"),
        needs_reorder=("needs", "sum"),
    ).reset_index()
    out["value"] = out["value"].round(2)
    return out.sort_values("value", ascending=False).reset_index(drop=True)


def alert_items(df, level="all"):
    """
    Items for the Low Stock Alert page.
      all     -> quantity <= reorder level (includes out of stock)
      out     -> out of stock only
      reorder -> reorder required only (not zero)
      watch   -> 'Low Stock' status (above reorder level, but under 2x)
    Most urgent first (lowest quantity compared with reorder level).
    """
    if level == "out":
        sel = df[df["status"] == STATUS_OUT]
    elif level == "reorder":
        sel = df[df["status"] == STATUS_REORDER]
    elif level == "watch":
        sel = df[df["status"] == STATUS_LOW]
    else:
        sel = df[df["qty"] <= df["reorder"]]
    sel = sel.copy()
    ratio = sel["qty"] / sel["reorder"].where(sel["reorder"] > 0, 1)
    sel = sel.assign(_ratio=ratio).sort_values(["_ratio", "item_id"]).drop(columns="_ratio")
    # How many units would bring the item back up to its maximum stock level
    sel["suggested_order"] = (sel["max_stock"] - sel["qty"]).clip(lower=0)
    return sel


# ---------------------------------------------------------------------------
# Searching, sorting, pagination
# ---------------------------------------------------------------------------
SEARCH_COLUMNS = ["item_id", "name", "category", "brand", "supplier"]
SORTABLE = ["item_id", "name", "category", "brand", "qty", "cost", "value", "status"]


def search_items(df, query):
    """
    Case-insensitive search across Item ID, Item Name, Category, Brand and Supplier.
    Every word you type must appear somewhere, so 'tata resistor' narrows down.
    """
    query = (query or "").strip().lower()
    if not query or df.empty:
        return df
    haystack = df[SEARCH_COLUMNS].agg(" | ".join, axis=1).str.lower()
    mask = pd.Series(True, index=df.index)
    for word in query.split():
        mask &= haystack.str.contains(word, regex=False)
    return df[mask]


def filter_items(df, query="", category="", status=""):
    """Apply search + category filter + status filter."""
    df = search_items(df, query)
    if category:
        df = df[df["category"] == category]
    if status:
        df = df[df["status"] == status]
    return df


def sort_items(df, column="item_id", order="asc"):
    """Sort by one of the sortable columns (text sorts ignore upper/lower case)."""
    if column not in SORTABLE:
        column = "item_id"
    ascending = order != "desc"

    def key(series):
        if series.name == "status":
            return series.map(STATUS_ORDER)
        if series.name in TEXT_COLUMNS:
            return series.str.lower()
        return series

    # item_id is the tie-breaker so the order is always stable
    by = [column] if column == "item_id" else [column, "item_id"]
    return df.sort_values(by, ascending=ascending, key=key, kind="mergesort")


def paginate(df, page, per_page):
    """Return (rows_for_this_page, paging_info)."""
    total = len(df)
    pages = max(1, -(-total // per_page))  # ceiling division
    page = min(max(1, page), pages)
    start = (page - 1) * per_page
    rows = df.iloc[start:start + per_page].to_dict("records")

    # Page numbers to show, with None meaning "..."
    numbers = sorted({1, pages, page - 1, page, page + 1} & set(range(1, pages + 1)))
    window = []
    for n in numbers:
        if window and n - window[-1] > 1:
            window.append(None)
        window.append(n)

    info = {
        "page": page, "pages": pages, "total": total, "per_page": per_page,
        "first": start + 1 if total else 0, "last": min(start + per_page, total),
        "window": window,
    }
    return rows, info


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def _clean_text(form, key, label, errors, required=True, max_len=100):
    value = " ".join(str(form.get(key, "")).split())  # trim + collapse spaces
    if not value:
        if required:
            errors[key] = f"{label} is required."
        return ""
    if len(value) > max_len:
        errors[key] = f"{label} must be {max_len} characters or fewer."
    return value


def _clean_int(form, key, label, errors, default=None, required=True, min_val=0, maximum=10_000_000):
    raw = str(form.get(key, "")).strip()
    if raw == "":
        if default is not None:
            return default
        if required:
            errors[key] = f"{label} is required."
        return None
    try:
        number = int(raw)
    except ValueError:
        errors[key] = f"{label} must be a whole number."
        return None
    if min_val is not None and number < min_val:
        if min_val == 0:
            errors[key] = f"{label} cannot be negative."
        else:
            errors[key] = f"{label} cannot be less than {min_val}."
        return None
    if number > maximum:
        errors[key] = f"{label} is too large."
        return None
    return number


def validate_new_item(form):
    """
    Check the Add Item form. Returns (clean_record, errors).
    errors is a dict {field_name: message}; it is empty when everything is valid.
    The Item ID is NOT taken from the form: the server generates it.

    Keep only 8 required fields:
      1. name (Item Name)
      2. category (Category)
      3. brand (Brand)
      4. unit (Unit)
      5. qty (Quantity in Stock, >= 0)
      6. cost (Unit Cost in INR, >= 0)
      7. supplier (Supplier)
      8. warehouse (Warehouse)

    Automatic defaults for other fields:
      - spec: Optional ("")
      - reorder: Automatically set to 50
      - min_stock: Automatically set to 10
      - max_stock: Automatically set to 1000
      - bin: Optional, defaults to "Not Assigned"
      - received: Automatically set to today's date
      - remarks: Optional ("")
    """
    errors = {}
    rec = {
        "name": _clean_text(form, "name", "Item name", errors, required=True, max_len=100),
        "category": _clean_text(form, "category", "Category", errors, required=True, max_len=60),
        "spec": _clean_text(form, "spec", "Specification", errors, required=False, max_len=100),
        "brand": _clean_text(form, "brand", "Brand", errors, required=True, max_len=60),
        "unit": _clean_text(form, "unit", "Unit", errors, required=True, max_len=20),
        "supplier": _clean_text(form, "supplier", "Supplier", errors, required=True, max_len=100),
        "warehouse": _clean_text(form, "warehouse", "Warehouse", errors, required=True, max_len=60),
        "remarks": _clean_text(form, "remarks", "Remarks", errors, required=False, max_len=200),
    }

    # Bin location: optional, default to "Not Assigned"
    raw_bin = _clean_text(form, "bin", "Bin location", errors, required=False, max_len=30)
    rec["bin"] = raw_bin if raw_bin else "Not Assigned"

    # Quantity in stock: required, zero or greater
    rec["qty"] = _clean_int(form, "qty", "Quantity in stock", errors, required=True, min_val=0)

    # Automatic / optional stock levels
    rec["reorder"] = _clean_int(form, "reorder", "Reorder level", errors, default=50, required=False, min_val=0)
    rec["min_stock"] = _clean_int(form, "min_stock", "Minimum stock", errors, default=10, required=False, min_val=0)
    rec["max_stock"] = _clean_int(form, "max_stock", "Maximum stock", errors, default=1000, required=False, min_val=0)

    # Unit cost: required, zero or greater
    raw_cost = str(form.get("cost", "")).strip().replace(",", "")
    cost = None
    if raw_cost == "":
        errors["cost"] = "Unit cost is required."
    else:
        try:
            cost = round(float(raw_cost), 2)
            if cost != cost or cost in (float("inf"), float("-inf")):
                raise ValueError
            if cost < 0:
                errors["cost"] = "Unit cost must be zero or greater."
            elif cost > 100_000_000:
                errors["cost"] = "Unit cost is too large."
            else:
                rec["cost"] = cost
        except ValueError:
            errors["cost"] = "Unit cost must be a valid number, e.g. 125.50."
    rec["cost"] = cost

    # Stock limits logic
    if (rec["min_stock"] is not None and rec["max_stock"] is not None
            and rec["max_stock"] < rec["min_stock"]):
        errors["max_stock"] = "Maximum stock cannot be less than minimum stock."
    if (rec["reorder"] is not None and rec["max_stock"] is not None
            and rec["max_stock"] < rec["reorder"]):
        errors["max_stock"] = "Maximum stock cannot be less than the reorder level."
    if (rec["qty"] is not None and rec["max_stock"] is not None
            and rec["qty"] > rec["max_stock"]):
        errors["qty"] = f"Quantity cannot be higher than maximum stock ({rec['max_stock']})."

    # Last received date (defaults to today)
    raw_date = str(form.get("received", "")).strip()
    if not raw_date:
        rec["received"] = datetime.combine(date.today(), datetime.min.time())
    else:
        try:
            parsed = datetime.strptime(raw_date, "%Y-%m-%d")
            if parsed.date() > date.today():
                errors["received"] = "Received date cannot be in the future."
            rec["received"] = parsed
        except ValueError:
            errors["received"] = "Enter a valid date."
            rec["received"] = None

    if not errors:
        rec["value"] = compute_value(rec["qty"], rec["cost"])
        rec["status"] = compute_status(rec["qty"], rec["reorder"])
    return rec, errors


def validate_edit_item(form):
    """
    Check the Edit Item form where any field (including automatic defaults)
    can be modified by the user.
    """
    errors = {}
    rec = {
        "name": _clean_text(form, "name", "Item name", errors, required=True, max_len=100),
        "category": _clean_text(form, "category", "Category", errors, required=True, max_len=60),
        "spec": _clean_text(form, "spec", "Specification", errors, required=False, max_len=100),
        "brand": _clean_text(form, "brand", "Brand", errors, required=True, max_len=60),
        "unit": _clean_text(form, "unit", "Unit", errors, required=True, max_len=20),
        "supplier": _clean_text(form, "supplier", "Supplier", errors, required=True, max_len=100),
        "warehouse": _clean_text(form, "warehouse", "Warehouse", errors, required=True, max_len=60),
        "remarks": _clean_text(form, "remarks", "Remarks", errors, required=False, max_len=200),
    }

    raw_bin = _clean_text(form, "bin", "Bin location", errors, required=False, max_len=30)
    rec["bin"] = raw_bin if raw_bin else "Not Assigned"

    rec["qty"] = _clean_int(form, "qty", "Quantity in stock", errors, required=True, min_val=0)
    rec["reorder"] = _clean_int(form, "reorder", "Reorder level", errors, default=50, required=True, min_val=0)
    rec["min_stock"] = _clean_int(form, "min_stock", "Minimum stock", errors, default=10, required=True, min_val=0)
    rec["max_stock"] = _clean_int(form, "max_stock", "Maximum stock", errors, default=1000, required=True, min_val=0)

    raw_cost = str(form.get("cost", "")).strip().replace(",", "")
    cost = None
    if raw_cost == "":
        errors["cost"] = "Unit cost is required."
    else:
        try:
            cost = round(float(raw_cost), 2)
            if cost != cost or cost in (float("inf"), float("-inf")):
                raise ValueError
            if cost < 0:
                errors["cost"] = "Unit cost must be zero or greater."
            elif cost > 100_000_000:
                errors["cost"] = "Unit cost is too large."
            else:
                rec["cost"] = cost
        except ValueError:
            errors["cost"] = "Unit cost must be a valid number, e.g. 125.50."
    rec["cost"] = cost

    if (rec["min_stock"] is not None and rec["max_stock"] is not None
            and rec["max_stock"] < rec["min_stock"]):
        errors["max_stock"] = "Maximum stock cannot be less than minimum stock."
    if (rec["reorder"] is not None and rec["max_stock"] is not None
            and rec["max_stock"] < rec["reorder"]):
        errors["max_stock"] = "Maximum stock cannot be less than the reorder level."
    if (rec["qty"] is not None and rec["max_stock"] is not None
            and rec["qty"] > rec["max_stock"]):
        errors["qty"] = f"Quantity cannot be higher than maximum stock ({rec['max_stock']})."

    raw_date = str(form.get("received", "")).strip()
    if not raw_date:
        rec["received"] = datetime.combine(date.today(), datetime.min.time())
    else:
        try:
            parsed = datetime.strptime(raw_date, "%Y-%m-%d")
            if parsed.date() > date.today():
                errors["received"] = "Received date cannot be in the future."
            rec["received"] = parsed
        except ValueError:
            errors["received"] = "Enter a valid date."
            rec["received"] = None

    if not errors:
        rec["value"] = compute_value(rec["qty"], rec["cost"])
        rec["status"] = compute_status(rec["qty"], rec["reorder"])
    return rec, errors


def apply_stock_change(current_qty, action, amount_text):
    """
    Check a Stock In / Stock Out request.
    Returns (new_quantity, error_message). One of them is always None.
    """
    if action not in ("in", "out"):
        return None, "Choose Stock In or Stock Out."
    try:
        amount = int(str(amount_text).strip())
    except ValueError:
        return None, "Quantity must be a whole number."
    if amount <= 0:
        return None, "Quantity must be greater than 0."
    if amount > 10_000_000:
        return None, "Quantity is too large."
    if action == "out":
        if amount > current_qty:
            return None, f"Cannot remove {amount}: only {current_qty} in stock."
        return current_qty - amount, None
    return current_qty + amount, None
