"""
Inventory Stock Tracker - Flask application
-------------------------------------------
Run with:   python app.py      then open http://127.0.0.1:5000

All data lives in the Excel file (inventory.xlsx by default).
To use a different file, set the INVENTORY_FILE environment variable (see README).
"""

import os
import re
from datetime import date, datetime

import pandas as pd
from flask import (Flask, abort, flash, jsonify, redirect, render_template,
                   request, send_file, url_for)

from utils import inventory_logic as logic
from utils.excel_handler import ExcelError, ExcelHandler, build_report_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.environ.get("INVENTORY_FILE") or os.path.join(BASE_DIR, "inventory.xlsx")

app = Flask(__name__)
# Needed for flash messages. Change it if you ever put this app on a shared server.
app.secret_key = os.environ.get("SECRET_KEY", "inventory-stock-tracker-local-key")

excel = ExcelHandler(EXCEL_PATH)
excel.ensure_backup()

PER_PAGE_CHOICES = [10, 25, 50, 100]
UNIT_FALLBACK = ["Nos", "Box", "Pack", "Set", "Roll", "Mtr"]


# ---------------------------------------------------------------------------
# Template helpers (formatting)
# ---------------------------------------------------------------------------
def _indian_grouping(whole):
    """1234567 -> '12,34,567' (lakh / crore style used in India)."""
    if len(whole) <= 3:
        return whole
    head, tail = whole[:-3], whole[-3:]
    head = re.sub(r"(\d)(?=(\d\d)+$)", r"\1,", head)
    return f"{head},{tail}"


@app.template_filter("inr")
def inr(value, decimals=2):
    """Format as rupees: 1234567.5 -> ₹12,34,567.50"""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "—"
    sign = "-" if number < 0 else ""
    text = f"{abs(number):.{decimals}f}"
    whole, _, frac = text.partition(".")
    out = _indian_grouping(whole) + (f".{frac}" if frac else "")
    return f"{sign}₹{out}"


@app.template_filter("inr_short")
def inr_short(value):
    """Compact rupees for cards: 635019033.5 -> ₹63.50 Cr"""
    number = float(value or 0)
    if number >= 1e7:
        return f"₹{number / 1e7:.2f} Cr"
    if number >= 1e5:
        return f"₹{number / 1e5:.2f} L"
    return inr(number, 0)


@app.template_filter("num")
def num(value):
    try:
        return _indian_grouping(str(int(value)))
    except (TypeError, ValueError):
        return "—"


@app.template_filter("fdate")
def fdate(value):
    if value is None or pd.isna(value):
        return "—"
    return value.strftime("%d-%b-%Y")


@app.template_filter("badge")
def badge(status):
    """CSS class for a stock status badge."""
    return {
        logic.STATUS_OK: "badge-ok",
        logic.STATUS_LOW: "badge-low",
        logic.STATUS_REORDER: "badge-reorder",
        logic.STATUS_OUT: "badge-out",
    }.get(status, "badge-ok")


@app.template_global()
def qs(**overrides):
    """Build a query string from the current URL's parameters plus changes (used by links)."""
    params = request.args.to_dict()
    params.update(overrides)
    return {k: v for k, v in params.items() if v not in (None, "")}


@app.context_processor
def inject_globals():
    return {"today": date.today().isoformat(), "excel_name": os.path.basename(EXCEL_PATH)}


# ---------------------------------------------------------------------------
# Error pages
# ---------------------------------------------------------------------------
@app.errorhandler(ExcelError)
def handle_excel_error(error):
    return render_template("error.html", title="Excel file problem", message=str(error),
                           excel_problem=True), 500


@app.errorhandler(404)
def handle_not_found(error):
    return render_template("error.html", title="Page not found",
                           message="That page or item does not exist."), 404


@app.errorhandler(500)
def handle_server_error(error):
    return render_template("error.html", title="Something went wrong",
                           message="An unexpected error occurred. Check the terminal for details."), 500


def safe_next(target, fallback):
    """Only redirect to pages inside this app (prevents open-redirect tricks)."""
    if target and target.startswith("/") and not target.startswith("//"):
        return target
    return fallback


# ---------------------------------------------------------------------------
# 1. Dashboard
# ---------------------------------------------------------------------------
@app.route("/")
def dashboard():
    df = excel.read()
    stats = logic.dashboard_stats(df)
    categories = logic.category_summary(df)
    urgent = logic.alert_items(df, "all").head(8).to_dict("records")
    chart_data = {
        "status": logic.status_distribution(df),
        "categories": [{"label": r.category, "value": float(r.value)}
                       for r in categories.head(8).itertuples()],
    }
    return render_template("dashboard.html", stats=stats, urgent=urgent, chart_data=chart_data)


# ---------------------------------------------------------------------------
# 2 + 3. Inventory list with search, filters, sorting and pagination
# ---------------------------------------------------------------------------
@app.route("/inventory")
def inventory():
    df = excel.read()
    query = request.args.get("q", "").strip()
    category = request.args.get("category", "")
    status = request.args.get("status", "")
    sort = request.args.get("sort", "item_id")
    order = "desc" if request.args.get("order") == "desc" else "asc"
    per_page = request.args.get("per_page", type=int)
    per_page = per_page if per_page in PER_PAGE_CHOICES else 25
    page = request.args.get("page", 1, type=int)

    categories = sorted(df["category"].unique())
    result = logic.filter_items(df, query, category, status)
    result = logic.sort_items(result, sort, order)
    rows, paging = logic.paginate(result, page, per_page)

    return render_template(
        "inventory.html", rows=rows, paging=paging, categories=categories,
        statuses=[logic.STATUS_OK, logic.STATUS_LOW, logic.STATUS_REORDER, logic.STATUS_OUT],
        q=query, category=category, status=status, sort=sort, order=order,
        per_page=per_page, per_page_choices=PER_PAGE_CHOICES,
        filtered_total=int(result["value"].sum()) if len(result) else 0,
    )


@app.route("/api/items/search")
def api_item_search():
    """Small JSON endpoint used by the Update Stock search box."""
    df = excel.read()
    matches = logic.search_items(df, request.args.get("q", ""))
    matches = matches.sort_values("item_id").head(8)
    return jsonify([
        {"item_id": r.item_id, "name": r.name, "brand": r.brand,
         "qty": int(r.qty), "unit": r.unit, "status": r.status}
        for r in matches.itertuples()
    ])


# ---------------------------------------------------------------------------
# 4. Add new item
# ---------------------------------------------------------------------------
def _form_options(df):
    """Suggestions for the drop-downs / auto-complete lists on the Add Item form."""
    return {
        "categories": sorted(df["category"].unique()),
        "brands": sorted(df["brand"].unique()),
        "suppliers": sorted(df["supplier"].unique()),
        "warehouses": sorted(df["warehouse"].unique()),
        "units": sorted(df["unit"].unique()) or UNIT_FALLBACK,
    }


@app.route("/add", methods=["GET", "POST"])
def add_item():
    df = excel.read()
    errors, form = {}, {}

    if request.method == "POST":
        form = request.form.to_dict()
        record, errors = logic.validate_new_item(form)
        if not errors:
            try:
                with excel.lock:  # read -> generate ID -> write must not be interrupted
                    df = excel.read()
                    record["item_id"] = logic.next_item_id(df)
                    excel.add_item(record)
                flash(f"Added {record['name']} ({record['item_id']}) successfully.", "success")
                return redirect(url_for("inventory"))
            except ExcelError as exc:
                flash(str(exc), "error")
        else:
            flash("Please fix the highlighted fields.", "error")
    else:
        # Prepopulate smart defaults for automatic/optional fields
        form = {
            "reorder": 50,
            "min_stock": 10,
            "max_stock": 1000,
            "bin": "Not Assigned",
            "received": date.today().isoformat(),
        }

    return render_template("add_item.html", next_id=logic.next_item_id(df), form=form,
                           errors=errors, **_form_options(df))


# ---------------------------------------------------------------------------
# 4b. Edit existing item (allows modifying automatic values later)
# ---------------------------------------------------------------------------
@app.route("/edit/<item_id>", methods=["GET", "POST"])
def edit_item(item_id):
    df = excel.read()
    item_id_str = str(item_id).strip()
    found = df[df["item_id"].str.lower() == item_id_str.lower()]
    if found.empty:
        flash(f"Item '{item_id}' was not found.", "error")
        return redirect(url_for("inventory"))

    item = found.iloc[0].to_dict()
    errors, form = {}, {}

    if request.method == "POST":
        form = request.form.to_dict()
        record, errors = logic.validate_edit_item(form)
        if not errors:
            try:
                excel.update_item(item["item_id"], record)
                flash(f"Updated {record['name']} ({item['item_id']}) successfully.", "success")
                return redirect(url_for("inventory"))
            except ExcelError as exc:
                flash(str(exc), "error")
        else:
            flash("Please fix the highlighted fields.", "error")
    else:
        form = {
            "name": item.get("name", ""),
            "category": item.get("category", ""),
            "spec": item.get("spec", ""),
            "brand": item.get("brand", ""),
            "unit": item.get("unit", ""),
            "qty": item.get("qty", 0),
            "reorder": item.get("reorder", 50),
            "min_stock": item.get("min_stock", 10),
            "max_stock": item.get("max_stock", 1000),
            "cost": f"{float(item.get('cost', 0)):.2f}" if item.get("cost") is not None else "0.00",
            "supplier": item.get("supplier", ""),
            "warehouse": item.get("warehouse", ""),
            "bin": item.get("bin", "Not Assigned"),
            "remarks": item.get("remarks", ""),
            "received": item["received"].strftime("%Y-%m-%d") if pd.notna(item.get("received")) else date.today().isoformat(),
        }

    return render_template(
        "edit_item.html",
        item=item,
        form=form,
        errors=errors,
        **_form_options(df)
    )


# ---------------------------------------------------------------------------
# 5. Update stock (Stock In / Stock Out)
# ---------------------------------------------------------------------------
@app.route("/update-stock", methods=["GET", "POST"])
def update_stock():
    df = excel.read()
    item_id = (request.values.get("item_id") or "").strip()
    action = request.form.get("action", "in")
    item = None

    if item_id:
        found = df[df["item_id"].str.lower() == item_id.lower()]
        if found.empty:
            flash(f"No item found with ID '{item_id}'.", "error")
        else:
            item = found.iloc[0].to_dict()

    if request.method == "POST" and item:
        new_qty, error = logic.apply_stock_change(item["qty"], action, request.form.get("quantity", ""))
        if error:
            flash(error, "error")
        else:
            try:
                received = datetime.combine(date.today(), datetime.min.time()) if action == "in" else None
                excel.update_quantity(item["item_id"], new_qty, received)
                verb = "added to" if action == "in" else "removed from"
                amount = abs(new_qty - item["qty"])
                flash(f"{amount} {item['unit']} {verb} {item['name']}. New stock: {new_qty}.", "success")
                if action == "in" and new_qty > item["max_stock"]:
                    flash(f"Note: stock is now above the maximum level of {item['max_stock']}.", "warning")
                return redirect(url_for("update_stock", item_id=item["item_id"]))
            except ExcelError as exc:
                flash(str(exc), "error")

    return render_template("update_stock.html", item=item, item_id=item_id, action=action)


# ---------------------------------------------------------------------------
# 6. Delete item (POST only; the page asks for confirmation first)
# ---------------------------------------------------------------------------
@app.route("/delete/<item_id>", methods=["POST"])
def delete_item(item_id):
    df = excel.read()
    found = df[df["item_id"].str.lower() == item_id.lower()]
    destination = safe_next(request.form.get("next"), url_for("inventory"))
    if found.empty:
        flash(f"Item {item_id} was not found.", "error")
        return redirect(destination)
    name = found.iloc[0]["name"]
    canonical_item_id = found.iloc[0]["item_id"]
    try:
        excel.delete_item(canonical_item_id)
        flash(f"Deleted {item_id} ({name}).", "success")
    except ExcelError as exc:
        flash(str(exc), "error")
    return redirect(destination)


# ---------------------------------------------------------------------------
# 7. Low stock alerts
# ---------------------------------------------------------------------------
@app.route("/low-stock")
def low_stock():
    df = excel.read()
    level = request.args.get("level", "all")
    if level not in ("all", "out", "reorder", "watch"):
        level = "all"
    items = logic.alert_items(df, level).to_dict("records")
    counts = {
        "all": int((df["qty"] <= df["reorder"]).sum()),
        "out": int((df["status"] == logic.STATUS_OUT).sum()),
        "reorder": int((df["status"] == logic.STATUS_REORDER).sum()),
        "watch": int((df["status"] == logic.STATUS_LOW).sum()),
    }
    return render_template("low_stock.html", items=items, level=level, counts=counts)


# ---------------------------------------------------------------------------
# 8. Reports + Excel export
# ---------------------------------------------------------------------------
@app.route("/reports")
def reports():
    df = excel.read()
    stats = logic.dashboard_stats(df)
    categories = logic.category_summary(df)
    low = logic.alert_items(df, "all")
    low_only = low[low["status"] != logic.STATUS_OUT].to_dict("records")
    out = df[df["status"] == logic.STATUS_OUT].sort_values("item_id").to_dict("records")
    chart_data = {"categories": [{"label": r.category, "value": float(r.value)}
                                 for r in categories.itertuples()]}
    return render_template("reports.html", stats=stats, categories=categories.to_dict("records"),
                           low_items=low_only, out_items=out, chart_data=chart_data)


def _report_sheets(df, kind):
    """Build the DataFrames for an export. kind is: all, category, low_stock or out_of_stock."""
    sheets = {}
    if kind in ("all", "category"):
        cat = logic.category_summary(df).rename(columns={
            "category": "Category", "items": "Items", "qty": "Total Quantity",
            "value": "Stock Value (INR)", "needs_reorder": "Items At/Below Reorder Level"})
        sheets["Category Summary"] = cat
    if kind == "all":
        sheets["Valuation"] = pd.DataFrame({
            "Metric": ["Total items", "Total quantity", "Total inventory value (INR)"],
            "Value": [len(df), int(df["qty"].sum()), round(float(df["value"].sum()), 2)]})
    if kind in ("all", "low_stock"):
        low = logic.alert_items(df, "all")
        low = low[low["status"] != logic.STATUS_OUT]
        sheets["Low Stock"] = _alert_frame(low)
    if kind in ("all", "out_of_stock"):
        out = df[df["status"] == logic.STATUS_OUT].sort_values("item_id")
        sheets["Out of Stock"] = _alert_frame(out.assign(suggested_order=out["max_stock"] - out["qty"]))
    return sheets


def _alert_frame(frame):
    cols = ["item_id", "name", "category", "brand", "qty", "reorder", "max_stock",
            "suggested_order", "supplier", "warehouse", "bin", "status"]
    names = ["Item ID", "Item Name", "Category", "Brand", "Quantity in Stock", "Reorder Level",
             "Maximum Stock", "Suggested Order Qty", "Supplier", "Warehouse", "Bin Location", "Stock Status"]
    return frame[cols].rename(columns=dict(zip(cols, names)))


@app.route("/reports/export/<kind>")
def export_report(kind):
    if kind not in ("all", "category", "low_stock", "out_of_stock"):
        abort(404)
    df = excel.read()
    buffer = build_report_workbook(_report_sheets(df, kind))
    filename = f"inventory_report_{kind}_{date.today():%Y%m%d}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


if __name__ == "__main__":
    # debug=True reloads the app when you edit code and shows detailed errors.
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", 5000)), debug=True)
