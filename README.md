# Inventory Stock Tracker

A web-based inventory manager for your 600 electronic items. It runs on your own computer and uses your Excel file (`inventory.xlsx`) as its only database. There is no SQLite and no other database.

Built with Python, Flask, Pandas, OpenPyXL, HTML, CSS and JavaScript. The charts are drawn with plain SVG/CSS, so nothing needs to be downloaded from the internet and the app works offline.

---

## 1. Project structure

```
Inventory_Stock_Tracker/
├── app.py                  Flask app: all pages (routes)
├── inventory.xlsx          Your inventory (the database)
├── requirements.txt        Python packages needed
├── README.md               This file
├── templates/
│   ├── base.html           Sidebar, top bar, flash messages, delete dialog
│   ├── dashboard.html      Summary cards and charts
│   ├── inventory.html      Table with search, filters, sorting, paging
│   ├── add_item.html       Add New Item form
│   ├── update_stock.html   Stock In / Stock Out
│   ├── low_stock.html      Low-stock alerts
│   ├── reports.html        Reports + Excel export buttons
│   ├── error.html          Friendly error page (e.g. missing Excel file)
│   └── _macros.html        Small shared pieces (badges, pager, delete button)
├── static/
│   ├── css/style.css       Design
│   ├── js/script.js        Charts, delete confirmation, live previews
│   └── images/
└── utils/
    ├── excel_handler.py    Reads / writes the Excel file (Pandas + OpenPyXL)
    └── inventory_logic.py  Business rules: status, search, sort, validation
```

## 2. Install (VS Code)

You need **Python 3.9 or newer** ([python.org](https://www.python.org/downloads/); on Windows tick "Add Python to PATH").

1. Open VS Code, choose **File > Open Folder…** and select the `Inventory_Stock_Tracker` folder.
2. Open the terminal: **Terminal > New Terminal**.
3. Create a virtual environment (a private box for this project's packages):
   - Windows: `python -m venv venv` then `venv\Scripts\activate`
   - Mac / Linux: `python3 -m venv venv` then `source venv/bin/activate`
4. Install the packages: `pip install -r requirements.txt`
5. If VS Code asks which Python interpreter to use, pick the one inside `venv`.

## 3. Run

```
python app.py
```
(Mac / Linux: `python3 app.py`.) Then open **http://127.0.0.1:5000** in your browser. Stop the app with `Ctrl + C` in the terminal.

## 4. Connect your own Excel file

Your 600-item file is already included as `inventory.xlsx`, so it works straight away. To use a different file:

- **Simplest:** replace `inventory.xlsx` with your file (keep the name), while the app is stopped.
- **Or** keep the file anywhere and point to it:
  - Windows (PowerShell): `$env:INVENTORY_FILE="C:\path\to\my_inventory.xlsx"; python app.py`
  - Mac / Linux: `INVENTORY_FILE=/path/to/my_inventory.xlsx python3 app.py`

Your file needs a sheet named `Inventory` (or the first sheet is used) with these headers in row 1, in this order:
Item ID, Item Name, Category, Specification / Rating, Brand, Unit, Quantity in Stock, Reorder Level, Minimum Stock, Maximum Stock, Unit Cost (INR), Stock Value (INR), Supplier, Warehouse, Bin Location, Stock Status, Last Received Date, Remarks.

The first time the app runs it saves an untouched copy called `inventory_original_backup.xlsx` next to your file.

## 5. Features

- **Dashboard**: total items, total quantity, total value (in INR, lakh/crore style), in-stock, low-stock, reorder and out-of-stock cards (click one to see those items), a status donut chart, a value-by-category chart and the most urgent items.
- **All Inventory**: search, category filter, status filter, sortable columns, pagination (10/25/50/100 rows), responsive table.
- **Search**: case-insensitive across Item ID, Item Name, Category, Brand and Supplier. Several words must all match, so `molex connector` narrows down.
- **Add Item**: automatic unique Item ID (next after the highest, e.g. `ELC-0601`), validation of every required field, quantity and price, live Stock Value, saved to the Excel file.
- **Update Stock**: Stock In / Stock Out with live preview, negative stock blocked, value and status recalculated, saved immediately. Stock In also sets Last Received Date to today.
- **Delete**: confirmation dialog, then the row is removed from the Excel file.
- **Low Stock Alerts**: items at or below reorder level (Item ID, Name, Quantity, Reorder Level, Supplier, plus suggested order quantity), colour-coded, with tabs for Reorder required, Out of stock and a Watch list.
- **Reports**: category-wise inventory and value, low-stock report, out-of-stock report, total valuation, each exportable to Excel, plus one "Export all" workbook.
- **Excel integration**: your formatting, Excel table, status colour rules and Summary sheet are preserved and kept up to date. Saves are atomic (written to a temporary file, then swapped in). Edits you make in Excel while the app is running are picked up on the next page load.

### How Stock Status is decided
These rules match all 600 statuses already in your file:

| Condition | Status |
|---|---|
| Quantity = 0 | Out of Stock |
| Quantity ≤ Reorder Level | Reorder Required |
| Quantity ≤ 2 × Reorder Level | Low Stock |
| Otherwise | In Stock |

To change the rules, edit `compute_status()` in `utils/inventory_logic.py`.

## 6. Troubleshooting

| Problem | Fix |
|---|---|
| `python` is not recognised | Reinstall Python and tick "Add Python to PATH", or try `py app.py` (Windows) / `python3 app.py`. |
| `ModuleNotFoundError: No module named 'flask'` | The virtual environment is not active or packages are not installed. Run `venv\Scripts\activate` (or `source venv/bin/activate`), then `pip install -r requirements.txt`. |
| PowerShell says scripts are disabled | Run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, or use Command Prompt instead. |
| "Could not save: the inventory file is open in Excel" | Close the file in Excel (Windows locks open files), then repeat the action. Nothing was changed. |
| "The inventory file was not found" | Put `inventory.xlsx` next to `app.py`, or set `INVENTORY_FILE` (section 4). |
| "The Excel file is missing these columns" | Check the row 1 headers match section 4 exactly. |
| `Address already in use` / port 5000 busy (common on Mac because of AirPlay) | Use another port: `PORT=5001 python3 app.py` (PowerShell: `$env:PORT=5001; python app.py`) and open http://127.0.0.1:5001. |
| Page looks unstyled | Hard-refresh with `Ctrl + F5` (`Cmd + Shift + R` on Mac). |
| Want to undo everything | Stop the app and copy `inventory_original_backup.xlsx` over `inventory.xlsx`. |

## 7. Notes

- Designed for one person or a small team on one computer. Changes are applied one at a time, but this is not a multi-user server.
- Value and status are always recalculated from quantity and unit cost, so they cannot go stale.
- Item names, brands and so on are free text; the Add Item form suggests existing values to keep the data consistent.
- Before sharing the app on a network, set a private `SECRET_KEY` environment variable and run without `debug=True`.
