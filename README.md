# Inventory Stock Tracker (MySQL Edition)

A robust web-based inventory management system for electronic items, backed by **MySQL Server 8.0**. 

Previously reliant on Excel for storage, this application has been fully migrated to use **MySQL** as its primary, production-grade data storage system while preserving the complete UI design, responsive dashboard, real-time stock operations, search/filtering, and report exports.

---

## 1. Key Highlights

- **Complete MySQL 8.0 Integration**: All inventory operations (add, view, edit, stock updates, delete, search, dashboard metrics, and low-stock alerts) execute directly on MySQL using parameterized SQL queries.
- **Automated Database & Table Creation**: The application automatically creates the `inventory_tracker` database and `inventory` table on startup if they do not exist.
- **Secure Configuration**: Database credentials are fully configurable via `.env` environment variables.
- **Excel Data Migration**: A safe, idempotent migration script imports all records from `inventory.xlsx` into MySQL with duplicate detection without touching or modifying the original file.
- **Offline Charts & Clean UI**: SVG/CSS-based dashboards, instant live previews, responsive tables, and stock level badges.
- **Excel Report Downloads**: Exporting reports to `.xlsx` remains available as an on-demand reporting feature powered directly by live MySQL queries.

---

## 2. Project Structure

```
Inventory_Stock_Tracker/
├── .env.example                 Example database configuration template
├── .env                         Local environment variables (credentials)
├── .gitignore                   Git ignore file (ignores .env, venv, pycache)
├── app.py                       Flask web application (routes, views, error handlers)
├── inventory.xlsx               Original Excel workbook (preserved)
├── inventory_updated.xlsx       Latest exported Excel file generated from MySQL
├── inventory_original_backup.xlsx Preserved original workbook backup
├── export_to_excel.py           CLI script to export all MySQL records to inventory_updated.xlsx
├── migrate_excel_to_mysql.py    Migration script to import Excel records into MySQL
├── requirements.txt             Python package dependencies
├── run.bat                      One-click launch script for Windows
├── schema.sql                   Complete MySQL 8.0 database DDL schema
├── test_inventory_mysql.py      Automated end-to-end test suite
├── static/
│   ├── css/style.css            Design & styling
│   ├── js/script.js             Charts, modals, live stock calculators
│   └── images/
├── templates/
│   ├── base.html                Layout shell, navigation sidebar, top bar
│   ├── dashboard.html           Summary cards, category & status charts
│   ├── inventory.html           Inventory table with search, filters, pagination
│   ├── add_item.html            Add new product form
│   ├── edit_item.html           Edit product details
│   ├── update_stock.html        Stock In / Stock Out movement form
│   ├── low_stock.html           Urgent low-stock and reorder alerts
│   ├── reports.html             Valuation, category shares, export buttons
│   ├── error.html               User-friendly error diagnostics
│   └── _macros.html             Reusable template macros
└── utils/
    ├── db_handler.py            MySQL database handler (CRUD, queries, pool)
    ├── excel_handler.py         Export report workbook generator
    └── inventory_logic.py       Core business rules, validation, formulas
```

---

## 3. Database Table Structure (`inventory`)

The MySQL database `inventory_tracker` stores all items in the `inventory` table:

| Column | MySQL Data Type | Description |
|---|---|---|
| `item_id` | `VARCHAR(50) PRIMARY KEY` | Unique product identifier (e.g., `ELC-0001`) |
| `name` | `VARCHAR(150) NOT NULL` | Product name |
| `category` | `VARCHAR(100) NOT NULL` | Category (e.g., Resistors, Capacitors) |
| `spec` | `VARCHAR(150) NULL` | Specification / Rating |
| `brand` | `VARCHAR(100) NOT NULL` | Manufacturer or brand |
| `unit` | `VARCHAR(50) NOT NULL` | Unit of measure (Nos, Box, Pack, etc.) |
| `qty` | `INT NOT NULL` | Quantity in stock (>= 0) |
| `reorder` | `INT NOT NULL` | Reorder threshold level |
| `min_stock` | `INT NOT NULL` | Minimum allowable stock |
| `max_stock` | `INT NOT NULL` | Maximum warehouse capacity |
| `cost` | `DECIMAL(12, 2) NOT NULL` | Unit cost in INR |
| `value` | `DECIMAL(15, 2) NOT NULL` | Total stock value (`qty * cost`) |
| `supplier` | `VARCHAR(150) NOT NULL` | Primary vendor / supplier |
| `warehouse` | `VARCHAR(100) NOT NULL` | Warehouse name |
| `bin` | `VARCHAR(50) NOT NULL` | Storage bin / shelf location |
| `status` | `VARCHAR(50) NOT NULL` | Stock status (In Stock, Low Stock, Reorder Required, Out of Stock) |
| `received` | `DATE NULL` | Date of last received stock movement |
| `remarks` | `VARCHAR(255) NULL` | Additional notes or comments |
| `created_at` | `TIMESTAMP DEFAULT CURRENT_TIMESTAMP` | Record creation timestamp |
| `updated_at` | `TIMESTAMP ON UPDATE CURRENT_TIMESTAMP` | Last updated timestamp |

Indexes are created on `category`, `status`, `brand`, `supplier`, `warehouse`, and `qty` for sub-millisecond query filtering.

---

## 4. Setup Instructions

### Step 1: Install MySQL Server 8.0 & MySQL Workbench
Ensure MySQL Server 8.0 is installed and the Windows MySQL service is running.

### Step 2: Configure Environment Variables
Copy `.env.example` to `.env` (or edit existing `.env`):
```ini
DB_HOST=localhost
DB_PORT=3306
DB_USER=root
DB_PASSWORD=your_mysql_password
DB_NAME=inventory_tracker
SECRET_KEY=inventory-stock-tracker-local-key
```

### Step 3: Install Python Dependencies
Open PowerShell or Command Prompt in the project folder:
```bash
python -m pip install -r requirements.txt
```

*(Optional) Using a virtual environment:*
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

### Step 4: (Optional) Manual Database Creation
The application and migration script automatically create the database and table for you. However, you can also run the provided schema in MySQL Workbench:
- Open MySQL Workbench.
- Open `schema.sql`.
- Execute the script.

### Step 5: Migrate Existing Excel Data to MySQL
Run the migration script to load all 600+ records from `inventory.xlsx` into MySQL:
```bash
python migrate_excel_to_mysql.py
```
**Output Summary Example:**
```
=================================================================
      Inventory Stock Tracker: Excel to MySQL Migration
=================================================================
[1/4] Found Excel file: .../inventory.xlsx
[2/4] Connecting to MySQL and verifying database structure...
      Connected to MySQL database: 'inventory_tracker' at localhost:3306
[3/4] Reading records from Excel...
      Total records found in Excel: 601
[4/4] Importing records into MySQL...

=================================================================
                     MIGRATION SUMMARY
=================================================================
  • Source Excel File:        inventory.xlsx
  • Target MySQL Database:    inventory_tracker
  • Total Excel Records:      601
  • Successfully Imported:    601
  • Skipped (Already Exist):  0
  • Failed Records:           0

[NOTE] The original Excel file was NOT modified or deleted.
=================================================================
```

### Step 6: Start the Flask Application
Run:
```bash
python app.py
```
Or double-click `run.bat` on Windows.

Open your web browser and navigate to:
**http://127.0.0.1:5000**

---

## 5. Application Features

1. **Dashboard (`/`)**:
   - Total items, total units, total inventory valuation (in INR with Lakh/Crore formatting).
   - Stock health breakdown (In Stock, Low Stock, Reorder Required, Out of Stock).
   - Category distribution charts and urgent action table.
2. **All Inventory (`/inventory`)**:
   - Multi-keyword live search across ID, Name, Category, Brand, and Supplier.
   - Category and Status drop-down filters with automatic submission.
   - Dynamic sorting by any column (ascending/descending).
   - Configurable pagination (10, 25, 50, 100 rows per page).
3. **Add Item (`/add`)**:
   - Automatically determines the next sequential Item ID (e.g. `ELC-0602`).
   - Server-side validation with instant live stock valuation preview.
   - Direct parameterized insertion into MySQL.
4. **Update Stock (`/update-stock`)**:
   - Live autocomplete lookup for existing items.
   - Stock In (adds quantity, updates Last Received Date to today).
   - Stock Out (deducts quantity, prevents negative stock).
   - Automatically recomputes stock status and stock value.
5. **Edit Item (`/edit/<item_id>`)**:
   - Modifies existing product specifications, locations, limits, or costs.
6. **Delete Item (`/delete/<item_id>`)**:
   - Secure modal confirmation before executing parameterized SQL deletion.
7. **Low Stock Alerts (`/low-stock`)**:
   - Lists products requiring reorder with suggested order quantities calculated dynamically.
   - Filter tabs: All, Reorder Required, Out of Stock, and Watch list.
8. **Reports & Exports (`/reports`)**:
   - Valuation breakdown, category shares, and low-stock summaries.
   - One-click Excel `.xlsx` report generator querying live MySQL data.

---

## 6. Running the Automated Test Suite

A comprehensive test suite is included to verify all database queries, data integrity, and web endpoints:

```bash
python test_inventory_mysql.py
```

Tests performed:
- [x] MySQL server connectivity and authentication.
- [x] Verification of 601 preserved records.
- [x] Dashboard metrics and aggregation accuracy.
- [x] Multi-keyword search and category/status filtering.
- [x] Auto-generation of next Item ID.
- [x] Product insertion and retrieval.
- [x] Stock In and Stock Out movements and status updates.
- [x] Editing item details.
- [x] Safe deletion of records.
- [x] Flask HTTP routes and JSON API autocomplete.
- [x] Excel report downloads from MySQL data.

---

## 7. Troubleshooting

| Issue | Resolution |
|---|---|
| **Access Denied for user 'root'** | Verify `DB_PASSWORD` in your `.env` file matches your MySQL root password. |
| **Can't connect to MySQL server on 'localhost:3306'** | Ensure MySQL Server 8.0 is running in Windows Services (`services.msc` -> `MySQL80`). |
| **Table 'inventory_tracker.inventory' doesn't exist** | Run `python migrate_excel_to_mysql.py` or start the app with `python app.py` (which initializes tables automatically). |
| **Missing Python packages** | Run `pip install -r requirements.txt`. |
