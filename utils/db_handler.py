"""
db_handler.py
-------------
MySQL 8.0 database integration for the Inventory Stock Tracker.
Replaces Excel file storage with MySQL.

Handles:
- Automatic database & table initialization
- Thread-safe connection management
- Parameterized CRUD operations (preventing SQL injection)
- Search, filter, sorting, and pagination queries
- Dashboard statistics, category summaries, and low-stock alerts
"""

import os
import re
from datetime import date, datetime
from decimal import Decimal
import mysql.connector
from mysql.connector import errorcode
from dotenv import load_dotenv
import pandas as pd

from .inventory_logic import (
    STATUS_LOW, STATUS_OK, STATUS_OUT, STATUS_REORDER,
    compute_status, compute_value, ID_PREFIX
)

# Search current directory and parent directory for .env
load_dotenv()
_base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_env_path = os.path.join(_base_dir, ".env")
if os.path.isfile(_env_path):
    load_dotenv(_env_path)
_parent_env = os.path.join(os.path.dirname(_base_dir), ".env")
if os.path.isfile(_parent_env):
    load_dotenv(_parent_env)


class DatabaseError(Exception):
    """Plain-language database exception for user-facing messages."""
    pass


class DBHandler:
    def __init__(self):
        self.host = os.getenv("DB_HOST", "localhost")
        self.port = int(os.getenv("DB_PORT", 3306))
        self.user = os.getenv("DB_USER", "root")
        self.password = os.getenv("DB_PASSWORD", "")
        self.db_name = os.getenv("DB_NAME", "inventory_tracker")
        self._initialized = False

    def get_connection(self, include_db=True):
        """Create a new MySQL connection. Raises DatabaseError on failure."""
        config = {
            "host": self.host,
            "port": self.port,
            "user": self.user,
            "password": self.password,
            "autocommit": True,
            "charset": "utf8mb4",
            "use_pure": True,
        }
        if include_db:
            config["database"] = self.db_name

        try:
            conn = mysql.connector.connect(**config)
            return conn
        except mysql.connector.Error as err:
            if err.errno == errorcode.ER_ACCESS_DENIED_ERROR:
                raise DatabaseError(
                    f"MySQL Access Denied for user '{self.user}'. "
                    "Please verify DB_USER and DB_PASSWORD in your .env file."
                )
            elif err.errno == errorcode.ER_BAD_DB_ERROR:
                if include_db:
                    # Database does not exist yet; try creating it
                    self.init_db()
                    return mysql.connector.connect(**config)
                raise DatabaseError(f"Database '{self.db_name}' does not exist.")
            elif err.errno in (errorcode.CR_CONN_HOST_ERROR, 2003, 10013):
                raise DatabaseError(
                    f"Could not connect to MySQL Server at {self.host}:{self.port}. "
                    "Please verify that MySQL Server 8.0 service is running."
                )
            else:
                raise DatabaseError(f"MySQL Error ({err.errno}): {err.msg}")

    def init_db(self):
        """Automatically create the database and inventory table if they do not exist."""
        try:
            # 1. Connect without DB to ensure database exists
            conn = self.get_connection(include_db=False)
            cursor = conn.cursor()
            cursor.execute(
                f"CREATE DATABASE IF NOT EXISTS `{self.db_name}` "
                "CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;"
            )
            cursor.close()
            conn.close()

            # 2. Connect with DB to ensure table and indexes exist
            conn = self.get_connection(include_db=True)
            cursor = conn.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS `inventory` (
                    `item_id` VARCHAR(50) NOT NULL,
                    `name` VARCHAR(150) NOT NULL,
                    `category` VARCHAR(100) NOT NULL,
                    `spec` VARCHAR(150) DEFAULT NULL,
                    `brand` VARCHAR(100) NOT NULL,
                    `unit` VARCHAR(50) NOT NULL DEFAULT 'Nos',
                    `qty` INT NOT NULL DEFAULT 0,
                    `reorder` INT NOT NULL DEFAULT 50,
                    `min_stock` INT NOT NULL DEFAULT 10,
                    `max_stock` INT NOT NULL DEFAULT 1000,
                    `cost` DECIMAL(12, 2) NOT NULL DEFAULT 0.00,
                    `value` DECIMAL(15, 2) NOT NULL DEFAULT 0.00,
                    `supplier` VARCHAR(150) NOT NULL,
                    `warehouse` VARCHAR(100) NOT NULL,
                    `bin` VARCHAR(50) NOT NULL DEFAULT 'Not Assigned',
                    `status` VARCHAR(50) NOT NULL DEFAULT 'In Stock',
                    `received` DATE DEFAULT NULL,
                    `remarks` VARCHAR(255) DEFAULT NULL,
                    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    `updated_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    PRIMARY KEY (`item_id`),
                    INDEX `idx_category` (`category`),
                    INDEX `idx_status` (`status`),
                    INDEX `idx_brand` (`brand`),
                    INDEX `idx_supplier` (`supplier`),
                    INDEX `idx_warehouse` (`warehouse`),
                    INDEX `idx_qty` (`qty`)
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
            """)
            cursor.close()
            conn.close()
            self._initialized = True
        except Exception as exc:
            if isinstance(exc, DatabaseError):
                raise
            raise DatabaseError(f"Database initialization failed: {exc}")

    def _ensure_initialized(self):
        if not self._initialized:
            self.init_db()

    # ------------------------------------------------------------------
    # Item CRUD Operations
    # ------------------------------------------------------------------
    def get_item(self, item_id):
        """Retrieve a single item by Item ID (case-insensitive)."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)
            cursor.execute(
                "SELECT * FROM `inventory` WHERE LOWER(`item_id`) = LOWER(%s) LIMIT 1",
                (str(item_id).strip(),)
            )
            row = cursor.fetchone()
            cursor.close()
            if row:
                row = self._format_row(row)
            return row
        finally:
            conn.close()

    def add_item(self, record):
        """Insert a new inventory item."""
        self._ensure_initialized()
        existing = self.get_item(record["item_id"])
        if existing:
            raise DatabaseError(f"Item ID '{record['item_id']}' already exists.")

        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            query = """
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
            params = (
                record["item_id"],
                record["name"],
                record["category"],
                record.get("spec") or None,
                record["brand"],
                record.get("unit", "Nos"),
                int(record.get("qty", 0)),
                int(record.get("reorder", 50)),
                int(record.get("min_stock", 10)),
                int(record.get("max_stock", 1000)),
                float(record.get("cost", 0.0)),
                float(record.get("value", 0.0)),
                record["supplier"],
                record["warehouse"],
                record.get("bin", "Not Assigned"),
                record.get("status", STATUS_OK),
                record.get("received") or None,
                record.get("remarks") or None,
            )
            cursor.execute(query, params)
            cursor.close()
        except mysql.connector.Error as err:
            raise DatabaseError(f"Failed to add item: {err.msg}")
        finally:
            conn.close()

    def update_item(self, item_id, record):
        """Update an existing item's details."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            query = """
                UPDATE `inventory` SET
                    `name` = %s,
                    `category` = %s,
                    `spec` = %s,
                    `brand` = %s,
                    `unit` = %s,
                    `qty` = %s,
                    `reorder` = %s,
                    `min_stock` = %s,
                    `max_stock` = %s,
                    `cost` = %s,
                    `value` = %s,
                    `supplier` = %s,
                    `warehouse` = %s,
                    `bin` = %s,
                    `status` = %s,
                    `received` = %s,
                    `remarks` = %s
                WHERE LOWER(`item_id`) = LOWER(%s)
            """
            params = (
                record["name"],
                record["category"],
                record.get("spec") or None,
                record["brand"],
                record.get("unit", "Nos"),
                int(record.get("qty", 0)),
                int(record.get("reorder", 50)),
                int(record.get("min_stock", 10)),
                int(record.get("max_stock", 1000)),
                float(record.get("cost", 0.0)),
                float(record.get("value", 0.0)),
                record["supplier"],
                record["warehouse"],
                record.get("bin", "Not Assigned"),
                record.get("status", STATUS_OK),
                record.get("received") or None,
                record.get("remarks") or None,
                str(item_id).strip(),
            )
            cursor.execute(query, params)
            if cursor.rowcount == 0:
                raise DatabaseError(f"Item '{item_id}' not found.")
            cursor.close()
        except mysql.connector.Error as err:
            raise DatabaseError(f"Failed to update item: {err.msg}")
        finally:
            conn.close()

    def update_quantity(self, item_id, new_qty, received=None):
        """Update item stock quantity, recalculating status and stock value."""
        self._ensure_initialized()
        item = self.get_item(item_id)
        if not item:
            raise DatabaseError(f"Item '{item_id}' was not found in the database.")

        cost = float(item.get("cost", 0.0))
        reorder = int(item.get("reorder", 50))
        new_val = compute_value(new_qty, cost)
        new_status = compute_status(new_qty, reorder)

        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            if received is not None:
                query = """
                    UPDATE `inventory`
                    SET `qty` = %s, `value` = %s, `status` = %s, `received` = %s
                    WHERE LOWER(`item_id`) = LOWER(%s)
                """
                params = (new_qty, new_val, new_status, received, str(item_id).strip())
            else:
                query = """
                    UPDATE `inventory`
                    SET `qty` = %s, `value` = %s, `status` = %s
                    WHERE LOWER(`item_id`) = LOWER(%s)
                """
                params = (new_qty, new_val, new_status, str(item_id).strip())

            cursor.execute(query, params)
            cursor.close()
        except mysql.connector.Error as err:
            raise DatabaseError(f"Failed to update quantity: {err.msg}")
        finally:
            conn.close()

    def delete_item(self, item_id):
        """Delete an item from inventory."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute(
                "DELETE FROM `inventory` WHERE LOWER(`item_id`) = LOWER(%s)",
                (str(item_id).strip(),)
            )
            if cursor.rowcount == 0:
                raise DatabaseError(f"Item '{item_id}' was not found in the database.")
            cursor.close()
        except mysql.connector.Error as err:
            raise DatabaseError(f"Failed to delete item: {err.msg}")
        finally:
            conn.close()

    def next_item_id(self):
        """Determine next free Item ID (e.g. ELC-0602)."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT `item_id` FROM `inventory` WHERE `item_id` LIKE 'ELC-%'")
            ids = [row[0] for row in cursor.fetchall()]
            cursor.close()

            highest = 0
            for item_id in ids:
                match = re.fullmatch(re.escape(ID_PREFIX) + r"(\d+)", str(item_id))
                if match:
                    highest = max(highest, int(match.group(1)))
            existing = set(ids)
            number = highest + 1
            while f"{ID_PREFIX}{number:04d}" in existing:
                number += 1
            return f"{ID_PREFIX}{number:04d}"
        finally:
            conn.close()

    # ------------------------------------------------------------------
    # Search, Filter, Sort, Paginate
    # ------------------------------------------------------------------
    def search_and_filter(self, query="", category="", status="", sort="item_id", order="asc", page=1, per_page=25):
        """Filter, sort, and paginate inventory items directly using SQL."""
        self._ensure_initialized()
        conditions = []
        params = []

        query = (query or "").strip()
        if query:
            # Multi-word search: each word must match anywhere in search columns
            for word in query.split():
                like_pattern = f"%{word}%"
                conditions.append(
                    "(CONCAT_WS(' ', `item_id`, `name`, `category`, `brand`, `supplier`) LIKE %s)"
                )
                params.append(like_pattern)

        if category:
            conditions.append("`category` = %s")
            params.append(category)

        if status:
            conditions.append("`status` = %s")
            params.append(status)

        where_clause = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        # Safe sorting column mapping
        allowed_sorts = {
            "item_id": "`item_id`",
            "name": "`name`",
            "category": "`category`",
            "brand": "`brand`",
            "qty": "`qty`",
            "cost": "`cost`",
            "value": "`value`",
            "status": "FIELD(`status`, 'Out of Stock', 'Reorder Required', 'Low Stock', 'In Stock')",
        }
        sort_col = allowed_sorts.get(sort, "`item_id`")
        direction = "DESC" if order == "desc" else "ASC"

        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)

            # 1. Total matching count and valuation sum
            agg_query = f"SELECT COUNT(*) AS total, COALESCE(SUM(`value`), 0) AS total_val FROM `inventory` {where_clause}"
            cursor.execute(agg_query, tuple(params))
            agg_res = cursor.fetchone()
            total = agg_res["total"] if agg_res else 0
            total_val = float(agg_res["total_val"]) if agg_res else 0.0

            # 2. Pagination calculation
            pages = max(1, -(-total // per_page))
            page = min(max(1, page), pages)
            offset = (page - 1) * per_page

            # 3. Fetch paginated records
            order_clause = f"ORDER BY {sort_col} {direction}, `item_id` ASC"
            data_query = f"SELECT * FROM `inventory` {where_clause} {order_clause} LIMIT %s OFFSET %s"
            data_params = list(params) + [per_page, offset]
            cursor.execute(data_query, tuple(data_params))
            rows = cursor.fetchall()
            cursor.close()

            formatted_rows = [self._format_row(r) for r in rows]

            # Build pagination window
            numbers = sorted({1, pages, page - 1, page, page + 1} & set(range(1, pages + 1)))
            window = []
            for n in numbers:
                if window and n - window[-1] > 1:
                    window.append(None)
                window.append(n)

            paging_info = {
                "page": page,
                "pages": pages,
                "total": total,
                "per_page": per_page,
                "first": offset + 1 if total else 0,
                "last": min(offset + per_page, total),
                "window": window,
            }

            return formatted_rows, paging_info, int(round(total_val))
        finally:
            conn.close()

    def api_search_items(self, query, limit=8):
        """Autocomplete search for Update Stock."""
        self._ensure_initialized()
        query = (query or "").strip()
        if not query:
            return []

        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)
            words = query.split()
            conditions = []
            params = []
            for word in words:
                conditions.append(
                    "CONCAT_WS(' ', `item_id`, `name`, `category`, `brand`, `supplier`) LIKE %s"
                )
                params.append(f"%{word}%")

            where_clause = "WHERE " + " AND ".join(conditions)
            sql = f"""
                SELECT `item_id`, `name`, `brand`, `qty`, `unit`, `status`
                FROM `inventory`
                {where_clause}
                ORDER BY `item_id` ASC
                LIMIT %s
            """
            params.append(limit)
            cursor.execute(sql, tuple(params))
            rows = cursor.fetchall()
            cursor.close()
            return [
                {
                    "item_id": r["item_id"],
                    "name": r["name"],
                    "brand": r["brand"],
                    "qty": int(r["qty"]),
                    "unit": r["unit"],
                    "status": r["status"]
                }
                for r in rows
            ]
        finally:
            conn.close()

    # ------------------------------------------------------------------
    # Dashboard and Reports Queries
    # ------------------------------------------------------------------
    def get_dashboard_stats(self):
        """Aggregate stats for Dashboard and Reports."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)
            cursor.execute("""
                SELECT
                    COUNT(*) AS total_items,
                    COALESCE(SUM(`qty`), 0) AS total_qty,
                    COALESCE(SUM(`value`), 0) AS total_value,
                    COALESCE(SUM(CASE WHEN `status` = 'In Stock' THEN 1 ELSE 0 END), 0) AS in_stock,
                    COALESCE(SUM(CASE WHEN `status` = 'Low Stock' THEN 1 ELSE 0 END), 0) AS low_stock,
                    COALESCE(SUM(CASE WHEN `status` = 'Reorder Required' THEN 1 ELSE 0 END), 0) AS reorder,
                    COALESCE(SUM(CASE WHEN `status` = 'Out of Stock' THEN 1 ELSE 0 END), 0) AS out_of_stock
                FROM `inventory`
            """)
            stats = cursor.fetchone()
            cursor.close()
            return {
                "total_items": int(stats["total_items"]),
                "total_qty": int(stats["total_qty"]),
                "total_value": float(stats["total_value"]),
                "in_stock": int(stats["in_stock"]),
                "low_stock": int(stats["low_stock"]),
                "reorder": int(stats["reorder"]),
                "out_of_stock": int(stats["out_of_stock"]),
            }
        finally:
            conn.close()

    def get_status_distribution(self):
        """List of {label, value} for the status chart."""
        stats = self.get_dashboard_stats()
        return [
            {"label": STATUS_OK, "value": stats["in_stock"]},
            {"label": STATUS_LOW, "value": stats["low_stock"]},
            {"label": STATUS_REORDER, "value": stats["reorder"]},
            {"label": STATUS_OUT, "value": stats["out_of_stock"]},
        ]

    def get_category_summary(self):
        """Category aggregation for dashboard and reports."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)
            cursor.execute("""
                SELECT
                    `category`,
                    COUNT(*) AS `items`,
                    COALESCE(SUM(`qty`), 0) AS `qty`,
                    COALESCE(SUM(`value`), 0) AS `value`,
                    COALESCE(SUM(CASE WHEN `qty` <= `reorder` THEN 1 ELSE 0 END), 0) AS `needs_reorder`
                FROM `inventory`
                GROUP BY `category`
                ORDER BY `value` DESC
            """)
            rows = cursor.fetchall()
            cursor.close()
            for r in rows:
                r["items"] = int(r["items"])
                r["qty"] = int(r["qty"])
                r["value"] = round(float(r["value"]), 2)
                r["needs_reorder"] = int(r["needs_reorder"])
            return rows
        finally:
            conn.close()

    def get_alert_items(self, level="all", limit=None):
        """Fetch alert items with suggested order calculation."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)
            where = ""
            if level == "out":
                where = "WHERE `status` = 'Out of Stock'"
            elif level == "reorder":
                where = "WHERE `status` = 'Reorder Required'"
            elif level == "watch":
                where = "WHERE `status` = 'Low Stock'"
            else:  # all
                where = "WHERE `qty` <= `reorder`"

            limit_clause = f"LIMIT {int(limit)}" if limit else ""
            sql = f"""
                SELECT *,
                       GREATEST(0, `max_stock` - `qty`) AS `suggested_order`,
                       (`qty` / NULLIF(`reorder`, 0)) AS `_ratio`
                FROM `inventory`
                {where}
                ORDER BY `_ratio` ASC, `item_id` ASC
                {limit_clause}
            """
            cursor.execute(sql)
            rows = cursor.fetchall()
            cursor.close()
            return [self._format_row(r) for r in rows]
        finally:
            conn.close()

    def get_alert_counts(self):
        """Counts for Low Stock tabs."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)
            cursor.execute("""
                SELECT
                    COALESCE(SUM(CASE WHEN `qty` <= `reorder` THEN 1 ELSE 0 END), 0) AS `all`,
                    COALESCE(SUM(CASE WHEN `status` = 'Out of Stock' THEN 1 ELSE 0 END), 0) AS `out`,
                    COALESCE(SUM(CASE WHEN `status` = 'Reorder Required' THEN 1 ELSE 0 END), 0) AS `reorder`,
                    COALESCE(SUM(CASE WHEN `status` = 'Low Stock' THEN 1 ELSE 0 END), 0) AS `watch`
                FROM `inventory`
            """)
            row = cursor.fetchone()
            cursor.close()
            return {
                "all": int(row["all"]),
                "out": int(row["out"]),
                "reorder": int(row["reorder"]),
                "watch": int(row["watch"]),
            }
        finally:
            conn.close()

    def get_form_options(self):
        """Distinct values for dropdown suggestions."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            def _get_distinct(col):
                cursor.execute(f"SELECT DISTINCT `{col}` FROM `inventory` WHERE `{col}` IS NOT NULL AND `{col}` != '' ORDER BY `{col}` ASC")
                return [r[0] for r in cursor.fetchall()]

            options = {
                "categories": _get_distinct("category"),
                "brands": _get_distinct("brand"),
                "suppliers": _get_distinct("supplier"),
                "warehouses": _get_distinct("warehouse"),
                "units": _get_distinct("unit") or ["Nos", "Box", "Pack", "Set", "Roll", "Mtr"],
            }
            cursor.close()
            return options
        finally:
            conn.close()

    def get_all_dataframe(self):
        """Retrieve the entire inventory as a pandas DataFrame."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor(dictionary=True)
            cursor.execute("SELECT * FROM `inventory` ORDER BY `item_id` ASC")
            rows = cursor.fetchall()
            cursor.close()
            if not rows:
                return pd.DataFrame()
            df = pd.DataFrame([self._format_row(r) for r in rows])
            df["received"] = pd.to_datetime(df["received"], errors="coerce")
            return df
        finally:
            conn.close()

    def count_records(self):
        """Total number of inventory records in the database."""
        self._ensure_initialized()
        conn = self.get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM `inventory`")
            count = cursor.fetchone()[0]
            cursor.close()
            return int(count)
        finally:
            conn.close()

    @staticmethod
    def _format_row(row):
        """Ensure correct Python data types for template and logic compatibility."""
        row["qty"] = int(row.get("qty") or 0)
        row["reorder"] = int(row.get("reorder") or 0)
        row["min_stock"] = int(row.get("min_stock") or 0)
        row["max_stock"] = int(row.get("max_stock") or 0)
        row["cost"] = float(row.get("cost") or 0.0)
        row["value"] = float(row.get("value") or 0.0)
        if "suggested_order" in row and row["suggested_order"] is not None:
            row["suggested_order"] = int(row["suggested_order"])

        # Format date if string or datetime
        rec = row.get("received")
        if isinstance(rec, (datetime, date)):
            pass
        elif rec:
            try:
                row["received"] = datetime.strptime(str(rec)[:10], "%Y-%m-%d").date()
            except ValueError:
                pass
        return row

