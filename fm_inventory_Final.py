"""
FM INVENTORY - Facilities Management Stock Inventory Management App
Single-file Python + SQLite + PySide6 desktop application.

Default login after first run:
    Username: admin
    Password: admin123

Required packages:
    pip install PySide6 pandas openpyxl reportlab matplotlib bcrypt

SQLite setup:
    No server, username, or password is required.
    Run: python fm_inventory_sqlite_app.py

This file creates fm_inventory.db in the same folder as the script, including schema,
default master data, default roles, admin user, sample items, transactions, audit logs,
and then launches the desktop app.
"""

from __future__ import annotations

import csv
import os
import sys
import shutil
import tempfile
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Optional, Tuple

import sqlite3

try:
    import bcrypt
except Exception:
    bcrypt = None

try:
    import pandas as pd
except Exception:
    pd = None

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas
except Exception:
    canvas = None
    A4 = landscape = None

try:
    from PySide6.QtCore import Qt, QSize
    from PySide6.QtGui import QFont, QAction
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
        QLabel, QPushButton, QLineEdit, QComboBox, QDateEdit, QTextEdit, QSpinBox,
        QDoubleSpinBox, QTableWidget, QTableWidgetItem, QMessageBox, QFileDialog,
        QFrame, QStackedWidget, QFormLayout, QDialog, QDialogButtonBox, QHeaderView,
        QCheckBox, QTabWidget, QScrollArea, QGroupBox, QInputDialog, QMenu
    )
except Exception as e:
    print("PySide6 is required. Install with: pip install PySide6")
    raise

def get_app_data_dir():
    """Return a permanent writable folder for the SQLite database.

    In PyInstaller --onefile mode, __file__ points to a temporary extraction
    folder that is deleted when the EXE closes. Storing fm_inventory.db there
    causes all data to disappear. This keeps the database permanently under
    the user's AppData folder when running as an EXE, and beside the script
    during normal Python development.
    """
    if getattr(sys, "frozen", False):
        base = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "StoreInventoryManagement")
    else:
        base = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    os.makedirs(base, exist_ok=True)
    return base


BASE_DIR = get_app_data_dir()
DB_PATH = os.path.join(BASE_DIR, "fm_inventory.db")


APP_TITLE = "STORE INVENTORY MANAGEMENT"
MATERIAL_TYPES = ["Critical Spare", "Fast Moving Spare", "General Spare"]
ADJUSTMENT_TYPES = ["Increase", "Decrease"]
DEFAULT_CATEGORIES = ["Electrical", "Plumbing", "HVAC", "Civil", "General", "Cleaning", "Other"]
DEFAULT_UNITS = ["Nos", "Kg", "Litre", "Meter", "Box", "Set"]
DEFAULT_DEPARTMENTS = ["Facilities", "Electrical", "Plumbing", "HVAC", "Civil", "Cleaning", "Administration"]
DEFAULT_LOCATIONS = ["Main Store", "Electrical Store", "HVAC Store", "Plumbing Store", "Site Store - A", "Site Store - B"]
DEFAULT_RACKS = ["R-01", "R-02", "R-03", "R-04", "R-05"]

STYLE = """
* { font-family: Segoe UI, Arial; font-size: 13px; }
QMainWindow, QWidget { background: #f5f7fb; color: #0b1833; }
#sidebar { background: #102a43; color: #ffffff; border: none; }
#sideTitle { background: transparent; color: #ffffff; font-size: 16px; font-weight: 900; letter-spacing: .2px; padding: 4px 2px; }
#sideSub { background: transparent; color: #c8d8e8; font-size: 12px; padding: 0 2px 8px 2px; }
#sectionLabel { background: transparent; color:#90cdf4; font-size:11px; font-weight:900; letter-spacing:1.1px; padding: 16px 4px 6px 4px; }
QPushButton#nav { color: #eaf6ff; background: transparent; border: none; text-align: left; padding: 12px 14px; border-radius: 10px; font-weight: 700; }
QPushButton#nav:hover { background: #1f3f5f; color: #ffffff; }
QPushButton#nav[active="true"] { background: #2f80ed; color: white; font-weight: 900; }
QPushButton { background: #1769ff; color: white; border: none; padding: 10px 16px; border-radius: 9px; font-weight: 800; }
QPushButton:hover { background: #095bd8; }
QPushButton#softBtn { background: #edf4ff; color: #0b3b86; border: 1px solid #cfe0ff; }
QPushButton#greenBtn { background: #0aa36d; color: white; }
QPushButton#orangeBtn { background: #ff7a1a; color: white; }
QPushButton#redBtn { background: #dc3545; color: white; }
QLineEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox, QTextEdit { background: white; color:#0b1833; border: 1px solid #cbd8ea; border-radius: 8px; padding: 8px; selection-background-color:#1769ff; }
QLineEdit:focus, QComboBox:focus, QDateEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QTextEdit:focus { border: 1px solid #1769ff; }
QTableWidget { background: white; color:#0b1833; gridline-color: #e7edf6; border: 1px solid #d8e3f2; border-radius: 10px; alternate-background-color: #f9fbff; }
QTableWidget::item { padding: 8px; }
QHeaderView::section { background: #eaf1fb; color: #071a3a; padding: 10px; border: none; border-bottom: 1px solid #d6e2f0; font-weight: 900; }
QFrame#card, QGroupBox { background: white; border: 1px solid #dce6f3; border-radius: 14px; }
QFrame#warn { background: #fff8ec; border: 1px solid #ffdba8; border-radius: 14px; }
QFrame#dangerCard { background: #fff1f3; border: 1px solid #ffc8cf; border-radius: 14px; }
QFrame#topbar { background: transparent; border: none; }
QLabel { background: transparent; }
QLabel#metric { font-size: 28px; font-weight: 900; color: #071a3a; }
QLabel#small { color: #64748b; font-size: 12px; }
QLabel#title { color:#071a3a; font-size: 26px; font-weight: 900; }
QLabel#subtitle { color:#53657e; font-size: 13px; }
QGroupBox { margin-top: 12px; padding-top: 16px; font-weight: 900; }
QGroupBox::title { subcontrol-origin: margin; left: 14px; padding: 0 7px; color:#071a3a; }
QTabWidget::pane { border: 1px solid #dfe7f2; border-radius: 12px; background: white; }
QTabBar::tab { padding: 10px 16px; background: #eaf1fb; border-top-left-radius: 8px; border-top-right-radius: 8px; margin-right: 3px; }
QTabBar::tab:selected { background: white; color: #1769ff; font-weight: 900; }
"""


def today_iso() -> str:
    return date.today().isoformat()


def add_months(src_date: date, months: int) -> date:
    """Return src_date plus whole calendar months, clamping day to month-end."""
    month = src_date.month - 1 + months
    year = src_date.year + month // 12
    month = month % 12 + 1
    days_in_month = [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]
    return date(year, month, min(src_date.day, days_in_month))


def parse_iso_date(value: str) -> date:
    return datetime.strptime(str(value), "%Y-%m-%d").date()


def money(v: Any) -> str:
    try:
        return f"OMR {float(v):,.3f}"
    except Exception:
        return "OMR 0.000"


def hash_password(password: str) -> str:
    if bcrypt:
        return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    import hashlib
    return "sha256$" + hashlib.sha256(password.encode()).hexdigest()


def verify_password(password: str, hashed: str) -> bool:
    if hashed.startswith("sha256$"):
        import hashlib
        return hashlib.sha256(password.encode()).hexdigest() == hashed.split("$", 1)[1]
    if bcrypt:
        return bcrypt.checkpw(password.encode(), hashed.encode())
    return False


class Database:
    """SQLite database wrapper. No SQLite server, username, or password is needed."""

    def __init__(self, db_path: str):
        self.db_path = db_path
        self.create_schema()
        self.seed_defaults()

    def connect(self):
        con = sqlite3.connect(self.db_path)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA foreign_keys = ON")
        return con

    def _sql(self, sql: str) -> str:
        # Compatibility layer for the earlier SQLite-style code.
        sql = sql.replace("%s", "?")
        sql = sql.replace("INSERT IGNORE", "INSERT OR IGNORE")
        sql = sql.replace("CURDATE()", "DATE('now')")
        return sql

    def _clean_params(self, params: Tuple = ()) -> Tuple:
        # SQLite does not accept Decimal objects returned/used by some widgets/calculations.
        # Convert every DB-bound value to SQLite-safe primitive types before execution.
        cleaned = []
        for value in tuple(params or ()):
            if isinstance(value, Decimal):
                cleaned.append(float(value))
            elif hasattr(value, "isoformat") and value.__class__.__name__ in {"date", "datetime"}:
                cleaned.append(value.isoformat())
            elif value is True:
                cleaned.append(1)
            elif value is False:
                cleaned.append(0)
            else:
                cleaned.append(value)
        return tuple(cleaned)

    def execute(self, sql: str, params: Tuple = (), commit: bool = True) -> int:
        sql = self._sql(sql)
        params = self._clean_params(params)
        con = self.connect()
        try:
            cur = con.cursor()
            cur.execute(sql, params)
            last_id = cur.lastrowid
            if commit:
                con.commit()
            return last_id
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()

    def query(self, sql: str, params: Tuple = ()) -> List[Dict[str, Any]]:
        sql = self._sql(sql)
        params = self._clean_params(params)
        con = self.connect()
        try:
            cur = con.cursor()
            cur.execute(sql, params)
            return [dict(r) for r in cur.fetchall()]
        finally:
            con.close()

    def one(self, sql: str, params: Tuple = ()) -> Optional[Dict[str, Any]]:
        rows = self.query(sql, params)
        return rows[0] if rows else None

    def create_schema(self):
        statements = [
            """CREATE TABLE IF NOT EXISTS roles (
                id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, description TEXT,
                is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
                full_name TEXT, role_id INTEGER, is_active INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(role_id) REFERENCES roles(id))""",
            """CREATE TABLE IF NOT EXISTS permissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT, role_id INTEGER, module TEXT, can_view INTEGER DEFAULT 1,
                can_add INTEGER DEFAULT 0, can_edit INTEGER DEFAULT 0, can_delete INTEGER DEFAULT 0,
                FOREIGN KEY(role_id) REFERENCES roles(id))""",
            """CREATE TABLE IF NOT EXISTS departments (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS categories (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS units (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS suppliers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, contact TEXT, email TEXT, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS locations (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS racks (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, location_id INTEGER NULL, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(location_id) REFERENCES locations(id))""",
            """CREATE TABLE IF NOT EXISTS inventory_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT, material_id TEXT UNIQUE NOT NULL, item_description TEXT NOT NULL,
                category_id INTEGER, department_id INTEGER, quantity_available REAL DEFAULT 0, unit_id INTEGER,
                minimum_stock_level REAL DEFAULT 0, material_type TEXT, supplier_id INTEGER NULL,
                cost_per_unit REAL DEFAULT 0, procurement_date TEXT NULL, purchase_order_no TEXT, delivery_order_no TEXT,
                warranty_details TEXT, expiry_date TEXT NULL, rack_id INTEGER NULL, location_id INTEGER NULL, remarks TEXT,
                is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(category_id) REFERENCES categories(id), FOREIGN KEY(department_id) REFERENCES departments(id), FOREIGN KEY(unit_id) REFERENCES units(id),
                FOREIGN KEY(supplier_id) REFERENCES suppliers(id), FOREIGN KEY(rack_id) REFERENCES racks(id), FOREIGN KEY(location_id) REFERENCES locations(id))""",
            """CREATE TABLE IF NOT EXISTS stock_inward (
                id INTEGER PRIMARY KEY AUTOINCREMENT, inward_no TEXT UNIQUE, material_id TEXT, item_id INTEGER, quantity REAL, unit_cost REAL, total_cost REAL,
                procurement_date TEXT, supplier_id INTEGER NULL, po_no TEXT, do_no TEXT, remarks TEXT, created_by INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(item_id) REFERENCES inventory_items(id), FOREIGN KEY(supplier_id) REFERENCES suppliers(id), FOREIGN KEY(created_by) REFERENCES users(id))""",
            """CREATE TABLE IF NOT EXISTS inventory_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT, batch_no TEXT UNIQUE NOT NULL, inward_id INTEGER NULL, item_id INTEGER NOT NULL, material_id TEXT NOT NULL,
                quantity_received REAL DEFAULT 0, quantity_available REAL DEFAULT 0, expiry_date TEXT NOT NULL,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(inward_id) REFERENCES stock_inward(id), FOREIGN KEY(item_id) REFERENCES inventory_items(id))""",
            """CREATE TABLE IF NOT EXISTS stock_outward (
                id INTEGER PRIMARY KEY AUTOINCREMENT, siv_no TEXT UNIQUE, issue_date TEXT, material_id TEXT, item_id INTEGER, quantity_issued REAL,
                department_id INTEGER, issued_to TEXT, issued_by TEXT, purpose TEXT, remarks TEXT, created_by INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(item_id) REFERENCES inventory_items(id), FOREIGN KEY(department_id) REFERENCES departments(id), FOREIGN KEY(created_by) REFERENCES users(id))""",
            """CREATE TABLE IF NOT EXISTS stock_outward_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT, outward_id INTEGER NOT NULL, batch_id INTEGER NOT NULL, material_id TEXT NOT NULL, quantity_issued REAL DEFAULT 0,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(outward_id) REFERENCES stock_outward(id), FOREIGN KEY(batch_id) REFERENCES inventory_batches(id))""",
            """CREATE TABLE IF NOT EXISTS stock_adjustments (
                id INTEGER PRIMARY KEY AUTOINCREMENT, adjustment_no TEXT UNIQUE, adjustment_date TEXT, material_id TEXT, item_id INTEGER, current_stock REAL,
                adjustment_type TEXT, adjustment_quantity REAL, reason TEXT, approved_by TEXT, remarks TEXT, created_by INTEGER,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(item_id) REFERENCES inventory_items(id), FOREIGN KEY(created_by) REFERENCES users(id))""",
            """CREATE TABLE IF NOT EXISTS stock_transfers (
                id INTEGER PRIMARY KEY AUTOINCREMENT, transfer_no TEXT UNIQUE, transfer_date TEXT, material_id TEXT, item_id INTEGER, quantity_transferred REAL,
                from_department_id INTEGER, to_department_id INTEGER, from_location_id INTEGER NULL, to_location_id INTEGER NULL, from_rack_id INTEGER NULL, to_rack_id INTEGER NULL,
                requested_by TEXT, approved_by TEXT, remarks TEXT, created_by INTEGER, created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(item_id) REFERENCES inventory_items(id), FOREIGN KEY(from_department_id) REFERENCES departments(id), FOREIGN KEY(to_department_id) REFERENCES departments(id), FOREIGN KEY(created_by) REFERENCES users(id))""",
            """CREATE TABLE IF NOT EXISTS report_templates (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, report_type TEXT, filters_json TEXT, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS report_template_columns (id INTEGER PRIMARY KEY AUTOINCREMENT, template_id INTEGER, column_name TEXT, display_order INTEGER, is_visible INTEGER DEFAULT 1, FOREIGN KEY(template_id) REFERENCES report_templates(id))""",
            """CREATE TABLE IF NOT EXISTS app_settings (setting_key TEXT PRIMARY KEY, setting_value TEXT, updated_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS custom_fields (id INTEGER PRIMARY KEY AUTOINCREMENT, field_label TEXT, field_type TEXT DEFAULT 'Text', is_required INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP)""",
            """CREATE TABLE IF NOT EXISTS audit_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NULL, action TEXT, module TEXT, record_ref TEXT, details TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(user_id) REFERENCES users(id))""",
            "CREATE INDEX IF NOT EXISTS idx_inventory_mat ON inventory_items(material_id)",
            "CREATE INDEX IF NOT EXISTS idx_inventory_cat ON inventory_items(category_id)",
            "CREATE INDEX IF NOT EXISTS idx_inventory_dept ON inventory_items(department_id)",
            "CREATE INDEX IF NOT EXISTS idx_inventory_supplier ON inventory_items(supplier_id)",
            "CREATE INDEX IF NOT EXISTS idx_inventory_expiry ON inventory_items(expiry_date)",
            "CREATE INDEX IF NOT EXISTS idx_batches_item ON inventory_batches(item_id)",
            "CREATE INDEX IF NOT EXISTS idx_batches_expiry ON inventory_batches(expiry_date)",
            "CREATE INDEX IF NOT EXISTS idx_batches_available ON inventory_batches(quantity_available)",
            "CREATE INDEX IF NOT EXISTS idx_inward_date ON stock_inward(procurement_date)",
            "CREATE INDEX IF NOT EXISTS idx_outward_date ON stock_outward(issue_date)",
            "CREATE INDEX IF NOT EXISTS idx_adj_date ON stock_adjustments(adjustment_date)",
            "CREATE INDEX IF NOT EXISTS idx_transfer_date ON stock_transfers(transfer_date)",
            "CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_logs(created_at)",
        ]
        for s in statements:
            self.execute(s)

    def seed_defaults(self):
        for role in ["Admin", "Store Keeper", "Manager", "Viewer"]:
            self.execute("INSERT OR IGNORE INTO roles(name, description) VALUES(?,?)", (role, role))
        admin_role = self.one("SELECT id FROM roles WHERE name='Admin'")["id"]
        if not self.one("SELECT id FROM users WHERE username=?", ("admin",)):
            self.execute("INSERT INTO users(username,password_hash,full_name,role_id) VALUES(?,?,?,?)", ("admin", hash_password("admin123"), "System Administrator", admin_role))
        for table, values in [("categories", DEFAULT_CATEGORIES), ("units", DEFAULT_UNITS), ("departments", DEFAULT_DEPARTMENTS), ("locations", DEFAULT_LOCATIONS), ("racks", DEFAULT_RACKS)]:
            for v in values:
                self.execute(f"INSERT OR IGNORE INTO {table}(name) VALUES(?)", (v,))
        self.execute("INSERT OR IGNORE INTO suppliers(name,contact,email) VALUES(?,?,?)", ("Default Supplier", "", ""))
        self.execute("INSERT OR IGNORE INTO app_settings(setting_key, setting_value) VALUES('app_name', ?), ('theme_color', ?)", (APP_TITLE, "#0d6efd"))
        # Do not seed testing/demo inventory. Keep a clean live database.
        # Also remove old demo rows if they were created by earlier versions.
        self.remove_sample_inventory()

    def next_code(self, prefix: str, table: str, column: str) -> str:
        row = self.one(f"SELECT {column} AS code FROM {table} WHERE {column} LIKE ? ORDER BY id DESC LIMIT 1", (f"{prefix}-%",))
        n = 1
        if row and row["code"]:
            try:
                n = int(str(row["code"]).split("-")[-1]) + 1
            except Exception:
                n = 1
        return f"{prefix}-{n:05d}"

    def id_by_name(self, table: str, name: str) -> Optional[int]:
        if not name:
            return None
        row = self.one(f"SELECT id FROM {table} WHERE name=?", (name,))
        return row["id"] if row else None

    def log(self, user_id: Optional[int], action: str, module: str, ref: str, details: str = ""):
        self.execute("INSERT INTO audit_logs(user_id,action,module,record_ref,details) VALUES(?,?,?,?,?)", (user_id, action, module, ref, details))

    def remove_sample_inventory(self):
        """Remove demo/testing stock inserted by early development versions.
        This does not touch user-uploaded/live inventory unless the item description
        exactly matches the original demo list.
        """
        demo_names = [
            "LED Panel Light 18W",
            "PVC Pipe 1/2 inch",
            "AC Filter",
            "Disinfectant 5L",
            "Hand Wash 1L",
            "Sanitizer",
            "Hand Sanitizer",
        ]
        placeholders = ",".join(["?"] * len(demo_names))
        rows = self.query(f"SELECT id, material_id FROM inventory_items WHERE item_description IN ({placeholders})", tuple(demo_names))
        if not rows:
            return
        item_ids = [r["id"] for r in rows]
        material_ids = [r["material_id"] for r in rows]
        item_ph = ",".join(["?"] * len(item_ids))
        mat_ph = ",".join(["?"] * len(material_ids))
        for table in ["stock_inward", "stock_outward", "stock_adjustments", "stock_transfers"]:
            self.execute(f"DELETE FROM {table} WHERE item_id IN ({item_ph})", tuple(item_ids))
            self.execute(f"DELETE FROM {table} WHERE material_id IN ({mat_ph})", tuple(material_ids))
        self.execute(f"DELETE FROM audit_logs WHERE record_ref IN ({mat_ph})", tuple(material_ids))
        self.execute(f"DELETE FROM inventory_items WHERE id IN ({item_ph})", tuple(item_ids))

    def seed_sample_inventory(self):
        samples = [
            ("LED Panel Light 18W", "Electrical", "Facilities", 25, "Nos", 10, "Fast Moving Spare", 35.5, "Main Store", "R-01", None),
            ("PVC Pipe 1/2 inch", "Plumbing", "Plumbing", 80, "Meter", 20, "General Spare", 2.2, "Plumbing Store", "R-02", None),
            ("AC Filter", "HVAC", "HVAC", 8, "Nos", 10, "Critical Spare", 12.0, "HVAC Store", "R-03", None),
            ("Disinfectant 5L", "Cleaning", "Cleaning", 20, "Litre", 12, "Fast Moving Spare", 4.5, "Main Store", "R-04", (date.today()+timedelta(days=12)).isoformat()),
            ("Hand Wash 1L", "Cleaning", "Cleaning", 4, "Litre", 10, "Fast Moving Spare", 3.0, "Main Store", "R-05", (date.today()-timedelta(days=8)).isoformat()),
        ]
        sup = self.id_by_name("suppliers", "Default Supplier")
        for item, cat, dept, qty, unit, minq, typ, cost, loc, rack, exp in samples:
            mat = self.next_code("MAT", "inventory_items", "material_id")
            item_id = self.execute("""INSERT INTO inventory_items(material_id,item_description,category_id,department_id,quantity_available,unit_id,minimum_stock_level,material_type,supplier_id,cost_per_unit,procurement_date,expiry_date,rack_id,location_id,remarks)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (mat, item, self.id_by_name("categories", cat), self.id_by_name("departments", dept), qty, self.id_by_name("units", unit), minq, typ, sup, cost, today_iso(), exp, self.id_by_name("racks", rack), self.id_by_name("locations", loc), "Sample data"))
            inward = self.next_code("IN", "stock_inward", "inward_no")
            self.execute("INSERT INTO stock_inward(inward_no,material_id,item_id,quantity,unit_cost,total_cost,procurement_date,supplier_id,remarks) VALUES(?,?,?,?,?,?,?,?,?)", (inward, mat, item_id, qty, cost, float(qty)*float(cost), today_iso(), sup, "Opening balance"))


class LoginDialog(QDialog):
    def __init__(self, db: Database):
        super().__init__(); self.db = db; self.user = None
        self.setWindowTitle("FM Inventory Login"); self.setFixedWidth(380)
        v = QVBoxLayout(self)
        title = QLabel("FM INVENTORY"); title.setFont(QFont("Segoe UI", 22, QFont.Bold)); title.setAlignment(Qt.AlignCenter)
        sub = QLabel("Facilities Management Login"); sub.setAlignment(Qt.AlignCenter); sub.setObjectName("small")
        self.username = QLineEdit(); self.username.setPlaceholderText("Username")
        self.password = QLineEdit(); self.password.setPlaceholderText("Password"); self.password.setEchoMode(QLineEdit.Password)
        btn = QPushButton("Login"); btn.clicked.connect(self.try_login)
        hint = QLabel("Default: admin / admin123"); hint.setAlignment(Qt.AlignCenter); hint.setObjectName("small")
        v.addWidget(title); v.addWidget(sub); v.addSpacing(12); v.addWidget(self.username); v.addWidget(self.password); v.addWidget(btn); v.addWidget(hint)

    def try_login(self):
        row = self.db.one("SELECT u.*, r.name role FROM users u LEFT JOIN roles r ON r.id=u.role_id WHERE username=%s AND u.is_active=1", (self.username.text().strip(),))
        if row and verify_password(self.password.text(), row["password_hash"]):
            self.user = row; self.accept()
        else:
            QMessageBox.warning(self, "Login failed", "Invalid username or password.")


class Card(QFrame):
    def __init__(self, title: str, value: str, subtitle: str = "", kind: str = "card", icon: str = "▣"):
        super().__init__(); self.setObjectName(kind); self.setMinimumHeight(116)
        outer = QHBoxLayout(self); outer.setContentsMargins(18, 16, 18, 16); outer.setSpacing(14)
        badge = QLabel(icon); badge.setAlignment(Qt.AlignCenter); badge.setFixedSize(54, 54)
        badge.setStyleSheet("background:#eef5ff;border-radius:27px;font-size:25px;color:#0f6bff;font-weight:900;")
        txt = QVBoxLayout(); txt.setSpacing(5)
        t = QLabel(title); t.setStyleSheet("font-weight:850;color:#14213d;")
        val = QLabel(value); val.setObjectName("metric")
        sub = QLabel(subtitle); sub.setObjectName("small")
        txt.addWidget(t); txt.addWidget(val); txt.addWidget(sub)
        outer.addWidget(badge); outer.addLayout(txt, 1)


class ImportPreviewDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Import Preview - Existing Store Materials")
        self.resize(1100, 650)
        self.accepted_import = False
        v = QVBoxLayout(self)
        title = QLabel("Review Excel data before saving to inventory")
        title.setObjectName("title")
        note = QLabel("Material ID will be generated automatically. Existing material_id rows will update stock if present; blank material_id rows will create new items.")
        note.setObjectName("subtitle")
        v.addWidget(title); v.addWidget(note)
        self.table = QTableWidget(); self.table.setAlternatingRowColors(True); v.addWidget(self.table, 1)
        rows = df.fillna("").to_dict("records")
        cols = list(df.columns)
        self.table.setColumnCount(len(cols)); self.table.setHorizontalHeaderLabels(cols); self.table.setRowCount(min(len(rows), 300))
        for r, row in enumerate(rows[:300]):
            for c, col in enumerate(cols):
                self.table.setItem(r, c, QTableWidgetItem(str(row.get(col, ""))))
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        if len(rows) > 300:
            v.addWidget(QLabel(f"Showing first 300 rows only. Total rows in file: {len(rows)}"))
        h = QHBoxLayout(); h.addStretch()
        cancel = QPushButton("Cancel"); cancel.setObjectName("softBtn"); cancel.clicked.connect(self.reject)
        save = QPushButton("Import Now"); save.setObjectName("greenBtn"); save.clicked.connect(self._accept)
        h.addWidget(cancel); h.addWidget(save); v.addLayout(h)

    def _accept(self):
        self.accepted_import = True
        self.accept()



class ColumnSelectDialog(QDialog):
    """Small modal dialog used by Reports to choose visible report columns."""
    def __init__(self, columns: List[Tuple[str, str]], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Report Columns")
        self.resize(520, 560)
        self.selected_columns: List[Tuple[str, str]] = []
        v = QVBoxLayout(self)
        title = QLabel("Select columns to display in the report")
        title.setStyleSheet("font-size:16px;font-weight:900;color:#071a3a;")
        note = QLabel("Uncheck the fields you do not want in the generated table/export.")
        note.setObjectName("small")
        v.addWidget(title); v.addWidget(note)
        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        holder = QWidget(); self.check_layout = QVBoxLayout(holder)
        self.checks: List[Tuple[QCheckBox, Tuple[str, str]]] = []
        for label, key in columns:
            cb = QCheckBox(label)
            cb.setChecked(True)
            cb.setStyleSheet("padding:6px;font-weight:700;")
            self.check_layout.addWidget(cb)
            self.checks.append((cb, (label, key)))
        self.check_layout.addStretch()
        scroll.setWidget(holder); v.addWidget(scroll, 1)
        tools = QHBoxLayout()
        all_btn = QPushButton("Select All"); all_btn.setObjectName("softBtn")
        none_btn = QPushButton("Clear All"); none_btn.setObjectName("softBtn")
        all_btn.clicked.connect(lambda: [cb.setChecked(True) for cb, _ in self.checks])
        none_btn.clicked.connect(lambda: [cb.setChecked(False) for cb, _ in self.checks])
        tools.addWidget(all_btn); tools.addWidget(none_btn); tools.addStretch(); v.addLayout(tools)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept_selection)
        buttons.rejected.connect(self.reject)
        v.addWidget(buttons)

    def accept_selection(self):
        self.selected_columns = [col for cb, col in self.checks if cb.isChecked()]
        if not self.selected_columns:
            QMessageBox.warning(self, "Validation", "Select at least one column.")
            return
        self.accept()

class MainWindow(QMainWindow):
    def __init__(self, db: Database, user: Dict[str, Any]):
        super().__init__(); self.db = db; self.user = user
        self.setWindowTitle(APP_TITLE); self.resize(1450, 900)
        self.nav_buttons: List[QPushButton] = []
        root = QWidget(); self.setCentralWidget(root)
        h = QHBoxLayout(root); h.setContentsMargins(0,0,0,0); h.setSpacing(0)
        self.sidebar = self.build_sidebar(); h.addWidget(self.sidebar)
        self.stack = QStackedWidget(); h.addWidget(self.stack, 1)
        self.pages = [
            ("Dashboard", self.page_dashboard), ("Stock Inward", self.page_inward), ("Stock Outward", self.page_outward),
            ("Stock Adjustment", self.page_adjustment), ("Transfers", self.page_transfer), ("Reports", self.page_reports),
            ("Inventory", self.page_inventory), ("Settings", self.page_settings)
        ]
        for name, builder in self.pages: self.stack.addWidget(builder())
        self.set_page(0)

    def build_sidebar(self) -> QWidget:
        side = QFrame(); side.setObjectName("sidebar"); side.setFixedWidth(385)
        v = QVBoxLayout(side); v.setContentsMargins(16,22,16,18); v.setSpacing(6)
        title = QLabel("▥  STORE INVENTORY MANAGEMENT"); title.setObjectName("sideTitle"); title.setWordWrap(True); title.setMinimumWidth(340)
        sub = QLabel("DESIGNED & DEVELOPED - PAVAN KUMAR AKELLA"); sub.setObjectName("sideSub"); sub.setWordWrap(True); sub.setMinimumWidth(340)
        v.addWidget(title); v.addWidget(sub); v.addSpacing(18)
        sections = [
            ("MAIN", [("⌂", "Dashboard")]),
            ("STORE MANAGEMENT", [("↓", "Stock Inward"),("↑", "Stock Outward"),("↕", "Stock Adjustment"),("⇄", "Transfers")]),
            ("REPORTS", [("☷", "Reports")]),
            ("INVENTORY", [("▦", "Inventory")]),
            ("SETTINGS", [("⚙", "Settings")])
        ]
        page_index = 0
        for sec, items in sections:
            lab = QLabel(sec); lab.setObjectName("sectionLabel"); v.addWidget(lab)
            for icon, item in items:
                b = QPushButton(f"  {icon}   {item}"); b.setObjectName("nav"); b.setProperty("active", False)
                b.clicked.connect(lambda _, i=page_index: self.set_page(i)); self.nav_buttons.append(b); v.addWidget(b); page_index += 1
        v.addStretch(); out = QPushButton("  ↪   Logout"); out.setObjectName("nav"); out.clicked.connect(self.close); v.addWidget(out)
        return side

    def set_page(self, i: int):
        self.stack.setCurrentIndex(i)
        for n,b in enumerate(self.nav_buttons): b.setProperty("active", n==i); b.style().unpolish(b); b.style().polish(b)
        if i == 0:
            self.refresh_dashboard()
        elif i == 1 and hasattr(self, "clear_inward_form"):
            self.clear_inward_form()
            self.refresh_inward_table()
        elif i == 2 and hasattr(self, "clear_outward_form"):
            self.clear_outward_form()
            self.refresh_outward_table()
        elif i == 3 and hasattr(self, "clear_adjustment_form"):
            self.clear_adjustment_form()
            self.refresh_adjustment_table()
        elif i == 4 and hasattr(self, "clear_transfer_form"):
            self.clear_transfer_form()
            self.refresh_transfer_table()
        elif i == 5 and hasattr(self, "generate_report"):
            self.generate_report()

    def master_names(self, table: str) -> List[str]:
        return [r["name"] for r in self.db.query(f"SELECT name FROM {table} WHERE is_active=1 ORDER BY name")]

    def item_rows(self, search: str = "") -> List[Dict[str, Any]]:
        q = """SELECT i.*, c.name category, d.name department, u.name unit, s.name supplier, l.name location, r.name rack,
            (i.quantity_available*i.cost_per_unit) total_value FROM inventory_items i
            LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN units u ON u.id=i.unit_id
            LEFT JOIN suppliers s ON s.id=i.supplier_id LEFT JOIN locations l ON l.id=i.location_id LEFT JOIN racks r ON r.id=i.rack_id
            WHERE i.is_active=1"""
        params: Tuple = ()
        if search:
            like = f"%{search}%"; q += " AND (i.material_id LIKE %s OR i.item_description LIKE %s OR c.name LIKE %s OR d.name LIKE %s OR s.name LIKE %s OR l.name LIKE %s OR r.name LIKE %s)"; params = (like,)*7
        q += " ORDER BY i.updated_at DESC"
        return self.db.query(q, params)

    def table_fill(self, table: QTableWidget, rows: List[Dict[str, Any]], columns: List[Tuple[str,str]]):
        table.setColumnCount(len(columns)); table.setHorizontalHeaderLabels([c[0] for c in columns]); table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            for c, (_, key) in enumerate(columns):
                val = row.get(key, "")
                if isinstance(val, Decimal): val = f"{float(val):.2f}"
                if isinstance(val, (datetime, date)): val = val.isoformat()
                item = QTableWidgetItem("" if val is None else str(val)); item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                table.setItem(r,c,item)
        table.setWordWrap(True)
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table.horizontalHeader().setStretchLastSection(True)

    def base_page(self, title: str) -> Tuple[QWidget, QVBoxLayout]:
        w = QWidget(); v = QVBoxLayout(w); v.setContentsMargins(26,22,26,22); v.setSpacing(14)
        top_frame = QFrame(); top_frame.setObjectName("topbar")
        top = QHBoxLayout(top_frame); top.setContentsMargins(0,0,0,0)
        lab = QLabel(title); lab.setObjectName("title")
        top.addWidget(lab); top.addStretch()
        user = QLabel(f"  👤  {self.user.get('full_name') or self.user.get('username')}  "); user.setStyleSheet("background:white;border:1px solid #dfe7f2;border-radius:18px;padding:8px 12px;font-weight:700;")
        top.addWidget(user)
        v.addWidget(top_frame); return w,v

    def page_dashboard(self):
        w,v = self.base_page("Dashboard")
        self.dash_grid = QGridLayout(); v.addLayout(self.dash_grid)
        return w

    def refresh_dashboard(self):
        while self.dash_grid.count():
            item = self.dash_grid.takeAt(0); widget = item.widget();
            if widget: widget.deleteLater()
        m = self.db.one("""SELECT COUNT(*) items, COALESCE(SUM(quantity_available*cost_per_unit),0) stock_value,
            SUM(CASE WHEN quantity_available <= minimum_stock_level THEN 1 ELSE 0 END) low_stock
            FROM inventory_items WHERE is_active=1""") or {}
        exp_stats = self.db.one("""SELECT
            SUM(CASE WHEN expiry_date BETWEEN DATE('now') AND DATE('now','+30 day') AND quantity_available>0 THEN 1 ELSE 0 END) expiring,
            SUM(CASE WHEN expiry_date < DATE('now') AND quantity_available>0 THEN 1 ELSE 0 END) expired
            FROM inventory_batches""") or {}
        m["expiring"] = exp_stats.get("expiring") or 0
        m["expired"] = exp_stats.get("expired") or 0
        inward = self.db.one("SELECT COALESCE(SUM(quantity),0) q FROM stock_inward")["q"]
        outward = self.db.one("SELECT COALESCE(SUM(quantity_issued),0) q FROM stock_outward")["q"]
        cards = [("Total Items", str(m.get("items",0)), "All items in inventory", "card"), ("Total Stock Inward", f"{float(inward):,.0f}", "Total inward quantity", "card"), ("Total Stock Outward", f"{float(outward):,.0f}", "Total outward quantity", "card"), ("Total Stock Value", money(m.get("stock_value",0)), "Overall inventory value", "card"), ("Low Stock Items", str(m.get("low_stock") or 0), "Require attention", "warn"), ("To Be Expired This Month", str(m.get("expiring") or 0), "Items to expire", "dangerCard"), ("Expired Items", str(m.get("expired") or 0), "Already expired", "warn")]
        for idx,c in enumerate(cards): self.dash_grid.addWidget(Card(*c), idx//4, idx%4)
        low = QTableWidget(); self.table_fill(low, self.db.query("""SELECT material_id,item_description,quantity_available,minimum_stock_level FROM inventory_items WHERE is_active=1 AND quantity_available<=minimum_stock_level ORDER BY quantity_available LIMIT 8"""), [("Material ID","material_id"),("Item","item_description"),("Stock","quantity_available"),("Reorder Level","minimum_stock_level")])
        box = QGroupBox("Recent Low Stock Items"); bv=QVBoxLayout(box); bv.addWidget(low); self.dash_grid.addWidget(box,2,0,1,2)
        exp = QTableWidget(); self.table_fill(exp, self.db.query("""SELECT b.material_id,i.item_description,b.batch_no,b.quantity_available,b.expiry_date,CAST(julianday(b.expiry_date)-julianday('now') AS INTEGER) days_left
            FROM inventory_batches b JOIN inventory_items i ON i.id=b.item_id
            WHERE i.is_active=1 AND b.quantity_available>0 ORDER BY b.expiry_date LIMIT 8"""), [("Material ID","material_id"),("Item","item_description"),("Batch","batch_no"),("Qty Left","quantity_available"),("Expiry","expiry_date"),("Days","days_left")])
        box2 = QGroupBox("Expiry Items"); b2=QVBoxLayout(box2); b2.addWidget(exp); self.dash_grid.addWidget(box2,2,2,1,2)
        qa = QGroupBox("Quick Actions"); ql=QHBoxLayout(qa)
        for txt, idx in [("Stock Inward",1),("Stock Outward",2),("Generate Reports",5),("Inventory List",6)]:
            b=QPushButton(txt); b.clicked.connect(lambda _, i=idx:self.set_page(i)); ql.addWidget(b)
        self.dash_grid.addWidget(qa,3,0,1,4)

    def add_combo(self, form: QFormLayout, label: str, values: List[str]) -> QComboBox:
        c=QComboBox(); c.addItems(values); form.addRow(label,c); return c

    def set_combo_text(self, combo: QComboBox, text: str):
        idx = combo.findText(text or "")
        if idx >= 0:
            combo.setCurrentIndex(idx)

    def make_blank_spinbox(self, spin: QDoubleSpinBox, decimals: int = 2, prefix: str = ""):
        """Make a numeric field visually blank when its value is zero."""
        spin.setRange(0, 9999999)
        spin.setDecimals(decimals)
        spin.setPrefix(prefix)
        spin.setSpecialValueText("")
        spin.setValue(0)

    def clear_spinbox(self, spin: QDoubleSpinBox):
        spin.setValue(0)
        spin.clear()

    def make_blank_date(self, edit: QDateEdit):
        """Keep date visually blank, but open the calendar on the current date/year."""
        edit.setCalendarPopup(True)
        edit.setMinimumDate(date(2000, 1, 1))
        edit.setMaximumDate(date(2099, 12, 31))
        edit.setDisplayFormat("dd/MM/yyyy")
        edit.setDate(date.today())
        edit.setProperty("blank_date", True)
        edit.setDisplayFormat(" ")

        def mark_selected(*_):
            edit.setProperty("blank_date", False)
            edit.setDisplayFormat("dd/MM/yyyy")

        # Selecting a day from the popup must display the chosen date.
        if edit.calendarWidget():
            edit.calendarWidget().clicked.connect(mark_selected)
        edit.editingFinished.connect(mark_selected)

    def clear_date(self, edit: QDateEdit):
        """Clear display, while keeping the popup calendar positioned at today."""
        edit.blockSignals(True)
        edit.setDate(date.today())
        edit.setProperty("blank_date", True)
        edit.setDisplayFormat(" ")
        edit.blockSignals(False)

    def date_value_or_today(self, edit: QDateEdit) -> str:
        """Use today's date if a required date field is left blank."""
        if edit.property("blank_date") is True:
            return today_iso()
        return edit.date().toPython().isoformat()

    def date_value_or_none(self, edit: QDateEdit) -> Optional[str]:
        if edit.property("blank_date") is True:
            return None
        return edit.date().toPython().isoformat()

    def refresh_item_combos(self):
        rows = self.item_rows()
        targets = []
        for name in ["in_existing", "out_item", "adj_item", "tr_item"]:
            if hasattr(self, name):
                targets.append((name, getattr(self, name)))
        for name, combo in targets:
            old = combo.currentData()
            combo.blockSignals(True)
            combo.clear()
            if name == "in_existing":
                combo.addItem("-- New Material --", None)
            else:
                combo.addItem("", None)
            for r in rows:
                combo.addItem(f"{r['material_id']} - {r['item_description']} | Stock: {r['quantity_available']}", r['material_id'])
            if old:
                idx = combo.findData(old)
                if idx >= 0: combo.setCurrentIndex(idx)
            combo.blockSignals(False)

    def page_inward(self):
        w, v = self.base_page("Stock Inward")
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.NoFrame)
        holder = QWidget(); outer = QVBoxLayout(holder); outer.setContentsMargins(0,0,0,0); outer.setSpacing(14)

        box = QGroupBox("Add / Increase Stock")
        grid = QGridLayout(box); grid.setContentsMargins(18,22,18,18); grid.setHorizontalSpacing(18); grid.setVerticalSpacing(12)

        self.in_desc = QLineEdit(); self.in_desc.setPlaceholderText("Full material description")
        self.in_existing = QComboBox(); self.in_existing.addItem("-- New Material --", None)
        self.in_existing.setEditable(True)
        self.in_existing.setInsertPolicy(QComboBox.NoInsert)
        self.in_existing.lineEdit().setPlaceholderText("Type Material ID or item description...")
        for r in self.item_rows():
            self.in_existing.addItem(f"{r['material_id']} - {r['item_description']}", r['material_id'])
        self.in_existing.setMinimumWidth(520)
        self.in_existing.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.in_existing.currentIndexChanged.connect(self.load_existing_for_inward)
        self.in_existing.activated.connect(self.load_existing_for_inward)

        self.in_cat = QComboBox(); self.in_cat.addItem(""); self.in_cat.addItems(self.master_names("categories"))
        self.in_dept = QComboBox(); self.in_dept.addItem(""); self.in_dept.addItems(self.master_names("departments"))
        self.in_unit = QComboBox(); self.in_unit.addItem(""); self.in_unit.addItems(self.master_names("units"))
        self.in_type = QComboBox(); self.in_type.addItem(""); self.in_type.addItems(MATERIAL_TYPES)
        self.in_supplier = QComboBox(); self.in_supplier.addItem(""); self.in_supplier.addItems(self.master_names("suppliers"))
        self.in_location = QComboBox(); self.in_location.addItem(""); self.in_location.addItems(self.master_names("locations"))
        self.in_rack = QComboBox(); self.in_rack.addItem(""); self.in_rack.addItems(self.master_names("racks"))
        self.in_qty = QDoubleSpinBox(); self.make_blank_spinbox(self.in_qty, 2)
        self.in_min = QDoubleSpinBox(); self.make_blank_spinbox(self.in_min, 2)
        self.in_cost = QDoubleSpinBox(); self.make_blank_spinbox(self.in_cost, 3, "OMR ")
        self.in_date = QDateEdit(); self.make_blank_date(self.in_date)
        self.in_exp = QDateEdit(); self.make_blank_date(self.in_exp)
        self.in_exp_na = QCheckBox("Expiry Not Applicable"); self.in_exp_na.toggled.connect(self.toggle_inward_expiry_date)
        self.in_po = QLineEdit(); self.in_do = QLineEdit(); self.in_warranty = QTextEdit(); self.in_remarks = QTextEdit()
        self.in_warranty.setFixedHeight(70); self.in_remarks.setFixedHeight(70)

        fields = [
            ("Existing Material", self.in_existing), ("Item Description *", self.in_desc),
            ("Category *", self.in_cat), ("Department *", self.in_dept),
            ("Unit *", self.in_unit), ("Material Type *", self.in_type),
            ("Supplier", self.in_supplier), ("Location", self.in_location),
            ("Rack No. *", self.in_rack), ("Quantity *", self.in_qty),
            ("Minimum Stock Level *", self.in_min), ("Cost Per Unit *", self.in_cost),
            ("Procurement Date *", self.in_date), ("Expiry Date", self.in_exp),
            ("MR No.", self.in_po), ("DO No.", self.in_do),
            ("Warranty Details", self.in_warranty), ("Remarks", self.in_remarks),
        ]
        for i, (lab, wid) in enumerate(fields):
            row = i // 2; col = (i % 2) * 2
            label = QLabel(lab); label.setStyleSheet("font-weight:800;color:#263b5e;")
            grid.addWidget(label, row, col); grid.addWidget(wid, row, col+1)
        grid.addWidget(QLabel(""), 9, 0); grid.addWidget(self.in_exp_na, 9, 1)
        grid.setColumnStretch(1, 1); grid.setColumnStretch(3, 1)

        batch_box = QGroupBox("Expiry Batch Details - use when one inward quantity has multiple expiry dates")
        batch_layout = QVBoxLayout(batch_box)
        batch_top = QHBoxLayout()
        self.batch_qty = QDoubleSpinBox(); self.make_blank_spinbox(self.batch_qty, 2)
        self.batch_date = QDateEdit(); self.make_blank_date(self.batch_date)
        self.batch_add = QPushButton("Add Expiry Batch"); self.batch_add.setObjectName("softBtn"); self.batch_add.clicked.connect(self.add_inward_batch_row)
        self.batch_remove = QPushButton("Remove Selected Batch"); self.batch_remove.setObjectName("redBtn"); self.batch_remove.clicked.connect(self.remove_selected_inward_batch)
        batch_top.addWidget(QLabel("Batch Qty")); batch_top.addWidget(self.batch_qty)
        batch_top.addWidget(QLabel("Expiry Date")); batch_top.addWidget(self.batch_date)
        batch_top.addWidget(self.batch_add); batch_top.addWidget(self.batch_remove); batch_top.addStretch()
        self.batch_table = QTableWidget(); self.batch_table.setColumnCount(2); self.batch_table.setHorizontalHeaderLabels(["Quantity", "Expiry Date"]); self.batch_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); self.batch_table.setFixedHeight(150)
        batch_note = QLabel("Example: for 10 paint drums, add 4 with one expiry date and 6 with another. Batch total must equal inward quantity."); batch_note.setObjectName("small")
        batch_layout.addWidget(batch_note); batch_layout.addLayout(batch_top); batch_layout.addWidget(self.batch_table)
        outer.addWidget(box); outer.addWidget(batch_box)
        self.in_batch_box = batch_box

        actions = QHBoxLayout(); actions.addStretch()
        clear = QPushButton("Clear Form"); clear.setObjectName("softBtn"); clear.clicked.connect(self.clear_inward_form)
        save = QPushButton("Save Stock Inward"); save.setObjectName("greenBtn"); save.clicked.connect(self.save_inward)
        actions.addWidget(clear); actions.addWidget(save)
        outer.addLayout(actions)

        recent = QGroupBox("Recent Stock Inward Entries")
        rv = QVBoxLayout(recent); self.in_table = QTableWidget(); rv.addWidget(self.in_table)
        outer.addWidget(recent)
        scroll.setWidget(holder); v.addWidget(scroll, 1)
        self.refresh_inward_table(); self.clear_inward_form(); return w

    def add_inward_batch_row(self):
        if not hasattr(self, "batch_table"):
            return
        qty = float(self.batch_qty.value())
        exp = self.date_value_or_none(self.batch_date)
        if qty <= 0:
            QMessageBox.warning(self, "Batch Validation", "Enter a batch quantity greater than zero.")
            return
        if not exp:
            QMessageBox.warning(self, "Batch Validation", "Select the expiry date for this batch.")
            return
        procurement_iso = self.date_value_or_today(self.in_date) if hasattr(self, "in_date") else today_iso()
        min_expiry_iso = add_months(parse_iso_date(procurement_iso), 6).isoformat()
        if exp < min_expiry_iso:
            QMessageBox.warning(self, "Expiry Date Alert", f"Expiry date must be at least 6 months after procurement date.\n\nProcurement Date: {procurement_iso}\nMinimum Allowed Expiry: {min_expiry_iso}\nSelected Expiry: {exp}")
            return
        r = self.batch_table.rowCount()
        self.batch_table.insertRow(r)
        self.batch_table.setItem(r, 0, QTableWidgetItem(f"{qty:g}"))
        self.batch_table.setItem(r, 1, QTableWidgetItem(exp))
        self.clear_spinbox(self.batch_qty)
        self.clear_date(self.batch_date)

    def remove_selected_inward_batch(self):
        if hasattr(self, "batch_table") and self.batch_table.currentRow() >= 0:
            self.batch_table.removeRow(self.batch_table.currentRow())

    def inward_batch_entries(self) -> List[Tuple[float, str]]:
        entries = []
        if not hasattr(self, "batch_table"):
            return entries
        for r in range(self.batch_table.rowCount()):
            qitem = self.batch_table.item(r, 0)
            ditem = self.batch_table.item(r, 1)
            try:
                qty = float(qitem.text()) if qitem else 0
            except Exception:
                qty = 0
            exp = ditem.text().strip() if ditem else ""
            if qty > 0 and exp:
                entries.append((qty, exp))
        return entries

    def create_expiry_batch(self, inward_id: int, item_id: int, material_id: str, qty: float, expiry_date: str):
        batch_no = self.db.next_code("BAT", "inventory_batches", "batch_no")
        self.db.execute("""INSERT INTO inventory_batches(batch_no,inward_id,item_id,material_id,quantity_received,quantity_available,expiry_date)
            VALUES(%s,%s,%s,%s,%s,%s,%s)""", (batch_no, inward_id, item_id, material_id, float(qty), float(qty), expiry_date))

    def issue_from_expiry_batches(self, outward_id: int, item_id: int, material_id: str, qty: float):
        batches = self.db.query("""SELECT id,batch_no,quantity_available,expiry_date FROM inventory_batches
            WHERE item_id=%s AND quantity_available>0 ORDER BY expiry_date ASC, id ASC""", (item_id,))
        if not batches:
            return
        available = sum(float(b.get('quantity_available') or 0) for b in batches)
        if qty > available + 1e-9:
            raise ValueError(f"Batch-wise expiry stock available is only {available:g}. Cannot issue {qty:g}.")
        remaining = float(qty)
        for b in batches:
            if remaining <= 0:
                break
            take = min(remaining, float(b.get('quantity_available') or 0))
            if take <= 0:
                continue
            self.db.execute("UPDATE inventory_batches SET quantity_available=quantity_available-%s, updated_at=CURRENT_TIMESTAMP WHERE id=%s", (take, b['id']))
            self.db.execute("INSERT INTO stock_outward_batches(outward_id,batch_id,material_id,quantity_issued) VALUES(%s,%s,%s,%s)", (outward_id, b['id'], material_id, take))
            remaining -= take

    def toggle_inward_expiry_date(self, checked: bool):
        """Disable and visually grey the expiry date/batch controls when expiry is not applicable."""
        if not hasattr(self, "in_exp"):
            return
        self.in_exp.setEnabled(not checked)
        for obj_name in ["batch_qty", "batch_date", "batch_add", "batch_remove", "batch_table"]:
            if hasattr(self, obj_name):
                getattr(self, obj_name).setEnabled(not checked)
        if checked:
            self.clear_date(self.in_exp)
            if hasattr(self, "batch_table"):
                self.batch_table.setRowCount(0)
            self.in_exp.setStyleSheet("background:#e5e7eb;color:#6b7280;border:1px solid #cbd5e1;border-radius:8px;padding:8px;")
            if hasattr(self, "in_batch_box"):
                self.in_batch_box.setStyleSheet("background:#f3f4f6;color:#6b7280;border:1px solid #d1d5db;border-radius:14px;")
        else:
            self.in_exp.setStyleSheet("")
            if hasattr(self, "in_batch_box"):
                self.in_batch_box.setStyleSheet("")

    def selected_material_id(self, combo: QComboBox) -> Optional[str]:
        data = combo.currentData()
        if data:
            return str(data)
        text = combo.currentText().strip()
        if not text or text.startswith("--"):
            return None
        token = text.split(" - ", 1)[0].strip()
        row = self.db.one("SELECT material_id FROM inventory_items WHERE material_id=%s AND is_active=1", (token,))
        if row:
            return row["material_id"]
        like = f"%{text}%"
        row = self.db.one("SELECT material_id FROM inventory_items WHERE is_active=1 AND (material_id LIKE %s OR item_description LIKE %s) ORDER BY material_id LIMIT 1", (like, like))
        return row["material_id"] if row else None

    def load_existing_for_inward(self):
        mat=self.selected_material_id(self.in_existing)
        if not mat:
            self.clear_inward_form(keep_selection=True)
            return
        r=self.db.one("""SELECT i.*, c.name category, d.name department, u.name unit, s.name supplier, l.name location, r.name rack
            FROM inventory_items i
            LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id
            LEFT JOIN units u ON u.id=i.unit_id LEFT JOIN suppliers s ON s.id=i.supplier_id
            LEFT JOIN locations l ON l.id=i.location_id LEFT JOIN racks r ON r.id=i.rack_id
            WHERE i.material_id=%s""",(mat,))
        if r:
            self.in_desc.setText(r.get('item_description') or '')
            self.set_combo_text(self.in_cat, r.get('category') or '')
            self.set_combo_text(self.in_dept, r.get('department') or '')
            self.set_combo_text(self.in_unit, r.get('unit') or '')
            self.set_combo_text(self.in_type, r.get('material_type') or '')
            self.set_combo_text(self.in_rack, r.get('rack') or '')
            self.in_min.setValue(float(r.get('minimum_stock_level') or 0))
            self.clear_spinbox(self.in_qty)
            self.clear_spinbox(self.in_cost)
            self.in_supplier.setCurrentIndex(0 if self.in_supplier.count() else -1)
            self.in_location.setCurrentIndex(0 if self.in_location.count() else -1)
            self.in_po.clear(); self.in_do.clear(); self.in_warranty.clear(); self.in_remarks.clear()
            self.in_exp_na.setChecked(True); self.toggle_inward_expiry_date(True)

    def clear_inward_form(self, keep_selection: bool = False):
        if not hasattr(self, 'in_desc'):
            return
        if not keep_selection and hasattr(self, 'in_existing'):
            self.in_existing.blockSignals(True)
            self.in_existing.setCurrentIndex(0)
            self.in_existing.blockSignals(False)
        self.in_desc.clear()
        self.clear_spinbox(self.in_qty)
        self.clear_spinbox(self.in_min)
        self.clear_spinbox(self.in_cost)
        for c in [self.in_cat,self.in_dept,self.in_unit,self.in_type,self.in_supplier,self.in_location,self.in_rack]:
            if c.count(): c.setCurrentIndex(0)
        self.clear_date(self.in_date)
        self.clear_date(self.in_exp)
        self.in_exp_na.setChecked(False)
        self.toggle_inward_expiry_date(False)
        self.in_po.clear(); self.in_do.clear(); self.in_warranty.clear(); self.in_remarks.clear()
        if hasattr(self, "batch_table"):
            self.batch_table.setRowCount(0)
        if hasattr(self, "batch_qty"):
            self.clear_spinbox(self.batch_qty)
        if hasattr(self, "batch_date"):
            self.clear_date(self.batch_date)

    def save_inward(self):
        try:
            if not self.in_desc.text().strip():
                QMessageBox.warning(self, "Validation", "Item Description is required.")
                return
            qty = float(self.in_qty.value())
            cost = float(self.in_cost.value())
            if qty <= 0:
                QMessageBox.warning(self, "Validation", "Quantity must be greater than zero.")
                return
            total = qty * cost
            mat = self.selected_material_id(self.in_existing)
            procurement_iso = self.date_value_or_today(self.in_date)
            min_expiry_iso = add_months(parse_iso_date(procurement_iso), 6).isoformat()
            expiry_not_applicable = self.in_exp_na.isChecked()
            batch_entries = []
            exp = None
            if not expiry_not_applicable:
                batch_entries = self.inward_batch_entries()
                # Backward-compatible shortcut: if user entered only one expiry date, create one batch for total qty.
                single_exp = self.date_value_or_none(self.in_exp)
                if not batch_entries and single_exp:
                    batch_entries = [(qty, single_exp)]
                if not batch_entries:
                    QMessageBox.warning(self, "Expiry Required", "Expiry is applicable. Add at least one expiry batch or select one expiry date.")
                    return
                invalid_expiries = [d for _, d in batch_entries if d < min_expiry_iso]
                if invalid_expiries:
                    QMessageBox.warning(self, "Expiry Date Alert", f"Expiry date must be at least 6 months after procurement date.\n\nProcurement Date: {procurement_iso}\nMinimum Allowed Expiry: {min_expiry_iso}\nInvalid Expiry: {invalid_expiries[0]}")
                    return
                batch_total = sum(q for q, _ in batch_entries)
                if abs(batch_total - qty) > 0.001:
                    QMessageBox.warning(self, "Batch Quantity Mismatch", f"Total batch quantity must equal inward quantity.\n\nInward quantity: {qty:g}\nBatch total: {batch_total:g}")
                    return
                exp = min(d for _, d in batch_entries)

            if mat:
                item = self.db.one("SELECT id,quantity_available FROM inventory_items WHERE material_id=%s", (mat,))
                if not item:
                    QMessageBox.warning(self, "Validation", "Selected material was not found in inventory.")
                    return
                item_id = item['id']
                self.db.execute(
                    """UPDATE inventory_items
                       SET quantity_available = quantity_available + %s,
                           cost_per_unit = CASE WHEN %s > 0 THEN %s ELSE cost_per_unit END,
                           minimum_stock_level = %s,
                           expiry_date = COALESCE(%s, expiry_date),
                           remarks = %s,
                           updated_at = CURRENT_TIMESTAMP
                       WHERE id = %s""",
                    (qty, cost, cost, float(self.in_min.value()), exp, self.in_remarks.toPlainText(), item_id)
                )
            else:
                mat = self.db.next_code("MAT", "inventory_items", "material_id")
                item_id = self.db.execute(
                    """INSERT INTO inventory_items(
                        material_id,item_description,category_id,department_id,quantity_available,unit_id,
                        minimum_stock_level,material_type,supplier_id,cost_per_unit,procurement_date,
                        purchase_order_no,delivery_order_no,warranty_details,expiry_date,rack_id,location_id,remarks)
                       VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                    (mat, self.in_desc.text().strip(), self.db.id_by_name('categories', self.in_cat.currentText()),
                     self.db.id_by_name('departments', self.in_dept.currentText()), qty,
                     self.db.id_by_name('units', self.in_unit.currentText()), float(self.in_min.value()),
                     self.in_type.currentText(), self.db.id_by_name('suppliers', self.in_supplier.currentText()),
                     cost, self.date_value_or_today(self.in_date), self.in_po.text(), self.in_do.text(),
                     self.in_warranty.toPlainText(), exp, self.db.id_by_name('racks', self.in_rack.currentText()),
                     self.db.id_by_name('locations', self.in_location.currentText()), self.in_remarks.toPlainText())
                )

            inward = self.db.next_code("IN", "stock_inward", "inward_no")
            inward_id = self.db.execute(
                """INSERT INTO stock_inward(
                    inward_no,material_id,item_id,quantity,unit_cost,total_cost,procurement_date,
                    supplier_id,po_no,do_no,remarks,created_by)
                   VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (inward, mat, item_id, qty, cost, total, self.date_value_or_today(self.in_date),
                 self.db.id_by_name('suppliers', self.in_supplier.currentText()), self.in_po.text(),
                 self.in_do.text(), self.in_remarks.toPlainText(), self.user['id'])
            )
            for bqty, bexp in batch_entries:
                self.create_expiry_batch(inward_id, item_id, mat, bqty, bexp)
            self.db.log(self.user['id'], "ADD", "Stock Inward", mat, f"Qty {qty}; batches {len(batch_entries)}")
            QMessageBox.information(self, "Saved", f"Stock inward saved successfully.\n\nMaterial ID: {mat}\nQuantity added: {qty:g}\nExpiry batches: {len(batch_entries)}\nTotal Value: {money(total)}")
            self.refresh_inward_table()
            self.refresh_item_combos()
            self.clear_inward_form()
            if hasattr(self, 'inv_table'):
                self.refresh_inventory()
            self.refresh_dashboard()
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", f"Stock inward was not saved.\n\nReason:\n{e}")

    def refresh_inward_table(self):
        rows=self.db.query("SELECT inward_no,material_id,quantity,unit_cost,total_cost,procurement_date,created_at FROM stock_inward ORDER BY id DESC LIMIT 100")
        self.table_fill(self.in_table, rows, [("Inward No","inward_no"),("Material ID","material_id"),("Qty","quantity"),("Unit Cost (OMR)","unit_cost"),("Total (OMR)","total_cost"),("Procurement Date","procurement_date"),("Created","created_at")])

    def page_outward(self):
        w,v=self.base_page("Stock Outward")
        box=QGroupBox("Issue Stock")
        grid=QGridLayout(box); grid.setContentsMargins(18,22,18,18); grid.setHorizontalSpacing(18); grid.setVerticalSpacing(12)
        v.addWidget(box)
        self.out_item=QComboBox(); [self.out_item.addItem(f"{r['material_id']} - {r['item_description']} | Stock: {r['quantity_available']}", r['material_id']) for r in self.item_rows()]
        self.out_item.setMinimumWidth(360)
        self.out_item.currentIndexChanged.connect(self.load_existing_for_outward)
        self.out_item.setCurrentIndex(-1)
        self.out_siv=QLineEdit(); self.out_siv.setPlaceholderText("Enter SIV # manually")
        self.out_date=QDateEdit(); self.make_blank_date(self.out_date)
        self.out_qty=QDoubleSpinBox(); self.make_blank_spinbox(self.out_qty, 2)
        self.out_dept=QComboBox(); self.out_dept.addItem(""); self.out_dept.addItems(self.master_names("departments"))
        self.out_to=QLineEdit(); self.out_by=QLineEdit(); self.out_remarks=QTextEdit(); self.out_remarks.setFixedHeight(80)
        fields=[("Material ID *", self.out_item), ("SIV # *", self.out_siv), ("Issue Date *", self.out_date), ("Quantity Issued *", self.out_qty), ("Department *", self.out_dept), ("Issued To *", self.out_to), ("Issued By *", self.out_by), ("Remarks", self.out_remarks)]
        for i,(lab,wid) in enumerate(fields):
            row=i//2; col=(i%2)*2
            label=QLabel(lab); label.setStyleSheet("font-weight:800;color:#263b5e;")
            grid.addWidget(label,row,col); grid.addWidget(wid,row,col+1)
        grid.setColumnStretch(1,1); grid.setColumnStretch(3,1)
        actions=QHBoxLayout(); actions.addStretch()
        clear=QPushButton("Clear Form"); clear.setObjectName("softBtn"); clear.clicked.connect(self.clear_outward_form)
        b=QPushButton("Save Stock Outward"); b.setObjectName("greenBtn"); b.clicked.connect(self.save_outward)
        actions.addWidget(clear); actions.addWidget(b); v.addLayout(actions)
        self.out_table=QTableWidget(); v.addWidget(self.out_table); self.refresh_outward_table(); self.clear_outward_form(); return w

    def load_existing_for_outward(self):
        mat=self.out_item.currentData()
        if not mat: return
        r=self.db.one("""SELECT d.name department FROM inventory_items i LEFT JOIN departments d ON d.id=i.department_id WHERE i.material_id=%s""",(mat,))
        if r: self.set_combo_text(self.out_dept, r.get('department') or '')
        self.out_siv.clear(); self.clear_spinbox(self.out_qty); self.out_to.clear(); self.out_by.clear(); self.out_remarks.clear()

    def clear_outward_form(self):
        if not hasattr(self, 'out_siv'):
            return
        if hasattr(self,'out_item') and self.out_item.count():
            self.out_item.blockSignals(True)
            self.out_item.setCurrentIndex(-1)
            self.out_item.blockSignals(False)
        self.out_siv.clear()
        self.clear_date(self.out_date)
        self.clear_spinbox(self.out_qty)
        if self.out_dept.count(): self.out_dept.setCurrentIndex(0)
        self.out_to.clear(); self.out_by.clear(); self.out_remarks.clear()

    def save_outward(self):
        try:
            mat=self.out_item.currentData(); item=self.db.one("SELECT id,quantity_available FROM inventory_items WHERE material_id=%s",(mat,)); qty=float(self.out_qty.value())
            siv=self.out_siv.text().strip()
            if not mat:
                QMessageBox.warning(self,"Validation","Select Material ID."); return
            if not siv:
                QMessageBox.warning(self,"Validation","SIV # is required and must be entered manually."); return
            if not item or qty <= 0 or qty > float(item['quantity_available'] or 0):
                QMessageBox.warning(self,"Blocked","Quantity issued cannot exceed available stock."); return
            if not self.out_to.text().strip() or not self.out_by.text().strip():
                QMessageBox.warning(self,"Validation","Issued To and Issued By are required."); return
            # If expiry batches exist, issue by FEFO: First Expiry, First Out.
            batch_total_row = self.db.one("SELECT COALESCE(SUM(quantity_available),0) q FROM inventory_batches WHERE item_id=%s AND quantity_available>0", (item['id'],))
            batch_total = float((batch_total_row or {}).get('q') or 0)
            if batch_total > 0 and qty > batch_total + 1e-9:
                QMessageBox.warning(self,"Blocked",f"Batch-wise expiry stock available is only {batch_total:g}. Cannot issue {qty:g}."); return
            self.db.execute("UPDATE inventory_items SET quantity_available=quantity_available-%s, updated_at=CURRENT_TIMESTAMP WHERE id=%s",(qty,item['id']))
            outward_id = self.db.execute("INSERT INTO stock_outward(siv_no,issue_date,material_id,item_id,quantity_issued,department_id,issued_to,issued_by,purpose,remarks,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(siv,self.date_value_or_today(self.out_date),mat,item['id'],qty,self.db.id_by_name('departments',self.out_dept.currentText()),self.out_to.text(),self.out_by.text(),"",self.out_remarks.toPlainText(),self.user['id']))
            if batch_total > 0:
                self.issue_from_expiry_batches(outward_id, item['id'], mat, qty)
            self.db.log(self.user['id'],"ISSUE","Stock Outward",mat,f"SIV {siv}, Qty {qty}")
            QMessageBox.information(self,"Stock Updated",f"Stock outward saved successfully.\nSIV #: {siv}\nQuantity issued: {qty:g}\nExpiry method: {'FEFO batch deduction' if batch_total > 0 else 'Non-expiry stock deduction'}")
            self.refresh_outward_table(); self.refresh_item_combos(); self.clear_outward_form()
            if hasattr(self, "inv_table"): self.refresh_inventory()
            self.refresh_dashboard()
        except Exception as e:
            QMessageBox.critical(self,"Save Failed",f"Stock outward was not saved.\n\nReason:\n{e}")

    def refresh_outward_table(self):
        rows=self.db.query("SELECT siv_no,issue_date,material_id,quantity_issued,issued_to,issued_by,remarks FROM stock_outward ORDER BY id DESC LIMIT 100")
        self.table_fill(self.out_table,rows,[("SIV #","siv_no"),("Date","issue_date"),("Material ID","material_id"),("Qty","quantity_issued"),("Issued To","issued_to"),("Issued By","issued_by"),("Remarks","remarks")])

    def page_adjustment(self):
        w,v=self.base_page("Stock Adjustment")
        box=QGroupBox("Manual Stock Correction")
        form=QFormLayout(box); form.setContentsMargins(18,22,18,18); form.setSpacing(12)
        v.addWidget(box)
        self.adj_item=QComboBox(); self.adj_item.setMinimumWidth(420)
        self.adj_item.addItem("", None)
        for r in self.item_rows():
            self.adj_item.addItem(f"{r['material_id']} - {r['item_description']} | Stock: {r['quantity_available']}", r['material_id'])
        self.adj_date=QDateEdit(); self.make_blank_date(self.adj_date)
        self.adj_type=QComboBox(); self.adj_type.addItem(""); self.adj_type.addItems(ADJUSTMENT_TYPES)
        self.adj_qty=QDoubleSpinBox(); self.make_blank_spinbox(self.adj_qty, 2)
        self.adj_reason=QLineEdit(); self.adj_approved=QLineEdit(); self.adj_remarks=QTextEdit(); self.adj_remarks.setFixedHeight(80)
        for lab,wid in [("Material ID *",self.adj_item),("Date *",self.adj_date),("Adjustment Type *",self.adj_type),("Adjustment Quantity *",self.adj_qty),("Reason *",self.adj_reason),("Approved By",self.adj_approved),("Remarks",self.adj_remarks)]: form.addRow(lab,wid)
        actions=QHBoxLayout(); actions.addStretch()
        clear=QPushButton("Clear Form"); clear.setObjectName("softBtn"); clear.clicked.connect(self.clear_adjustment_form)
        b=QPushButton("Save Adjustment"); b.setObjectName("greenBtn"); b.clicked.connect(self.save_adjustment)
        actions.addWidget(clear); actions.addWidget(b); v.addLayout(actions)
        self.adj_table=QTableWidget(); v.addWidget(self.adj_table); self.refresh_adjustment_table(); self.clear_adjustment_form(); return w

    def clear_adjustment_form(self):
        if not hasattr(self, 'adj_item'):
            return
        self.adj_item.blockSignals(True); self.adj_item.setCurrentIndex(0); self.adj_item.blockSignals(False)
        self.clear_date(self.adj_date)
        if self.adj_type.count(): self.adj_type.setCurrentIndex(0)
        self.clear_spinbox(self.adj_qty)
        self.adj_reason.clear(); self.adj_approved.clear(); self.adj_remarks.clear()

    def save_adjustment(self):
        try:
            mat=self.adj_item.currentData()
            if not mat:
                QMessageBox.warning(self,"Validation","Select a Material ID before saving stock adjustment."); return
            item=self.db.one("SELECT id,quantity_available FROM inventory_items WHERE material_id=%s",(mat,))
            if not item:
                QMessageBox.warning(self,"Validation","Selected material was not found."); return
            qty=Decimal(str(self.adj_qty.value()))
            if qty <= 0:
                QMessageBox.warning(self,"Validation","Adjustment quantity must be greater than zero."); return
            if not self.adj_type.currentText().strip():
                QMessageBox.warning(self,"Validation","Adjustment type is required."); return
            if not self.adj_reason.text().strip(): QMessageBox.warning(self,"Validation","Reason is required."); return
            if self.adj_type.currentText()=="Decrease" and qty>Decimal(str(item['quantity_available'])): QMessageBox.warning(self,"Blocked","Decrease cannot exceed current stock."); return
            sign = 1 if self.adj_type.currentText()=="Increase" else -1; no=self.db.next_code("ADJ","stock_adjustments","adjustment_no")
            self.db.execute("UPDATE inventory_items SET quantity_available=quantity_available+%s, updated_at=CURRENT_TIMESTAMP WHERE id=%s",(qty*sign,item['id']))
            self.db.execute("INSERT INTO stock_adjustments(adjustment_no,adjustment_date,material_id,item_id,current_stock,adjustment_type,adjustment_quantity,reason,approved_by,remarks,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(no,self.date_value_or_today(self.adj_date),mat,item['id'],item['quantity_available'],self.adj_type.currentText(),qty,self.adj_reason.text(),self.adj_approved.text(),self.adj_remarks.toPlainText(),self.user['id']))
            self.db.log(self.user['id'],"ADJUST", "Stock Adjustment", mat, f"{self.adj_type.currentText()} {qty}")
            QMessageBox.information(self,"Saved",f"Adjustment saved: {no}")
            self.refresh_adjustment_table(); self.refresh_item_combos(); self.clear_adjustment_form()
            if hasattr(self, 'inv_table'): self.refresh_inventory()
            self.refresh_dashboard()
        except Exception as e:
            QMessageBox.critical(self,"Save Failed",f"Stock adjustment was not saved.\n\nReason:\n{e}")

    def refresh_adjustment_table(self):
        rows=self.db.query("SELECT adjustment_no,adjustment_date,material_id,current_stock,adjustment_type,adjustment_quantity,reason,approved_by FROM stock_adjustments ORDER BY id DESC LIMIT 100")
        self.table_fill(self.adj_table,rows,[("Adjustment No","adjustment_no"),("Date","adjustment_date"),("Material ID","material_id"),("Old Stock","current_stock"),("Type","adjustment_type"),("Qty","adjustment_quantity"),("Reason","reason"),("Approved By","approved_by")])

    def page_transfer(self):
        w,v=self.base_page("Transfers")
        box=QGroupBox("Transfer Stock")
        form=QFormLayout(box); form.setContentsMargins(18,22,18,18); form.setSpacing(12)
        v.addWidget(box)
        self.tr_item=QComboBox(); self.tr_item.setMinimumWidth(420); self.tr_item.addItem("", None)
        for r in self.item_rows():
            self.tr_item.addItem(f"{r['material_id']} - {r['item_description']} | Stock: {r['quantity_available']}", r['material_id'])
        self.tr_date=QDateEdit(); self.make_blank_date(self.tr_date)
        self.tr_qty=QDoubleSpinBox(); self.make_blank_spinbox(self.tr_qty, 2)
        self.tr_from_dept=QComboBox(); self.tr_from_dept.addItem(""); self.tr_from_dept.addItems(self.master_names("departments"))
        self.tr_to_dept=QComboBox(); self.tr_to_dept.addItem(""); self.tr_to_dept.addItems(self.master_names("departments"))
        self.tr_from_loc=QComboBox(); self.tr_from_loc.addItem(""); self.tr_from_loc.addItems(self.master_names("locations"))
        self.tr_to_loc=QComboBox(); self.tr_to_loc.addItem(""); self.tr_to_loc.addItems(self.master_names("locations"))
        self.tr_from_rack=QComboBox(); self.tr_from_rack.addItem(""); self.tr_from_rack.addItems(self.master_names("racks"))
        self.tr_to_rack=QComboBox(); self.tr_to_rack.addItem(""); self.tr_to_rack.addItems(self.master_names("racks"))
        self.tr_req=QLineEdit(); self.tr_app=QLineEdit(); self.tr_rem=QTextEdit(); self.tr_rem.setFixedHeight(80)
        for lab,wid in [("Material ID *",self.tr_item),("Transfer Date *",self.tr_date),("Quantity *",self.tr_qty),("From Department *",self.tr_from_dept),("To Department *",self.tr_to_dept),("From Location",self.tr_from_loc),("To Location",self.tr_to_loc),("From Rack",self.tr_from_rack),("To Rack",self.tr_to_rack),("Requested By",self.tr_req),("Approved By",self.tr_app),("Remarks",self.tr_rem)]: form.addRow(lab,wid)
        actions=QHBoxLayout(); actions.addStretch()
        clear=QPushButton("Clear Form"); clear.setObjectName("softBtn"); clear.clicked.connect(self.clear_transfer_form)
        b=QPushButton("Save Transfer"); b.setObjectName("greenBtn"); b.clicked.connect(self.save_transfer)
        actions.addWidget(clear); actions.addWidget(b); v.addLayout(actions)
        self.tr_table=QTableWidget(); v.addWidget(self.tr_table); self.refresh_transfer_table(); self.clear_transfer_form(); return w

    def clear_transfer_form(self):
        if not hasattr(self, 'tr_item'):
            return
        for c in [self.tr_item,self.tr_from_dept,self.tr_to_dept,self.tr_from_loc,self.tr_to_loc,self.tr_from_rack,self.tr_to_rack]:
            c.blockSignals(True); c.setCurrentIndex(0); c.blockSignals(False)
        self.clear_date(self.tr_date); self.clear_spinbox(self.tr_qty)
        self.tr_req.clear(); self.tr_app.clear(); self.tr_rem.clear()

    def save_transfer(self):
        try:
            mat=self.tr_item.currentData()
            if not mat:
                QMessageBox.warning(self,"Validation","Select a Material ID before saving transfer."); return
            item=self.db.one("SELECT id,quantity_available FROM inventory_items WHERE material_id=%s",(mat,))
            if not item:
                QMessageBox.warning(self,"Validation","Selected material was not found."); return
            qty=Decimal(str(self.tr_qty.value()))
            if qty <= 0:
                QMessageBox.warning(self,"Validation","Transfer quantity must be greater than zero."); return
            if qty>Decimal(str(item['quantity_available'])): QMessageBox.warning(self,"Blocked","Transfer quantity cannot exceed available stock."); return
            if not self.tr_from_dept.currentText().strip() or not self.tr_to_dept.currentText().strip():
                QMessageBox.warning(self,"Validation","From Department and To Department are required."); return
            no=self.db.next_code("TR","stock_transfers","transfer_no")
            self.db.execute("UPDATE inventory_items SET department_id=%s, location_id=%s, rack_id=%s, updated_at=CURRENT_TIMESTAMP WHERE id=%s",(self.db.id_by_name('departments',self.tr_to_dept.currentText()),self.db.id_by_name('locations',self.tr_to_loc.currentText()),self.db.id_by_name('racks',self.tr_to_rack.currentText()),item['id']))
            self.db.execute("INSERT INTO stock_transfers(transfer_no,transfer_date,material_id,item_id,quantity_transferred,from_department_id,to_department_id,from_location_id,to_location_id,from_rack_id,to_rack_id,requested_by,approved_by,remarks,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(no,self.date_value_or_today(self.tr_date),mat,item['id'],qty,self.db.id_by_name('departments',self.tr_from_dept.currentText()),self.db.id_by_name('departments',self.tr_to_dept.currentText()),self.db.id_by_name('locations',self.tr_from_loc.currentText()),self.db.id_by_name('locations',self.tr_to_loc.currentText()),self.db.id_by_name('racks',self.tr_from_rack.currentText()),self.db.id_by_name('racks',self.tr_to_rack.currentText()),self.tr_req.text(),self.tr_app.text(),self.tr_rem.toPlainText(),self.user['id']))
            self.db.log(self.user['id'],"TRANSFER","Transfers",mat,f"Qty {qty}")
            QMessageBox.information(self,"Saved",f"Transfer saved: {no}")
            self.refresh_transfer_table(); self.refresh_item_combos(); self.clear_transfer_form()
            if hasattr(self, 'inv_table'): self.refresh_inventory()
            self.refresh_dashboard()
        except Exception as e:
            QMessageBox.critical(self,"Save Failed",f"Stock transfer was not saved.\n\nReason:\n{e}")

    def refresh_transfer_table(self):
        rows=self.db.query("SELECT transfer_no,transfer_date,material_id,quantity_transferred,requested_by,approved_by FROM stock_transfers ORDER BY id DESC LIMIT 100")
        self.table_fill(self.tr_table,rows,[("Transfer No","transfer_no"),("Date","transfer_date"),("Material ID","material_id"),("Qty","quantity_transferred"),("Requested By","requested_by"),("Approved By","approved_by")])

    def page_inventory(self):
        w,v=self.base_page("Inventory Management")
        actions = QFrame(); actions.setObjectName("card")
        av = QVBoxLayout(actions); av.setContentsMargins(16,14,16,14)

        head = QHBoxLayout()
        title = QLabel("Store Materials Register"); title.setStyleSheet("font-size:17px;font-weight:900;")
        hint = QLabel("Search live inventory, upload corrected inventory data, edit/delete selected materials, and export store data."); hint.setObjectName("small")
        th = QVBoxLayout(); th.addWidget(title); th.addWidget(hint); head.addLayout(th); head.addStretch()

        self.inv_search=QLineEdit(); self.inv_search.setPlaceholderText("Search Material ID, item, category, department, supplier, rack, location...")
        self.inv_search.setMinimumWidth(520)
        self.inv_search.returnPressed.connect(self.refresh_inventory); head.addWidget(self.inv_search, 1)
        b=QPushButton("Search"); b.clicked.connect(self.refresh_inventory); head.addWidget(b)
        av.addLayout(head)

        btns = QHBoxLayout(); btns.setSpacing(8)
        down=QPushButton("Download Inventory"); down.clicked.connect(self.export_complete_inventory)
        imp=QPushButton("Upload Inventory"); imp.setObjectName("greenBtn"); imp.clicked.connect(self.import_excel)
        exp=QPushButton("Export Current View"); exp.setObjectName("softBtn"); exp.clicked.connect(lambda:self.export_table(self.inv_table,"inventory_current_view"))
        edit=QPushButton("Edit Selected Item"); edit.setObjectName("orangeBtn"); edit.clicked.connect(self.edit_selected_inventory)
        delete=QPushButton("Delete Selected Item"); delete.setObjectName("redBtn"); delete.clicked.connect(self.delete_selected_inventory)
        for x in [down, imp, exp, edit, delete]:
            x.setMinimumWidth(170)
            btns.addWidget(x)
        btns.addStretch(); av.addLayout(btns)
        v.addWidget(actions)
        self.inv_table=QTableWidget(); self.inv_table.setAlternatingRowColors(True); v.addWidget(self.inv_table, 1); self.refresh_inventory(); return w

    def refresh_inventory(self):
        self.inv_rows = self.item_rows(self.inv_search.text().strip() if hasattr(self,'inv_search') else "")
        self.table_fill(self.inv_table,self.inv_rows,[("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min. Level","minimum_stock_level"),("Type","material_type"),("Rack No.","rack")])
        widths = [150, 460, 160, 180, 110, 110, 130, 210, 130]
        for i, w in enumerate(widths):
            self.inv_table.setColumnWidth(i, w)
        self.inv_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.inv_table.setSelectionMode(QTableWidget.SingleSelection)
        self.inv_table.resizeRowsToContents()

    def delete_selected_inventory(self):
        row = self.inv_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Selection", "Search/select one item first.")
            return
        mat_item = self.inv_table.item(row, 0)
        desc_item = self.inv_table.item(row, 1)
        if not mat_item:
            QMessageBox.warning(self, "Selection", "Invalid selection.")
            return
        mat = mat_item.text()
        desc = desc_item.text() if desc_item else ""
        confirm = QMessageBox.question(
            self,
            "Confirm Delete",
            f"Delete/deactivate this inventory item?\n\nMaterial ID: {mat}\nDescription: {desc}\n\nThis will remove it from active inventory reports but keep old transaction records for reference.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if confirm != QMessageBox.Yes:
            return
        try:
            self.db.execute("UPDATE inventory_items SET is_active=0, updated_at=CURRENT_TIMESTAMP WHERE material_id=%s", (mat,))
            self.db.log(self.user['id'], "DELETE", "Inventory", mat, "Inventory item deactivated from Inventory screen")
            QMessageBox.information(self, "Deleted", f"Selected item deleted/deactivated successfully.\n\nMaterial ID: {mat}")
            self.refresh_inventory()
            self.refresh_item_combos()
            self.refresh_dashboard()
        except Exception as e:
            QMessageBox.critical(self, "Delete Failed", f"The selected item could not be deleted.\n\nReason:\n{e}")

    def edit_selected_inventory(self):
        row = self.inv_table.currentRow()
        if row < 0:
            QMessageBox.warning(self,"Selection","Search/select one item first."); return
        mat = self.inv_table.item(row,0).text()
        data = self.db.one("""SELECT i.*, c.name category, d.name department, u.name unit, s.name supplier, l.name location, r.name rack
            FROM inventory_items i
            LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id
            LEFT JOIN units u ON u.id=i.unit_id LEFT JOIN suppliers s ON s.id=i.supplier_id
            LEFT JOIN locations l ON l.id=i.location_id LEFT JOIN racks r ON r.id=i.rack_id
            WHERE i.material_id=%s""", (mat,))
        if not data: return
        dlg=QDialog(self); dlg.setWindowTitle(f"Edit Inventory Item - {mat}"); dlg.resize(700,520)
        form=QFormLayout(dlg)
        desc=QLineEdit(data.get('item_description') or '')
        cat=QComboBox(); cat.addItems(self.master_names('categories')); self.set_combo_text(cat,data.get('category') or '')
        dept=QComboBox(); dept.addItems(self.master_names('departments')); self.set_combo_text(dept,data.get('department') or '')
        stock=QDoubleSpinBox(); stock.setRange(0,9999999); stock.setDecimals(2); stock.setValue(float(data.get('quantity_available') or 0))
        unit=QComboBox(); unit.addItems(self.master_names('units')); self.set_combo_text(unit,data.get('unit') or '')
        minq=QDoubleSpinBox(); minq.setRange(0,9999999); minq.setDecimals(2); minq.setValue(float(data.get('minimum_stock_level') or 0))
        typ=QComboBox(); typ.addItems(MATERIAL_TYPES); self.set_combo_text(typ,data.get('material_type') or '')
        rack=QComboBox(); rack.addItems(self.master_names('racks')); self.set_combo_text(rack,data.get('rack') or '')
        loc=QComboBox(); loc.addItems(self.master_names('locations')); self.set_combo_text(loc,data.get('location') or '')
        supplier=QComboBox(); supplier.addItems(self.master_names('suppliers')); self.set_combo_text(supplier,data.get('supplier') or '')
        cost=QDoubleSpinBox(); cost.setRange(0,9999999); cost.setDecimals(3); cost.setPrefix('OMR '); cost.setValue(float(data.get('cost_per_unit') or 0))
        remarks=QTextEdit(data.get('remarks') or ''); remarks.setFixedHeight(70)
        for lab,wid in [("Material ID",QLabel(mat)),("Item Description",desc),("Category",cat),("Department",dept),("Stock",stock),("Unit",unit),("Min. Level",minq),("Type",typ),("Rack No.",rack),("Location",loc),("Supplier",supplier),("Cost/Unit",cost),("Remarks",remarks)]: form.addRow(lab,wid)
        buttons=QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel); form.addRow(buttons)
        buttons.rejected.connect(dlg.reject)
        def save_edit():
            if not desc.text().strip(): QMessageBox.warning(dlg,"Validation","Item description is required."); return
            self.db.execute("""UPDATE inventory_items SET item_description=%s, category_id=%s, department_id=%s, quantity_available=%s, unit_id=%s, minimum_stock_level=%s, material_type=%s, rack_id=%s, location_id=%s, supplier_id=%s, cost_per_unit=%s, remarks=%s, updated_at=CURRENT_TIMESTAMP WHERE material_id=%s""",
                (desc.text().strip(), self.db.id_by_name('categories',cat.currentText()), self.db.id_by_name('departments',dept.currentText()), float(stock.value()), self.db.id_by_name('units',unit.currentText()), float(minq.value()), typ.currentText(), self.db.id_by_name('racks',rack.currentText()), self.db.id_by_name('locations',loc.currentText()), self.db.id_by_name('suppliers',supplier.currentText()), float(cost.value()), remarks.toPlainText(), mat))
            self.db.log(self.user['id'],"EDIT","Inventory",mat,"Inventory item edited from Inventory screen")
            QMessageBox.information(dlg,"Updated","Inventory item updated successfully.")
            dlg.accept(); self.refresh_inventory(); self.refresh_item_combos(); self.refresh_dashboard()
        buttons.accepted.connect(save_edit)
        dlg.exec()

    def inventory_export_rows(self):
        return self.item_rows("")

    def download_inventory_template(self):
        if pd is None:
            QMessageBox.warning(self,"Missing package","Install pandas and openpyxl first."); return
        path,_=QFileDialog.getSaveFileName(self,"Save Excel Upload Template","fm_inventory_upload_template.xlsx","Excel (*.xlsx)")
        if not path: return
        if not path.lower().endswith('.xlsx'): path += '.xlsx'
        cols = [
            "material_id", "item_description", "category", "department", "quantity_available", "unit",
            "minimum_stock_level", "material_type", "supplier", "cost_per_unit", "procurement_date",
            "purchase_order_no", "delivery_order_no", "warranty_details", "expiry_date", "rack", "location", "remarks"
        ]
        sample = [{
            "material_id": "", "item_description": "LED Panel Light 18W", "category": "Electrical", "department": "Facilities",
            "quantity_available": 20, "unit": "Nos", "minimum_stock_level": 5, "material_type": "Fast Moving Spare",
            "supplier": "Default Supplier", "cost_per_unit": 3.5, "procurement_date": today_iso(),
            "purchase_order_no": "PO-001", "delivery_order_no": "DO-001", "warranty_details": "12 months",
            "expiry_date": "", "rack": "R-01", "location": "Main Store", "remarks": "Opening upload"
        }]
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            pd.DataFrame(sample, columns=cols).to_excel(writer, sheet_name="Inventory Upload", index=False)
            pd.DataFrame({"Allowed Categories": self.master_names("categories")}).to_excel(writer, sheet_name="Categories", index=False)
            pd.DataFrame({"Allowed Departments": self.master_names("departments")}).to_excel(writer, sheet_name="Departments", index=False)
            pd.DataFrame({"Allowed Units": self.master_names("units")}).to_excel(writer, sheet_name="Units", index=False)
            pd.DataFrame({"Allowed Material Types": MATERIAL_TYPES}).to_excel(writer, sheet_name="Material Types", index=False)
        QMessageBox.information(self,"Template Saved",f"Excel upload template saved:\n{path}")

    def _ensure_master(self, table: str, name: str) -> Optional[int]:
        name = str(name or "").strip()
        if not name: return None
        existing = self.db.id_by_name(table, name)
        if existing: return existing
        return self.db.execute(f"INSERT OR IGNORE INTO {table}(name) VALUES(%s)", (name,)) or self.db.id_by_name(table, name)

    def import_excel(self):
        if pd is None:
            QMessageBox.warning(self,"Missing package","Install pandas and openpyxl first."); return
        path,_=QFileDialog.getOpenFileName(self,"Upload Inventory / Corrected Store Data","","Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            df=pd.read_excel(path).fillna("")
        except Exception as e:
            QMessageBox.critical(self,"Import Failed",f"Could not read Excel file.\n\n{e}"); return
        df.columns=[str(c).strip().lower().replace(" ","_") for c in df.columns]
        required=['item_description','category','department','quantity_available','unit','minimum_stock_level','material_type','cost_per_unit']
        missing=[c for c in required if c not in df.columns]
        if missing:
            QMessageBox.warning(self,"Invalid Excel Format", "Missing required columns:\n" + ", ".join(missing) + "\n\nUse the same columns as the downloaded inventory file."); return
        if df.empty:
            QMessageBox.warning(self,"Empty File","The selected Excel file has no rows."); return
        preview=ImportPreviewDialog(df, self)
        if preview.exec()!=QDialog.Accepted or not preview.accepted_import: return
        if QMessageBox.question(self, "Confirm Inventory Override", "Uploading inventory will DELETE the existing inventory items and all stock transaction records, then replace them with the uploaded file.\n\nContinue?") != QMessageBox.Yes:
            return
        try:
            for tbl in ["stock_outward_batches", "stock_outward", "inventory_batches", "stock_inward", "stock_adjustments", "stock_transfers", "inventory_items"]:
                self.db.execute(f"DELETE FROM {tbl}")
            self.db.execute("DELETE FROM sqlite_sequence WHERE name IN ('stock_outward_batches','stock_outward','inventory_batches','stock_inward','stock_adjustments','stock_transfers','inventory_items')")
        except Exception as e:
            QMessageBox.critical(self, "Override Failed", f"Could not clear existing inventory/transactions.\n\nReason:\n{e}")
            return
        created=updated=skipped=0; errors=[]
        for idx,row in df.iterrows():
            try:
                desc=str(row.get('item_description','')).strip()
                if not desc:
                    skipped+=1; continue
                qty=float(row.get('quantity_available') or 0); minq=float(row.get('minimum_stock_level') or 0); cost=float(row.get('cost_per_unit') or 0)
                cat_id=self._ensure_master('categories', row.get('category'))
                dept_id=self._ensure_master('departments', row.get('department'))
                unit_id=self._ensure_master('units', row.get('unit'))
                supplier_id=self._ensure_master('suppliers', row.get('supplier') or 'Default Supplier')
                rack_id=self._ensure_master('racks', row.get('rack') or 'R-01')
                loc_id=self._ensure_master('locations', row.get('location') or 'Main Store')
                exp=str(row.get('expiry_date','')).strip() or None
                proc=str(row.get('procurement_date','')).strip() or today_iso()
                mat=str(row.get('material_id','')).strip()
                existing = self.db.one("SELECT id, quantity_available FROM inventory_items WHERE material_id=%s AND is_active=1", (mat,)) if mat else None
                if existing:
                    self.db.execute("""UPDATE inventory_items SET item_description=%s, category_id=%s, department_id=%s, quantity_available=%s,
                        unit_id=%s, minimum_stock_level=%s, material_type=%s, supplier_id=%s, cost_per_unit=%s, procurement_date=%s,
                        purchase_order_no=%s, delivery_order_no=%s, warranty_details=%s, expiry_date=%s, rack_id=%s, location_id=%s, remarks=%s, updated_at=CURRENT_TIMESTAMP WHERE id=%s""",
                        (desc,cat_id,dept_id,qty,unit_id,minq,str(row.get('material_type','General Spare')),supplier_id,cost,proc,str(row.get('purchase_order_no','')),str(row.get('delivery_order_no','')),str(row.get('warranty_details','')),exp,rack_id,loc_id,str(row.get('remarks','Excel import update')),existing['id']))
                    item_id=existing['id']; updated+=1
                else:
                    mat=self.db.next_code("MAT","inventory_items","material_id")
                    item_id=self.db.execute("""INSERT INTO inventory_items(material_id,item_description,category_id,department_id,quantity_available,unit_id,minimum_stock_level,material_type,supplier_id,cost_per_unit,procurement_date,purchase_order_no,delivery_order_no,warranty_details,expiry_date,rack_id,location_id,remarks)
                        VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                        (mat,desc,cat_id,dept_id,qty,unit_id,minq,str(row.get('material_type','General Spare')),supplier_id,cost,proc,str(row.get('purchase_order_no','')),str(row.get('delivery_order_no','')),str(row.get('warranty_details','')),exp,rack_id,loc_id,str(row.get('remarks','Excel import'))))
                    created+=1
                inward_no=self.db.next_code("IN","stock_inward","inward_no")
                inward_id = self.db.execute("INSERT INTO stock_inward(inward_no,material_id,item_id,quantity,unit_cost,total_cost,procurement_date,supplier_id,po_no,do_no,remarks,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                    (inward_no,mat,item_id,qty,cost,qty*cost,proc,supplier_id,str(row.get('purchase_order_no','')),str(row.get('delivery_order_no','')),"Excel stock upload",self.user['id']))
                if exp:
                    self.create_expiry_batch(inward_id, item_id, mat, qty, exp)
            except Exception as e:
                errors.append(f"Row {idx+2}: {e}")
        self.db.log(self.user['id'],"IMPORT","Inventory","Excel Upload",f"Created {created}, updated {updated}, skipped {skipped}")
        msg=f"Inventory override completed.\n\nCreated materials: {created}\nUpdated duplicate material IDs in upload: {updated}\nSkipped blank rows: {skipped}"
        if errors: msg += "\n\nErrors:\n" + "\n".join(errors[:8])
        QMessageBox.information(self,"Import Completed",msg); self.refresh_inventory(); self.refresh_item_combos(); self.refresh_dashboard()

    def export_complete_inventory(self):
        if not hasattr(self, 'inv_table'):
            return
        rows=self.inventory_export_rows()
        cols=[("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min Level","minimum_stock_level"),("Type","material_type"),("Supplier","supplier"),("Cost/Unit","cost_per_unit"),("Total Value","total_value"),("Expiry","expiry_date"),("Location","location"),("Rack","rack"),("Remarks","remarks"),("Created","created_at"),("Updated","updated_at")]
        temp=QTableWidget(); self.table_fill(temp, rows, cols); self.export_table(temp,"complete_store_inventory")

    def page_reports(self):
        w,v=self.base_page("Reports")
        controls_box = QGroupBox("Report Filters")
        controls = QGridLayout(controls_box)
        controls.setContentsMargins(18, 22, 18, 18)
        controls.setHorizontalSpacing(18)
        controls.setVerticalSpacing(14)

        self.rep_type=QComboBox()
        self.rep_type.addItems([
            "Balance Stock Report",
            "Complete Inventory Report",
            "Low Stock List",
            "Expiring Materials List",
            "Expired Materials List",
            "Stock Inward Report",
            "Stock Outward Report",
            "Transfer List",
            "Department-wise Stock Report",
            "Category-wise Stock Report",
            "Supplier-wise Purchase Report",
            "Store Location-wise Stock Report"
        ])
        self.rep_department=QComboBox(); self.rep_department.addItem("All"); self.rep_department.addItems(self.master_names("departments"))
        self.rep_category=QComboBox(); self.rep_category.addItem("All"); self.rep_category.addItems(self.master_names("categories"))
        self.rep_supplier=QComboBox(); self.rep_supplier.addItem("All"); self.rep_supplier.addItems(self.master_names("suppliers"))
        self.rep_from=QDateEdit(); self.rep_from.setCalendarPopup(True); self.rep_from.setDisplayFormat("dd/MM/yyyy"); self.rep_from.setDate(date(date.today().year, date.today().month, 1))
        self.rep_to=QDateEdit(); self.rep_to.setCalendarPopup(True); self.rep_to.setDisplayFormat("dd/MM/yyyy"); self.rep_to.setDate(date.today())
        for c in [self.rep_type,self.rep_department,self.rep_category,self.rep_supplier]:
            c.setMinimumWidth(330); c.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        gen=QPushButton("Generate"); gen.clicked.connect(self.generate_report)
        ex=QPushButton("Export CSV/Excel/PDF"); ex.clicked.connect(lambda:self.export_table(self.rep_table,"report"))

        controls.addWidget(QLabel("Report"), 0, 0); controls.addWidget(self.rep_type, 0, 1)
        controls.addWidget(QLabel("Department"), 0, 2); controls.addWidget(self.rep_department, 0, 3)
        controls.addWidget(QLabel("Category"), 1, 0); controls.addWidget(self.rep_category, 1, 1)
        controls.addWidget(QLabel("Supplier"), 1, 2); controls.addWidget(self.rep_supplier, 1, 3)
        controls.addWidget(QLabel("Date From"), 2, 0); controls.addWidget(self.rep_from, 2, 1)
        controls.addWidget(QLabel("Date To"), 2, 2); controls.addWidget(self.rep_to, 2, 3)
        controls.addWidget(gen, 3, 1); controls.addWidget(ex, 3, 3)
        controls.setColumnStretch(1, 1); controls.setColumnStretch(3, 1)
        v.addWidget(controls_box)

        self.rep_table=QTableWidget(); v.addWidget(self.rep_table, 1)
        return w

    def _filtered_inventory_rows_for_reports(self):
        rows = self.item_rows()
        dept = self.rep_department.currentText() if hasattr(self, 'rep_department') else 'All'
        cat = self.rep_category.currentText() if hasattr(self, 'rep_category') else 'All'
        sup = self.rep_supplier.currentText() if hasattr(self, 'rep_supplier') else 'All'
        if dept != 'All': rows = [r for r in rows if (r.get('department') or '') == dept]
        if cat != 'All': rows = [r for r in rows if (r.get('category') or '') == cat]
        if sup != 'All': rows = [r for r in rows if (r.get('supplier') or '') == sup]
        if hasattr(self, 'rep_from') and hasattr(self, 'rep_to'):
            start = self.rep_from.date().toPython().isoformat(); end = self.rep_to.date().toPython().isoformat()
            rows = [r for r in rows if not r.get('procurement_date') or start <= str(r.get('procurement_date')) <= end]
        return rows

    def _append_inventory_filter_sql(self, sql: str, params: list, date_column: Optional[str] = None):
        dept = self.rep_department.currentText() if hasattr(self, 'rep_department') else 'All'
        cat = self.rep_category.currentText() if hasattr(self, 'rep_category') else 'All'
        sup = self.rep_supplier.currentText() if hasattr(self, 'rep_supplier') else 'All'
        if dept != 'All':
            sql += " AND d.name=%s"; params.append(dept)
        if cat != 'All':
            sql += " AND c.name=%s"; params.append(cat)
        if sup != 'All':
            sql += " AND s.name=%s"; params.append(sup)
        if date_column and hasattr(self, 'rep_from') and hasattr(self, 'rep_to'):
            sql += f" AND {date_column} BETWEEN %s AND %s"
            params.extend([self.rep_from.date().toPython().isoformat(), self.rep_to.date().toPython().isoformat()])
        return sql, params

    def choose_report_columns(self, cols: List[Tuple[str, str]]) -> Optional[List[Tuple[str, str]]]:
        dlg = ColumnSelectDialog(cols, self)
        if dlg.exec() == QDialog.Accepted:
            return dlg.selected_columns
        return None

    def generate_report(self):
        typ=self.rep_type.currentText()
        base_cols=[("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min. Level","minimum_stock_level"),("Type","material_type"),("Supplier","supplier"),("Rack No.","rack"),("Location","location"),("Cost/Unit (OMR)","cost_per_unit"),("Total Value (OMR)","total_value")]
        if typ in ["Balance Stock Report", "Complete Inventory Report", "Department-wise Stock Report", "Category-wise Stock Report", "Supplier-wise Purchase Report", "Store Location-wise Stock Report"]:
            rows=self._filtered_inventory_rows_for_reports(); cols=base_cols
        elif typ=="Low Stock List":
            rows=[r for r in self._filtered_inventory_rows_for_reports() if float(r.get('quantity_available') or 0) <= float(r.get('minimum_stock_level') or 0)]
            cols=[("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min Level","minimum_stock_level"),("Type","material_type"),("Supplier","supplier"),("Rack No.","rack"),("Location","location")]
        elif typ=="Expiring Materials List":
            sql="""SELECT b.material_id,i.item_description,b.batch_no,b.quantity_received,b.quantity_available,c.name category,d.name department,s.name supplier,l.name location,r.name rack,b.expiry_date,CAST(julianday(b.expiry_date)-julianday('now') AS INTEGER) days_left
                   FROM inventory_batches b JOIN inventory_items i ON i.id=b.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN suppliers s ON s.id=i.supplier_id
                   LEFT JOIN locations l ON l.id=i.location_id LEFT JOIN racks r ON r.id=i.rack_id
                   WHERE i.is_active=1 AND b.quantity_available>0 AND b.expiry_date BETWEEN DATE('now') AND DATE('now','+30 day')"""
            sql, params = self._append_inventory_filter_sql(sql, [], "b.expiry_date")
            rows=self.db.query(sql, tuple(params)); cols=[("Material ID","material_id"),("Item Description","item_description"),("Batch No","batch_no"),("Batch Qty Received","quantity_received"),("Batch Qty Left","quantity_available"),("Category","category"),("Department","department"),("Supplier","supplier"),("Location","location"),("Rack No.","rack"),("Expiry","expiry_date"),("Days Left","days_left")]
        elif typ=="Expired Materials List":
            sql="""SELECT b.material_id,i.item_description,b.batch_no,b.quantity_received,b.quantity_available,c.name category,d.name department,s.name supplier,l.name location,r.name rack,b.expiry_date,CAST(julianday('now')-julianday(b.expiry_date) AS INTEGER) days_overdue
                   FROM inventory_batches b JOIN inventory_items i ON i.id=b.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN suppliers s ON s.id=i.supplier_id
                   LEFT JOIN locations l ON l.id=i.location_id LEFT JOIN racks r ON r.id=i.rack_id
                   WHERE i.is_active=1 AND b.quantity_available>0 AND b.expiry_date < DATE('now')"""
            sql, params = self._append_inventory_filter_sql(sql, [], "b.expiry_date")
            rows=self.db.query(sql, tuple(params)); cols=[("Material ID","material_id"),("Item Description","item_description"),("Batch No","batch_no"),("Batch Qty Received","quantity_received"),("Batch Qty Left","quantity_available"),("Category","category"),("Department","department"),("Supplier","supplier"),("Location","location"),("Rack No.","rack"),("Expiry","expiry_date"),("Overdue","days_overdue")]
        elif typ=="Stock Inward Report":
            sql="""SELECT si.inward_no,si.material_id,i.item_description,c.name category,d.name department,s.name supplier,si.quantity,si.unit_cost,si.total_cost,si.procurement_date,si.po_no,si.do_no,si.remarks
                   FROM stock_inward si LEFT JOIN inventory_items i ON i.id=si.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN suppliers s ON s.id=si.supplier_id
                   WHERE 1=1"""
            sql, params = self._append_inventory_filter_sql(sql, [], "si.procurement_date")
            rows=self.db.query(sql, tuple(params)); cols=[("Inward No","inward_no"),("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Supplier","supplier"),("Qty","quantity"),("Unit Cost (OMR)","unit_cost"),("Total (OMR)","total_cost"),("Date","procurement_date"),("MR No","po_no"),("DO No","do_no"),("Remarks","remarks")]
        elif typ=="Stock Outward Report":
            sql="""SELECT so.siv_no,so.issue_date,so.material_id,i.item_description,c.name category,d.name department,s.name supplier,so.quantity_issued,so.issued_to,so.issued_by,so.remarks
                   FROM stock_outward so LEFT JOIN inventory_items i ON i.id=so.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=so.department_id LEFT JOIN suppliers s ON s.id=i.supplier_id
                   WHERE 1=1"""
            sql, params = self._append_inventory_filter_sql(sql, [], "so.issue_date")
            rows=self.db.query(sql, tuple(params)); cols=[("SIV #","siv_no"),("Date","issue_date"),("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Supplier","supplier"),("Qty","quantity_issued"),("Issued To","issued_to"),("Issued By","issued_by"),("Remarks","remarks")]
        elif typ=="Transfer List":
            sql="""SELECT st.transfer_no,st.transfer_date,st.material_id,i.item_description,c.name category,fd.name from_department,td.name to_department,s.name supplier,st.quantity_transferred,st.requested_by,st.approved_by,st.remarks
                   FROM stock_transfers st LEFT JOIN inventory_items i ON i.id=st.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN departments fd ON fd.id=st.from_department_id LEFT JOIN departments td ON td.id=st.to_department_id LEFT JOIN suppliers s ON s.id=i.supplier_id
                   WHERE 1=1"""
            sql, params = self._append_inventory_filter_sql(sql, [], "st.transfer_date")
            rows=self.db.query(sql, tuple(params)); cols=[("Transfer No","transfer_no"),("Date","transfer_date"),("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("From Department","from_department"),("To Department","to_department"),("Supplier","supplier"),("Qty","quantity_transferred"),("Requested By","requested_by"),("Approved By","approved_by"),("Remarks","remarks")]
        else:
            rows=self._filtered_inventory_rows_for_reports(); cols=base_cols
        chosen = self.choose_report_columns(cols)
        if chosen is None:
            return
        self.table_fill(self.rep_table,rows,chosen)

    def export_table(self, table: QTableWidget, basename: str):
        path,_=QFileDialog.getSaveFileName(self,"Export",f"{basename}_{datetime.now().strftime('%Y%m%d_%H%M')}","Excel (*.xlsx);;CSV (*.csv);;PDF (*.pdf)")
        if not path: return
        headers=[table.horizontalHeaderItem(c).text() for c in range(table.columnCount())]
        rows=[[table.item(r,c).text() if table.item(r,c) else "" for c in range(table.columnCount())] for r in range(table.rowCount())]
        if path.endswith('.xlsx'):
            if pd is None: QMessageBox.warning(self,"Missing package","Install pandas openpyxl."); return
            pd.DataFrame(rows, columns=headers).to_excel(path,index=False)
        elif path.endswith('.pdf'):
            if canvas is None: QMessageBox.warning(self,"Missing package","Install reportlab."); return
            c=canvas.Canvas(path, pagesize=landscape(A4)); w,h=landscape(A4); y=h-35; c.setFont("Helvetica-Bold",12); c.drawString(30,y,APP_TITLE); y-=25; c.setFont("Helvetica",7); c.drawString(30,y," | ".join(headers)); y-=14
            for row in rows[:200]:
                c.drawString(30,y," | ".join(row)[:180]); y-=12
                if y<35: c.showPage(); y=h-35; c.setFont("Helvetica",7)
            c.save()
        else:
            if not path.endswith('.csv'): path += '.csv'
            with open(path,'w',newline='',encoding='utf-8') as f: wr=csv.writer(f); wr.writerow(headers); wr.writerows(rows)
        QMessageBox.information(self,"Exported",f"Saved: {path}")

    def page_settings(self):
        w,v=self.base_page("Settings"); tabs=QTabWidget(); v.addWidget(tabs)
        for table,title in [("departments","Departments"),("categories","Categories"),("units","Units"),("suppliers","Suppliers"),("locations","Locations"),("racks","Racks")]: tabs.addTab(self.master_tab(table), title)
        backup=QWidget(); bl=QVBoxLayout(backup)
        b1=QPushButton("Backup Database"); b1.setObjectName("softBtn"); b1.clicked.connect(self.backup_db)
        b2=QPushButton("Restore Database"); b2.setObjectName("orangeBtn"); b2.clicked.connect(self.restore_db)
        bl.addWidget(QLabel("Use Backup before major uploads or corrections.")); bl.addWidget(b1); bl.addWidget(b2); bl.addStretch(); tabs.addTab(backup,"Backup / Restore")
        return w

    def refresh_master_dropdowns(self):
        """Refresh master-data combo boxes after Settings add/edit/delete/upload."""
        def reset_combo(attr: str, values: List[str], include_all: bool = False, include_blank: bool = False):
            if not hasattr(self, attr):
                return
            combo = getattr(self, attr)
            old = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            if include_all:
                combo.addItem("All")
            if include_blank:
                combo.addItem("")
            combo.addItems(values)
            idx = combo.findText(old)
            if idx >= 0:
                combo.setCurrentIndex(idx)
            elif combo.count() > 0:
                combo.setCurrentIndex(0)
            combo.blockSignals(False)

        cats = self.master_names("categories")
        depts = self.master_names("departments")
        units = self.master_names("units")
        sups = self.master_names("suppliers")
        locs = self.master_names("locations")
        racks = self.master_names("racks")
        for attr in ["in_cat"]: reset_combo(attr, cats, include_blank=True)
        for attr in ["in_dept", "out_dept", "tr_from_dept", "tr_to_dept"]: reset_combo(attr, depts, include_blank=True)
        for attr in ["in_unit"]: reset_combo(attr, units, include_blank=True)
        for attr in ["in_supplier"]: reset_combo(attr, sups, include_blank=True)
        for attr in ["in_location", "tr_from_loc", "tr_to_loc"]: reset_combo(attr, locs, include_blank=True)
        for attr in ["in_rack", "tr_from_rack", "tr_to_rack"]: reset_combo(attr, racks, include_blank=True)
        reset_combo("rep_department", depts, include_all=True)
        reset_combo("rep_category", cats, include_all=True)
        reset_combo("rep_supplier", sups, include_all=True)
        self.refresh_item_combos()
        if hasattr(self, "inv_table"):
            self.refresh_inventory()

    def page_settings(self):
        w,v=self.base_page("Settings"); tabs=QTabWidget(); v.addWidget(tabs)
        for table,title in [("departments","Departments"),("categories","Categories"),("units","Units"),("suppliers","Suppliers"),("locations","Locations"),("racks","Racks")]: tabs.addTab(self.master_tab(table), title)
        backup=QWidget(); bl=QVBoxLayout(backup)
        b1=QPushButton("Backup Database"); b1.setObjectName("softBtn"); b1.clicked.connect(self.backup_db)
        b2=QPushButton("Restore Database"); b2.setObjectName("orangeBtn"); b2.clicked.connect(self.restore_db)
        bl.addWidget(QLabel("Use Backup before major uploads or corrections.")); bl.addWidget(b1); bl.addWidget(b2); bl.addStretch(); tabs.addTab(backup,"Backup / Restore")
        return w

    def master_tab(self, table: str):
        w = QWidget(); v = QVBoxLayout(w)
        h = QHBoxLayout()
        inp = QLineEdit(); inp.setPlaceholderText(f"{table[:-1].title()} name")
        add = QPushButton("Add"); add.setObjectName("greenBtn")
        edit = QPushButton("Edit Selected"); edit.setObjectName("softBtn")
        delete = QPushButton("Delete Selected"); delete.setObjectName("redBtn")
        download = QPushButton("Download"); download.setObjectName("softBtn")
        upload = QPushButton("Upload Corrected Data"); upload.setObjectName("softBtn")
        h.addWidget(inp); h.addWidget(add); h.addWidget(edit); h.addWidget(delete); h.addWidget(download); h.addWidget(upload); v.addLayout(h)
        tw = QTableWidget(); tw.setSelectionBehavior(QTableWidget.SelectRows); tw.setSelectionMode(QTableWidget.SingleSelection); v.addWidget(tw)

        def display_columns():
            if table == "suppliers":
                return [("Name", "name"), ("Contact", "contact"), ("Email", "email"), ("Active", "is_active"), ("Created", "created_at")]
            if table == "racks":
                return [("Name", "name"), ("Location", "location"), ("Active", "is_active"), ("Created", "created_at")]
            return [("Name", "name"), ("Active", "is_active"), ("Created", "created_at")]

        def get_rows():
            if table == "racks":
                return self.db.query("""SELECT r.id,r.name,l.name location,r.is_active,r.created_at FROM racks r LEFT JOIN locations l ON l.id=r.location_id ORDER BY r.name""")
            if table == "suppliers":
                return self.db.query("SELECT id,name,contact,email,is_active,created_at FROM suppliers ORDER BY name")
            return self.db.query(f"SELECT id,name,is_active,created_at FROM {table} ORDER BY name")

        def refresh():
            rows = get_rows()
            cols = display_columns()
            tw.setColumnCount(len(cols)); tw.setHorizontalHeaderLabels([c[0] for c in cols]); tw.setRowCount(len(rows))
            for r, row in enumerate(rows):
                for c, (_, key) in enumerate(cols):
                    value = row.get(key, "")
                    item = QTableWidgetItem("" if value is None else str(value))
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
                    if c == 0:
                        item.setData(Qt.UserRole, row.get("id"))
                    tw.setItem(r, c, item)
            tw.setWordWrap(True); tw.resizeColumnsToContents(); tw.resizeRowsToContents()
            tw.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive); tw.horizontalHeader().setStretchLastSection(True)

        def selected_id_name():
            r = tw.currentRow()
            if r < 0 or not tw.item(r,0): return None, None
            return tw.item(r,0).data(Qt.UserRole), tw.item(r,0).text()

        def after_master_change(message: str):
            refresh(); self.refresh_master_dropdowns(); QMessageBox.information(self, "Updated", message)

        def do_add():
            name = inp.text().strip()
            if not name:
                QMessageBox.warning(self,"Validation","Enter a name first."); return
            self.db.execute(f"INSERT OR IGNORE INTO {table}(name,is_active) VALUES(%s,1)",(name,))
            inp.clear(); after_master_change(f"{name} added.")

        def do_edit():
            row_id, old = selected_id_name()
            if not row_id: QMessageBox.warning(self,"Selection","Select a row first."); return
            new, ok = QInputDialog.getText(self,"Edit",f"Rename {old} to:", text=old)
            if ok and new.strip():
                self.db.execute(f"UPDATE {table} SET name=%s, updated_at=CURRENT_TIMESTAMP WHERE id=%s",(new.strip(), row_id))
                after_master_change(f"Updated to {new.strip()}.")

        def do_delete():
            row_id, old = selected_id_name()
            if not row_id: QMessageBox.warning(self,"Selection","Select a row first."); return
            if QMessageBox.question(self,"Confirm Delete",f"Delete '{old}' permanently?\n\nIf this master value is already used in inventory/transactions, SQLite may block deletion to protect existing records.") == QMessageBox.Yes:
                try:
                    self.db.execute(f"DELETE FROM {table} WHERE id=%s",(row_id,))
                    after_master_change(f"{old} deleted.")
                except Exception as e:
                    QMessageBox.critical(self,"Delete Failed",f"Could not delete '{old}'.\n\nIt is probably already used in inventory or transaction records.\n\nReason:\n{e}")

        def do_download():
            rows = get_rows()
            for row in rows:
                row.pop("id", None)
            path,_=QFileDialog.getSaveFileName(self, "Download Master Data", f"{table}_master_data.xlsx", "Excel (*.xlsx);;CSV (*.csv)")
            if not path: return
            try:
                if path.lower().endswith(".csv"):
                    if not rows:
                        rows=[{k:"" for _, k in display_columns()}]
                    with open(path,'w',newline='',encoding='utf-8') as f:
                        wr=csv.DictWriter(f, fieldnames=list(rows[0].keys())); wr.writeheader(); wr.writerows(rows)
                else:
                    if not path.lower().endswith('.xlsx'): path += '.xlsx'
                    if pd is None:
                        QMessageBox.warning(self,"Missing Package","Install pandas and openpyxl to export Excel."); return
                    pd.DataFrame(rows).to_excel(path,index=False)
                QMessageBox.information(self,"Downloaded",f"Master data downloaded:\n{path}")
            except Exception as e:
                QMessageBox.critical(self,"Download Failed",str(e))

        def do_upload():
            if pd is None:
                QMessageBox.warning(self,"Missing Package","Install pandas and openpyxl first."); return
            path,_=QFileDialog.getOpenFileName(self,"Upload Corrected Master Data","","Excel/CSV (*.xlsx *.csv)")
            if not path: return
            try:
                df = pd.read_csv(path) if path.lower().endswith('.csv') else pd.read_excel(path)
                df.columns = [str(c).strip().lower() for c in df.columns]
                if 'name' not in df.columns:
                    QMessageBox.warning(self,"Invalid File","The uploaded file must contain a 'name' column."); return
                added=updated=skipped=0
                for _, row in df.fillna('').iterrows():
                    name = str(row.get('name','')).strip()
                    if not name:
                        skipped += 1; continue
                    active = int(float(row.get('is_active', 1) or 1))
                    existing = self.db.one(f"SELECT id FROM {table} WHERE name=%s", (name,))
                    if table == 'suppliers':
                        contact = str(row.get('contact','')).strip(); email = str(row.get('email','')).strip()
                        if existing:
                            self.db.execute("UPDATE suppliers SET contact=%s,email=%s,is_active=%s,updated_at=CURRENT_TIMESTAMP WHERE id=%s", (contact,email,active,existing['id'])); updated += 1
                        else:
                            self.db.execute("INSERT INTO suppliers(name,contact,email,is_active) VALUES(%s,%s,%s,%s)", (name,contact,email,active)); added += 1
                    elif table == 'racks':
                        loc_name = str(row.get('location','')).strip()
                        loc_id = self._ensure_master('locations', loc_name) if loc_name else None
                        if existing:
                            self.db.execute("UPDATE racks SET location_id=%s,is_active=%s,updated_at=CURRENT_TIMESTAMP WHERE id=%s", (loc_id,active,existing['id'])); updated += 1
                        else:
                            self.db.execute("INSERT INTO racks(name,location_id,is_active) VALUES(%s,%s,%s)", (name,loc_id,active)); added += 1
                    else:
                        if existing:
                            self.db.execute(f"UPDATE {table} SET is_active=%s,updated_at=CURRENT_TIMESTAMP WHERE id=%s", (active,existing['id'])); updated += 1
                        else:
                            self.db.execute(f"INSERT INTO {table}(name,is_active) VALUES(%s,%s)", (name,active)); added += 1
                after_master_change(f"Upload completed. Added: {added}, Updated: {updated}, Skipped blank rows: {skipped}")
            except Exception as e:
                QMessageBox.critical(self,"Upload Failed",str(e))

        add.clicked.connect(do_add); edit.clicked.connect(do_edit); delete.clicked.connect(do_delete); download.clicked.connect(do_download); upload.clicked.connect(do_upload)
        refresh(); return w

    def backup_db(self):
        path,_=QFileDialog.getSaveFileName(self,"Backup Database","fm_inventory_backup.db","SQLite DB (*.db)")
        if not path: return
        if not path.endswith('.db'): path += '.db'
        try:
            shutil.copy2(DB_PATH, path)
            QMessageBox.information(self,"Backup", f"Backup completed:\n{path}")
        except Exception as e:
            QMessageBox.critical(self,"Backup Failed", str(e))

    def restore_db(self):
        path,_=QFileDialog.getOpenFileName(self,"Restore Database From Backup","","SQLite DB (*.db)")
        if not path: return
        if QMessageBox.question(self, "Confirm Restore", "Restoring will replace the current database with the selected backup. The application will close after restore.\n\nContinue?") != QMessageBox.Yes:
            return
        try:
            backup_current = DB_PATH + ".before_restore"
            if os.path.exists(DB_PATH):
                shutil.copy2(DB_PATH, backup_current)
            shutil.copy2(path, DB_PATH)
            QMessageBox.information(self, "Restore Completed", "Database restored successfully. Restart the application now.")
            QApplication.quit()
        except Exception as e:
            QMessageBox.critical(self, "Restore Failed", str(e))

    def page_audit(self):
        w,v=self.base_page("Audit Logs"); self.audit_table=QTableWidget(); v.addWidget(self.audit_table); rows=self.db.query("SELECT a.created_at,u.username,a.action,a.module,a.record_ref,a.details FROM audit_logs a LEFT JOIN users u ON u.id=a.user_id ORDER BY a.id DESC LIMIT 300"); self.table_fill(self.audit_table,rows,[("Date","created_at"),("User","username"),("Action","action"),("Module","module"),("Record","record_ref"),("Details","details")]); return w




# ========================= MARKET-GRADE UPGRADE PATCH V19 =========================
# Adds: persistent barcode fields, item ledger, audit log access, role permission checks,
# QR/barcode label PDF support, and professional PDF report export.

_ORIG_DB_CREATE_SCHEMA = Database.create_schema
_ORIG_DB_SEED_DEFAULTS = Database.seed_defaults
_ORIG_SAVE_INWARD = MainWindow.save_inward
_ORIG_SAVE_OUTWARD = MainWindow.save_outward
_ORIG_SAVE_ADJUSTMENT = MainWindow.save_adjustment
_ORIG_SAVE_TRANSFER = MainWindow.save_transfer
_ORIG_DELETE_SELECTED_INVENTORY = MainWindow.delete_selected_inventory
_ORIG_IMPORT_EXCEL = MainWindow.import_excel


def _v19_db_create_schema(self):
    _ORIG_DB_CREATE_SCHEMA(self)
    # Safe migrations for existing databases.
    for coldef in [
        "barcode TEXT",
        "manufacturer_part_no TEXT",
        "model_no TEXT",
    ]:
        try:
            self.execute(f"ALTER TABLE inventory_items ADD COLUMN {coldef}")
        except Exception:
            pass
    # Permission rows are kept normalized so future UI restrictions can be expanded.
    try:
        self.execute("CREATE INDEX IF NOT EXISTS idx_inventory_barcode ON inventory_items(barcode)")
    except Exception:
        pass


def _v19_db_seed_defaults(self):
    _ORIG_DB_SEED_DEFAULTS(self)
    # Default role permissions. These do not delete user data; they only seed missing permissions.
    modules = ["Dashboard", "Inventory", "Stock Inward", "Stock Outward", "Stock Adjustment", "Transfers", "Reports", "Settings", "Audit Logs"]
    role_rules = {
        "Admin":        dict(view=1, add=1, edit=1, delete=1),
        "Store Keeper": dict(view=1, add=1, edit=1, delete=0),
        "Manager":      dict(view=1, add=0, edit=0, delete=0),
        "Viewer":       dict(view=1, add=0, edit=0, delete=0),
    }
    # Managers can export/view reports; settings/audit are view only. Viewers are strictly read-only.
    for role, rule in role_rules.items():
        role_row = self.one("SELECT id FROM roles WHERE name=%s", (role,))
        if not role_row:
            continue
        for module in modules:
            r = dict(rule)
            if role == "Store Keeper" and module in ["Settings", "Audit Logs"]:
                r = dict(view=0, add=0, edit=0, delete=0)
            if role == "Manager" and module in ["Stock Inward", "Stock Outward", "Stock Adjustment", "Transfers", "Inventory"]:
                r = dict(view=1, add=0, edit=0, delete=0)
            if role == "Viewer" and module in ["Settings", "Audit Logs"]:
                r = dict(view=0, add=0, edit=0, delete=0)
            existing = self.one("SELECT id FROM permissions WHERE role_id=%s AND module=%s", (role_row['id'], module))
            if existing:
                self.execute("UPDATE permissions SET can_view=%s, can_add=%s, can_edit=%s, can_delete=%s WHERE id=%s", (r['view'], r['add'], r['edit'], r['delete'], existing['id']))
            else:
                self.execute("INSERT INTO permissions(role_id,module,can_view,can_add,can_edit,can_delete) VALUES(%s,%s,%s,%s,%s,%s)", (role_row['id'], module, r['view'], r['add'], r['edit'], r['delete']))

Database.create_schema = _v19_db_create_schema
Database.seed_defaults = _v19_db_seed_defaults


def _v19_role_name(self):
    return str(self.user.get('role') or '').strip() or 'Viewer'


def _v19_has_permission(self, module: str, action: str = 'view') -> bool:
    role = _v19_role_name(self)
    if role == 'Admin':
        return True
    col = {'view':'can_view','add':'can_add','edit':'can_edit','delete':'can_delete'}.get(action, 'can_view')
    row = self.db.one(f"""SELECT p.{col} allowed
                       FROM permissions p JOIN roles r ON r.id=p.role_id
                       WHERE r.name=%s AND p.module=%s""", (role, module))
    return bool(row and int(row.get('allowed') or 0) == 1)


def _v19_require_permission(self, module: str, action: str = 'view') -> bool:
    if _v19_has_permission(self, module, action):
        return True
    QMessageBox.warning(self, "Permission Blocked", f"Your role ({_v19_role_name(self)}) does not have {action} permission for {module}.")
    return False

MainWindow.has_permission = _v19_has_permission
MainWindow.require_permission = _v19_require_permission


def _v19_item_rows(self, search: str = "") -> List[Dict[str, Any]]:
    q = """SELECT i.*, c.name category, d.name department, u.name unit, s.name supplier, l.name location, r.name rack,
            (i.quantity_available*i.cost_per_unit) total_value FROM inventory_items i
            LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN units u ON u.id=i.unit_id
            LEFT JOIN suppliers s ON s.id=i.supplier_id LEFT JOIN locations l ON l.id=i.location_id LEFT JOIN racks r ON r.id=i.rack_id
            WHERE i.is_active=1"""
    params: Tuple = ()
    if search:
        like = f"%{search}%"
        q += " AND (i.material_id LIKE %s OR i.barcode LIKE %s OR i.item_description LIKE %s OR c.name LIKE %s OR d.name LIKE %s OR s.name LIKE %s OR l.name LIKE %s OR r.name LIKE %s)"
        params = (like,)*8
    q += " ORDER BY i.updated_at DESC"
    return self.db.query(q, params)

MainWindow.item_rows = _v19_item_rows


def _v19_page_inventory(self):
    w,v=self.base_page("Inventory Management")
    actions = QFrame(); actions.setObjectName("card")
    av = QVBoxLayout(actions); av.setContentsMargins(16,14,16,14)

    head = QHBoxLayout()
    title = QLabel("Store Materials Register"); title.setStyleSheet("font-size:17px;font-weight:900;")
    hint = QLabel("Search by material ID, barcode/QR, item, category, department, supplier, rack, or location."); hint.setObjectName("small")
    th = QVBoxLayout(); th.addWidget(title); th.addWidget(hint); head.addLayout(th); head.addStretch()
    self.inv_search=QLineEdit(); self.inv_search.setPlaceholderText("Search Material ID, barcode, item, category, department, supplier, rack, location...")
    self.inv_search.setMinimumWidth(520); self.inv_search.returnPressed.connect(self.refresh_inventory); head.addWidget(self.inv_search, 1)
    b=QPushButton("Search"); b.clicked.connect(self.refresh_inventory); head.addWidget(b)
    av.addLayout(head)

    btns = QHBoxLayout(); btns.setSpacing(8)
    down=QPushButton("Download Inventory"); down.clicked.connect(self.export_complete_inventory)
    imp=QPushButton("Upload Inventory"); imp.setObjectName("greenBtn"); imp.clicked.connect(self.import_excel)
    exp=QPushButton("Export Current View"); exp.setObjectName("softBtn"); exp.clicked.connect(lambda:self.export_table(self.inv_table,"inventory_current_view"))
    edit=QPushButton("Edit Selected Item"); edit.setObjectName("orangeBtn"); edit.clicked.connect(self.edit_selected_inventory)
    delete=QPushButton("Delete Selected Item"); delete.setObjectName("redBtn"); delete.clicked.connect(self.delete_selected_inventory)
    ledger=QPushButton("Item Ledger"); ledger.setObjectName("softBtn"); ledger.clicked.connect(self.show_selected_item_ledger)
    label=QPushButton("Barcode / QR Labels"); label.setObjectName("softBtn"); label.clicked.connect(self.export_selected_barcode_label)
    for x in [down, imp, exp, edit, delete, ledger, label]:
        x.setMinimumWidth(150); btns.addWidget(x)
    btns.addStretch(); av.addLayout(btns)
    v.addWidget(actions)
    self.inv_table=QTableWidget(); self.inv_table.setAlternatingRowColors(True); v.addWidget(self.inv_table, 1); self.refresh_inventory(); return w

MainWindow.page_inventory = _v19_page_inventory


def _v19_refresh_inventory(self):
    self.inv_rows = self.item_rows(self.inv_search.text().strip() if hasattr(self,'inv_search') else "")
    cols=[("Material ID","material_id"),("Barcode / QR","barcode"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min. Level","minimum_stock_level"),("Type","material_type"),("Rack No.","rack"),("Location","location")]
    self.table_fill(self.inv_table,self.inv_rows,cols)
    widths = [145, 160, 430, 150, 165, 100, 100, 120, 190, 120, 150]
    for i, width in enumerate(widths):
        self.inv_table.setColumnWidth(i, width)
    self.inv_table.setSelectionBehavior(QTableWidget.SelectRows)
    self.inv_table.setSelectionMode(QTableWidget.SingleSelection)
    self.inv_table.resizeRowsToContents()

MainWindow.refresh_inventory = _v19_refresh_inventory


def _v19_selected_material_from_inventory(self) -> Optional[str]:
    if not hasattr(self, 'inv_table'):
        return None
    row = self.inv_table.currentRow()
    if row < 0 or not self.inv_table.item(row,0):
        QMessageBox.warning(self, "Selection", "Search/select one inventory item first.")
        return None
    return self.inv_table.item(row,0).text()

MainWindow.selected_material_from_inventory = _v19_selected_material_from_inventory


def _v19_show_selected_item_ledger(self):
    mat = self.selected_material_from_inventory()
    if not mat: return
    dlg=QDialog(self); dlg.setWindowTitle(f"Item Ledger - {mat}"); dlg.resize(1200,700)
    v=QVBoxLayout(dlg)
    item=self.db.one("""SELECT i.material_id,i.barcode,i.item_description,c.name category,d.name department,i.quantity_available,u.name unit
                       FROM inventory_items i LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN units u ON u.id=i.unit_id
                       WHERE i.material_id=%s""", (mat,)) or {}
    title=QLabel(f"Item Ledger: {mat} - {item.get('item_description','')}")
    title.setStyleSheet("font-size:18px;font-weight:900;color:#071a3a;")
    sub=QLabel(f"Barcode/QR: {item.get('barcode') or '-'} | Current Stock: {item.get('quantity_available',0)} {item.get('unit') or ''} | Category: {item.get('category') or ''} | Department: {item.get('department') or ''}")
    sub.setObjectName("small")
    v.addWidget(title); v.addWidget(sub)
    rows=[]
    for r in self.db.query("SELECT procurement_date date, 'INWARD' type, inward_no ref_no, quantity in_qty, 0 out_qty, 0 adj_qty, total_cost value, remarks, created_at FROM stock_inward WHERE material_id=%s", (mat,)):
        rows.append(r)
    for r in self.db.query("SELECT issue_date date, 'OUTWARD' type, siv_no ref_no, 0 in_qty, quantity_issued out_qty, 0 adj_qty, 0 value, remarks, created_at FROM stock_outward WHERE material_id=%s", (mat,)):
        rows.append(r)
    for r in self.db.query("SELECT adjustment_date date, adjustment_type type, adjustment_no ref_no, 0 in_qty, 0 out_qty, adjustment_quantity adj_qty, 0 value, reason remarks, created_at FROM stock_adjustments WHERE material_id=%s", (mat,)):
        rows.append(r)
    for r in self.db.query("SELECT transfer_date date, 'TRANSFER' type, transfer_no ref_no, 0 in_qty, 0 out_qty, quantity_transferred adj_qty, 0 value, remarks, created_at FROM stock_transfers WHERE material_id=%s", (mat,)):
        rows.append(r)
    rows.sort(key=lambda x: str(x.get('date') or '') + str(x.get('created_at') or ''))
    balance=0.0
    ledger=[]
    for r in rows:
        typ=str(r.get('type') or '')
        inn=float(r.get('in_qty') or 0); out=float(r.get('out_qty') or 0); adj=float(r.get('adj_qty') or 0)
        if typ == 'Decrease': balance -= adj
        elif typ == 'Increase': balance += adj
        elif typ == 'OUTWARD': balance -= out
        elif typ == 'TRANSFER': balance = balance
        else: balance += inn
        rr=dict(r); rr['balance']=round(balance,3); ledger.append(rr)
    table=QTableWidget(); v.addWidget(table,1)
    self.table_fill(table, ledger, [("Date","date"),("Type","type"),("Ref No","ref_no"),("In","in_qty"),("Out","out_qty"),("Adjustment/Transfer Qty","adj_qty"),("Value","value"),("Balance","balance"),("Remarks","remarks")])
    h=QHBoxLayout(); h.addStretch(); ex=QPushButton("Export Ledger"); ex.clicked.connect(lambda:self.export_table(table, f"ledger_{mat}")); close=QPushButton("Close"); close.setObjectName("softBtn"); close.clicked.connect(dlg.accept); h.addWidget(ex); h.addWidget(close); v.addLayout(h)
    dlg.exec()

MainWindow.show_selected_item_ledger = _v19_show_selected_item_ledger


def _v19_export_selected_barcode_label(self):
    mat = self.selected_material_from_inventory()
    if not mat: return
    row = self.db.one("SELECT material_id, barcode, item_description FROM inventory_items WHERE material_id=%s", (mat,))
    if not row: return
    if canvas is None:
        QMessageBox.warning(self,"Missing package","Install reportlab first."); return
    path,_=QFileDialog.getSaveFileName(self,"Save Barcode / QR Label PDF",f"barcode_label_{mat}.pdf","PDF (*.pdf)")
    if not path: return
    if not path.lower().endswith('.pdf'): path += '.pdf'
    try:
        c=canvas.Canvas(path, pagesize=A4); w,h=A4
        y=h-80; c.setFont("Helvetica-Bold",18); c.drawString(50,y,APP_TITLE); y-=35
        c.setFont("Helvetica-Bold",13); c.drawString(50,y,"Inventory Barcode / QR Label"); y-=35
        c.setFont("Helvetica",11)
        c.drawString(50,y,f"Material ID: {row.get('material_id')}"); y-=24
        c.drawString(50,y,f"Barcode/QR: {row.get('barcode') or row.get('material_id')}"); y-=24
        c.drawString(50,y,f"Description: {str(row.get('item_description') or '')[:90]}"); y-=40
        # Text-based QR/barcode fallback: works without extra packages/scanners can use the printed value.
        c.setFont("Courier-Bold",28); c.drawString(50,y, f"*{row.get('barcode') or row.get('material_id')}*")
        y-=55; c.setFont("Helvetica",9); c.drawString(50,y,"Use this value in the Inventory search bar or Material ID/barcode search fields.")
        c.save(); QMessageBox.information(self,"Saved",f"Barcode/QR label PDF saved:\n{path}")
    except Exception as e:
        QMessageBox.critical(self,"Export Failed",str(e))

MainWindow.export_selected_barcode_label = _v19_export_selected_barcode_label


def _v19_edit_selected_inventory(self):
    if not self.require_permission('Inventory','edit'): return
    mat = self.selected_material_from_inventory()
    if not mat: return
    data = self.db.one("""SELECT i.*, c.name category, d.name department, u.name unit, s.name supplier, l.name location, r.name rack
            FROM inventory_items i
            LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id
            LEFT JOIN units u ON u.id=i.unit_id LEFT JOIN suppliers s ON s.id=i.supplier_id
            LEFT JOIN locations l ON l.id=i.location_id LEFT JOIN racks r ON r.id=i.rack_id
            WHERE i.material_id=%s""", (mat,))
    if not data: return
    dlg=QDialog(self); dlg.setWindowTitle(f"Edit Inventory Item - {mat}"); dlg.resize(760,620)
    form=QFormLayout(dlg)
    barcode=QLineEdit(data.get('barcode') or ''); barcode.setPlaceholderText("Barcode / QR value / manufacturer code")
    mpn=QLineEdit(data.get('manufacturer_part_no') or '')
    model=QLineEdit(data.get('model_no') or '')
    desc=QLineEdit(data.get('item_description') or '')
    cat=QComboBox(); cat.addItems(self.master_names('categories')); self.set_combo_text(cat,data.get('category') or '')
    dept=QComboBox(); dept.addItems(self.master_names('departments')); self.set_combo_text(dept,data.get('department') or '')
    stock=QDoubleSpinBox(); stock.setRange(0,9999999); stock.setDecimals(2); stock.setValue(float(data.get('quantity_available') or 0))
    unit=QComboBox(); unit.addItems(self.master_names('units')); self.set_combo_text(unit,data.get('unit') or '')
    minq=QDoubleSpinBox(); minq.setRange(0,9999999); minq.setDecimals(2); minq.setValue(float(data.get('minimum_stock_level') or 0))
    typ=QComboBox(); typ.addItems(MATERIAL_TYPES); self.set_combo_text(typ,data.get('material_type') or '')
    rack=QComboBox(); rack.addItems(self.master_names('racks')); self.set_combo_text(rack,data.get('rack') or '')
    loc=QComboBox(); loc.addItems(self.master_names('locations')); self.set_combo_text(loc,data.get('location') or '')
    supplier=QComboBox(); supplier.addItems(self.master_names('suppliers')); self.set_combo_text(supplier,data.get('supplier') or '')
    cost=QDoubleSpinBox(); cost.setRange(0,9999999); cost.setDecimals(3); cost.setPrefix('OMR '); cost.setValue(float(data.get('cost_per_unit') or 0))
    remarks=QTextEdit(data.get('remarks') or ''); remarks.setFixedHeight(70)
    for lab,wid in [("Material ID",QLabel(mat)),("Barcode / QR",barcode),("Manufacturer Part No.",mpn),("Model No.",model),("Item Description",desc),("Category",cat),("Department",dept),("Stock",stock),("Unit",unit),("Min. Level",minq),("Type",typ),("Rack No.",rack),("Location",loc),("Supplier",supplier),("Cost/Unit",cost),("Remarks",remarks)]: form.addRow(lab,wid)
    buttons=QDialogButtonBox(QDialogButtonBox.Save|QDialogButtonBox.Cancel); form.addRow(buttons)
    buttons.rejected.connect(dlg.reject)
    def save_edit():
        if not desc.text().strip(): QMessageBox.warning(dlg,"Validation","Item description is required."); return
        self.db.execute("""UPDATE inventory_items SET barcode=%s, manufacturer_part_no=%s, model_no=%s, item_description=%s, category_id=%s, department_id=%s, quantity_available=%s, unit_id=%s, minimum_stock_level=%s, material_type=%s, rack_id=%s, location_id=%s, supplier_id=%s, cost_per_unit=%s, remarks=%s, updated_at=CURRENT_TIMESTAMP WHERE material_id=%s""",
            (barcode.text().strip(), mpn.text().strip(), model.text().strip(), desc.text().strip(), self.db.id_by_name('categories',cat.currentText()), self.db.id_by_name('departments',dept.currentText()), float(stock.value()), self.db.id_by_name('units',unit.currentText()), float(minq.value()), typ.currentText(), self.db.id_by_name('racks',rack.currentText()), self.db.id_by_name('locations',loc.currentText()), self.db.id_by_name('suppliers',supplier.currentText()), float(cost.value()), remarks.toPlainText(), mat))
        self.db.log(self.user['id'],"EDIT","Inventory",mat,"Inventory item edited including barcode/QR fields")
        QMessageBox.information(dlg,"Updated","Inventory item updated successfully.")
        dlg.accept(); self.refresh_inventory(); self.refresh_item_combos(); self.refresh_dashboard()
    buttons.accepted.connect(save_edit)
    dlg.exec()

MainWindow.edit_selected_inventory = _v19_edit_selected_inventory


def _v19_inventory_export_rows(self):
    return self.item_rows("")

MainWindow.inventory_export_rows = _v19_inventory_export_rows


def _v19_export_complete_inventory(self):
    rows=self.inventory_export_rows()
    cols=[("Material ID","material_id"),("Barcode / QR","barcode"),("Manufacturer Part No.","manufacturer_part_no"),("Model No.","model_no"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min Level","minimum_stock_level"),("Type","material_type"),("Supplier","supplier"),("Cost/Unit","cost_per_unit"),("Total Value","total_value"),("Expiry","expiry_date"),("Location","location"),("Rack","rack"),("Remarks","remarks"),("Created","created_at"),("Updated","updated_at")]
    temp=QTableWidget(); self.table_fill(temp, rows, cols); self.export_table(temp,"complete_store_inventory")

MainWindow.export_complete_inventory = _v19_export_complete_inventory


def _v19_import_excel(self):
    if not self.require_permission('Inventory','edit'): return
    if pd is None:
        QMessageBox.warning(self,"Missing package","Install pandas and openpyxl first."); return
    path,_=QFileDialog.getOpenFileName(self,"Upload Inventory / Corrected Store Data","","Excel Files (*.xlsx *.xls)")
    if not path: return
    try:
        df=pd.read_excel(path).fillna("")
    except Exception as e:
        QMessageBox.critical(self,"Import Failed",f"Could not read Excel file.\n\n{e}"); return
    df.columns=[str(c).strip().lower().replace("/","_").replace(" ","_") for c in df.columns]
    # Accept both old and new column names.
    aliases = {'barcode___qr':'barcode', 'barcode_qr':'barcode', 'barcode':'barcode', 'mr_no':'purchase_order_no'}
    df = df.rename(columns={k:v for k,v in aliases.items() if k in df.columns})
    required=['item_description','category','department','quantity_available','unit','minimum_stock_level','material_type','cost_per_unit']
    missing=[c for c in required if c not in df.columns]
    if missing:
        QMessageBox.warning(self,"Invalid Excel Format", "Missing required columns:\n" + ", ".join(missing)); return
    if df.empty:
        QMessageBox.warning(self,"Empty File","The selected Excel file has no rows."); return
    preview=ImportPreviewDialog(df, self)
    if preview.exec()!=QDialog.Accepted or not preview.accepted_import: return
    if QMessageBox.question(self, "Confirm Inventory Override", "Uploading inventory will DELETE existing inventory and stock transaction records, then replace them with the uploaded file.\n\nContinue?") != QMessageBox.Yes:
        return
    try:
        for tbl in ["stock_outward_batches", "stock_outward", "inventory_batches", "stock_inward", "stock_adjustments", "stock_transfers", "inventory_items"]:
            self.db.execute(f"DELETE FROM {tbl}")
        self.db.execute("DELETE FROM sqlite_sequence WHERE name IN ('stock_outward_batches','stock_outward','inventory_batches','stock_inward','stock_adjustments','stock_transfers','inventory_items')")
    except Exception as e:
        QMessageBox.critical(self, "Override Failed", f"Could not clear existing inventory/transactions.\n\nReason:\n{e}")
        return
    created=skipped=0; errors=[]
    for idx,row in df.iterrows():
        try:
            desc=str(row.get('item_description','')).strip()
            if not desc:
                skipped+=1; continue
            qty=float(row.get('quantity_available') or 0); minq=float(row.get('minimum_stock_level') or 0); cost=float(row.get('cost_per_unit') or 0)
            cat_id=self._ensure_master('categories', row.get('category'))
            dept_id=self._ensure_master('departments', row.get('department'))
            unit_id=self._ensure_master('units', row.get('unit'))
            supplier_id=self._ensure_master('suppliers', row.get('supplier') or 'Default Supplier')
            rack_id=self._ensure_master('racks', row.get('rack') or 'R-01')
            loc_id=self._ensure_master('locations', row.get('location') or 'Main Store')
            exp=str(row.get('expiry_date','')).strip() or None
            proc=str(row.get('procurement_date','')).strip() or today_iso()
            mat=str(row.get('material_id','')).strip() or self.db.next_code("MAT","inventory_items","material_id")
            barcode=str(row.get('barcode','')).strip()
            mpn=str(row.get('manufacturer_part_no','')).strip()
            model=str(row.get('model_no','')).strip()
            item_id=self.db.execute("""INSERT INTO inventory_items(material_id,barcode,manufacturer_part_no,model_no,item_description,category_id,department_id,quantity_available,unit_id,minimum_stock_level,material_type,supplier_id,cost_per_unit,procurement_date,purchase_order_no,delivery_order_no,warranty_details,expiry_date,rack_id,location_id,remarks)
                        VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                        (mat,barcode,mpn,model,desc,cat_id,dept_id,qty,unit_id,minq,str(row.get('material_type','General Spare')),supplier_id,cost,proc,str(row.get('purchase_order_no','')),str(row.get('delivery_order_no','')),str(row.get('warranty_details','')),exp,rack_id,loc_id,str(row.get('remarks','Excel import'))))
            inward_no=self.db.next_code("IN","stock_inward","inward_no")
            inward_id = self.db.execute("INSERT INTO stock_inward(inward_no,material_id,item_id,quantity,unit_cost,total_cost,procurement_date,supplier_id,po_no,do_no,remarks,created_by) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",
                    (inward_no,mat,item_id,qty,cost,qty*cost,proc,supplier_id,str(row.get('purchase_order_no','')),str(row.get('delivery_order_no','')),"Excel stock upload",self.user['id']))
            if exp and qty > 0:
                self.create_expiry_batch(inward_id, item_id, mat, qty, exp)
            created+=1
        except Exception as e:
            errors.append(f"Row {idx+2}: {e}")
    self.db.log(self.user['id'],"IMPORT","Inventory","Excel Upload",f"Override import created {created}, skipped {skipped}")
    msg=f"Inventory override completed.\n\nCreated materials: {created}\nSkipped blank rows: {skipped}"
    if errors: msg += "\n\nErrors:\n" + "\n".join(errors[:10])
    QMessageBox.information(self,"Import Completed",msg); self.refresh_inventory(); self.refresh_item_combos(); self.refresh_master_dropdowns(); self.refresh_dashboard()

MainWindow.import_excel = _v19_import_excel


def _v19_delete_selected_inventory(self):
    if not self.require_permission('Inventory','delete'): return
    return _ORIG_DELETE_SELECTED_INVENTORY(self)
MainWindow.delete_selected_inventory = _v19_delete_selected_inventory


def _v19_save_inward(self):
    if not self.require_permission('Stock Inward','add'): return
    return _ORIG_SAVE_INWARD(self)
MainWindow.save_inward = _v19_save_inward


def _v19_save_outward(self):
    if not self.require_permission('Stock Outward','add'): return
    return _ORIG_SAVE_OUTWARD(self)
MainWindow.save_outward = _v19_save_outward


def _v19_save_adjustment(self):
    if not self.require_permission('Stock Adjustment','add'): return
    return _ORIG_SAVE_ADJUSTMENT(self)
MainWindow.save_adjustment = _v19_save_adjustment


def _v19_save_transfer(self):
    if not self.require_permission('Transfers','add'): return
    return _ORIG_SAVE_TRANSFER(self)
MainWindow.save_transfer = _v19_save_transfer


def _v19_generate_report(self):
    # Existing report logic is preserved. We only add barcode-capable inventory columns by patching the base table after generation when applicable.
    return MainWindow.__dict__.get('_v18_generate_report_original', None)(self) if False else _ORIG_GENERATE_REPORT(self)

# Keep original report generation but add barcode to inventory-style base reports by replacing generate_report fully.
_ORIG_GENERATE_REPORT = MainWindow.generate_report


def _v19_generate_report_full(self):
    typ=self.rep_type.currentText()
    base_cols=[("Material ID","material_id"),("Barcode / QR","barcode"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min. Level","minimum_stock_level"),("Type","material_type"),("Supplier","supplier"),("Rack No.","rack"),("Location","location"),("Cost/Unit (OMR)","cost_per_unit"),("Total Value (OMR)","total_value")]
    if typ in ["Balance Stock Report", "Complete Inventory Report", "Department-wise Stock Report", "Category-wise Stock Report", "Supplier-wise Purchase Report", "Store Location-wise Stock Report"]:
        rows=self._filtered_inventory_rows_for_reports(); cols=base_cols
    else:
        # Use original report generation for transaction/expiry reports, then return.
        return _ORIG_GENERATE_REPORT(self)
    chosen = self.choose_report_columns(cols)
    if chosen is None: return
    self.table_fill(self.rep_table, rows, chosen)

MainWindow.generate_report = _v19_generate_report_full


def _v19_export_table(self, table: QTableWidget, basename: str):
    path,_=QFileDialog.getSaveFileName(self,"Export",f"{basename}_{datetime.now().strftime('%Y%m%d_%H%M')}","Excel (*.xlsx);;CSV (*.csv);;PDF (*.pdf)")
    if not path: return
    headers=[table.horizontalHeaderItem(c).text() for c in range(table.columnCount())]
    rows=[[table.item(r,c).text() if table.item(r,c) else "" for c in range(table.columnCount())] for r in range(table.rowCount())]
    if path.endswith('.xlsx'):
        if pd is None: QMessageBox.warning(self,"Missing package","Install pandas openpyxl."); return
        pd.DataFrame(rows, columns=headers).to_excel(path,index=False)
    elif path.endswith('.pdf'):
        if canvas is None: QMessageBox.warning(self,"Missing package","Install reportlab."); return
        try:
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet
            if not path.lower().endswith('.pdf'): path += '.pdf'
            doc=SimpleDocTemplate(path, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=24, bottomMargin=24)
            styles=getSampleStyleSheet(); elems=[]
            elems.append(Paragraph(f"<b>{APP_TITLE}</b>", styles['Title']))
            elems.append(Paragraph(f"<b>Report:</b> {basename.replace('_',' ').title()} &nbsp;&nbsp; <b>Generated:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')} &nbsp;&nbsp; <b>User:</b> {self.user.get('username','')}", styles['Normal']))
            if hasattr(self, 'rep_type'):
                elems.append(Paragraph(f"<b>Filters:</b> Report={self.rep_type.currentText()}, Department={getattr(self,'rep_department').currentText()}, Category={getattr(self,'rep_category').currentText()}, Supplier={getattr(self,'rep_supplier').currentText()}", styles['Normal']))
            elems.append(Spacer(1,10))
            data=[headers]+rows
            # Trim very wide text for printable PDF but keep Excel/CSV full.
            data=[[str(x)[:38] for x in row] for row in data]
            t=Table(data, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#102a43')),('TEXTCOLOR',(0,0),(-1,0),colors.white),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7),
                ('GRID',(0,0),(-1,-1),0.25,colors.HexColor('#d5dde8')),('VALIGN',(0,0),(-1,-1),'TOP'),
                ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, colors.HexColor('#f5f7fb')])
            ]))
            elems.append(t); elems.append(Spacer(1,20)); elems.append(Paragraph("Prepared by: ____________________ &nbsp;&nbsp;&nbsp;&nbsp; Checked by: ____________________ &nbsp;&nbsp;&nbsp;&nbsp; Approved by: ____________________", styles['Normal']))
            doc.build(elems)
        except Exception as e:
            QMessageBox.critical(self,"PDF Export Failed",str(e)); return
    else:
        if not path.endswith('.csv'): path += '.csv'
        with open(path,'w',newline='',encoding='utf-8') as f: wr=csv.writer(f); wr.writerow(headers); wr.writerows(rows)
    QMessageBox.information(self,"Exported",f"Saved: {path}")

MainWindow.export_table = _v19_export_table


def _v19_page_settings(self):
    if not self.has_permission('Settings','view'):
        w,v=self.base_page("Settings")
        v.addWidget(QLabel("You do not have permission to view Settings.")); return w
    w,v=self.base_page("Settings"); tabs=QTabWidget(); v.addWidget(tabs)
    for table,title in [("departments","Departments"),("categories","Categories"),("units","Units"),("suppliers","Suppliers"),("locations","Locations"),("racks","Racks")]: tabs.addTab(self.master_tab(table), title)
    backup=QWidget(); bl=QVBoxLayout(backup)
    b1=QPushButton("Backup Database"); b1.setObjectName("softBtn"); b1.clicked.connect(self.backup_db)
    b2=QPushButton("Restore Database"); b2.setObjectName("orangeBtn"); b2.clicked.connect(self.restore_db)
    bl.addWidget(QLabel("Use Backup before major uploads or corrections.")); bl.addWidget(b1); bl.addWidget(b2); bl.addStretch(); tabs.addTab(backup,"Backup / Restore")
    # Audit Log tab
    audit=QWidget(); av=QVBoxLayout(audit)
    at=QTableWidget(); av.addWidget(at)
    rows=self.db.query("SELECT a.created_at,u.username,a.action,a.module,a.record_ref,a.details FROM audit_logs a LEFT JOIN users u ON u.id=a.user_id ORDER BY a.id DESC LIMIT 500")
    self.table_fill(at, rows, [("Date","created_at"),("User","username"),("Action","action"),("Module","module"),("Record","record_ref"),("Details","details")])
    ex=QPushButton("Export Audit Logs"); ex.clicked.connect(lambda:self.export_table(at,"audit_logs")); av.addWidget(ex)
    tabs.addTab(audit,"Audit Logs")
    # User/Roles view tab
    ur=QWidget(); uv=QVBoxLayout(ur)
    note=QLabel("Role-based permissions are active. Default roles: Admin, Store Keeper, Manager, Viewer."); note.setObjectName('small'); uv.addWidget(note)
    ut=QTableWidget(); uv.addWidget(ut)
    rows=self.db.query("SELECT u.username,u.full_name,r.name role,u.is_active,u.created_at FROM users u LEFT JOIN roles r ON r.id=u.role_id ORDER BY u.username")
    self.table_fill(ut, rows, [("Username","username"),("Full Name","full_name"),("Role","role"),("Active","is_active"),("Created","created_at")])
    tabs.addTab(ur,"Users / Roles")
    return w

MainWindow.page_settings = _v19_page_settings

# ======================= END MARKET-GRADE UPGRADE PATCH V19 =======================


# ======================= V20 VISIBLE MARKET UPGRADE FIX =======================
# This final patch makes the V19 upgrades visible in the main navigation and UI.

_V20_ORIG_BUILD_SIDEBAR = MainWindow.build_sidebar
_V20_ORIG_SET_PAGE = MainWindow.set_page


def _v20_build_sidebar(self) -> QWidget:
    side = QFrame(); side.setObjectName("sidebar"); side.setFixedWidth(400)
    v = QVBoxLayout(side); v.setContentsMargins(16,22,16,18); v.setSpacing(6)
    title = QLabel("▥  STORE INVENTORY MANAGEMENT"); title.setObjectName("sideTitle"); title.setWordWrap(True); title.setMinimumWidth(360)
    sub = QLabel("DESIGNED & DEVELOPED - PAVAN KUMAR AKELLA"); sub.setObjectName("sideSub"); sub.setWordWrap(True); sub.setMinimumWidth(360)
    v.addWidget(title); v.addWidget(sub); v.addSpacing(18)
    sections = [
        ("MAIN", [("⌂", "Dashboard")]),
        ("STORE MANAGEMENT", [("↓", "Stock Inward"),("↑", "Stock Outward"),("↕", "Stock Adjustment"),("⇄", "Transfers")]),
        ("REPORTS", [("☷", "Reports")]),
        ("INVENTORY", [("▦", "Inventory"),("📒", "Item Ledger")]),
        ("CONTROL", [("⚙", "Settings"),("🧾", "Audit Logs")])
    ]
    page_index = 0
    self.nav_buttons = []
    for sec, items in sections:
        lab = QLabel(sec); lab.setObjectName("sectionLabel"); v.addWidget(lab)
        for icon, item in items:
            b = QPushButton(f"  {icon}   {item}"); b.setObjectName("nav"); b.setProperty("active", False)
            b.clicked.connect(lambda _, i=page_index: self.set_page(i))
            self.nav_buttons.append(b); v.addWidget(b); page_index += 1
    v.addStretch(); out = QPushButton("  ↪   Logout"); out.setObjectName("nav"); out.clicked.connect(self.close); v.addWidget(out)
    return side


MainWindow.build_sidebar = _v20_build_sidebar


def _v20_mainwindow_init(self, db: Database, user: Dict[str, Any]):
    QMainWindow.__init__(self)
    self.db = db; self.user = user
    self.setWindowTitle(APP_TITLE); self.resize(1500, 920)
    self.nav_buttons: List[QPushButton] = []
    root = QWidget(); self.setCentralWidget(root)
    h = QHBoxLayout(root); h.setContentsMargins(0,0,0,0); h.setSpacing(0)
    self.sidebar = self.build_sidebar(); h.addWidget(self.sidebar)
    self.stack = QStackedWidget(); h.addWidget(self.stack, 1)
    self.pages = [
        ("Dashboard", self.page_dashboard),
        ("Stock Inward", self.page_inward),
        ("Stock Outward", self.page_outward),
        ("Stock Adjustment", self.page_adjustment),
        ("Transfers", self.page_transfer),
        ("Reports", self.page_reports),
        ("Inventory", self.page_inventory),
        ("Item Ledger", self.page_item_ledger),
        ("Settings", self.page_settings),
        ("Audit Logs", self.page_audit_logs),
    ]
    for name, builder in self.pages:
        try:
            self.stack.addWidget(builder())
        except Exception as e:
            page = QWidget(); lay = QVBoxLayout(page); lay.addWidget(QLabel(f"Could not load {name}: {e}")); self.stack.addWidget(page)
    self.set_page(0)


MainWindow.__init__ = _v20_mainwindow_init


def _v20_set_page(self, i: int):
    self.stack.setCurrentIndex(i)
    for n,b in enumerate(self.nav_buttons):
        b.setProperty("active", n==i); b.style().unpolish(b); b.style().polish(b)
    name = self.pages[i][0] if hasattr(self, 'pages') and i < len(self.pages) else ""
    if name == "Dashboard" and hasattr(self, "refresh_dashboard"):
        self.refresh_dashboard()
    elif name == "Stock Inward" and hasattr(self, "clear_inward_form"):
        self.clear_inward_form(); self.refresh_inward_table()
    elif name == "Stock Outward" and hasattr(self, "clear_outward_form"):
        self.clear_outward_form(); self.refresh_outward_table()
    elif name == "Stock Adjustment" and hasattr(self, "clear_adjustment_form"):
        self.clear_adjustment_form(); self.refresh_adjustment_table()
    elif name == "Transfers" and hasattr(self, "clear_transfer_form"):
        self.clear_transfer_form(); self.refresh_transfer_table()
    elif name == "Reports" and hasattr(self, "generate_report"):
        self.generate_report()
    elif name == "Inventory" and hasattr(self, "refresh_inventory"):
        self.refresh_inventory()
    elif name == "Audit Logs" and hasattr(self, "refresh_audit_logs"):
        self.refresh_audit_logs()


MainWindow.set_page = _v20_set_page


def _v20_page_audit_logs(self):
    w, v = self.base_page("Audit Logs")
    if hasattr(self, 'has_permission') and not self.has_permission('Audit Logs','view'):
        v.addWidget(QLabel("You do not have permission to view Audit Logs.")); return w
    bar = QHBoxLayout()
    self.audit_search = QLineEdit(); self.audit_search.setPlaceholderText("Search action, module, material ID, user, details...")
    self.audit_search.returnPressed.connect(self.refresh_audit_logs)
    b = QPushButton("Search"); b.clicked.connect(self.refresh_audit_logs)
    ex = QPushButton("Export Audit Logs"); ex.setObjectName("softBtn"); ex.clicked.connect(lambda: self.export_table(self.audit_table, "audit_logs"))
    bar.addWidget(self.audit_search, 1); bar.addWidget(b); bar.addWidget(ex)
    v.addLayout(bar)
    self.audit_table = QTableWidget(); v.addWidget(self.audit_table, 1)
    self.refresh_audit_logs()
    return w


def _v20_refresh_audit_logs(self):
    search = getattr(self, 'audit_search', None).text().strip() if hasattr(self, 'audit_search') else ""
    q = """SELECT a.created_at,u.username,a.action,a.module,a.record_ref,a.details
           FROM audit_logs a LEFT JOIN users u ON u.id=a.user_id"""
    params = ()
    if search:
        like = f"%{search}%"
        q += " WHERE a.action LIKE %s OR a.module LIKE %s OR a.record_ref LIKE %s OR a.details LIKE %s OR u.username LIKE %s"
        params = (like, like, like, like, like)
    q += " ORDER BY a.id DESC LIMIT 1000"
    rows = self.db.query(q, params)
    self.table_fill(self.audit_table, rows, [("Date","created_at"),("User","username"),("Action","action"),("Module","module"),("Record","record_ref"),("Details","details")])
    self.audit_table.setColumnWidth(0, 170); self.audit_table.setColumnWidth(1, 130); self.audit_table.setColumnWidth(2, 110); self.audit_table.setColumnWidth(3, 150); self.audit_table.setColumnWidth(4, 160); self.audit_table.setColumnWidth(5, 520)


MainWindow.page_audit_logs = _v20_page_audit_logs
MainWindow.refresh_audit_logs = _v20_refresh_audit_logs


def _v20_page_item_ledger(self):
    w, v = self.base_page("Item Ledger")
    card = QFrame(); card.setObjectName("card"); cl = QVBoxLayout(card); cl.setContentsMargins(16,14,16,14)
    row = QHBoxLayout()
    self.ledger_search = QLineEdit(); self.ledger_search.setPlaceholderText("Search/select material ID, barcode/QR, or item description...")
    self.ledger_search.returnPressed.connect(self.refresh_item_ledger_page)
    self.ledger_combo = QComboBox(); self.ledger_combo.setEditable(True); self.ledger_combo.setInsertPolicy(QComboBox.NoInsert); self.ledger_combo.setMinimumWidth(520)
    self.refresh_ledger_combo()
    gen = QPushButton("Generate Ledger"); gen.clicked.connect(self.refresh_item_ledger_page)
    exp = QPushButton("Export Ledger"); exp.setObjectName("softBtn"); exp.clicked.connect(lambda: self.export_table(self.ledger_table, "item_ledger"))
    row.addWidget(self.ledger_search, 1); row.addWidget(self.ledger_combo, 1); row.addWidget(gen); row.addWidget(exp)
    cl.addLayout(row)
    self.ledger_summary = QLabel("Select one item and click Generate Ledger."); self.ledger_summary.setObjectName("small"); cl.addWidget(self.ledger_summary)
    v.addWidget(card)
    self.ledger_table = QTableWidget(); v.addWidget(self.ledger_table, 1)
    return w


def _v20_refresh_ledger_combo(self):
    if not hasattr(self, 'ledger_combo'): return
    self.ledger_combo.blockSignals(True); self.ledger_combo.clear(); self.ledger_combo.addItem("-- Select Material --", None)
    for r in self.item_rows(getattr(self, 'ledger_search', None).text().strip() if hasattr(self, 'ledger_search') else "")[:500]:
        label = f"{r.get('material_id')} | {r.get('barcode') or '-'} | {r.get('item_description')} | Stock: {r.get('quantity_available')}"
        self.ledger_combo.addItem(label, r.get('material_id'))
    self.ledger_combo.blockSignals(False)


def _v20_refresh_item_ledger_page(self):
    self.refresh_ledger_combo()
    mat = self.ledger_combo.currentData()
    if not mat:
        # Try direct typed material/barcode search.
        key = self.ledger_search.text().strip()
        if key:
            r = self.db.one("SELECT material_id FROM inventory_items WHERE material_id=%s OR barcode=%s OR item_description LIKE %s LIMIT 1", (key, key, f"%{key}%"))
            mat = r['material_id'] if r else None
    if not mat:
        QMessageBox.warning(self, "Selection", "Select or search one material first."); return
    item = self.db.one("""SELECT i.*, c.name category, d.name department, u.name unit, r.name rack, l.name location
                           FROM inventory_items i
                           LEFT JOIN categories c ON c.id=i.category_id
                           LEFT JOIN departments d ON d.id=i.department_id
                           LEFT JOIN units u ON u.id=i.unit_id
                           LEFT JOIN racks r ON r.id=i.rack_id
                           LEFT JOIN locations l ON l.id=i.location_id
                           WHERE i.material_id=%s""", (mat,)) or {}
    rows=[]
    balance=0.0
    # Opening/current baseline is shown first when no transaction history is available.
    tx=[]
    for r in self.db.query("SELECT procurement_date AS dt, inward_no AS ref, quantity AS in_qty, 0 AS out_qty, 0 AS adj_qty, 'Inward' AS type, remarks FROM stock_inward WHERE material_id=%s ORDER BY procurement_date, id", (mat,)):
        tx.append(r)
    for r in self.db.query("SELECT issue_date AS dt, siv_no AS ref, 0 AS in_qty, quantity_issued AS out_qty, 0 AS adj_qty, 'Outward' AS type, remarks FROM stock_outward WHERE material_id=%s ORDER BY issue_date, id", (mat,)):
        tx.append(r)
    for r in self.db.query("SELECT adjustment_date AS dt, adjustment_no AS ref, 0 AS in_qty, 0 AS out_qty, CASE WHEN adjustment_type='Increase' THEN adjustment_quantity ELSE -adjustment_quantity END AS adj_qty, 'Adjustment' AS type, reason AS remarks FROM stock_adjustments WHERE material_id=%s ORDER BY adjustment_date, id", (mat,)):
        tx.append(r)
    for r in self.db.query("SELECT transfer_date AS dt, transfer_no AS ref, 0 AS in_qty, 0 AS out_qty, 0 AS adj_qty, 'Transfer' AS type, remarks FROM stock_transfers WHERE material_id=%s ORDER BY transfer_date, id", (mat,)):
        tx.append(r)
    tx.sort(key=lambda x: str(x.get('dt') or ''))
    for r in tx:
        inq=float(r.get('in_qty') or 0); outq=float(r.get('out_qty') or 0); adj=float(r.get('adj_qty') or 0)
        balance += inq - outq + adj
        rows.append({'date':r.get('dt'), 'type':r.get('type'), 'ref':r.get('ref'), 'in_qty':inq if inq else '', 'out_qty':outq if outq else '', 'adjustment':adj if adj else '', 'balance':balance, 'remarks':r.get('remarks')})
    if not rows:
        rows.append({'date':'-', 'type':'Current Balance', 'ref':mat, 'in_qty':'', 'out_qty':'', 'adjustment':'', 'balance':item.get('quantity_available',0), 'remarks':'No transaction history found'})
    self.ledger_summary.setText(f"{mat} | {item.get('item_description','')} | Barcode/QR: {item.get('barcode') or '-'} | Current Stock: {item.get('quantity_available',0)} {item.get('unit') or ''}")
    self.table_fill(self.ledger_table, rows, [("Date","date"),("Type","type"),("Ref No.","ref"),("In","in_qty"),("Out","out_qty"),("Adjustment","adjustment"),("Running Balance","balance"),("Remarks","remarks")])
    self.ledger_table.setColumnWidth(0, 130); self.ledger_table.setColumnWidth(1, 130); self.ledger_table.setColumnWidth(2, 160); self.ledger_table.setColumnWidth(7, 420)


MainWindow.page_item_ledger = _v20_page_item_ledger
MainWindow.refresh_ledger_combo = _v20_refresh_ledger_combo
MainWindow.refresh_item_ledger_page = _v20_refresh_item_ledger_page

# Add barcode/QR field to Stock Inward for new materials and existing item visibility.
_V20_ORIG_PAGE_INWARD = MainWindow.page_inward
_V20_ORIG_SAVE_INWARD = MainWindow.save_inward
_V20_ORIG_LOAD_EXISTING_INWARD = MainWindow.load_existing_for_inward


def _v20_page_inward(self):
    page = _V20_ORIG_PAGE_INWARD(self)
    # Place a Barcode/QR input above remarks if the generated form exists.
    try:
        if not hasattr(self, 'in_barcode'):
            self.in_barcode = QLineEdit(); self.in_barcode.setPlaceholderText("Barcode / QR / manufacturer code")
            # Add visibly near top of the form by finding the first QGroupBox layout.
            group = page.findChild(QGroupBox)
            lay = group.layout() if group else None
            if isinstance(lay, QGridLayout):
                label = QLabel("Barcode / QR"); label.setStyleSheet("font-weight:800;color:#263b5e;")
                lay.addWidget(label, 10, 0); lay.addWidget(self.in_barcode, 10, 1)
    except Exception:
        pass
    return page


def _v20_load_existing_for_inward(self):
    _V20_ORIG_LOAD_EXISTING_INWARD(self)
    try:
        mat = self.in_existing.currentData()
        r = self.db.one("SELECT barcode FROM inventory_items WHERE material_id=%s", (mat,)) if mat else None
        if hasattr(self, 'in_barcode'):
            self.in_barcode.setText((r or {}).get('barcode') or '')
    except Exception:
        pass


def _v20_save_inward(self):
    barcode_value = getattr(self, 'in_barcode', None).text().strip() if hasattr(self, 'in_barcode') else ''
    before_mat = self.in_existing.currentData() if hasattr(self, 'in_existing') else None
    result = _V20_ORIG_SAVE_INWARD(self)
    # If a barcode was entered, save it against the existing material or the newest item inserted.
    if barcode_value:
        try:
            mat = before_mat
            if not mat:
                row = self.db.one("SELECT material_id FROM inventory_items ORDER BY id DESC LIMIT 1")
                mat = row['material_id'] if row else None
            if mat:
                self.db.execute("UPDATE inventory_items SET barcode=%s, updated_at=CURRENT_TIMESTAMP WHERE material_id=%s", (barcode_value, mat))
        except Exception:
            pass
    return result

MainWindow.page_inward = _v20_page_inward
MainWindow.load_existing_for_inward = _v20_load_existing_for_inward
MainWindow.save_inward = _v20_save_inward

# Make feature presence obvious in Settings too.
_V20_ORIG_PAGE_SETTINGS = MainWindow.page_settings

def _v20_page_settings_visible(self):
    w = _V20_ORIG_PAGE_SETTINGS(self)
    try:
        layout = w.layout()
        banner = QLabel("Active modules: Item Ledger | Audit Logs | Barcode/QR Search & Labels | Role Permissions | Professional PDF Reports")
        banner.setStyleSheet("background:#eaf1fb;color:#102a43;border:1px solid #cbd8ea;border-radius:8px;padding:10px;font-weight:800;")
        layout.insertWidget(0, banner)
    except Exception:
        pass
    return w

MainWindow.page_settings = _v20_page_settings_visible

# ======================= END V20 VISIBLE MARKET UPGRADE FIX =======================


# ======================= V21 BARCODE SCANNER WORKFLOW UPGRADE =======================
# Adds full barcode workflow for wireless scanners:
# - generate missing barcodes for all items
# - print barcode labels for selected/all items
# - scan/enter barcode in inward/outward/adjustment/transfer
# - inventory upload/download retains barcode fields
# - duplicate barcode detection warning

def _v21_generate_barcode_value(material_id: str) -> str:
    safe = ''.join(ch for ch in str(material_id or '').upper() if ch.isalnum() or ch in ['-', '_'])
    return f"SIM-{safe}" if safe else f"SIM-{datetime.now().strftime('%Y%m%d%H%M%S')}"


def _v21_unique_barcode(self, material_id: str) -> str:
    base = _v21_generate_barcode_value(material_id)
    code = base
    n = 1
    while self.db.one("SELECT material_id FROM inventory_items WHERE barcode=%s AND material_id<>%s AND is_active=1", (code, material_id)):
        n += 1
        code = f"{base}-{n}"
    return code


def _v21_find_material_by_barcode(self, code: str) -> Optional[Dict[str, Any]]:
    code = str(code or '').strip()
    if not code:
        return None
    return self.db.one("""SELECT material_id, barcode, item_description, quantity_available
                       FROM inventory_items
                       WHERE is_active=1 AND (barcode=%s OR material_id=%s)
                       LIMIT 1""", (code, code))


def _v21_select_material_in_combo(self, combo: QComboBox, material_id: str) -> bool:
    if not combo or not material_id:
        return False
    idx = combo.findData(material_id)
    if idx < 0:
        try:
            self.refresh_item_combos()
        except Exception:
            pass
        idx = combo.findData(material_id)
    if idx >= 0:
        combo.setCurrentIndex(idx)
        return True
    return False


def _v21_generate_missing_barcodes(self, silent: bool = False):
    try:
        rows = self.db.query("SELECT material_id FROM inventory_items WHERE is_active=1 AND (barcode IS NULL OR TRIM(barcode)='') ORDER BY material_id")
        count = 0
        for r in rows:
            mat = r.get('material_id')
            if not mat:
                continue
            self.db.execute("UPDATE inventory_items SET barcode=%s, updated_at=CURRENT_TIMESTAMP WHERE material_id=%s", (_v21_unique_barcode(self, mat), mat))
            count += 1
        if hasattr(self, 'inv_table'):
            self.refresh_inventory()
        if not silent:
            QMessageBox.information(self, "Barcode Generation", f"Missing barcode generation completed.\n\nItems updated: {count}")
        return count
    except Exception as e:
        if not silent:
            QMessageBox.critical(self, "Barcode Generation Failed", str(e))
        return 0

MainWindow.generate_missing_barcodes = _v21_generate_missing_barcodes


def _v21_check_duplicate_barcodes(self, show_ok: bool = False):
    rows = self.db.query("""SELECT barcode, COUNT(*) cnt, GROUP_CONCAT(material_id, ', ') materials
                          FROM inventory_items
                          WHERE is_active=1 AND barcode IS NOT NULL AND TRIM(barcode)<>''
                          GROUP BY barcode HAVING COUNT(*)>1""")
    if rows:
        msg = "Duplicate barcode values found. Scanning will be unsafe until corrected.\n\n"
        msg += "\n".join([f"{r['barcode']}: {r['materials']}" for r in rows[:20]])
        QMessageBox.warning(self, "Duplicate Barcodes", msg)
        return False
    if show_ok:
        QMessageBox.information(self, "Barcode Check", "No duplicate barcodes found.")
    return True

MainWindow.check_duplicate_barcodes = _v21_check_duplicate_barcodes


# Ensure imported / existing data gets barcode values automatically after startup seed/migration.
_V21_ORIG_DB_SEED_DEFAULTS = Database.seed_defaults

def _v21_db_seed_defaults(self):
    _V21_ORIG_DB_SEED_DEFAULTS(self)
    try:
        rows = self.query("SELECT material_id FROM inventory_items WHERE is_active=1 AND (barcode IS NULL OR TRIM(barcode)='')")
        for r in rows:
            mat = r.get('material_id')
            if mat:
                base = _v21_generate_barcode_value(mat)
                code = base; n = 1
                while self.one("SELECT material_id FROM inventory_items WHERE barcode=%s AND material_id<>%s AND is_active=1", (code, mat)):
                    n += 1; code = f"{base}-{n}"
                self.execute("UPDATE inventory_items SET barcode=%s WHERE material_id=%s", (code, mat))
    except Exception:
        pass

Database.seed_defaults = _v21_db_seed_defaults


# Inventory page: add clear barcode workflow buttons visibly.
_V21_ORIG_PAGE_INVENTORY = MainWindow.page_inventory

def _v21_page_inventory(self):
    page = _V21_ORIG_PAGE_INVENTORY(self)
    try:
        layout = page.layout()
        bar = QFrame(); bar.setObjectName("card")
        hl = QHBoxLayout(bar); hl.setContentsMargins(14,10,14,10); hl.setSpacing(8)
        note = QLabel("Barcode workflow: generate missing barcodes, print labels, then scan items in inward/outward forms.")
        note.setObjectName("small"); hl.addWidget(note, 1)
        gen = QPushButton("Generate Missing Barcodes"); gen.setObjectName("greenBtn"); gen.clicked.connect(lambda: self.generate_missing_barcodes(False))
        chk = QPushButton("Check Duplicate Barcodes"); chk.setObjectName("softBtn"); chk.clicked.connect(lambda: self.check_duplicate_barcodes(True))
        all_labels = QPushButton("Print All Barcode Labels"); all_labels.setObjectName("softBtn"); all_labels.clicked.connect(self.export_all_barcode_labels)
        sel_labels = QPushButton("Print Selected Barcode Label"); sel_labels.setObjectName("softBtn"); sel_labels.clicked.connect(self.export_selected_barcode_label)
        for b in [gen, chk, all_labels, sel_labels]:
            b.setMinimumWidth(180); hl.addWidget(b)
        # Insert below the existing inventory button row but before the table.
        layout.insertWidget(2, bar)
    except Exception:
        pass
    return page

MainWindow.page_inventory = _v21_page_inventory


# Barcode PDF label helpers.
def _v23_wrap_label_text(text: str, max_chars: int = 34, max_lines: int = 2) -> List[str]:
    """Small word-wrap helper for fixed barcode labels."""
    words = str(text or "").replace("\n", " ").split()
    lines: List[str] = []
    current = ""
    for word in words:
        candidate = (current + " " + word).strip()
        if len(candidate) <= max_chars:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word[:max_chars]
        if len(lines) >= max_lines:
            break
    if current and len(lines) < max_lines:
        lines.append(current)
    return lines[:max_lines]


def _v21_draw_code128_label(c, row, x, y, w, h):
    """Draw one fixed 70mm x 35mm cut-and-paste label.

    Layout is locked for A4 3 columns x 8 rows. Coordinates are exact, so the
    PDF can be printed, cut along the guide borders, and pasted on racks/items.
    """
    from reportlab.lib.units import mm
    barcode_value = str(row.get('barcode') or row.get('material_id') or '').strip()
    material_id = str(row.get('material_id') or '').strip()
    desc = str(row.get('item_description') or '').strip()
    rack = str(row.get('rack') or '').strip()
    loc = str(row.get('location') or '').strip()

    # Outer cut guide. Light grey is visible but not visually heavy.
    try:
        c.setStrokeColorRGB(0.72, 0.72, 0.72)
        c.setLineWidth(0.35)
    except Exception:
        pass
    c.rect(x, y, w, h, stroke=1, fill=0)

    pad_x = 3.0 * mm
    top_y = y + h - 4.0 * mm

    # Header
    try:
        c.setFillColorRGB(0.05, 0.10, 0.20)
    except Exception:
        pass
    c.setFont("Helvetica-Bold", 6.5)
    c.drawCentredString(x + w / 2, top_y, "STORE INVENTORY MANAGEMENT")

    # Barcode centered inside fixed barcode zone.
    barcode_drawn = False
    try:
        from reportlab.graphics.barcode import code128
        barcode_max_w = w - (2 * pad_x)
        # Code128 supports scalable barWidth. Start with readable width and shrink if needed.
        bar_width = 0.30 * mm
        bc = code128.Code128(barcode_value, barHeight=11.0 * mm, barWidth=bar_width, humanReadable=False)
        while bc.width > barcode_max_w and bar_width > 0.16 * mm:
            bar_width -= 0.02 * mm
            bc = code128.Code128(barcode_value, barHeight=11.0 * mm, barWidth=bar_width, humanReadable=False)
        bx = x + (w - bc.width) / 2
        by = y + 16.5 * mm
        bc.drawOn(c, bx, by)
        barcode_drawn = True
    except Exception:
        barcode_drawn = False

    if not barcode_drawn:
        # Fallback keeps label usable even if barcode renderer fails.
        c.setFont("Courier-Bold", 15)
        c.drawCentredString(x + w / 2, y + 21.0 * mm, f"*{barcode_value}*")

    # Barcode value / material ID line.
    try:
        c.setFillColorRGB(0.00, 0.00, 0.00)
    except Exception:
        pass
    c.setFont("Helvetica-Bold", 6.8)
    c.drawCentredString(x + w / 2, y + 14.0 * mm, barcode_value[:38])

    c.setFont("Helvetica-Bold", 6.6)
    c.drawString(x + pad_x, y + 10.2 * mm, f"ID: {material_id[:24]}")

    # Description wraps in two lines, clipped to stay inside label.
    c.setFont("Helvetica", 5.8)
    desc_lines = _v23_wrap_label_text(desc, max_chars=44, max_lines=2)
    line_y = y + 7.6 * mm
    for line in desc_lines:
        c.drawString(x + pad_x, line_y, line)
        line_y -= 2.7 * mm

    # Rack/location footer.
    footer = ""
    if rack:
        footer += f"Rack: {rack}"
    if loc:
        footer += ("  |  " if footer else "") + f"Loc: {loc}"
    c.setFont("Helvetica-Bold", 5.6)
    c.drawString(x + pad_x, y + 1.8 * mm, footer[:52])


def _v21_export_barcode_labels(self, rows: List[Dict[str, Any]], default_name: str):
    if canvas is None:
        QMessageBox.warning(self, "Missing package", "ReportLab is required for PDF label export.")
        return
    if not rows:
        QMessageBox.warning(self, "No Items", "No inventory items found for barcode label printing.")
        return
    self.generate_missing_barcodes(True)
    # Re-read rows after barcode generation so blanks are filled.
    mats = [r.get('material_id') for r in rows if r.get('material_id')]
    if mats:
        placeholders = ','.join(['?']*len(mats))
        rows = self.db.query(f"""SELECT i.material_id,i.barcode,i.item_description,r.name rack,l.name location
                              FROM inventory_items i
                              LEFT JOIN racks r ON r.id=i.rack_id LEFT JOIN locations l ON l.id=i.location_id
                              WHERE i.material_id IN ({placeholders}) AND i.is_active=1
                              ORDER BY i.material_id""", tuple(mats))
    path,_ = QFileDialog.getSaveFileName(self, "Save Barcode Labels PDF", default_name, "PDF (*.pdf)")
    if not path:
        return
    if not path.lower().endswith('.pdf'):
        path += '.pdf'
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        width, height = A4
        c = canvas.Canvas(path, pagesize=A4)

        # Fixed A4 cut-label layout: 3 columns x 8 rows, exactly 70mm x 35mm.
        # A4 width = 210mm, so 3 x 70mm fills the page width exactly.
        # A4 height = 297mm; 8 x 35mm = 280mm, so we center with 8.5mm top/bottom margin.
        label_w = 70 * mm
        label_h = 35 * mm
        cols = 3
        rows_per_page = 8
        top_margin = (height - rows_per_page * label_h) / 2
        left_margin = (width - cols * label_w) / 2

        for idx, row in enumerate(rows):
            pos = idx % (cols * rows_per_page)
            page_row = pos // cols
            page_col = pos % cols
            if idx > 0 and pos == 0:
                c.showPage()
            x = left_margin + page_col * label_w
            y = height - top_margin - (page_row + 1) * label_h
            _v21_draw_code128_label(c, row, x, y, label_w, label_h)

        c.save()
        QMessageBox.information(
            self,
            "Barcode Labels Saved",
            "Barcode label PDF saved successfully.\n\n"
            "Layout: A4 paper, 3 columns x 8 rows, 70mm x 35mm per label.\n"
            "Print at 100% / Actual Size. Do not use Fit to Page.\n\n"
            f"File:\n{path}"
        )
    except Exception as e:
        QMessageBox.critical(self, "Barcode Label Failed", str(e))


def _v21_export_all_barcode_labels(self):
    rows = self.db.query("""SELECT material_id, barcode, item_description FROM inventory_items
                          WHERE is_active=1 ORDER BY material_id""")
    _v21_export_barcode_labels(self, rows, "all_inventory_barcode_labels.pdf")

MainWindow.export_all_barcode_labels = _v21_export_all_barcode_labels


def _v21_export_selected_barcode_label(self):
    mat = self.selected_material_from_inventory() if hasattr(self, 'selected_material_from_inventory') else None
    if not mat:
        return
    row = self.db.one("SELECT material_id, barcode, item_description FROM inventory_items WHERE material_id=%s AND is_active=1", (mat,))
    if not row:
        QMessageBox.warning(self, "Not Found", "Selected item was not found.")
        return
    _v21_export_barcode_labels(self, [row], f"barcode_label_{mat}.pdf")

MainWindow.export_selected_barcode_label = _v21_export_selected_barcode_label


# Scanner inputs in Stock Inward.
_V21_ORIG_PAGE_INWARD = MainWindow.page_inward

def _v21_page_inward(self):
    page = _V21_ORIG_PAGE_INWARD(self)
    try:
        layout = page.layout()
        scan_box = QFrame(); scan_box.setObjectName("card")
        hl = QHBoxLayout(scan_box); hl.setContentsMargins(14,10,14,10); hl.setSpacing(8)
        lab = QLabel("Scan Barcode for Inward"); lab.setStyleSheet("font-weight:900;color:#102a43;")
        self.in_scan_barcode = QLineEdit(); self.in_scan_barcode.setPlaceholderText("Click here and scan item barcode. Scanner should press Enter automatically.")
        self.in_scan_barcode.returnPressed.connect(self.scan_barcode_for_inward)
        btn = QPushButton("Load Item"); btn.setObjectName("softBtn"); btn.clicked.connect(self.scan_barcode_for_inward)
        hl.addWidget(lab); hl.addWidget(self.in_scan_barcode, 1); hl.addWidget(btn)
        layout.insertWidget(1, scan_box)
    except Exception:
        pass
    return page

MainWindow.page_inward = _v21_page_inward


def _v21_scan_barcode_for_inward(self):
    code = self.in_scan_barcode.text().strip() if hasattr(self, 'in_scan_barcode') else ''
    row = _v21_find_material_by_barcode(self, code)
    if not row:
        QMessageBox.warning(self, "Barcode Not Found", f"No active inventory item found for barcode/material ID:\n{code}")
        return
    if _v21_select_material_in_combo(self, self.in_existing, row['material_id']):
        self.load_existing_for_inward()
        QMessageBox.information(self, "Item Loaded", f"Loaded item for inward:\n{row['material_id']} - {row.get('item_description','')}")

MainWindow.scan_barcode_for_inward = _v21_scan_barcode_for_inward

_V21_ORIG_CLEAR_INWARD = MainWindow.clear_inward_form

def _v21_clear_inward_form(self, *args, **kwargs):
    result = _V21_ORIG_CLEAR_INWARD(self, *args, **kwargs)
    if hasattr(self, 'in_scan_barcode'):
        self.in_scan_barcode.clear()
    return result

MainWindow.clear_inward_form = _v21_clear_inward_form


# Scanner inputs in Stock Outward.
_V21_ORIG_PAGE_OUTWARD = MainWindow.page_outward

def _v21_page_outward(self):
    page = _V21_ORIG_PAGE_OUTWARD(self)
    try:
        layout = page.layout()
        scan_box = QFrame(); scan_box.setObjectName("card")
        hl = QHBoxLayout(scan_box); hl.setContentsMargins(14,10,14,10); hl.setSpacing(8)
        lab = QLabel("Scan Barcode for Outward"); lab.setStyleSheet("font-weight:900;color:#102a43;")
        self.out_scan_barcode = QLineEdit(); self.out_scan_barcode.setPlaceholderText("Click here and scan item barcode. Scanner should press Enter automatically.")
        self.out_scan_barcode.returnPressed.connect(self.scan_barcode_for_outward)
        btn = QPushButton("Load Item"); btn.setObjectName("softBtn"); btn.clicked.connect(self.scan_barcode_for_outward)
        hl.addWidget(lab); hl.addWidget(self.out_scan_barcode, 1); hl.addWidget(btn)
        layout.insertWidget(1, scan_box)
    except Exception:
        pass
    return page

MainWindow.page_outward = _v21_page_outward


def _v21_scan_barcode_for_outward(self):
    code = self.out_scan_barcode.text().strip() if hasattr(self, 'out_scan_barcode') else ''
    row = _v21_find_material_by_barcode(self, code)
    if not row:
        QMessageBox.warning(self, "Barcode Not Found", f"No active inventory item found for barcode/material ID:\n{code}")
        return
    if _v21_select_material_in_combo(self, self.out_item, row['material_id']):
        self.load_existing_for_outward()
        QMessageBox.information(self, "Item Loaded", f"Loaded item for outward:\n{row['material_id']} - {row.get('item_description','')}\nAvailable stock: {row.get('quantity_available',0)}")

MainWindow.scan_barcode_for_outward = _v21_scan_barcode_for_outward

_V21_ORIG_CLEAR_OUTWARD = MainWindow.clear_outward_form

def _v21_clear_outward_form(self, *args, **kwargs):
    result = _V21_ORIG_CLEAR_OUTWARD(self, *args, **kwargs)
    if hasattr(self, 'out_scan_barcode'):
        self.out_scan_barcode.clear()
    return result

MainWindow.clear_outward_form = _v21_clear_outward_form


# Scanner support for adjustment/transfer, placed above the form where possible.
_V21_ORIG_PAGE_ADJUSTMENT = MainWindow.page_adjustment

def _v21_page_adjustment(self):
    page = _V21_ORIG_PAGE_ADJUSTMENT(self)
    try:
        layout = page.layout()
        scan_box = QFrame(); scan_box.setObjectName("card")
        hl = QHBoxLayout(scan_box); hl.setContentsMargins(14,10,14,10)
        self.adj_scan_barcode = QLineEdit(); self.adj_scan_barcode.setPlaceholderText("Scan barcode to select material for adjustment")
        self.adj_scan_barcode.returnPressed.connect(self.scan_barcode_for_adjustment)
        btn = QPushButton("Load Item"); btn.setObjectName("softBtn"); btn.clicked.connect(self.scan_barcode_for_adjustment)
        hl.addWidget(QLabel("Scan Barcode")); hl.addWidget(self.adj_scan_barcode, 1); hl.addWidget(btn)
        layout.insertWidget(1, scan_box)
    except Exception:
        pass
    return page

MainWindow.page_adjustment = _v21_page_adjustment


def _v21_scan_barcode_for_adjustment(self):
    code = self.adj_scan_barcode.text().strip() if hasattr(self, 'adj_scan_barcode') else ''
    row = _v21_find_material_by_barcode(self, code)
    if not row:
        QMessageBox.warning(self, "Barcode Not Found", f"No active item found for barcode/material ID:\n{code}")
        return
    _v21_select_material_in_combo(self, self.adj_item, row['material_id'])

MainWindow.scan_barcode_for_adjustment = _v21_scan_barcode_for_adjustment

_V21_ORIG_PAGE_TRANSFER = MainWindow.page_transfer

def _v21_page_transfer(self):
    page = _V21_ORIG_PAGE_TRANSFER(self)
    try:
        layout = page.layout()
        scan_box = QFrame(); scan_box.setObjectName("card")
        hl = QHBoxLayout(scan_box); hl.setContentsMargins(14,10,14,10)
        self.tr_scan_barcode = QLineEdit(); self.tr_scan_barcode.setPlaceholderText("Scan barcode to select material for transfer")
        self.tr_scan_barcode.returnPressed.connect(self.scan_barcode_for_transfer)
        btn = QPushButton("Load Item"); btn.setObjectName("softBtn"); btn.clicked.connect(self.scan_barcode_for_transfer)
        hl.addWidget(QLabel("Scan Barcode")); hl.addWidget(self.tr_scan_barcode, 1); hl.addWidget(btn)
        layout.insertWidget(1, scan_box)
    except Exception:
        pass
    return page

MainWindow.page_transfer = _v21_page_transfer


def _v21_scan_barcode_for_transfer(self):
    code = self.tr_scan_barcode.text().strip() if hasattr(self, 'tr_scan_barcode') else ''
    row = _v21_find_material_by_barcode(self, code)
    if not row:
        QMessageBox.warning(self, "Barcode Not Found", f"No active item found for barcode/material ID:\n{code}")
        return
    _v21_select_material_in_combo(self, self.tr_item, row['material_id'])

MainWindow.scan_barcode_for_transfer = _v21_scan_barcode_for_transfer


# Import/export barcode behavior: after upload, generate missing barcodes and warn on duplicates.
_V21_ORIG_IMPORT_EXCEL = MainWindow.import_excel

def _v21_import_excel(self):
    result = _V21_ORIG_IMPORT_EXCEL(self)
    try:
        self.generate_missing_barcodes(True)
        self.check_duplicate_barcodes(False)
        if hasattr(self, 'inv_table'):
            self.refresh_inventory()
        self.refresh_item_combos()
    except Exception:
        pass
    return result

MainWindow.import_excel = _v21_import_excel

# Barcode duplicate protection when editing inventory.
_V21_ORIG_EDIT_SELECTED_INVENTORY = MainWindow.edit_selected_inventory

def _v21_edit_selected_inventory(self):
    # Existing edit dialog already contains Barcode / QR field. Duplicate check is performed after dialog closes.
    result = _V21_ORIG_EDIT_SELECTED_INVENTORY(self)
    try:
        self.check_duplicate_barcodes(False)
    except Exception:
        pass
    return result

MainWindow.edit_selected_inventory = _v21_edit_selected_inventory

# Generate missing barcode after inward save for new item if barcode was not manually entered.
_V21_ORIG_SAVE_INWARD = MainWindow.save_inward

def _v21_save_inward(self):
    result = _V21_ORIG_SAVE_INWARD(self)
    try:
        self.generate_missing_barcodes(True)
        if hasattr(self, 'inv_table'):
            self.refresh_inventory()
    except Exception:
        pass
    return result

MainWindow.save_inward = _v21_save_inward

# ===================== END V21 BARCODE SCANNER WORKFLOW UPGRADE =====================


# ===================== V22 EXPIRED ALERTS + ADVANCED ANALYTICS + RACK VISUALIZATION =====================
try:
    from PySide6.QtCore import QTimer
    from PySide6.QtWidgets import QSizePolicy
except Exception:
    QTimer = None


def _v22_metric_row(self, title: str, query: str, params: Tuple = (), suffix: str = "") -> str:
    try:
        row = self.db.one(query, params) or {}
        val = list(row.values())[0] if row else 0
        if val is None:
            val = 0
        if suffix == "money":
            return money(val)
        if isinstance(val, float):
            return f"{val:,.0f}"
        return f"{val:,}" if isinstance(val, int) else str(val)
    except Exception:
        return "0"


def _v22_make_table(self, rows, columns, min_h=220):
    tbl = QTableWidget()
    tbl.setAlternatingRowColors(True)
    tbl.setMinimumHeight(min_h)
    self.table_fill(tbl, rows, columns)
    try:
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.resizeRowsToContents()
    except Exception:
        pass
    return tbl


def _v22_refresh_dashboard(self):
    while self.dash_grid.count():
        item = self.dash_grid.takeAt(0)
        widget = item.widget()
        if widget:
            widget.deleteLater()

    metrics = self.db.one("""SELECT COUNT(*) items,
        COALESCE(SUM(quantity_available*cost_per_unit),0) stock_value,
        SUM(CASE WHEN quantity_available <= minimum_stock_level THEN 1 ELSE 0 END) low_stock,
        SUM(CASE WHEN quantity_available <= 0 THEN 1 ELSE 0 END) zero_stock
        FROM inventory_items WHERE is_active=1""") or {}
    batch = self.db.one("""SELECT
        COALESCE(SUM(CASE WHEN expiry_date IS NOT NULL AND expiry_date < DATE('now') AND quantity_available>0 THEN 1 ELSE 0 END),0) expired,
        COALESCE(SUM(CASE WHEN expiry_date IS NOT NULL AND expiry_date BETWEEN DATE('now') AND DATE('now','+30 day') AND quantity_available>0 THEN 1 ELSE 0 END),0) exp30,
        COALESCE(SUM(CASE WHEN expiry_date IS NOT NULL AND expiry_date BETWEEN DATE('now','+31 day') AND DATE('now','+90 day') AND quantity_available>0 THEN 1 ELSE 0 END),0) exp90
        FROM inventory_batches""") or {}
    inward_month = self.db.one("""SELECT COALESCE(SUM(quantity),0) q FROM stock_inward
        WHERE procurement_date BETWEEN DATE('now','start of month') AND DATE('now')""") or {"q":0}
    outward_month = self.db.one("""SELECT COALESCE(SUM(quantity_issued),0) q FROM stock_outward
        WHERE issue_date BETWEEN DATE('now','start of month') AND DATE('now')""") or {"q":0}
    rack_count = self.db.one("SELECT COUNT(*) c FROM racks WHERE is_active=1") or {"c":0}

    cards = [
        ("Total Items", str(metrics.get("items", 0) or 0), "Active materials", "card", "▦"),
        ("Stock Value", money(metrics.get("stock_value", 0)), "Current inventory value", "card", "OMR"),
        ("Inward This Month", f"{float(inward_month.get('q') or 0):,.0f}", "Monthly received qty", "card", "↓"),
        ("Outward This Month", f"{float(outward_month.get('q') or 0):,.0f}", "Monthly issued qty", "card", "↑"),
        ("Low Stock", str(metrics.get("low_stock", 0) or 0), "At/below min level", "warn", "⚠"),
        ("Out of Stock", str(metrics.get("zero_stock", 0) or 0), "Zero balance items", "dangerCard", "0"),
        ("Expired Batches", str(batch.get("expired", 0) or 0), "Blocked for issue", "dangerCard", "⛔"),
        ("Expiring 30 Days", str(batch.get("exp30", 0) or 0), "Immediate action", "warn", "⌛"),
        ("Expiring 90 Days", str(batch.get("exp90", 0) or 0), "Upcoming expiry", "card", "90"),
        ("Active Racks", str(rack_count.get("c", 0) or 0), "Rack master count", "card", "▤"),
    ]
    for idx, c in enumerate(cards):
        self.dash_grid.addWidget(Card(*c), idx // 5, idx % 5)

    expired_rows = self.db.query("""SELECT b.material_id, i.item_description, b.batch_no,
        b.quantity_available, b.expiry_date, CAST(julianday('now')-julianday(b.expiry_date) AS INTEGER) expired_days
        FROM inventory_batches b JOIN inventory_items i ON i.id=b.item_id
        WHERE i.is_active=1 AND b.quantity_available>0 AND b.expiry_date IS NOT NULL AND b.expiry_date < DATE('now')
        ORDER BY b.expiry_date ASC LIMIT 12""")
    box_exp = QGroupBox("Expired Items Alert - Immediate Action Required")
    bx = QVBoxLayout(box_exp)
    bx.addWidget(_v22_make_table(self, expired_rows, [("Material ID","material_id"),("Item","item_description"),("Batch","batch_no"),("Qty Left","quantity_available"),("Expiry","expiry_date"),("Expired Days","expired_days")], 230))
    self.dash_grid.addWidget(box_exp, 2, 0, 1, 3)

    cat_rows = self.db.query("""SELECT COALESCE(c.name,'Unassigned') category,
        COUNT(i.id) items, COALESCE(SUM(i.quantity_available),0) qty,
        COALESCE(SUM(i.quantity_available*i.cost_per_unit),0) value
        FROM inventory_items i LEFT JOIN categories c ON c.id=i.category_id
        WHERE i.is_active=1 GROUP BY c.name ORDER BY value DESC LIMIT 10""")
    box_cat = QGroupBox("Category-wise Stock Value")
    bc = QVBoxLayout(box_cat)
    bc.addWidget(_v22_make_table(self, cat_rows, [("Category","category"),("Items","items"),("Qty","qty"),("Value","value")], 230))
    self.dash_grid.addWidget(box_cat, 2, 3, 1, 2)

    fast_rows = self.db.query("""SELECT o.material_id, i.item_description, COALESCE(SUM(o.quantity_issued),0) issued_qty
        FROM stock_outward o LEFT JOIN inventory_items i ON i.id=o.item_id
        GROUP BY o.material_id, i.item_description ORDER BY issued_qty DESC LIMIT 10""")
    box_fast = QGroupBox("Fast Moving Items - Based on Outward Quantity")
    bf = QVBoxLayout(box_fast)
    bf.addWidget(_v22_make_table(self, fast_rows, [("Material ID","material_id"),("Item","item_description"),("Issued Qty","issued_qty")], 220))
    self.dash_grid.addWidget(box_fast, 3, 0, 1, 2)

    rack_rows = self.db.query("""SELECT COALESCE(r.name,'No Rack') rack, COALESCE(l.name,'') location,
        COUNT(i.id) items, COALESCE(SUM(i.quantity_available),0) qty,
        COALESCE(SUM(i.quantity_available*i.cost_per_unit),0) value
        FROM inventory_items i LEFT JOIN racks r ON r.id=i.rack_id LEFT JOIN locations l ON l.id=i.location_id
        WHERE i.is_active=1 GROUP BY r.name,l.name ORDER BY items DESC LIMIT 12""")
    box_rack = QGroupBox("Rack Occupancy Summary")
    br = QVBoxLayout(box_rack)
    br.addWidget(_v22_make_table(self, rack_rows, [("Rack","rack"),("Location","location"),("Items","items"),("Qty","qty"),("Value","value")], 220))
    self.dash_grid.addWidget(box_rack, 3, 2, 1, 3)

    qa = QGroupBox("Quick Actions")
    ql = QHBoxLayout(qa)
    for txt, page_name in [("Stock Inward","Stock Inward"),("Stock Outward","Stock Outward"),("Expiry Reports","Reports"),("Rack Management","Rack Management"),("Inventory","Inventory")]:
        b = QPushButton(txt)
        def go(_, p=page_name):
            for idx, (name, _) in enumerate(self.pages):
                if name == p:
                    self.set_page(idx); break
        b.clicked.connect(go)
        ql.addWidget(b)
    self.dash_grid.addWidget(qa, 4, 0, 1, 5)

MainWindow.refresh_dashboard = _v22_refresh_dashboard


def _v22_show_expired_alerts(self):
    try:
        row = self.db.one("""SELECT COUNT(*) c, COALESCE(SUM(quantity_available),0) qty
            FROM inventory_batches
            WHERE expiry_date IS NOT NULL AND expiry_date < DATE('now') AND quantity_available>0""") or {}
        count = int(row.get('c') or 0)
        qty = float(row.get('qty') or 0)
        if count > 0:
            sample = self.db.query("""SELECT b.material_id, i.item_description, b.batch_no, b.quantity_available, b.expiry_date
                FROM inventory_batches b JOIN inventory_items i ON i.id=b.item_id
                WHERE b.expiry_date IS NOT NULL AND b.expiry_date < DATE('now') AND b.quantity_available>0
                ORDER BY b.expiry_date ASC LIMIT 8""")
            lines = [f"{r['material_id']} | {r['item_description']} | {r['batch_no']} | Qty {r['quantity_available']} | Exp {r['expiry_date']}" for r in sample]
            msg = f"Expired stock found. Do not issue these batches without admin decision.\n\nExpired batches: {count}\nExpired quantity: {qty:g}\n\n" + "\n".join(lines)
            QMessageBox.warning(self, "Expired Items Alert", msg)
    except Exception:
        pass

MainWindow.show_expired_alerts = _v22_show_expired_alerts


def _v22_page_rack_management(self):
    w, v = self.base_page("Rack Management Visualization")
    controls = QFrame(); controls.setObjectName("card")
    ch = QHBoxLayout(controls); ch.setContentsMargins(16,14,16,14)
    ch.addWidget(QLabel("Location:"))
    self.rack_location_filter = QComboBox(); self.rack_location_filter.setMinimumWidth(260)
    self.rack_location_filter.addItem("All Locations", None)
    for r in self.db.query("SELECT id,name FROM locations WHERE is_active=1 ORDER BY name"):
        self.rack_location_filter.addItem(r['name'], r['id'])
    ch.addWidget(self.rack_location_filter)
    refresh = QPushButton("Refresh Rack View"); refresh.clicked.connect(self.refresh_rack_visualization); ch.addWidget(refresh)
    ch.addStretch()
    self.rack_summary_label = QLabel("Select a rack to view items."); self.rack_summary_label.setObjectName("small"); ch.addWidget(self.rack_summary_label)
    v.addWidget(controls)

    split = QHBoxLayout(); split.setSpacing(14)
    self.rack_scroll = QScrollArea(); self.rack_scroll.setWidgetResizable(True); self.rack_scroll.setMinimumWidth(430)
    self.rack_cards_holder = QWidget(); self.rack_cards_layout = QGridLayout(self.rack_cards_holder); self.rack_cards_layout.setSpacing(12)
    self.rack_scroll.setWidget(self.rack_cards_holder); split.addWidget(self.rack_scroll, 1)

    right = QGroupBox("Selected Rack Items")
    rv = QVBoxLayout(right)
    self.rack_items_table = QTableWidget(); self.rack_items_table.setAlternatingRowColors(True)
    rv.addWidget(self.rack_items_table)
    split.addWidget(right, 2)
    v.addLayout(split, 1)
    self.refresh_rack_visualization()
    return w


def _v22_refresh_rack_visualization(self):
    while self.rack_cards_layout.count():
        item = self.rack_cards_layout.takeAt(0)
        widget = item.widget()
        if widget:
            widget.deleteLater()
    loc_id = self.rack_location_filter.currentData() if hasattr(self, 'rack_location_filter') else None
    params = []
    where = "WHERE r.is_active=1"
    if loc_id:
        where += " AND (r.location_id=? OR i.location_id=?)"
        params.extend([loc_id, loc_id])
    rows = self.db.query(f"""SELECT r.id rack_id, r.name rack, COALESCE(l.name,'') location,
        COUNT(i.id) item_count, COALESCE(SUM(i.quantity_available),0) stock_qty,
        COALESCE(SUM(i.quantity_available*i.cost_per_unit),0) stock_value,
        COALESCE(SUM(CASE WHEN i.quantity_available<=i.minimum_stock_level THEN 1 ELSE 0 END),0) low_items,
        COALESCE(SUM(CASE WHEN b.expiry_date IS NOT NULL AND b.expiry_date < DATE('now') AND b.quantity_available>0 THEN 1 ELSE 0 END),0) expired_batches
        FROM racks r
        LEFT JOIN locations l ON l.id=r.location_id
        LEFT JOIN inventory_items i ON i.rack_id=r.id AND i.is_active=1
        LEFT JOIN inventory_batches b ON b.item_id=i.id AND b.quantity_available>0
        {where}
        GROUP BY r.id,r.name,l.name ORDER BY r.name""", tuple(params))
    if not rows:
        lab = QLabel("No racks found. Add racks in Settings first."); lab.setStyleSheet("padding:20px;font-weight:900;color:#64748b;")
        self.rack_cards_layout.addWidget(lab,0,0)
        return
    for idx, r in enumerate(rows):
        btn = QPushButton()
        btn.setMinimumHeight(115)
        btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        btn.setText(f"{r['rack']}\n{r.get('location') or 'No location'}\nItems: {r['item_count']} | Qty: {float(r['stock_qty'] or 0):g}\nValue: {money(r['stock_value'])}")
        if int(r.get('expired_batches') or 0) > 0:
            btn.setStyleSheet("QPushButton{background:#dc3545;color:white;border-radius:14px;text-align:left;padding:12px;font-weight:900;} QPushButton:hover{background:#bd2130;}")
        elif int(r.get('low_items') or 0) > 0:
            btn.setStyleSheet("QPushButton{background:#ffb020;color:#172033;border-radius:14px;text-align:left;padding:12px;font-weight:900;} QPushButton:hover{background:#f59e0b;}")
        elif int(r.get('item_count') or 0) == 0:
            btn.setStyleSheet("QPushButton{background:#e5e7eb;color:#334155;border-radius:14px;text-align:left;padding:12px;font-weight:900;} QPushButton:hover{background:#d1d5db;}")
        else:
            btn.setStyleSheet("QPushButton{background:#0aa36d;color:white;border-radius:14px;text-align:left;padding:12px;font-weight:900;} QPushButton:hover{background:#087f56;}")
        btn.clicked.connect(lambda _, rack_id=r['rack_id'], rack_name=r['rack']: self.load_rack_items(rack_id, rack_name))
        self.rack_cards_layout.addWidget(btn, idx//3, idx%3)
    self.rack_summary_label.setText("Green = normal | Orange = low-stock item | Red = expired batch | Grey = empty")
    if rows:
        self.load_rack_items(rows[0]['rack_id'], rows[0]['rack'])


def _v22_load_rack_items(self, rack_id, rack_name):
    rows = self.db.query("""SELECT i.material_id, i.item_description, COALESCE(c.name,'') category,
        COALESCE(d.name,'') department, i.quantity_available, COALESCE(u.name,'') unit,
        i.minimum_stock_level, i.material_type,
        COALESCE(MIN(CASE WHEN b.quantity_available>0 THEN b.expiry_date END),'') nearest_expiry
        FROM inventory_items i
        LEFT JOIN categories c ON c.id=i.category_id
        LEFT JOIN departments d ON d.id=i.department_id
        LEFT JOIN units u ON u.id=i.unit_id
        LEFT JOIN inventory_batches b ON b.item_id=i.id
        WHERE i.is_active=1 AND i.rack_id=?
        GROUP BY i.id ORDER BY i.item_description""", (rack_id,))
    self.table_fill(self.rack_items_table, rows, [
        ("Material ID","material_id"),("Item Description","item_description"),("Category","category"),
        ("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min Level","minimum_stock_level"),
        ("Type","material_type"),("Nearest Expiry","nearest_expiry")
    ])
    try:
        self.rack_items_table.setColumnWidth(1, 420)
        self.rack_summary_label.setText(f"Selected Rack: {rack_name} | Items: {len(rows)}")
    except Exception:
        pass

MainWindow.page_rack_management = _v22_page_rack_management
MainWindow.refresh_rack_visualization = _v22_refresh_rack_visualization
MainWindow.load_rack_items = _v22_load_rack_items


def _v22_build_sidebar(self) -> QWidget:
    side = QFrame(); side.setObjectName("sidebar"); side.setFixedWidth(385)
    v = QVBoxLayout(side); v.setContentsMargins(16,22,16,18); v.setSpacing(6)
    title = QLabel("▥  STORE INVENTORY MANAGEMENT"); title.setObjectName("sideTitle"); title.setWordWrap(True); title.setMinimumWidth(360)
    sub = QLabel("DESIGNED & DEVELOPED - PAVAN KUMAR AKELLA"); sub.setObjectName("sideSub"); sub.setWordWrap(True); sub.setMinimumWidth(360)
    v.addWidget(title); v.addWidget(sub); v.addSpacing(18)
    sections = [
        ("MAIN", [("⌂", "Dashboard")]),
        ("STORE MANAGEMENT", [("↓", "Stock Inward"),("↑", "Stock Outward"),("↕", "Stock Adjustment"),("⇄", "Transfers")]),
        ("REPORTS", [("☷", "Reports")]),
        ("INVENTORY", [("▦", "Inventory"),("📒", "Item Ledger"),("▤", "Rack Management")]),
        ("CONTROL", [("⚙", "Settings"),("🧾", "Audit Logs")])
    ]
    page_index = 0
    self.nav_buttons = []
    for sec, items in sections:
        lab = QLabel(sec); lab.setObjectName("sectionLabel"); v.addWidget(lab)
        for icon, item in items:
            b = QPushButton(f"  {icon}   {item}"); b.setObjectName("nav"); b.setProperty("active", False)
            b.clicked.connect(lambda _, i=page_index: self.set_page(i))
            self.nav_buttons.append(b); v.addWidget(b); page_index += 1
    v.addStretch(); out = QPushButton("  ↪   Logout"); out.setObjectName("nav"); out.clicked.connect(self.close); v.addWidget(out)
    return side

MainWindow.build_sidebar = _v22_build_sidebar


def _v22_mainwindow_init(self, db: Database, user: Dict[str, Any]):
    QMainWindow.__init__(self)
    self.db = db; self.user = user
    self.setWindowTitle(APP_TITLE); self.resize(1550, 930)
    self.nav_buttons: List[QPushButton] = []
    root = QWidget(); self.setCentralWidget(root)
    h = QHBoxLayout(root); h.setContentsMargins(0,0,0,0); h.setSpacing(0)
    self.sidebar = self.build_sidebar(); h.addWidget(self.sidebar)
    self.stack = QStackedWidget(); h.addWidget(self.stack, 1)
    self.pages = [
        ("Dashboard", self.page_dashboard),
        ("Stock Inward", self.page_inward),
        ("Stock Outward", self.page_outward),
        ("Stock Adjustment", self.page_adjustment),
        ("Transfers", self.page_transfer),
        ("Reports", self.page_reports),
        ("Inventory", self.page_inventory),
        ("Item Ledger", self.page_item_ledger),
        ("Rack Management", self.page_rack_management),
        ("Settings", self.page_settings),
        ("Audit Logs", self.page_audit_logs),
    ]
    for name, builder in self.pages:
        try:
            self.stack.addWidget(builder())
        except Exception as e:
            page = QWidget(); lay = QVBoxLayout(page); lay.addWidget(QLabel(f"Could not load {name}: {e}")); self.stack.addWidget(page)
    self.set_page(0)
    if QTimer:
        QTimer.singleShot(900, self.show_expired_alerts)

MainWindow.__init__ = _v22_mainwindow_init


def _v22_set_page(self, i: int):
    self.stack.setCurrentIndex(i)
    for n,b in enumerate(self.nav_buttons):
        b.setProperty("active", n==i); b.style().unpolish(b); b.style().polish(b)
    name = self.pages[i][0] if hasattr(self, 'pages') and i < len(self.pages) else ""
    if name == "Dashboard" and hasattr(self, "refresh_dashboard"):
        self.refresh_dashboard()
    elif name == "Stock Inward" and hasattr(self, "clear_inward_form"):
        self.clear_inward_form(); self.refresh_inward_table()
    elif name == "Stock Outward" and hasattr(self, "clear_outward_form"):
        self.clear_outward_form(); self.refresh_outward_table()
    elif name == "Stock Adjustment" and hasattr(self, "clear_adjustment_form"):
        self.clear_adjustment_form(); self.refresh_adjustment_table()
    elif name == "Transfers" and hasattr(self, "clear_transfer_form"):
        self.clear_transfer_form(); self.refresh_transfer_table()
    elif name == "Reports" and hasattr(self, "generate_report"):
        self.generate_report()
    elif name == "Inventory" and hasattr(self, "refresh_inventory"):
        self.refresh_inventory()
    elif name == "Rack Management" and hasattr(self, "refresh_rack_visualization"):
        self.refresh_rack_visualization()
    elif name == "Audit Logs" and hasattr(self, "refresh_audit_logs"):
        self.refresh_audit_logs()

MainWindow.set_page = _v22_set_page

# ===================== END V22 EXPIRED ALERTS + ADVANCED ANALYTICS + RACK VISUALIZATION =====================

def main():
    app = QApplication(sys.argv); app.setStyleSheet(STYLE)
    try:
        db = Database(DB_PATH)
    except Exception as e:
        QMessageBox.critical(None, "Database Error", f"Could not create/open SQLite database.\n\n{e}\n\nCheck folder permission for fm_inventory.db.")
        return 1
    login = LoginDialog(db)
    if login.exec() != QDialog.Accepted: return 0
    win = MainWindow(db, login.user); win.show(); return app.exec()

if __name__ == "__main__":
    sys.exit(main())
