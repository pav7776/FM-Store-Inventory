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

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
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
#sideTitle { background: transparent; color: #ffffff; font-size: 18px; font-weight: 900; letter-spacing: .3px; padding: 4px 2px; }
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
        side = QFrame(); side.setObjectName("sidebar"); side.setFixedWidth(310)
        v = QVBoxLayout(side); v.setContentsMargins(16,22,16,18); v.setSpacing(6)
        title = QLabel("▥  STORE INVENTORY MANAGEMENT"); title.setObjectName("sideTitle")
        sub = QLabel("DESIGNED & DEVELOPED - PAVAN KUMAR AKELLA"); sub.setObjectName("sideSub"); sub.setWordWrap(True)
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
        inward = self.db.one("SELECT COALESCE(SUM(quantity),0) q FROM stock_inward WHERE strftime('%m', procurement_date)=strftime('%m','now') AND strftime('%Y', procurement_date)=strftime('%Y','now')")["q"]
        outward = self.db.one("SELECT COALESCE(SUM(quantity_issued),0) q FROM stock_outward WHERE strftime('%m', issue_date)=strftime('%m','now') AND strftime('%Y', issue_date)=strftime('%Y','now')")["q"]
        cards = [("Total Items", str(m.get("items",0)), "All items in inventory", "card"), ("Total Stock Inward", f"{float(inward):,.0f}", "This month", "card"), ("Total Stock Outward", f"{float(outward):,.0f}", "This month", "card"), ("Total Stock Value", money(m.get("stock_value",0)), "Overall inventory value", "card"), ("Low Stock Items", str(m.get("low_stock") or 0), "Require attention", "warn"), ("To Be Expired This Month", str(m.get("expiring") or 0), "Items to expire", "dangerCard"), ("Expired Items", str(m.get("expired") or 0), "Already expired", "warn")]
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
        self.in_existing = QComboBox(); self.in_existing.addItem("-- New Material --")
        for r in self.item_rows():
            self.in_existing.addItem(f"{r['material_id']} - {r['item_description']}", r['material_id'])
        self.in_existing.setMinimumWidth(360)
        self.in_existing.currentIndexChanged.connect(self.load_existing_for_inward)

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
            ("PO No.", self.in_po), ("DO No.", self.in_do),
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

    def load_existing_for_inward(self):
        mat=self.in_existing.currentData()
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
            mat = self.in_existing.currentData()
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
        hint = QLabel("Upload initial stock, search live inventory, edit selected material, and download complete store data."); hint.setObjectName("small")
        th = QVBoxLayout(); th.addWidget(title); th.addWidget(hint); head.addLayout(th); head.addStretch()
        self.inv_search=QLineEdit(); self.inv_search.setPlaceholderText("Search Material ID, item, category, department, supplier, rack, location..."); self.inv_search.setFixedWidth(430)
        self.inv_search.returnPressed.connect(self.refresh_inventory); head.addWidget(self.inv_search)
        b=QPushButton("Search"); b.clicked.connect(self.refresh_inventory); head.addWidget(b)
        av.addLayout(head)
        btns = QHBoxLayout()
        tmpl=QPushButton("Download Excel Upload Template"); tmpl.setObjectName("softBtn"); tmpl.clicked.connect(self.download_inventory_template)
        imp=QPushButton("Upload Initial Stock Excel"); imp.setObjectName("greenBtn"); imp.clicked.connect(self.import_excel)
        down=QPushButton("Download Complete Inventory"); down.clicked.connect(self.export_complete_inventory)
        edit=QPushButton("Edit Selected Item"); edit.setObjectName("orangeBtn"); edit.clicked.connect(self.edit_selected_inventory)
        delete=QPushButton("Delete Selected Item"); delete.setObjectName("redBtn"); delete.clicked.connect(self.delete_selected_inventory)
        exp=QPushButton("Export Current View"); exp.setObjectName("softBtn"); exp.clicked.connect(lambda:self.export_table(self.inv_table,"inventory_current_view"))
        for x in [tmpl, imp, down, edit, delete, exp]: btns.addWidget(x)
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
        path,_=QFileDialog.getOpenFileName(self,"Upload Existing Store Materials / Stock Details","","Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            df=pd.read_excel(path).fillna("")
        except Exception as e:
            QMessageBox.critical(self,"Import Failed",f"Could not read Excel file.\n\n{e}"); return
        df.columns=[str(c).strip().lower().replace(" ","_") for c in df.columns]
        required=['item_description','category','department','quantity_available','unit','minimum_stock_level','material_type','cost_per_unit']
        missing=[c for c in required if c not in df.columns]
        if missing:
            QMessageBox.warning(self,"Invalid Excel Format", "Missing required columns:\n" + ", ".join(missing) + "\n\nUse 'Download Excel Upload Template' first."); return
        if df.empty:
            QMessageBox.warning(self,"Empty File","The selected Excel file has no rows."); return
        preview=ImportPreviewDialog(df, self)
        if preview.exec()!=QDialog.Accepted or not preview.accepted_import: return
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
        msg=f"Import completed.\n\nCreated new materials: {created}\nUpdated existing materials: {updated}\nSkipped blank rows: {skipped}"
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
        controls.setHorizontalSpacing(16)
        controls.setVerticalSpacing(12)

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
        for c in [self.rep_type,self.rep_department,self.rep_category,self.rep_supplier]:
            c.setMinimumWidth(280); c.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self.rep_from=QDateEdit(); self.rep_from.setCalendarPopup(True); self.rep_from.setDate(date.today().replace(day=1)); self.rep_from.setMinimumWidth(170)
        self.rep_to=QDateEdit(); self.rep_to.setCalendarPopup(True); self.rep_to.setDate(date.today()); self.rep_to.setMinimumWidth(170)
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
        self.generate_report(); return w

    def _filtered_inventory_rows_for_reports(self):
        rows = self.item_rows()
        dept = self.rep_department.currentText() if hasattr(self, 'rep_department') else 'All'
        cat = self.rep_category.currentText() if hasattr(self, 'rep_category') else 'All'
        sup = self.rep_supplier.currentText() if hasattr(self, 'rep_supplier') else 'All'
        if dept != 'All': rows = [r for r in rows if (r.get('department') or '') == dept]
        if cat != 'All': rows = [r for r in rows if (r.get('category') or '') == cat]
        if sup != 'All': rows = [r for r in rows if (r.get('supplier') or '') == sup]
        return rows

    def _append_inventory_filter_sql(self, sql: str, params: list):
        dept = self.rep_department.currentText() if hasattr(self, 'rep_department') else 'All'
        cat = self.rep_category.currentText() if hasattr(self, 'rep_category') else 'All'
        sup = self.rep_supplier.currentText() if hasattr(self, 'rep_supplier') else 'All'
        if dept != 'All':
            sql += " AND d.name=%s"; params.append(dept)
        if cat != 'All':
            sql += " AND c.name=%s"; params.append(cat)
        if sup != 'All':
            sql += " AND s.name=%s"; params.append(sup)
        return sql, params

    def generate_report(self):
        typ=self.rep_type.currentText(); start=self.rep_from.date().toPython().isoformat(); end=self.rep_to.date().toPython().isoformat()
        base_cols=[("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Unit","unit"),("Min. Level","minimum_stock_level"),("Type","material_type"),("Supplier","supplier"),("Rack No.","rack"),("Total Value (OMR)","total_value")]
        if typ in ["Balance Stock Report", "Complete Inventory Report", "Department-wise Stock Report", "Category-wise Stock Report", "Supplier-wise Purchase Report", "Store Location-wise Stock Report"]:
            rows=self._filtered_inventory_rows_for_reports(); cols=base_cols
        elif typ=="Low Stock List":
            rows=[r for r in self._filtered_inventory_rows_for_reports() if float(r.get('quantity_available') or 0) <= float(r.get('minimum_stock_level') or 0)]
            cols=[("Material ID","material_id"),("Item Description","item_description"),("Category","category"),("Department","department"),("Stock","quantity_available"),("Min Level","minimum_stock_level"),("Supplier","supplier")]
        elif typ=="Expiring Materials List":
            sql="""SELECT b.material_id,i.item_description,b.batch_no,b.quantity_received,b.quantity_available,c.name category,d.name department,s.name supplier,b.expiry_date,CAST(julianday(b.expiry_date)-julianday('now') AS INTEGER) days_left
                   FROM inventory_batches b JOIN inventory_items i ON i.id=b.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN suppliers s ON s.id=i.supplier_id
                   WHERE i.is_active=1 AND b.quantity_available>0 AND b.expiry_date BETWEEN DATE('now') AND DATE('now','+30 day')"""
            sql, params = self._append_inventory_filter_sql(sql, [])
            rows=self.db.query(sql, tuple(params)); cols=[("Material ID","material_id"),("Item","item_description"),("Batch No","batch_no"),("Batch Qty Received","quantity_received"),("Batch Qty Left","quantity_available"),("Category","category"),("Department","department"),("Supplier","supplier"),("Expiry","expiry_date"),("Days Left","days_left")]
        elif typ=="Expired Materials List":
            sql="""SELECT b.material_id,i.item_description,b.batch_no,b.quantity_received,b.quantity_available,c.name category,d.name department,s.name supplier,b.expiry_date,CAST(julianday('now')-julianday(b.expiry_date) AS INTEGER) days_overdue
                   FROM inventory_batches b JOIN inventory_items i ON i.id=b.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN suppliers s ON s.id=i.supplier_id
                   WHERE i.is_active=1 AND b.quantity_available>0 AND b.expiry_date < DATE('now')"""
            sql, params = self._append_inventory_filter_sql(sql, [])
            rows=self.db.query(sql, tuple(params)); cols=[("Material ID","material_id"),("Item","item_description"),("Batch No","batch_no"),("Batch Qty Received","quantity_received"),("Batch Qty Left","quantity_available"),("Category","category"),("Department","department"),("Supplier","supplier"),("Expiry","expiry_date"),("Overdue","days_overdue")]
        elif typ=="Stock Inward Report":
            sql="""SELECT si.inward_no,si.material_id,i.item_description,c.name category,d.name department,s.name supplier,si.quantity,si.unit_cost,si.total_cost,si.procurement_date
                   FROM stock_inward si LEFT JOIN inventory_items i ON i.id=si.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=i.department_id LEFT JOIN suppliers s ON s.id=si.supplier_id
                   WHERE si.procurement_date BETWEEN %s AND %s"""
            sql, params = self._append_inventory_filter_sql(sql, [start,end])
            rows=self.db.query(sql, tuple(params)); cols=[("Inward No","inward_no"),("Material ID","material_id"),("Item","item_description"),("Category","category"),("Department","department"),("Supplier","supplier"),("Qty","quantity"),("Unit Cost (OMR)","unit_cost"),("Total (OMR)","total_cost"),("Date","procurement_date")]
        elif typ=="Stock Outward Report":
            sql="""SELECT so.siv_no,so.issue_date,so.material_id,i.item_description,c.name category,d.name department,s.name supplier,so.quantity_issued,so.issued_to,so.issued_by,so.remarks
                   FROM stock_outward so LEFT JOIN inventory_items i ON i.id=so.item_id
                   LEFT JOIN categories c ON c.id=i.category_id LEFT JOIN departments d ON d.id=so.department_id LEFT JOIN suppliers s ON s.id=i.supplier_id
                   WHERE so.issue_date BETWEEN %s AND %s"""
            sql, params = self._append_inventory_filter_sql(sql, [start,end])
            rows=self.db.query(sql, tuple(params)); cols=[("SIV #","siv_no"),("Date","issue_date"),("Material ID","material_id"),("Item","item_description"),("Category","category"),("Department","department"),("Qty","quantity_issued"),("Issued To","issued_to"),("Issued By","issued_by"),("Remarks","remarks")]
        elif typ=="Transfer List":
            rows=self.db.query("SELECT transfer_no,transfer_date,material_id,quantity_transferred,requested_by,approved_by FROM stock_transfers WHERE transfer_date BETWEEN %s AND %s",(start,end))
            cols=[("Transfer No","transfer_no"),("Date","transfer_date"),("Material ID","material_id"),("Qty","quantity_transferred"),("Requested By","requested_by"),("Approved By","approved_by")]
        else:
            rows=self._filtered_inventory_rows_for_reports(); cols=base_cols
        self.table_fill(self.rep_table,rows,cols)

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
        backup=QWidget(); bl=QVBoxLayout(backup); b1=QPushButton("Backup SQLite Database File"); b1.clicked.connect(self.backup_db); bl.addWidget(b1); bl.addWidget(QLabel("Restore: close the app and replace fm_inventory.db with the backup .db file.")); bl.addStretch(); tabs.addTab(backup,"Backup / Restore")
        return w

    def master_tab(self, table: str):
        w = QWidget(); v = QVBoxLayout(w)
        h = QHBoxLayout()
        inp = QLineEdit(); inp.setPlaceholderText(f"{table[:-1].title()} name")
        add = QPushButton("Add"); add.setObjectName("greenBtn")
        edit = QPushButton("Edit Selected"); edit.setObjectName("softBtn")
        delete = QPushButton("Delete / Deactivate Selected"); delete.setObjectName("redBtn")
        h.addWidget(inp); h.addWidget(add); h.addWidget(edit); h.addWidget(delete); v.addLayout(h)
        tw = QTableWidget(); tw.setSelectionBehavior(QTableWidget.SelectRows); tw.setSelectionMode(QTableWidget.SingleSelection); v.addWidget(tw)

        def refresh():
            self.table_fill(tw, self.db.query(f"SELECT id,name,is_active,created_at FROM {table} ORDER BY name"), [("ID","id"),("Name","name"),("Active","is_active"),("Created","created_at")])

        def selected_id_name():
            r = tw.currentRow()
            if r < 0: return None, None
            return tw.item(r,0).text(), tw.item(r,1).text()

        def do_add():
            name = inp.text().strip()
            if not name:
                QMessageBox.warning(self,"Validation","Enter a name first."); return
            self.db.execute(f"INSERT OR IGNORE INTO {table}(name,is_active) VALUES(%s,1)",(name,))
            inp.clear(); refresh(); QMessageBox.information(self,"Saved",f"{name} added.")

        def do_edit():
            row_id, old = selected_id_name()
            if not row_id: QMessageBox.warning(self,"Selection","Select a row first."); return
            new, ok = QInputDialog.getText(self,"Edit",f"Rename {old} to:", text=old)
            if ok and new.strip():
                self.db.execute(f"UPDATE {table} SET name=%s, updated_at=CURRENT_TIMESTAMP WHERE id=%s",(new.strip(), row_id))
                refresh(); QMessageBox.information(self,"Updated",f"Updated to {new.strip()}.")

        def do_delete():
            row_id, old = selected_id_name()
            if not row_id: QMessageBox.warning(self,"Selection","Select a row first."); return
            if QMessageBox.question(self,"Confirm",f"Deactivate '{old}'? Existing transactions remain safe.") == QMessageBox.Yes:
                self.db.execute(f"UPDATE {table} SET is_active=0, updated_at=CURRENT_TIMESTAMP WHERE id=%s",(row_id,))
                refresh(); QMessageBox.information(self,"Updated",f"{old} deactivated.")

        add.clicked.connect(do_add); edit.clicked.connect(do_edit); delete.clicked.connect(do_delete)
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

    def page_audit(self):
        w,v=self.base_page("Audit Logs"); self.audit_table=QTableWidget(); v.addWidget(self.audit_table); rows=self.db.query("SELECT a.created_at,u.username,a.action,a.module,a.record_ref,a.details FROM audit_logs a LEFT JOIN users u ON u.id=a.user_id ORDER BY a.id DESC LIMIT 300"); self.table_fill(self.audit_table,rows,[("Date","created_at"),("User","username"),("Action","action"),("Module","module"),("Record","record_ref"),("Details","details")]); return w


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
