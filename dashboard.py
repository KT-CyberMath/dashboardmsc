import os
import re
import sys
import shutil
import sqlite3
import subprocess
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, ttk
from datetime import datetime
from email import policy
from email.parser import BytesParser
import pandas as pd

# Native OS drag-and-drop (dragging a file from Finder/Explorer/Mail/Outlook
# straight onto the app) needs the third-party tkinterdnd2 package — plain
# Tkinter can only do drag-and-drop *within* the app (which is what the
# existing "drag a subject/body into the to-do list" feature already uses).
# Install with: pip install tkinterdnd2
# If it's not installed, the app still runs fine — drag-and-drop from the OS
# just won't be available, and "Browse" buttons keep working as before.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

CURRENCY_SYMBOLS = {"EUR": "€", "USD": "$", "JPY": "¥"}
ATTACHMENT_TYPES = [
    ("contract", "Contract"),
    ("email", "Email"),
]
ATTACHMENT_FILETYPES = [
    ("PDF files", "*.pdf"),
    ("Word files", "*.doc *.docx"),
    ("Excel files", "*.xls *.xlsx"),
    ("Images", "*.jpg *.jpeg *.png"),
    ("All files", "*.*"),
]

# The database used to live next to the script (BASE_DIR = the .py file's
# folder). That breaks once this gets packaged into a standalone app with
# PyInstaller: a packaged .app/.exe unpacks to a temporary folder on every
# launch, so "next to the script" would be a different, throwaway location
# each time — meaning every launch would look like a fresh install with no
# saved data. Store the database in a stable, per-user app-data folder
# instead, which works identically whether this runs as a plain .py file or
# as a packaged app.
def get_app_data_dir():
    if sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    elif sys.platform.startswith("win"):
        base = os.getenv("APPDATA", os.path.expanduser("~"))
    else:
        base = os.path.expanduser("~/.local/share")

    app_dir = os.path.join(base, "DashboardMSC")
    os.makedirs(app_dir, exist_ok=True)
    return app_dir


BASE_DIR = get_app_data_dir()
DB_FILE = os.path.join(BASE_DIR, "dashboard.db")

# One-time migration: if there's an old dashboard.db sitting next to the
# script from before this change, and no database in the new location yet,
# copy the old one over so existing data isn't lost.
_LEGACY_DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dashboard.db")
if not os.path.exists(DB_FILE) and os.path.exists(_LEGACY_DB_FILE):
    try:
        shutil.copy2(_LEGACY_DB_FILE, DB_FILE)
    except OSError:
        pass

# ================= SUPPLIER FIELD VALIDATION =================
# Name: letters (Greek or Latin) and numbers only (plus spaces).
SUPPLIER_NAME_PATTERN = re.compile(r"^[A-Za-zΑ-Ωα-ωΆΈΉΊΌΎΏάέήίόύώϊϋΐΰ0-9\s]+$")
# Tel: numbers and phone-style symbols only (+, -, (, ), space, /, .).
SUPPLIER_TEL_PATTERN = re.compile(r"^[0-9+\-()/.\s]+$")
# Contract: both letters/numbers and symbols are allowed — just block empty
# after stripping.
SUPPLIER_CONTRACT_PATTERN = re.compile(r"^.+$")


def validate_supplier_fields(name, tel, contact):
    """Returns an error message string if invalid, or None if all fields
    that were actually filled in pass their format rules. Name is required;
    Tel and Contract are optional but must match their format if provided."""
    if not name:
        return "Το πεδίο Name είναι υποχρεωτικό."

    if not SUPPLIER_NAME_PATTERN.match(name):
        return "Το Name επιτρέπει μόνο γράμματα και αριθμούς."

    if tel and not SUPPLIER_TEL_PATTERN.match(tel):
        return "Το Tel επιτρέπει μόνο αριθμούς και σύμβολα (+, -, (, ), /, .)."

    if contact and not SUPPLIER_CONTRACT_PATTERN.match(contact):
        return "Το Contract δεν μπορεί να είναι κενό."

    return None


def to_title_case(text):
    """"john smith" -> "JOHN SMITH", "sales manager" -> "SALES MANAGER".
    Used for member Name/Surname/Job Title so entries stay consistent
    regardless of how someone typed them in. (Despite the name, this now
    upper-cases rather than title-cases — kept the name to avoid touching
    every call site.)"""
    return text.strip().upper()


def get_conn():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sender TEXT,
            subject TEXT NOT NULL,
            body TEXT,
            expanded INTEGER NOT NULL DEFAULT 0
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            text TEXT NOT NULL,
            date TEXT NOT NULL,
            done INTEGER NOT NULL DEFAULT 0
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS suppliers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            tel TEXT,
            contact TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS supplier_attachments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER NOT NULL,
            type_key TEXT NOT NULL,
            file_path TEXT,
            FOREIGN KEY (supplier_id) REFERENCES suppliers(id) ON DELETE CASCADE
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS supplier_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_id INTEGER NOT NULL,
            description TEXT NOT NULL,
            qty INTEGER NOT NULL,
            price REAL NOT NULL,
            currency TEXT NOT NULL,
            FOREIGN KEY (supplier_id) REFERENCES suppliers(id) ON DELETE CASCADE
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            job_title TEXT NOT NULL,
            note TEXT,
            created_date TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS member_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER NOT NULL,
            note_text TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY (member_id) REFERENCES members(id) ON DELETE CASCADE
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS task_assignees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            task_id INTEGER NOT NULL,
            member_id INTEGER NOT NULL,
            FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE,
            FOREIGN KEY (member_id) REFERENCES members(id) ON DELETE CASCADE
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS meetings (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            meeting_date TEXT NOT NULL,
            title TEXT NOT NULL,
            note TEXT,
            created_at TEXT NOT NULL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS stock_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_file TEXT,
            sheet_name TEXT,
            item_group TEXT NOT NULL,
            item_name TEXT,
            price REAL NOT NULL,
            imported_at TEXT NOT NULL
        )
    """)

    # Existing databases from before file attachments were added won't have
    # these columns — CREATE TABLE IF NOT EXISTS doesn't retroactively add
    # columns to a table that already exists, so add them here if missing.
    _add_column_if_missing(cur, "emails", "file_path", "TEXT")
    _add_column_if_missing(cur, "tasks", "attachment_path", "TEXT")
    # status: 'active' | 'completed' | 'dismissed' — tracks a task's state
    # beyond the old plain done/not-done checkbox, so a member's profile can
    # separate "still active" from "previously handled" (either finished or
    # dismissed), rather than just a single boolean.
    _add_column_if_missing(cur, "tasks", "status", "TEXT NOT NULL DEFAULT 'active'")
    # Legacy single-assignee column, kept only so old data isn't lost —
    # tasks can now be assigned to multiple people via task_assignees below.
    _add_column_if_missing(cur, "tasks", "assigned_member_id", "INTEGER")

    # One-time backfill: move any old single assigned_member_id values into
    # the new task_assignees table, so nothing gets silently dropped for
    # people upgrading from before multi-assignee support existed.
    cur.execute("""
        INSERT INTO task_assignees (task_id, member_id)
        SELECT t.id, t.assigned_member_id
        FROM tasks t
        WHERE t.assigned_member_id IS NOT NULL
          AND NOT EXISTS (
              SELECT 1 FROM task_assignees ta
              WHERE ta.task_id = t.id AND ta.member_id = t.assigned_member_id
          )
    """)

    # Stock items now require a supplier attached (so the stocklist can be
    # browsed by supplier). Old rows imported before this existed have no
    # way to know which supplier they belonged to, so — per an explicit
    # "wipe and start fresh" decision — the whole stocklist is cleared the
    # one time this column is first added, rather than leaving orphaned
    # unassigned rows around. Only fires once (isinstance check: the column
    # only gets "added" the first time this runs against a given database).
    stock_supplier_col_is_new = _add_column_if_missing(cur, "stock_items", "supplier_id", "INTEGER")
    if stock_supplier_col_is_new:
        cur.execute("DELETE FROM stock_items")

    conn.commit()
    conn.close()


def _add_column_if_missing(cur, table, column, coltype):
    """Returns True if the column was actually added (didn't exist before),
    False if it was already there. Some callers use this to run a one-time
    migration step only on the run where the column first appears."""
    try:
        cur.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coltype}")
        return True
    except sqlite3.OperationalError:
        return False  # column already exists


def clean_price_to_float(value):
    if value is None:
        return None

    s = str(value).strip()
    if s in ("", "nan", "None", "NaN"):
        return None

    s = (
        s.replace("€", "")
         .replace("$", "")
         .replace("£", "")
         .replace("¥", "")
         .replace("EUR", "")
         .replace(" ", "")
    )

    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(".", "")
        s = s.replace(",", ".")

    try:
        return float(s)
    except:
        return None


# ================= EMAILS =================
def get_all_emails():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM emails ORDER BY id DESC")
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def add_email_db(sender, subject, body, expanded=0, file_path=None):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO emails (sender, subject, body, expanded, file_path) VALUES (?, ?, ?, ?, ?)",
        (sender, subject, body, expanded, file_path)
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def update_email_db(email_id, sender, subject, body, expanded):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE emails SET sender=?, subject=?, body=?, expanded=? WHERE id=?",
        (sender, subject, body, expanded, email_id)
    )
    conn.commit()
    conn.close()


def delete_email_db(email_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM emails WHERE id=?", (email_id,))
    conn.commit()
    conn.close()


def set_email_expanded_db(email_id, expanded):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE emails SET expanded=? WHERE id=?", (expanded, email_id))
    conn.commit()
    conn.close()


# ================= TASKS =================
def get_all_tasks():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM tasks ORDER BY id DESC")
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def add_task_db(text, date, done=0):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("INSERT INTO tasks (text, date, done) VALUES (?, ?, ?)", (text, date, done))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def update_task_db(task_id, text):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE tasks SET text=? WHERE id=?", (text, task_id))
    conn.commit()
    conn.close()


def set_task_done_db(task_id, done):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE tasks SET done=? WHERE id=?", (done, task_id))
    conn.commit()
    conn.close()


def delete_task_db(task_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM tasks WHERE id=?", (task_id,))
    conn.commit()
    conn.close()


def set_task_attachment_db(task_id, file_path):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE tasks SET attachment_path=? WHERE id=?", (file_path, task_id))
    conn.commit()
    conn.close()


def set_task_status_db(task_id, status):
    """status: 'active' | 'completed' | 'dismissed'. Keeps the legacy
    `done` column in sync too (1 for completed/dismissed, 0 for active) so
    anything still reading `done` stays consistent."""
    conn = get_conn()
    cur = conn.cursor()
    done = 1 if status in ("completed", "dismissed") else 0
    cur.execute("UPDATE tasks SET status=?, done=? WHERE id=?", (status, done, task_id))
    conn.commit()
    conn.close()


def set_task_assignees_db(task_id, member_ids):
    """Replaces the full set of people assigned to a task with `member_ids`
    (a list, possibly empty for "unassigned"). Tasks can now be assigned to
    multiple people at once."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM task_assignees WHERE task_id=?", (task_id,))
    cur.executemany(
        "INSERT INTO task_assignees (task_id, member_id) VALUES (?, ?)",
        [(task_id, mid) for mid in member_ids]
    )
    # Keep the legacy single-assignee column loosely in sync (first person
    # in the list, or NULL) purely so any old code path reading it directly
    # doesn't see stale data — the real source of truth is task_assignees.
    cur.execute(
        "UPDATE tasks SET assigned_member_id=? WHERE id=?",
        (member_ids[0] if member_ids else None, task_id)
    )
    conn.commit()
    conn.close()


def get_task_assignees_db(task_id):
    """Returns the list of member dicts assigned to a task."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT m.* FROM members m
        JOIN task_assignees ta ON ta.member_id = m.id
        WHERE ta.task_id = ?
        ORDER BY m.first_name COLLATE NOCASE ASC, m.last_name COLLATE NOCASE ASC
        """,
        (task_id,)
    )
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def get_tasks_for_member_db(member_id, statuses):
    conn = get_conn()
    cur = conn.cursor()
    placeholders = ",".join("?" for _ in statuses)
    cur.execute(
        f"""
        SELECT t.* FROM tasks t
        JOIN task_assignees ta ON ta.task_id = t.id
        WHERE ta.member_id=? AND t.status IN ({placeholders})
        ORDER BY t.id DESC
        """,
        (member_id, *statuses)
    )
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


# ================= SUPPLIERS =================
def get_all_suppliers(order_by="date"):
    """order_by: "date" = registration order, newest first (id is assigned
    strictly in insertion order, so it doubles as a registration-date sort
    without needing a separate timestamp column). "name" = alphabetical."""
    conn = get_conn()
    cur = conn.cursor()

    if order_by == "name":
        cur.execute("SELECT * FROM suppliers ORDER BY name COLLATE NOCASE ASC")
    else:
        cur.execute("SELECT * FROM suppliers ORDER BY id DESC")

    supplier_rows = cur.fetchall()

    suppliers = []
    for row in supplier_rows:
        supplier = dict(row)

        cur.execute(
            "SELECT type_key, file_path FROM supplier_attachments WHERE supplier_id=?",
            (supplier["id"],)
        )
        attachments = {r["type_key"]: r["file_path"] for r in cur.fetchall()}

        cur.execute(
            "SELECT * FROM supplier_items WHERE supplier_id=? ORDER BY id DESC",
            (supplier["id"],)
        )
        items = [dict(r) for r in cur.fetchall()]

        supplier["attachments"] = attachments
        supplier["items"] = items
        suppliers.append(supplier)

    conn.close()
    return suppliers


def add_supplier_db(name, tel, contact):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO suppliers (name, tel, contact) VALUES (?, ?, ?)",
        (name, tel, contact)
    )
    supplier_id = cur.lastrowid

    for key in ["contract", "email", "quote"]:
        cur.execute(
            "INSERT INTO supplier_attachments (supplier_id, type_key, file_path) VALUES (?, ?, ?)",
            (supplier_id, key, None)
        )

    conn.commit()
    conn.close()
    return supplier_id


def update_supplier_db(supplier_id, name, tel, contact):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE suppliers SET name=?, tel=?, contact=? WHERE id=?",
        (name, tel, contact, supplier_id)
    )
    conn.commit()
    conn.close()


def delete_supplier_db(supplier_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM suppliers WHERE id=?", (supplier_id,))
    conn.commit()
    conn.close()


def set_supplier_attachment_db(supplier_id, type_key, file_path):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "UPDATE supplier_attachments SET file_path=? WHERE supplier_id=? AND type_key=?",
        (file_path, supplier_id, type_key)
    )
    conn.commit()
    conn.close()


def add_supplier_item_db(supplier_id, description, qty, price, currency):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO supplier_items (supplier_id, description, qty, price, currency) VALUES (?, ?, ?, ?, ?)",
        (supplier_id, description, qty, price, currency)
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def delete_supplier_item_db(item_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM supplier_items WHERE id=?", (item_id,))
    conn.commit()
    conn.close()


# ================= MEMBERS =================
def get_all_members(search=""):
    """Alphabetical by first name, then last name. `search` (optional)
    filters to names containing that text, case-insensitively."""
    conn = get_conn()
    cur = conn.cursor()
    if search:
        like = f"%{search}%"
        cur.execute(
            """
            SELECT * FROM members
            WHERE first_name LIKE ? OR last_name LIKE ?
            ORDER BY first_name COLLATE NOCASE ASC, last_name COLLATE NOCASE ASC
            """,
            (like, like)
        )
    else:
        cur.execute(
            "SELECT * FROM members ORDER BY first_name COLLATE NOCASE ASC, last_name COLLATE NOCASE ASC"
        )
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def get_member_by_id(member_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM members WHERE id=?", (member_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


def member_exists_db(first_name, last_name, exclude_id=None):
    """Case-insensitive check for whether a member with this exact
    first+last name combo is already registered. exclude_id lets a member
    keep their own name when just editing (renaming to the same name they
    already have shouldn't trigger "already exists")."""
    conn = get_conn()
    cur = conn.cursor()
    if exclude_id is not None:
        cur.execute(
            "SELECT 1 FROM members WHERE first_name=? COLLATE NOCASE AND last_name=? COLLATE NOCASE AND id != ? LIMIT 1",
            (first_name, last_name, exclude_id)
        )
    else:
        cur.execute(
            "SELECT 1 FROM members WHERE first_name=? COLLATE NOCASE AND last_name=? COLLATE NOCASE LIMIT 1",
            (first_name, last_name)
        )
    exists = cur.fetchone() is not None
    conn.close()
    return exists


def add_member_db(first_name, last_name, job_title, created_date):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO members (first_name, last_name, job_title, note, created_date) VALUES (?, ?, ?, ?, ?)",
        (first_name, last_name, job_title, "", created_date)
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def update_member_job_title_db(member_id, job_title):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE members SET job_title=? WHERE id=?", (job_title, member_id))
    conn.commit()
    conn.close()


def update_member_name_db(member_id, first_name, last_name):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE members SET first_name=?, last_name=? WHERE id=?", (first_name, last_name, member_id))
    conn.commit()
    conn.close()


def delete_member_db(member_id):
    conn = get_conn()
    cur = conn.cursor()
    # Unassign (rather than delete) any tasks assigned to this member —
    # deleting someone from the roster shouldn't silently delete their
    # to-do items, just detach them.
    cur.execute("UPDATE tasks SET assigned_member_id=NULL WHERE assigned_member_id=?", (member_id,))
    cur.execute("DELETE FROM members WHERE id=?", (member_id,))
    conn.commit()
    conn.close()


def add_member_note_db(member_id, note_text, created_at):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO member_notes (member_id, note_text, created_at) VALUES (?, ?, ?)",
        (member_id, note_text, created_at)
    )
    conn.commit()
    conn.close()


def get_member_notes_db(member_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM member_notes WHERE member_id=? ORDER BY id DESC", (member_id,))
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def delete_member_note_db(note_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM member_notes WHERE id=?", (note_id,))
    conn.commit()
    conn.close()


# ================= MEETINGS =================
def get_all_meetings():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM meetings ORDER BY meeting_date ASC, id DESC")
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def get_meetings_by_date(meeting_date):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM meetings WHERE meeting_date=? ORDER BY id DESC",
        (meeting_date,)
    )
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def add_meeting_db(meeting_date, title, note, created_at):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO meetings (meeting_date, title, note, created_at) VALUES (?, ?, ?, ?)",
        (meeting_date, title, note, created_at)
    )
    conn.commit()
    conn.close()


def delete_meeting_db(meeting_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM meetings WHERE id=?", (meeting_id,))
    conn.commit()
    conn.close()


# ================= STOCKLIST =================
def clear_stock_items_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM stock_items")
    conn.commit()
    conn.close()


def clear_stock_items_for_supplier_db(supplier_id):
    """Re-importing a supplier's price list should only replace THAT
    supplier's rows, not wipe every other supplier's data too."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM stock_items WHERE supplier_id=?", (supplier_id,))
    conn.commit()
    conn.close()


def add_stock_rows_db(rows):
    """rows: (source_file, sheet_name, item_group, item_name, price, imported_at, supplier_id)"""
    conn = get_conn()
    cur = conn.cursor()
    cur.executemany(
        """
        INSERT INTO stock_items (source_file, sheet_name, item_group, item_name, price, imported_at, supplier_id)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        rows
    )
    conn.commit()
    conn.close()


def get_stock_group_stats(supplier_id=None):
    """Per-group Min/Max/Average/Count. With supplier_id=None (the "Group"
    browse mode), this is combined across every supplier — same as the
    original single stocklist view. With a supplier_id given (used by the
    "Supplier" drill-down), it's scoped to just that supplier's rows."""
    conn = get_conn()
    cur = conn.cursor()
    if supplier_id is not None:
        cur.execute("""
            SELECT
                item_group,
                COUNT(*) AS item_count,
                MIN(price) AS min_price,
                MAX(price) AS max_price,
                AVG(price) AS avg_price
            FROM stock_items
            WHERE supplier_id=?
            GROUP BY item_group
            ORDER BY item_group COLLATE NOCASE ASC
        """, (supplier_id,))
    else:
        cur.execute("""
            SELECT
                item_group,
                COUNT(*) AS item_count,
                MIN(price) AS min_price,
                MAX(price) AS max_price,
                AVG(price) AS avg_price
            FROM stock_items
            GROUP BY item_group
            ORDER BY item_group COLLATE NOCASE ASC
        """)
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def get_stock_overall_count():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) AS cnt FROM stock_items")
    row = cur.fetchone()
    conn.close()
    return row["cnt"] if row else 0


def get_suppliers_with_stock_db():
    """Suppliers who currently have at least one stock_items row, with a
    count — the first level of the "Supplier" browse mode."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT s.id, s.name, COUNT(si.id) AS item_count
        FROM suppliers s
        JOIN stock_items si ON si.supplier_id = s.id
        GROUP BY s.id, s.name
        ORDER BY s.name COLLATE NOCASE ASC
    """)
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def get_stock_items_for_group_db(item_group, supplier_id=None):
    """Individual item rows within one group — the deepest level of the
    "Supplier" drill-down (Supplier -> Group -> these items)."""
    conn = get_conn()
    cur = conn.cursor()
    if supplier_id is not None:
        cur.execute("""
            SELECT * FROM stock_items
            WHERE item_group=? AND supplier_id=?
            ORDER BY item_name COLLATE NOCASE ASC
        """, (item_group, supplier_id))
    else:
        cur.execute("""
            SELECT * FROM stock_items
            WHERE item_group=?
            ORDER BY item_name COLLATE NOCASE ASC
        """, (item_group,))
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def get_all_stock_items_flat_db(search=""):
    """Every individual stock item row, joined with its supplier's name —
    the flat, searchable "Item" browse mode. `search` (optional) filters to
    item names or groups containing that text, case-insensitively."""
    conn = get_conn()
    cur = conn.cursor()
    if search:
        like = f"%{search}%"
        cur.execute("""
            SELECT si.*, s.name AS supplier_name
            FROM stock_items si
            LEFT JOIN suppliers s ON s.id = si.supplier_id
            WHERE si.item_name LIKE ? OR si.item_group LIKE ?
            ORDER BY si.item_name COLLATE NOCASE ASC
        """, (like, like))
    else:
        cur.execute("""
            SELECT si.*, s.name AS supplier_name
            FROM stock_items si
            LEFT JOIN suppliers s ON s.id = si.supplier_id
            ORDER BY si.item_name COLLATE NOCASE ASC
        """)
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def open_file(path):
    try:
        if not path:
            return
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.call(["open", path])
        else:
            subprocess.call(["xdg-open", path])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file:\n{e}")


def parse_dnd_files(data):
    """tkinterdnd2 hands drop events a single string that can contain
    multiple file paths, brace-quoted whenever a path has spaces in it
    (e.g. "{/Users/me/My Documents/note.eml} /Users/me/other.pdf"). This
    splits that safely and returns a plain list of paths."""
    root_widget = tk._default_root
    try:
        return list(root_widget.tk.splitlist(data))
    except Exception:
        return [p for p in data.split() if p]


def make_drop_target(widget, on_files_dropped):
    """Register `widget` as a drop target for files dragged in from outside
    the app (Finder/Explorer/Mail/Outlook), calling on_files_dropped(paths)
    with the list of dropped file paths. No-op if tkinterdnd2 isn't
    installed — the app degrades to Browse-button-only in that case."""
    if not DND_AVAILABLE:
        return

    widget.drop_target_register(DND_FILES)

    def handle_drop(event):
        paths = parse_dnd_files(event.data)
        if paths:
            on_files_dropped(paths)

    widget.dnd_bind("<<Drop>>", handle_drop)


def bind_mousewheel_scroll(canvas):
    """Lets the mouse wheel / trackpad two-finger scroll work over a
    scrollable list, instead of forcing the user to drag the scrollbar
    thumb by hand. Only active while the pointer is actually over this
    specific canvas (toggled on Enter/Leave), so scrolling one list doesn't
    also scroll some other list sitting elsewhere in the same window."""
    def _on_mousewheel(event):
        if sys.platform == "darwin":
            canvas.yview_scroll(int(-1 * event.delta), "units")
        else:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_wheel_up(event):
        canvas.yview_scroll(-1, "units")

    def _on_wheel_down(event):
        canvas.yview_scroll(1, "units")

    def _bind(event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        canvas.bind_all("<Button-4>", _on_wheel_up)
        canvas.bind_all("<Button-5>", _on_wheel_down)

    def _unbind(event):
        canvas.unbind_all("<MouseWheel>")
        canvas.unbind_all("<Button-4>")
        canvas.unbind_all("<Button-5>")

    canvas.bind("<Enter>", _bind)
    canvas.bind("<Leave>", _unbind)


def bind_canvas_stretch(canvas, window_id):
    """Every scrollable list in this app (emails, tasks, suppliers, members,
    meetings) is a Frame embedded inside a Canvas for scrolling. Canvas-
    embedded frames don't automatically resize when the canvas itself does —
    so when someone maximizes/fullscreens a window (the boss likes doing
    this), the cards would stay stuck at their original narrow width
    instead of stretching to fill the new space. This keeps the embedded
    frame's width in sync with the canvas's current width on every resize."""
    def _on_configure(event):
        canvas.itemconfig(window_id, width=event.width)
    canvas.bind("<Configure>", _on_configure, add="+")


# ================= SCALABLE FONTS =================
# Every font used in this app is one of these (size, style) combinations.
# They're built as shared tkinter.font.Font objects (not plain ("Arial", N)
# tuples) so that resizing/maximizing a window can grow or shrink ALL of
# them at once by just reconfiguring these objects — every widget using one
# updates automatically, without having to touch each widget individually.
FONT_BASE_SPECS = {
    "f10": (10, {}),
    "f11": (11, {}),
    "f8i": (8, {"slant": "italic"}),
    "f11b": (11, {"weight": "bold"}),
    "f14b": (14, {"weight": "bold"}),
    "f12b": (12, {"weight": "bold"}),
    "f9i": (9, {"slant": "italic"}),
    "f10i": (10, {"slant": "italic"}),
    "f9": (9, {}),
    "f7i": (7, {"slant": "italic"}),
    "f22b": (22, {"weight": "bold"}),
    "f13b": (13, {"weight": "bold"}),
    "f10b": (10, {"weight": "bold"}),
    "f11o": (11, {"overstrike": True}),
}

# Populated by setup_fonts(), which needs an actual Tk root to exist first —
# left empty here (rather than building the Font objects at import time) so
# this module can be read/compiled without a display/root available yet.
FONTS = {}

_current_font_scale = 1.0


def setup_fonts():
    """Create the shared Font objects. Must be called once, right after
    the Tk root window is created (root = tk.Tk() / TkinterDnD.Tk())."""
    import tkinter.font as tkfont
    for key, (size, extra) in FONT_BASE_SPECS.items():
        FONTS[key] = tkfont.Font(family="Arial", size=size, **extra)


def scale_fonts(scale):
    """Resize every shared font to `scale` × its original base size.
    Clamped to a sane range, and skips tiny changes so a slow window-drag
    resize doesn't reconfigure every font on every pixel of movement.
    Returns True if fonts were actually resized, False if skipped."""
    global _current_font_scale
    scale = max(0.85, min(scale, 2.6))
    if abs(scale - _current_font_scale) < 0.03:
        return False
    _current_font_scale = scale
    for key, (base_size, _extra) in FONT_BASE_SPECS.items():
        FONTS[key].configure(size=max(6, round(base_size * scale)))
    return True


def bind_font_scaling(win, base_width, base_height):
    """Scale every shared font proportionally as `win` is resized (e.g. the
    boss maximizing/fullscreening a window), based on how much bigger or
    smaller it currently is than its original size."""
    def _on_configure(event):
        if event.widget is not win:
            return
        # Notice immediately if the user manually toggled fullscreen/maximize
        # on this window, so that state can carry over to every other window
        # in the app (not just the one being closed right now).
        sync_global_maximize_state(win)
        w_ratio = event.width / base_width
        h_ratio = event.height / base_height
        changed = scale_fonts((w_ratio + h_ratio) / 2)
        if changed:
            # Reconfiguring a shared Font object tells every widget using it
            # that its required size changed, but Tk only actually re-runs
            # each frame's pack()/grid() layout on its next idle cycle — not
            # necessarily before the frame that's already mid-resize finishes
            # painting. Without forcing it here, buttons/labels can end up
            # showing bigger text clipped inside an old, smaller box (the
            # padding/frame around it hasn't caught up yet). update_idletasks
            # forces that re-layout to happen immediately instead of lagging
            # a frame behind.
            win.update_idletasks()
    win.bind("<Configure>", _on_configure, add="+")


def shade_color(hex_color, factor):
    """Lighten (factor > 0) or darken (factor < 0) a "#RRGGBB" color by
    `factor` (e.g. -0.15 = 15% darker). Used for flat-button hover states —
    a subtle color shift reads as "hover" the way a 3D bevel used to,
    without bringing back the dated raised/sunken look."""
    hex_color = hex_color.lstrip("#")
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

    def adjust(c):
        if factor >= 0:
            c = c + (255 - c) * factor
        else:
            c = c * (1 + factor)
        return max(0, min(255, round(c)))

    return f"#{adjust(r):02x}{adjust(g):02x}{adjust(b):02x}"


def make_button(parent, text, command, bg="#639922", fg="white",
                 font=None, width=None, height=None, padx=10, pady=4):
    """A colored, clickable Label styled to look like a button.

    On macOS, the native "Aqua" theme ignores custom bg/fg colors on
    tk.Button (a long-standing Tk platform limitation) — buttons always
    render as plain gray/white system buttons no matter what colors are
    configured. Windows' Tk theme does not have this restriction. Using a
    styled Label instead of a real Button sidesteps that limitation and
    renders identically on macOS, Windows, and Linux.

    Flat by design (no raised/sunken 3D bevel) — that beveled look is the
    single biggest "old Windows 95 app" visual cue. Hover is a subtle
    color shift instead of a bevel toggle, matching how flat/modern UIs
    typically indicate hover.

    font defaults to None (rather than a FONTS[...] lookup) because default
    argument values are evaluated once, when the module loads — at which
    point FONTS is still empty (it's only populated once a Tk root exists,
    much later at startup). Resolving the actual default font at call time
    instead means this keeps working AND stays scalable, since FONTS["f10"]
    will have been populated by then.
    """
    font = font or FONTS["f10"]
    hover_bg = shade_color(bg, -0.12)
    kwargs = dict(text=text, bg=bg, fg=fg, font=font, relief="flat",
                  borderwidth=0, cursor="hand2", padx=padx, pady=pady)
    if width is not None:
        kwargs["width"] = width
    if height is not None:
        kwargs["height"] = height

    btn = tk.Label(parent, **kwargs)
    btn.bind("<Button-1>", lambda event: command())
    btn.bind("<Enter>", lambda event: btn.config(bg=hover_bg))
    btn.bind("<Leave>", lambda event: btn.config(bg=bg))
    return btn


def make_icon_button(parent, text, command, fg="#A32D2D", hover_fg="#791F1F"):
    """A plain icon-only button — no filled color box behind it, just the
    glyph itself (e.g. 🗑️ for Delete). A solid red rectangle around a
    trash-can emoji looks heavy-handed since the icon already reads clearly
    as a destructive action on its own; this matches the parent's
    background instead and just changes color slightly on hover."""
    bg = parent.cget("bg")
    btn = tk.Label(parent, text=text, bg=bg, fg=fg, font=FONTS["f14b"],
                   cursor="hand2", padx=4, pady=0)
    btn.bind("<Button-1>", lambda event: command())
    btn.bind("<Enter>", lambda event: btn.config(fg=hover_fg))
    btn.bind("<Leave>", lambda event: btn.config(fg=fg))
    return btn


def open_assignee_picker(parent, current_ids, on_confirm):
    """Popup with a checkbox per registered member, so a task can be
    assigned to any number of people at once (not just one). Calls
    on_confirm(list_of_selected_member_ids) if OK is clicked."""
    dialog = tk.Toplevel(parent)
    dialog.title("Assign to")
    dialog.configure(bg="#EEF2F7")
    dialog.grab_set()
    center_window(dialog, 320, 420)

    tk.Label(dialog, text="Select people:", font=FONTS["f11b"], bg="#EEF2F7").pack(anchor="w", padx=12, pady=(12, 6))

    list_outer = tk.Frame(dialog, bg="#ffffff")
    list_outer.pack(fill="both", expand=True, padx=12)

    canvas = tk.Canvas(list_outer, bg="#ffffff", highlightthickness=0)
    scrollbar = tk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg="#ffffff")

    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    window_id = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    bind_canvas_stretch(canvas, window_id)
    bind_mousewheel_scroll(canvas)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    members = get_all_members()
    current_ids = set(current_ids)
    vars_by_id = {}

    if not members:
        tk.Label(inner, text="Δεν υπάρχουν εργαζόμενοι ακόμα.", bg="#ffffff", fg="#5B7A99", font=FONTS["f9i"]).pack(anchor="w", padx=6, pady=6)
    else:
        for m in members:
            var = tk.BooleanVar(value=m["id"] in current_ids)
            vars_by_id[m["id"]] = var
            tk.Checkbutton(
                inner, text=f"{m['first_name']} {m['last_name']}", variable=var,
                bg="#ffffff", anchor="w", font=FONTS["f10"]
            ).pack(fill="x", padx=6, pady=2)

    def confirm():
        selected = [mid for mid, var in vars_by_id.items() if var.get()]
        on_confirm(selected)
        dialog.destroy()

    btn_row = tk.Frame(dialog, bg="#EEF2F7")
    btn_row.pack(fill="x", padx=12, pady=10)
    make_button(btn_row, text="OK", bg="#639922", fg="white", command=confirm).pack(side="left")
    make_button(btn_row, text="Cancel", bg="#A32D2D", fg="white", command=dialog.destroy).pack(side="left", padx=(6, 0))


# Entry/Text fields elsewhere in this file are created with no explicit
# colors (e.g. tk.Entry(form_frame, width=22)). On some macOS + Tk 9.x
# setups, unstyled Entry/Text widgets pick up the system Dark Mode default
# (black background, black text/cursor) even while the app's own Frames and
# Labels stay correctly light-themed because THEY do set explicit colors.
# That's why entry boxes render as solid black rectangles. Patching the
# constructors here forces white/black defaults on every Entry and Text
# widget without having to edit each of the ~18 call sites individually —
# any call site that DOES pass its own bg/fg still wins (setdefault only
# fills in what's missing).
_original_entry_init = tk.Entry.__init__


def _patched_entry_init(self, master=None, **kwargs):
    kwargs.setdefault("bg", "white")
    kwargs.setdefault("fg", "#0F2A4A")
    kwargs.setdefault("insertbackground", "black")
    kwargs.setdefault("relief", "flat")
    kwargs.setdefault("highlightthickness", 1)
    kwargs.setdefault("highlightbackground", "#C8D4E2")
    kwargs.setdefault("highlightcolor", "#0F2A4A")
    _original_entry_init(self, master, **kwargs)


tk.Entry.__init__ = _patched_entry_init

_original_text_init = tk.Text.__init__


def _patched_text_init(self, master=None, **kwargs):
    kwargs.setdefault("bg", "white")
    kwargs.setdefault("fg", "#0F2A4A")
    kwargs.setdefault("insertbackground", "black")
    kwargs.setdefault("relief", "flat")
    kwargs.setdefault("highlightthickness", 1)
    kwargs.setdefault("highlightbackground", "#C8D4E2")
    kwargs.setdefault("highlightcolor", "#0F2A4A")
    _original_text_init(self, master, **kwargs)


tk.Text.__init__ = _patched_text_init

# Same root cause hits plain Labels and LabelFrame titles too: every Label
# call in this file sets bg= explicitly, but most don't set fg=, so text
# color falls back to this Mac/Tk combo's default — which renders as
# near-invisible pale/white text on the light backgrounds used everywhere
# here (e.g. "Name:", "Job Title:", the "New Employee" section title).
# Default fg to black the same way, without touching any bg or any fg a
# call site already sets explicitly.
_original_label_init = tk.Label.__init__


def _patched_label_init(self, master=None, **kwargs):
    kwargs.setdefault("fg", "#0F2A4A")
    _original_label_init(self, master, **kwargs)


tk.Label.__init__ = _patched_label_init

_original_labelframe_init = tk.LabelFrame.__init__


def _patched_labelframe_init(self, master=None, **kwargs):
    kwargs.setdefault("fg", "#0F2A4A")
    _original_labelframe_init(self, master, **kwargs)


tk.LabelFrame.__init__ = _patched_labelframe_init

# Same fix, same reason, for Checkbutton and Radiobutton — the "Assign to"
# checklist popup (and the Meetings "Show All"/"Search by Date" radios) use
# these with a text= label but no explicit fg=, so they hit the identical
# near-invisible-text bug.
_original_checkbutton_init = tk.Checkbutton.__init__


def _patched_checkbutton_init(self, master=None, **kwargs):
    kwargs.setdefault("fg", "#0F2A4A")
    _original_checkbutton_init(self, master, **kwargs)


tk.Checkbutton.__init__ = _patched_checkbutton_init

_original_radiobutton_init = tk.Radiobutton.__init__


def _patched_radiobutton_init(self, master=None, **kwargs):
    kwargs.setdefault("fg", "#0F2A4A")
    _original_radiobutton_init(self, master, **kwargs)


tk.Radiobutton.__init__ = _patched_radiobutton_init


def center_window(win, width, height):
    """Center a window on screen and lock its minimum size so the manual
    pack()/grid() layouts inside it don't get squashed if someone drags
    the window smaller than it was designed for."""
    win.update_idletasks()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()
    x = max((screen_w - width) // 2, 0)
    y = max((screen_h - height) // 3, 0)
    win.geometry(f"{width}x{height}+{x}+{y}")
    win.minsize(width, height)


def get_maximize_state(win):
    """Detects which kind of "big" state a window is in, if any:
    - "fullscreen": macOS's native green-button fullscreen (its own Space,
      hides the menu bar) — this is the "-fullscreen" wm attribute, NOT the
      same thing as state()=="zoomed". A window in native fullscreen still
      reports state()=="normal", which is why checking state() alone missed
      this case entirely.
    - "zoomed": an ordinary maximize-to-fill-the-screen (no dedicated Space).
    - None: neither.
    """
    try:
        if win.attributes("-fullscreen"):
            return "fullscreen"
    except tk.TclError:
        pass
    try:
        if win.state() == "zoomed":
            return "zoomed"
    except tk.TclError:
        pass
    return None


def apply_maximize_state(win, mode):
    if mode == "fullscreen":
        try:
            win.attributes("-fullscreen", True)
            return
        except tk.TclError:
            pass
    if mode == "zoomed":
        try:
            win.state("zoomed")
        except tk.TclError:
            pass


# Sticky, app-wide fullscreen: once ANY window goes fullscreen/maximized,
# every window opened afterwards (and every window navigated back to)
# picks up the same state, instead of each window having its own
# independent, easily-lost fullscreen state. Updated continuously by
# sync_global_maximize_state() (piggybacked on the font-scaling resize
# handler every window already has), and applied to freshly-created windows
# by apply_global_maximize_state_to().
_global_maximize_state = None


def sync_global_maximize_state(win):
    global _global_maximize_state
    current = get_maximize_state(win)
    if current is not None:
        _global_maximize_state = current


def apply_global_maximize_state_to(win):
    """Applies the sticky global fullscreen/maximize state to `win` right
    now, synchronously — no delay. Safe (and glitch-free) as long as `win`
    is still withdrawn/hidden when this is called: the window manager only
    ever has to handle ONE transition (hidden -> already-fullscreen-visible)
    instead of two (hidden -> normal-visible -> fullscreen-visible), which
    is what caused the earlier visible "flash small, then snap fullscreen"
    glitch."""
    if _global_maximize_state:
        apply_maximize_state(win, _global_maximize_state)


def go_back_to_dashboard(win):
    """Close a section window and bring the main dashboard menu back."""
    sync_global_maximize_state(win)
    win.destroy()
    # root is already withdrawn (it was hidden when this section window
    # opened) — apply the fullscreen state before revealing it so it
    # appears already in its final size, no flash.
    apply_global_maximize_state_to(root)
    root.deiconify()


def open_section_window(title, width, height):
    """Create a section window (Tasks/Suppliers/Members/Meetings/Stocklist):
    hides the main dashboard menu while it's open, and adds a "Back to
    Dashboard" bar at the top so the user can return without hunting for
    the window controls. Closing the window (the OS close button) does the
    same thing as clicking Back, so there's no way to strand the user with
    neither window visible.

    The window stays withdrawn (hidden) while it's being built — the caller
    (open_tasks/open_suppliers/etc.) adds all its own widgets after this
    returns — and is only revealed once, via after_idle, once that's fully
    done. That avoids the window ever being shown at its small default size
    for even a moment before snapping to fullscreen."""
    win = tk.Toplevel(root)
    win.withdraw()
    win.title(title)
    center_window(win, width, height)
    win.configure(bg="#EEF2F7")
    bind_font_scaling(win, width, height)
    apply_global_maximize_state_to(win)

    root.withdraw()
    win.protocol("WM_DELETE_WINDOW", lambda: go_back_to_dashboard(win))

    back_bar = tk.Frame(win, bg="#EEF2F7")
    back_bar.pack(fill="x", padx=10, pady=(10, 0))
    make_button(back_bar, text="← Back to Dashboard", bg="#5B7A99", fg="white",
                command=lambda: go_back_to_dashboard(win)).pack(side="left")

    win.after_idle(win.deiconify)

    return win


# ================= TASKS WINDOW =================
def open_tasks():
    win = open_section_window("Tasks", 850, 550)

    main_frame = tk.Frame(win, bg="#EEF2F7")
    main_frame.pack(fill="both", expand=True)

    left_frame = tk.Frame(main_frame, bg="#E3E9F0", width=400)
    left_frame.pack(side="left", fill="both", expand=True)

    right_frame = tk.Frame(main_frame, bg="#ffffff", width=400)
    right_frame.pack(side="right", fill="both", expand=True)

    tk.Label(left_frame, text="Emails", font=FONTS["f14b"], bg="#E3E9F0").pack(pady=(10, 5))

    new_email_btn_frame = tk.Frame(left_frame, bg="#E3E9F0")
    new_email_btn_frame.pack(pady=(0, 5))

    email_canvas = tk.Canvas(left_frame, bg="#E3E9F0", highlightthickness=0)
    email_scrollbar = tk.Scrollbar(left_frame, orient="vertical", command=email_canvas.yview)
    emails_container = tk.Frame(email_canvas, bg="#E3E9F0")

    emails_container.bind("<Configure>", lambda e: email_canvas.configure(scrollregion=email_canvas.bbox("all")))
    emails_window_id = email_canvas.create_window((0, 0), window=emails_container, anchor="nw")
    email_canvas.configure(yscrollcommand=email_scrollbar.set)
    bind_canvas_stretch(email_canvas, emails_window_id)
    bind_mousewheel_scroll(email_canvas)

    email_canvas.pack(side="left", fill="both", expand=True, padx=10)
    email_scrollbar.pack(side="right", fill="y")

    drag_data = {"text": ""}

    def on_drag_release(event):
        px, py = win.winfo_pointerxy()
        rx1 = right_frame.winfo_rootx()
        ry1 = right_frame.winfo_rooty()
        rx2 = rx1 + right_frame.winfo_width()
        ry2 = ry1 + right_frame.winfo_height()

        if rx1 <= px <= rx2 and ry1 <= py <= ry2 and drag_data["text"]:
            add_task_row(drag_data["text"])

        win.config(cursor="")
        drag_data["text"] = ""

    def on_drag_start(get_text):
        def handler(event):
            drag_data["text"] = get_text()
            win.config(cursor="hand2")
        return handler

    def render_email_card(email_dict):
        card = tk.Frame(emails_container, bg="#EEF2F7", relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#C8D4E2", pady=4, padx=6)
        card.pack(fill="x", pady=4, padx=2)

        def refresh_card():
            for widget in card.winfo_children():
                widget.destroy()

            # Subject gets its own full-width row so a long subject/filename
            # can never squeeze the Edit/Delete buttons off-screen — they
            # live on a separate row below with guaranteed space.
            subject_row = tk.Frame(card, bg="#EEF2F7")
            subject_row.pack(fill="x")

            arrow = "▼" if email_dict["expanded"] else "▶"
            subject_lbl = tk.Label(
                subject_row,
                text=f"{arrow} {email_dict['subject']}",
                font=FONTS["f11b"],
                bg="#EEF2F7",
                anchor="w",
                wraplength=340,
                cursor="hand2",
                justify="left"
            )
            subject_lbl.pack(fill="x")

            def toggle_expand(event=None):
                email_dict["expanded"] = 0 if email_dict["expanded"] else 1
                set_email_expanded_db(email_dict["id"], email_dict["expanded"])
                refresh_card()

            subject_lbl.bind("<Button-1>", toggle_expand)
            subject_lbl.bind("<ButtonPress-1>", on_drag_start(lambda: email_dict["subject"]), add="+")
            subject_lbl.bind("<ButtonRelease-1>", on_drag_release, add="+")

            btns_row = tk.Frame(card, bg="#EEF2F7")
            btns_row.pack(fill="x", pady=(4, 0))

            def edit_email_action():
                new_subject = simpledialog.askstring("Edit Email", "Θέμα:", initialvalue=email_dict["subject"])
                new_body = simpledialog.askstring("Edit Email", "Περιεχόμενο:", initialvalue=email_dict["body"])
                new_sender = simpledialog.askstring("Edit Email", "From:", initialvalue=email_dict["sender"])

                if new_subject:
                    email_dict["subject"] = new_subject
                if new_body is not None:
                    email_dict["body"] = new_body
                if new_sender is not None:
                    email_dict["sender"] = new_sender

                update_email_db(
                    email_dict["id"],
                    email_dict["sender"],
                    email_dict["subject"],
                    email_dict["body"],
                    email_dict["expanded"]
                )
                refresh_card()

            def delete_email_action():
                delete_email_db(email_dict["id"])
                card.destroy()

            make_button(btns_row, text="Edit", width=5, bg="#0F2A4A", fg="white", command=edit_email_action).pack(side="left", padx=2)
            make_icon_button(btns_row, text="🗑️", command=delete_email_action).pack(side="left", padx=2)

            if email_dict["expanded"]:
                tk.Label(card, text=f"Από: {email_dict['sender']}", font=FONTS["f9"], bg="#EEF2F7", anchor="w").pack(fill="x", pady=(4, 2))

                body_lbl = tk.Label(
                    card,
                    text=email_dict["body"],
                    font=FONTS["f10"],
                    bg="#FFFFFF",
                    anchor="w",
                    justify="left",
                    wraplength=320,
                    cursor="hand2",
                    relief="flat",
                    highlightthickness=1,
                    highlightbackground="#C8D4E2",
                    padx=8,
                    pady=6
                )
                body_lbl.pack(fill="x", pady=(0, 4))

                body_lbl.bind("<ButtonPress-1>", on_drag_start(lambda: email_dict["body"]))
                body_lbl.bind("<ButtonRelease-1>", on_drag_release)

                if email_dict.get("file_path"):
                    file_row = tk.Frame(card, bg="#EEF2F7")
                    file_row.pack(fill="x", pady=(0, 2))
                    tk.Label(file_row, text=f"📎 {os.path.basename(email_dict['file_path'])}",
                             font=FONTS["f8i"], bg="#EEF2F7", fg="#3D5A78").pack(side="left")
                    make_button(file_row, text="Open File", width=9, bg="#3D5A78", fg="white",
                                command=lambda: open_file(email_dict["file_path"])).pack(side="left", padx=(6, 0))

        refresh_card()

    def open_new_email_form():
        form = tk.Toplevel(win)
        form.title("Νέο Email")
        center_window(form, 380, 340)
        form.configure(bg="#EEF2F7")
        form.grab_set()

        tk.Label(form, text="Από (email):", font=FONTS["f10"], bg="#EEF2F7", anchor="w").pack(fill="x", padx=15, pady=(15, 2))
        from_entry = tk.Entry(form, font=FONTS["f11"])
        from_entry.insert(0, "user@example.com")
        from_entry.pack(fill="x", padx=15)

        tk.Label(form, text="Subject:", font=FONTS["f10"], bg="#EEF2F7", anchor="w").pack(fill="x", padx=15, pady=(15, 2))
        subject_entry = tk.Entry(form, font=FONTS["f11"])
        subject_entry.pack(fill="x", padx=15)

        tk.Label(form, text="Θέμα / Περιεχόμενο:", font=FONTS["f10"], bg="#EEF2F7", anchor="w").pack(fill="x", padx=15, pady=(15, 2))
        body_text = tk.Text(form, font=FONTS["f11"], height=6, wrap="word")
        body_text.pack(fill="both", expand=True, padx=15)

        def submit_email():
            sender = from_entry.get().strip() or "unknown@example.com"
            subject = subject_entry.get().strip()
            body = body_text.get("1.0", "end").strip()

            if not subject:
                messagebox.showwarning("Προσοχή", "Το Subject είναι υποχρεωτικό.")
                return

            email_id = add_email_db(sender, subject, body, 0)
            render_email_card({
                "id": email_id,
                "sender": sender,
                "subject": subject,
                "body": body,
                "expanded": 0
            })
            form.destroy()

        btn_frame = tk.Frame(form, bg="#EEF2F7")
        btn_frame.pack(pady=15)

        make_button(btn_frame, text="Add", bg="#639922", fg="white", width=10, command=submit_email).pack(side="left", padx=5)
        make_button(btn_frame, text="Cancel", bg="#A32D2D", fg="white", width=10, command=form.destroy).pack(side="left", padx=5)

    def import_email_file(path):
        subject = os.path.basename(path)
        sender = "imported"
        body = f"Imported from file: {os.path.basename(path)}\n\nClick \"Open File\" below to view the original."

        # .eml is a standard, parseable text format (unlike Outlook's binary
        # .msg or a PDF export) — read the real subject/sender/body out of it
        # so the card shows the actual email instead of just the filename.
        if path.lower().endswith(".eml"):
            try:
                with open(path, "rb") as f:
                    msg = BytesParser(policy=policy.default).parse(f)
                subject = msg.get("subject") or subject
                sender = msg.get("from") or sender
                body_part = msg.get_body(preferencelist=("plain", "html"))
                if body_part is not None:
                    body = body_part.get_content().strip()
            except Exception as e:
                messagebox.showwarning(
                    "Προσοχή",
                    f"Δεν ήταν δυνατή η ανάγνωση του περιεχομένου του email, θα προστεθεί μόνο ως συνημμένο.\n{e}"
                )

        email_id = add_email_db(sender, subject, body, 0, file_path=path)
        render_email_card({
            "id": email_id,
            "sender": sender,
            "subject": subject,
            "body": body,
            "expanded": 0,
            "file_path": path
        })

    def upload_email_file():
        path = filedialog.askopenfilename(
            title="Select email file",
            filetypes=[("Email files", "*.eml"), ("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            import_email_file(path)

    def handle_email_drop(paths):
        for path in paths:
            if os.path.isfile(path):
                import_email_file(path)

    email_btns_row = tk.Frame(new_email_btn_frame, bg="#E3E9F0")
    email_btns_row.pack()

    make_button(email_btns_row, text="+ Νέο Email", bg="#639922", fg="white", command=open_new_email_form).pack(side="left", padx=2)
    make_button(email_btns_row, text="📎 Upload Email", bg="#5B7A99", fg="white", command=upload_email_file).pack(side="left", padx=2)

    if DND_AVAILABLE:
        tk.Label(left_frame, text="…or drag an email file here", font=FONTS["f8i"],
                 bg="#E3E9F0", fg="#5B7A99").pack(pady=(2, 4))
        make_drop_target(email_canvas, handle_email_drop)
        make_drop_target(emails_container, handle_email_drop)
        make_drop_target(left_frame, handle_email_drop)

    for existing_email in get_all_emails():
        render_email_card(existing_email)

    tk.Label(right_frame, text="To-Do List", font=FONTS["f14b"], bg="#ffffff").pack(pady=(10, 5))

    entry_frame = tk.Frame(right_frame, bg="#ffffff")
    entry_frame.pack(pady=5, fill="x", padx=10)

    task_entry = tk.Entry(entry_frame, font=FONTS["f11"])
    task_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

    assign_row = tk.Frame(right_frame, bg="#ffffff")
    assign_row.pack(fill="x", padx=10, pady=(0, 5))

    UNASSIGNED_LABEL = "— Unassigned —"
    new_task_assignee_ids = []

    tk.Label(assign_row, text="Assign to:", bg="#ffffff", font=FONTS["f10"]).pack(side="left", padx=(0, 4))
    new_task_assignee_lbl = tk.Label(assign_row, text=UNASSIGNED_LABEL, bg="#ffffff", fg="#5B7A99", font=FONTS["f9i"])
    new_task_assignee_lbl.pack(side="left", padx=(0, 6))

    def names_for_ids(member_ids):
        if not member_ids:
            return UNASSIGNED_LABEL
        id_set = set(member_ids)
        names = [f"{m['first_name']} {m['last_name']}" for m in get_all_members() if m["id"] in id_set]
        return ", ".join(names) if names else UNASSIGNED_LABEL

    def pick_new_task_assignees():
        def on_confirm(selected_ids):
            new_task_assignee_ids.clear()
            new_task_assignee_ids.extend(selected_ids)
            new_task_assignee_lbl.config(text=names_for_ids(selected_ids))
        open_assignee_picker(win, new_task_assignee_ids, on_confirm)

    make_button(assign_row, text="Assign", bg="#5B7A99", fg="white", command=pick_new_task_assignees).pack(side="left")

    task_canvas = tk.Canvas(right_frame, bg="#ffffff", highlightthickness=0)
    task_scrollbar = tk.Scrollbar(right_frame, orient="vertical", command=task_canvas.yview)
    tasks_container = tk.Frame(task_canvas, bg="#ffffff")

    tasks_container.bind("<Configure>", lambda e: task_canvas.configure(scrollregion=task_canvas.bbox("all")))
    tasks_window_id = task_canvas.create_window((0, 0), window=tasks_container, anchor="nw")
    task_canvas.configure(yscrollcommand=task_scrollbar.set)
    bind_canvas_stretch(task_canvas, tasks_window_id)
    bind_mousewheel_scroll(task_canvas)

    task_canvas.pack(side="left", fill="both", expand=True, padx=10)
    task_scrollbar.pack(side="right", fill="y")

    def render_task_row(task_dict):
        row_outer = tk.Frame(tasks_container, bg="#ffffff", relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#C8D4E2", pady=4, padx=4)
        row_outer.pack(fill="x", pady=4, padx=2)

        def refresh_row():
            for widget in row_outer.winfo_children():
                widget.destroy()

            top_row = tk.Frame(row_outer, bg="#ffffff")
            top_row.pack(fill="x")

            done_var = tk.BooleanVar(value=bool(task_dict["done"]))

            def toggle_done():
                # Keeps status in sync with the checkbox so this stays
                # consistent with the Active/Previous split shown on a
                # member's own profile page (dismissed can only be set
                # from there, since this checkbox is just done/not-done).
                new_status = "completed" if done_var.get() else "active"
                set_task_status_db(task_dict["id"], new_status)
                task_dict["done"] = 1 if done_var.get() else 0
                task_dict["status"] = new_status
                refresh_row()

            tk.Checkbutton(top_row, variable=done_var, bg="#ffffff", command=toggle_done).pack(side="left")

            task_font = FONTS["f11o"] if task_dict["done"] else FONTS["f11"]
            task_color = "#7A93AC" if task_dict["done"] else "#0F2A4A"

            tk.Label(top_row, text=task_dict["text"], font=task_font, fg=task_color, bg="#ffffff", anchor="w", justify="left", wraplength=220).pack(side="left", fill="x", expand=True)

            def edit_task_action():
                new_text = simpledialog.askstring("Edit Task", "Επεξεργασία task:", initialvalue=task_dict["text"])
                if new_text:
                    task_dict["text"] = new_text
                    update_task_db(task_dict["id"], new_text)
                    refresh_row()

            def delete_task_action():
                delete_task_db(task_dict["id"])
                row_outer.destroy()

            make_button(top_row, text="Edit", width=5, bg="#0F2A4A", fg="white", command=edit_task_action).pack(side="left", padx=2)
            make_icon_button(top_row, text="🗑️", command=delete_task_action).pack(side="left", padx=2)

            bottom_row = tk.Frame(row_outer, bg="#ffffff")
            bottom_row.pack(fill="x", pady=(2, 0))
            tk.Label(bottom_row, text=f"📅 Καταχωρήθηκε: {task_dict['date']}", font=FONTS["f8i"], fg="#5B7A99", bg="#ffffff").pack(side="left")

            assignee_row = tk.Frame(row_outer, bg="#ffffff")
            assignee_row.pack(fill="x", pady=(2, 0))

            tk.Label(assignee_row, text="👤", bg="#ffffff", font=FONTS["f8i"]).pack(side="left")

            current_assignees = get_task_assignees_db(task_dict["id"])
            current_assignee_ids = [m["id"] for m in current_assignees]
            assignee_names_lbl = tk.Label(
                assignee_row,
                text=names_for_ids(current_assignee_ids),
                bg="#ffffff", fg="#5B7A99", font=FONTS["f8i"], anchor="w"
            )
            assignee_names_lbl.pack(side="left", padx=(2, 6))

            def pick_row_assignees(task_id=task_dict["id"], ids=current_assignee_ids):
                def on_confirm(selected_ids):
                    set_task_assignees_db(task_id, selected_ids)
                    refresh_row()
                open_assignee_picker(win, ids, on_confirm)

            make_button(assignee_row, text="Assign", width=6, bg="#5B7A99", fg="white", command=pick_row_assignees).pack(side="left")

            attach_row = tk.Frame(row_outer, bg="#ffffff")
            attach_row.pack(fill="x", pady=(4, 0))

            attachment_path = task_dict.get("attachment_path")
            attachment_name = os.path.basename(attachment_path) if attachment_path else "— none —"

            tk.Label(attach_row, text="📎 Email:", font=FONTS["f8i"], bg="#ffffff", fg="#5B7A99").pack(side="left")
            tk.Label(attach_row, text=attachment_name, font=FONTS["f8i"], bg="#ffffff",
                     fg="#3D5A78" if attachment_path else "#7A93AC").pack(side="left", padx=(4, 8))

            def set_attachment(path):
                set_task_attachment_db(task_dict["id"], path)
                task_dict["attachment_path"] = path
                refresh_row()

            def browse_attachment():
                chosen = filedialog.askopenfilename(
                    title="Select email file",
                    filetypes=[("Email files", "*.eml"), ("PDF files", "*.pdf"), ("All files", "*.*")]
                )
                if chosen:
                    set_attachment(chosen)

            def handle_attachment_drop(paths):
                for path in paths:
                    if os.path.isfile(path):
                        set_attachment(path)
                        break

            def open_attachment():
                if attachment_path:
                    open_file(attachment_path)
                else:
                    messagebox.showinfo("Email", "Δεν έχει επισυναφθεί email ακόμα.")

            make_button(attach_row, text="Browse", width=6, bg="#5B7A99", fg="white", command=browse_attachment).pack(side="left", padx=1)
            make_button(attach_row, text="Open", width=5, bg="#3D5A78", fg="white", command=open_attachment).pack(side="left", padx=1)

            if DND_AVAILABLE:
                tk.Label(attach_row, text="(or drop here)", font=FONTS["f7i"], bg="#ffffff", fg="#7A93AC").pack(side="left", padx=(6, 0))
                make_drop_target(row_outer, handle_attachment_drop)
                make_drop_target(attach_row, handle_attachment_drop)

        refresh_row()

    def add_task_row(text):
        date_str = datetime.now().strftime("%d/%m/%Y %H:%M")
        task_id = add_task_db(text, date_str, 0)

        if new_task_assignee_ids:
            set_task_assignees_db(task_id, new_task_assignee_ids)

        render_task_row({
            "id": task_id,
            "text": text,
            "date": date_str,
            "done": 0,
            "status": "active",
            "attachment_path": None
        })

    def add_task_action(event=None):
        text = task_entry.get().strip()
        if not text:
            messagebox.showwarning("Προσοχή", "Γράψε ένα task πρώτα.")
            return
        add_task_row(text)
        task_entry.delete(0, tk.END)
        new_task_assignee_ids.clear()
        new_task_assignee_lbl.config(text=UNASSIGNED_LABEL)

    make_button(entry_frame, text="Add Task", bg="#639922", fg="white", command=add_task_action).pack(side="left")
    task_entry.bind("<Return>", add_task_action)

    for existing_task in get_all_tasks():
        render_task_row(existing_task)


# ================= MEMBERS WINDOW =================
def open_members():
    win = open_section_window("Members", 850, 600)

    form_frame = tk.LabelFrame(win, text="New Employee", font=FONTS["f12b"], bg="#EEF2F7", padx=14, pady=12)
    form_frame.pack(fill="x", padx=12, pady=12)

    tk.Label(form_frame, text="Name:", bg="#EEF2F7").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=6)
    first_name_entry = tk.Entry(form_frame, width=20)
    first_name_entry.grid(row=0, column=1, padx=(0, 20), pady=6)

    tk.Label(form_frame, text="Surname:", bg="#EEF2F7").grid(row=0, column=2, sticky="w", padx=(0, 4), pady=6)
    last_name_entry = tk.Entry(form_frame, width=20)
    last_name_entry.grid(row=0, column=3, padx=0, pady=6)

    tk.Label(form_frame, text="Job Title:", bg="#EEF2F7").grid(row=1, column=0, sticky="w", padx=(0, 4), pady=6)
    job_title_entry = tk.Entry(form_frame, width=20)
    job_title_entry.grid(row=1, column=1, padx=(0, 20), pady=6)

    current_date = datetime.now().strftime("%d/%m/%Y")
    tk.Label(form_frame, text="Date:", bg="#EEF2F7").grid(row=1, column=2, sticky="w", padx=(0, 4), pady=6)
    tk.Label(form_frame, text=current_date, bg="#ffffff", relief="flat", highlightthickness=1, highlightbackground="#C8D4E2", width=18, anchor="w").grid(row=1, column=3, padx=0, pady=6)

    search_row = tk.Frame(win, bg="#EEF2F7")
    search_row.pack(fill="x", padx=14, pady=(4, 0))

    tk.Label(search_row, text="Employees", font=FONTS["f13b"], bg="#EEF2F7").pack(side="left")

    tk.Label(search_row, text="Search:", bg="#EEF2F7", font=FONTS["f10"]).pack(side="left", padx=(20, 4))
    search_entry = tk.Entry(search_row, width=22)
    search_entry.pack(side="left")

    list_outer = tk.Frame(win, bg="#EEF2F7")
    list_outer.pack(fill="both", expand=True, padx=12, pady=(8, 12))

    canvas = tk.Canvas(list_outer, bg="#ffffff", highlightthickness=0)
    scrollbar = tk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
    members_container = tk.Frame(canvas, bg="#ffffff")

    members_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    members_window_id = canvas.create_window((0, 0), window=members_container, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    bind_canvas_stretch(canvas, members_window_id)
    bind_mousewheel_scroll(canvas)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def refresh_members():
        for widget in members_container.winfo_children():
            widget.destroy()

        members = get_all_members(search=search_entry.get().strip())

        if not members:
            tk.Label(members_container, text="Δεν βρέθηκαν εργαζόμενοι.", bg="#ffffff", fg="#5B7A99", font=FONTS["f10"]).pack(anchor="w", padx=8, pady=8)
            return

        # One row per employee, stacked top to bottom, alphabetical
        # (get_all_members already orders by first/last name).
        for member in members:
            row = tk.Frame(members_container, bg="#EEF2F7", relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#C8D4E2", padx=10, pady=8)
            row.pack(fill="x", pady=3, padx=4)

            info = tk.Frame(row, bg="#EEF2F7", cursor="hand2")
            info.pack(side="left", fill="x", expand=True)

            name_lbl = tk.Label(info, text=f"{member['first_name']} {member['last_name']}",
                                 font=FONTS["f11b"], bg="#EEF2F7", anchor="w", cursor="hand2")
            name_lbl.pack(fill="x")
            job_lbl = tk.Label(info, text=member["job_title"], font=FONTS["f10"],
                                bg="#EEF2F7", fg="#5B7A99", anchor="w", cursor="hand2")
            job_lbl.pack(fill="x", pady=(2, 0))

            def open_profile(member_id=member["id"]):
                open_member_profile_window(member_id, win, refresh_members)

            info.bind("<Button-1>", lambda e, f=open_profile: f())
            name_lbl.bind("<Button-1>", lambda e, f=open_profile: f())
            job_lbl.bind("<Button-1>", lambda e, f=open_profile: f())

            def delete_action(member_id=member["id"]):
                if not messagebox.askyesno("Delete", f"Διαγραφή {member['first_name']} {member['last_name']};"):
                    return
                delete_member_db(member_id)
                refresh_members()

            make_icon_button(row, text="🗑️", command=delete_action).pack(side="right")

    def add_member_action():
        first_name = to_title_case(first_name_entry.get())
        last_name = to_title_case(last_name_entry.get())
        job_title = to_title_case(job_title_entry.get())
        created_date = datetime.now().strftime("%d/%m/%Y")

        if not first_name or not last_name or not job_title:
            messagebox.showwarning("Προσοχή", "Συμπλήρωσε Name, Surname και Job Title.")
            return

        if member_exists_db(first_name, last_name):
            messagebox.showwarning("Προσοχή", "User already exists.")
            return

        add_member_db(first_name, last_name, job_title, created_date)

        first_name_entry.delete(0, tk.END)
        last_name_entry.delete(0, tk.END)
        job_title_entry.delete(0, tk.END)

        refresh_members()

    make_button(form_frame, text="Add Member", bg="#639922", fg="white", command=add_member_action).grid(row=2, column=3, pady=(10, 0), sticky="e")

    search_entry.bind("<KeyRelease>", lambda e: refresh_members())

    refresh_members()


# ================= MEMBER PROFILE WINDOW =================
def open_member_profile_window(member_id, members_win, on_close_refresh):
    """Detail view for a single employee: dated notes log, and their tasks
    split into Active vs Previous (completed/dismissed). Opened from a
    click on a row in the Members list; hides the Members window while
    open and returns to it (re-running on_close_refresh so any changes —
    e.g. deleted notes/tasks — show up immediately) on Back/close."""
    member = get_member_by_id(member_id)
    if member is None:
        messagebox.showerror("Error", "Αυτός ο εργαζόμενος δεν υπάρχει πια.")
        return

    win = tk.Toplevel(members_win)
    win.withdraw()
    win.title(f"{member['first_name']} {member['last_name']}")
    center_window(win, 900, 650)
    win.configure(bg="#EEF2F7")
    bind_font_scaling(win, 900, 650)
    apply_global_maximize_state_to(win)

    members_win.withdraw()

    def go_back():
        sync_global_maximize_state(win)
        win.destroy()
        apply_global_maximize_state_to(members_win)
        members_win.deiconify()
        on_close_refresh()

    win.protocol("WM_DELETE_WINDOW", go_back)

    back_bar = tk.Frame(win, bg="#EEF2F7")
    back_bar.pack(fill="x", padx=10, pady=(10, 0))
    make_button(back_bar, text="← Back to Members", bg="#5B7A99", fg="white", command=go_back).pack(side="left")

    header = tk.Frame(win, bg="#EEF2F7")
    header.pack(fill="x", padx=14, pady=(10, 6))

    name_row = tk.Frame(header, bg="#EEF2F7")
    name_row.pack(anchor="w")

    name_lbl = tk.Label(name_row, text=f"{member['first_name']} {member['last_name']}", font=FONTS["f22b"], bg="#EEF2F7")
    name_lbl.pack(side="left")

    def edit_name_action():
        new_first = simpledialog.askstring("Edit Name", "Name:", initialvalue=member["first_name"])
        if new_first is None:
            return
        new_first = to_title_case(new_first)

        new_last = simpledialog.askstring("Edit Name", "Surname:", initialvalue=member["last_name"])
        if new_last is None:
            return
        new_last = to_title_case(new_last)

        if not new_first or not new_last:
            messagebox.showwarning("Προσοχή", "Το Name και το Surname δεν μπορούν να είναι κενά.")
            return

        if member_exists_db(new_first, new_last, exclude_id=member_id):
            messagebox.showwarning("Προσοχή", "User already exists.")
            return

        update_member_name_db(member_id, new_first, new_last)
        member["first_name"] = new_first
        member["last_name"] = new_last
        name_lbl.config(text=f"{new_first} {new_last}")
        win.title(f"{new_first} {new_last}")

    make_button(name_row, text="Edit", bg="#0F2A4A", fg="white", command=edit_name_action).pack(side="left", padx=(8, 0))

    job_title_row = tk.Frame(header, bg="#EEF2F7")
    job_title_row.pack(anchor="w")

    job_title_lbl = tk.Label(job_title_row, text=member["job_title"], font=FONTS["f12b"], bg="#EEF2F7", fg="#5B7A99")
    job_title_lbl.pack(side="left")

    def edit_job_title_action():
        new_title = simpledialog.askstring("Edit Job Title", "Job Title:", initialvalue=member["job_title"])
        if new_title is None:
            return
        new_title = to_title_case(new_title)
        if not new_title:
            messagebox.showwarning("Προσοχή", "Το Job Title δεν μπορεί να είναι κενό.")
            return
        update_member_job_title_db(member_id, new_title)
        member["job_title"] = new_title
        job_title_lbl.config(text=new_title)

    make_button(job_title_row, text="Edit", bg="#0F2A4A", fg="white", command=edit_job_title_action).pack(side="left", padx=(8, 0))

    tk.Label(header, text=f"Registered: {member['created_date']}", font=FONTS["f9i"], bg="#EEF2F7", fg="#7A93AC").pack(anchor="w", pady=(2, 0))

    body = tk.Frame(win, bg="#EEF2F7")
    body.pack(fill="both", expand=True, padx=14, pady=(4, 14))

    # -------- LEFT: dated notes --------
    notes_frame = tk.LabelFrame(body, text="Notes", font=FONTS["f12b"], bg="#EEF2F7", padx=10, pady=10)
    notes_frame.pack(side="left", fill="both", expand=True, padx=(0, 8))

    note_entry_row = tk.Frame(notes_frame, bg="#EEF2F7")
    note_entry_row.pack(fill="x", pady=(0, 8))

    note_text = tk.Text(note_entry_row, height=3, font=FONTS["f10"])
    note_text.pack(fill="x")

    notes_list_outer = tk.Frame(notes_frame, bg="#EEF2F7")
    notes_list_outer.pack(fill="both", expand=True)

    notes_canvas = tk.Canvas(notes_list_outer, bg="#ffffff", highlightthickness=0)
    notes_scroll = tk.Scrollbar(notes_list_outer, orient="vertical", command=notes_canvas.yview)
    notes_container = tk.Frame(notes_canvas, bg="#ffffff")

    notes_container.bind("<Configure>", lambda e: notes_canvas.configure(scrollregion=notes_canvas.bbox("all")))
    notes_window_id = notes_canvas.create_window((0, 0), window=notes_container, anchor="nw")
    notes_canvas.configure(yscrollcommand=notes_scroll.set)
    bind_canvas_stretch(notes_canvas, notes_window_id)
    bind_mousewheel_scroll(notes_canvas)

    notes_canvas.pack(side="left", fill="both", expand=True)
    notes_scroll.pack(side="right", fill="y")

    def refresh_notes():
        for widget in notes_container.winfo_children():
            widget.destroy()

        notes = get_member_notes_db(member_id)
        if not notes:
            tk.Label(notes_container, text="Δεν υπάρχουν σημειώσεις ακόμα.", bg="#ffffff", fg="#5B7A99", font=FONTS["f9i"]).pack(anchor="w", padx=6, pady=6)
            return

        for note in notes:
            box = tk.Frame(notes_container, bg="#EEF2F7", relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#C8D4E2", padx=6, pady=4)
            box.pack(fill="x", padx=4, pady=3)

            tk.Label(box, text=note["created_at"], font=FONTS["f8i"], bg="#EEF2F7", fg="#5B7A99", anchor="w").pack(fill="x")
            tk.Label(box, text=note["note_text"], font=FONTS["f10"], bg="#EEF2F7", anchor="w", justify="left", wraplength=260).pack(fill="x", pady=(2, 2))

            def delete_note_action(note_id=note["id"]):
                delete_member_note_db(note_id)
                refresh_notes()

            make_icon_button(box, text="🗑️", command=delete_note_action).pack(anchor="e")

    def add_note_action():
        text = note_text.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Προσοχή", "Γράψε κάτι πρώτα.")
            return
        add_member_note_db(member_id, text, datetime.now().strftime("%d/%m/%Y %H:%M"))
        note_text.delete("1.0", tk.END)
        refresh_notes()

    make_button(note_entry_row, text="Add Note", bg="#639922", fg="white", command=add_note_action).pack(anchor="e", pady=(4, 0))

    # -------- RIGHT: tasks (Active / Previous) --------
    tasks_frame = tk.LabelFrame(body, text="Tasks", font=FONTS["f12b"], bg="#EEF2F7", padx=10, pady=10)
    tasks_frame.pack(side="right", fill="both", expand=True, padx=(8, 0))

    task_view_mode = tk.StringVar(value="active")

    tasks_toggle_row = tk.Frame(tasks_frame, bg="#EEF2F7")
    tasks_toggle_row.pack(fill="x", pady=(0, 8))

    def set_mode(mode):
        task_view_mode.set(mode)
        refresh_tasks()

    tasks_list_outer = tk.Frame(tasks_frame, bg="#EEF2F7")
    tasks_list_outer.pack(fill="both", expand=True)

    tasks_canvas = tk.Canvas(tasks_list_outer, bg="#ffffff", highlightthickness=0)
    tasks_scroll = tk.Scrollbar(tasks_list_outer, orient="vertical", command=tasks_canvas.yview)
    member_tasks_container = tk.Frame(tasks_canvas, bg="#ffffff")

    member_tasks_container.bind("<Configure>", lambda e: tasks_canvas.configure(scrollregion=tasks_canvas.bbox("all")))
    tasks_window_id = tasks_canvas.create_window((0, 0), window=member_tasks_container, anchor="nw")
    tasks_canvas.configure(yscrollcommand=tasks_scroll.set)
    bind_canvas_stretch(tasks_canvas, tasks_window_id)
    bind_mousewheel_scroll(tasks_canvas)

    tasks_canvas.pack(side="left", fill="both", expand=True)
    tasks_scroll.pack(side="right", fill="y")

    STATUS_LABELS = {"completed": "✅ Completed", "dismissed": "🚫 Dismissed"}

    ACTIVE_BTN_COLOR = "#0F2A4A"
    PREVIOUS_BTN_COLOR = "#5B7A99"
    MUTED_TAB_COLOR = shade_color("#EEF2F7", -0.05)  # faint, for the unselected tab

    def refresh_tasks():
        for widget in member_tasks_container.winfo_children():
            widget.destroy()

        # Flat "selected tab" look: rebuild the toggle buttons each time
        # with full color for whichever mode is active and muted gray for
        # the other — a make_button's hover binding locks in its bg at
        # creation time, so reusing the same button widgets and just
        # .config()-ing their color would leave hover fighting the new
        # color; rebuilding fresh each time sidesteps that entirely, and
        # replaces the old raised/sunken bevel toggle with a modern flat-UI
        # "selected tab" convention.
        for widget in tasks_toggle_row.winfo_children():
            widget.destroy()

        is_active = task_view_mode.get() == "active"

        make_button(
            tasks_toggle_row, text="Active Tasks",
            bg=ACTIVE_BTN_COLOR if is_active else MUTED_TAB_COLOR,
            fg="white" if is_active else "#5B7A99",
            command=lambda: set_mode("active")
        ).pack(side="left", padx=(0, 4))

        make_button(
            tasks_toggle_row, text="Previous Tasks",
            bg=MUTED_TAB_COLOR if is_active else PREVIOUS_BTN_COLOR,
            fg="#5B7A99" if is_active else "white",
            command=lambda: set_mode("previous")
        ).pack(side="left")

        if is_active:
            tasks = get_tasks_for_member_db(member_id, ["active"])
        else:
            tasks = get_tasks_for_member_db(member_id, ["completed", "dismissed"])

        if not tasks:
            msg = "Δεν υπάρχουν ενεργά tasks." if task_view_mode.get() == "active" else "Δεν υπάρχουν προηγούμενα tasks."
            tk.Label(member_tasks_container, text=msg, bg="#ffffff", fg="#5B7A99", font=FONTS["f9i"]).pack(anchor="w", padx=6, pady=6)
            return

        for task in tasks:
            box = tk.Frame(member_tasks_container, bg="#EEF2F7", relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#C8D4E2", padx=8, pady=6)
            box.pack(fill="x", padx=4, pady=3)

            tk.Label(box, text=task["text"], font=FONTS["f10"], bg="#EEF2F7", anchor="w", justify="left", wraplength=260).pack(fill="x")
            tk.Label(box, text=f"📅 {task['date']}", font=FONTS["f8i"], bg="#EEF2F7", fg="#5B7A99", anchor="w").pack(fill="x", pady=(2, 0))

            actions = tk.Frame(box, bg="#EEF2F7")
            actions.pack(fill="x", pady=(4, 0))

            if task_view_mode.get() == "active":
                def complete_action(task_id=task["id"]):
                    set_task_status_db(task_id, "completed")
                    refresh_tasks()

                def dismiss_action(task_id=task["id"]):
                    set_task_status_db(task_id, "dismissed")
                    refresh_tasks()

                make_button(actions, text="✓ Complete", bg="#639922", fg="white", command=complete_action).pack(side="left", padx=(0, 4))
                make_button(actions, text="✕ Dismiss", bg="#A32D2D", fg="white", command=dismiss_action).pack(side="left")
            else:
                tk.Label(actions, text=STATUS_LABELS.get(task["status"], task["status"]), font=FONTS["f9i"], bg="#EEF2F7").pack(side="left")

                def reopen_action(task_id=task["id"]):
                    set_task_status_db(task_id, "active")
                    refresh_tasks()

                make_button(actions, text="↺ Reopen", bg="#5B7A99", fg="white", command=reopen_action).pack(side="right")

    refresh_notes()
    refresh_tasks()

    win.after_idle(win.deiconify)


# ================= MEETINGS WINDOW =================
def open_meetings():
    win = open_section_window("Meetings", 900, 600)

    left = tk.Frame(win, bg="#EEF2F7", width=260)
    left.pack(side="left", fill="y", padx=10, pady=10)

    right = tk.Frame(win, bg="#ffffff")
    right.pack(side="right", fill="both", expand=True, padx=10, pady=10)

    tk.Label(left, text="Add Meeting", font=FONTS["f14b"], bg="#EEF2F7").pack(anchor="w", pady=(0, 10))

    tk.Label(left, text="Date (YYYY-MM-DD):", bg="#EEF2F7", font=FONTS["f10"]).pack(anchor="w")
    date_entry = tk.Entry(left, width=24, font=FONTS["f11"])
    date_entry.pack(anchor="w", pady=(2, 10))
    date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

    tk.Label(left, text="Title:", bg="#EEF2F7", font=FONTS["f10"]).pack(anchor="w")
    title_entry = tk.Entry(left, width=24, font=FONTS["f11"])
    title_entry.pack(anchor="w", pady=(2, 10))

    tk.Label(left, text="Note:", bg="#EEF2F7", font=FONTS["f10"]).pack(anchor="w")
    note_text = tk.Text(left, width=28, height=8, font=FONTS["f10"])
    note_text.pack(anchor="w", pady=(2, 10))

    tk.Label(right, text="Meetings", font=FONTS["f14b"], bg="#ffffff").pack(anchor="w", padx=10, pady=(10, 8))

    filter_frame = tk.Frame(right, bg="#ffffff")
    filter_frame.pack(fill="x", padx=10, pady=(0, 8))

    view_mode = tk.StringVar(value="all")

    tk.Radiobutton(
        filter_frame,
        text="Show All",
        variable=view_mode,
        value="all",
        bg="#ffffff",
        command=lambda: apply_filter()
    ).pack(side="left", padx=(0, 10))

    tk.Radiobutton(
        filter_frame,
        text="Search by Date",
        variable=view_mode,
        value="date",
        bg="#ffffff",
        command=lambda: apply_filter()
    ).pack(side="left", padx=(0, 10))

    tk.Label(filter_frame, text="Date:", bg="#ffffff", font=FONTS["f10"]).pack(side="left")
    filter_entry = tk.Entry(filter_frame, width=16, font=FONTS["f11"])
    filter_entry.pack(side="left", padx=(6, 8))
    filter_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

    tk.Button(filter_frame, text="Apply", bg="#0F2A4A", fg="white", width=10, command=lambda: apply_filter()).pack(side="left")

    selected_date_label = tk.Label(right, text="", font=FONTS["f11b"], bg="#ffffff", fg="#0F2A4A")
    selected_date_label.pack(anchor="w", padx=10, pady=(0, 5))

    list_frame = tk.Frame(right, bg="#ffffff")
    list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    canvas = tk.Canvas(list_frame, bg="#ffffff", highlightthickness=0)
    scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    meetings_container = tk.Frame(canvas, bg="#ffffff")

    meetings_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    meetings_window_id = canvas.create_window((0, 0), window=meetings_container, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    bind_canvas_stretch(canvas, meetings_window_id)
    bind_mousewheel_scroll(canvas)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def render_meetings(meetings, label_text):
        for widget in meetings_container.winfo_children():
            widget.destroy()

        selected_date_label.config(text=label_text)

        if not meetings:
            tk.Label(
                meetings_container,
                text="Δεν υπάρχουν meetings για αυτό το φίλτρο.",
                bg="#ffffff",
                fg="#5B7A99",
                font=FONTS["f10"]
            ).pack(anchor="w", pady=6)
            return

        for meeting in meetings:
            box = tk.Frame(meetings_container, bg="#EEF2F7", relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#C8D4E2", padx=8, pady=6)
            box.pack(fill="x", pady=4)

            tk.Label(
                box,
                text=f"{meeting['meeting_date']} - {meeting['title']}",
                font=FONTS["f11b"],
                bg="#EEF2F7",
                anchor="w"
            ).pack(fill="x")

            tk.Label(
                box,
                text=meeting["note"] or "-",
                font=FONTS["f10"],
                bg="#EEF2F7",
                justify="left",
                anchor="w",
                wraplength=500
            ).pack(fill="x", pady=(4, 2))

            tk.Label(
                box,
                text=f"Καταχώριση: {meeting['created_at']}",
                font=FONTS["f8i"],
                bg="#EEF2F7",
                fg="#5B7A99"
            ).pack(anchor="w")

            def delete_action(meeting_id=meeting["id"]):
                delete_meeting_db(meeting_id)
                apply_filter()

            make_icon_button(box, text="🗑️", command=delete_action).pack(anchor="e", pady=(5, 0))

    def apply_filter(event=None):
        mode = view_mode.get()

        if mode == "all":
            meetings = get_all_meetings()
            render_meetings(meetings, "Showing all meetings")
            return

        selected_date = filter_entry.get().strip()

        try:
            datetime.strptime(selected_date, "%Y-%m-%d")
        except ValueError:
            selected_date_label.config(text="Λάθος μορφή ημερομηνίας. Βάλε YYYY-MM-DD")
            for widget in meetings_container.winfo_children():
                widget.destroy()
            return

        meetings = get_meetings_by_date(selected_date)
        render_meetings(meetings, f"Showing meetings for: {selected_date}")

    def add_meeting_action():
        meeting_date = date_entry.get().strip()
        title = title_entry.get().strip()
        note = note_text.get("1.0", "end").strip()
        created_at = datetime.now().strftime("%d/%m/%Y %H:%M")

        if not title:
            messagebox.showwarning("Προσοχή", "Το title του meeting είναι υποχρεωτικό.")
            return

        try:
            datetime.strptime(meeting_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Προσοχή", "Η ημερομηνία πρέπει να είναι στη μορφή YYYY-MM-DD.")
            return

        add_meeting_db(meeting_date, title, note, created_at)

        title_entry.delete(0, tk.END)
        note_text.delete("1.0", tk.END)

        filter_entry.delete(0, tk.END)
        filter_entry.insert(0, meeting_date)
        view_mode.set("date")

        apply_filter()

    tk.Button(left, text="Add Meeting", bg="#639922", fg="white", width=16, command=add_meeting_action).pack(anchor="w", pady=8)
    tk.Button(top_filter, text="View", bg="#0F2A4A", fg="white", width=10, command=view_selected_date).pack(side="left")

    filter_entry.bind("<Return>", apply_filter)

    apply_filter()


# ================= STOCKLIST WINDOW =================
def open_stocklist():
    win = open_section_window("Stocklist", 1100, 700)

    selected_file = {"path": None}
    excel_data = {"sheets": {}}

    top_frame = tk.LabelFrame(win, text="Import Excel Stocklist", font=FONTS["f12b"], bg="#EEF2F7", padx=10, pady=10)
    top_frame.pack(fill="x", padx=10, pady=10)

    row1 = tk.Frame(top_frame, bg="#EEF2F7")
    row1.pack(fill="x", pady=4)

    file_label = tk.Label(row1, text="No file selected", bg="#ffffff", relief="flat", highlightthickness=1, highlightbackground="#C8D4E2", anchor="w", width=70)
    file_label.pack(side="left", padx=(0, 8), fill="x", expand=True)

    sheet_var = tk.StringVar(value="")
    group_var = tk.StringVar(value="")
    price_var = tk.StringVar(value="")
    name_var = tk.StringVar(value="")
    import_supplier_var = tk.StringVar(value="")
    import_supplier_map = {}

    def update_option_menu(menu_widget, variable, values, default_empty=True):
        menu = menu_widget["menu"]
        menu.delete(0, "end")

        options = []
        if default_empty:
            options.append("")
        options.extend(values)

        for val in options:
            menu.add_command(label=val, command=lambda v=val: variable.set(v))

        if options:
            variable.set(options[0])
        else:
            variable.set("")

    def detect_columns(columns):
        cols_lower = {c.lower(): c for c in columns}

        group_candidates = [
            "title group",
            "group",
            "groups",
            "category",
            "product group",
            "item group",
            "family",
            "type",
            "brand",
            "department",
            "classification"
        ]
        price_candidates = [
            "finalprice",
            "final price",
            "price eur",
            "price",
            "unit price",
            "cost",
            "sale price",
            "stock price",
            "amount",
            "value",
            "net price"
        ]
        name_candidates = [
            "title",
            "item",
            "item name",
            "name",
            "description",
            "product",
            "product name",
            "group form title",
            "groupformtitle",
            "order form title",
            "orderformtitle",
            "form title"
        ]

        detected_group = ""
        detected_price = ""
        detected_name = ""

        for candidate in group_candidates:
            if candidate in cols_lower:
                detected_group = cols_lower[candidate]
                break

        for candidate in price_candidates:
            if candidate in cols_lower:
                detected_price = cols_lower[candidate]
                break

        for candidate in name_candidates:
            if candidate in cols_lower:
                detected_name = cols_lower[candidate]
                break

        if not detected_group and columns:
            detected_group = columns[0]

        if not detected_price and len(columns) >= 2:
            detected_price = columns[1]

        return detected_group, detected_price, detected_name

    row2 = tk.Frame(top_frame, bg="#EEF2F7")
    row2.pack(fill="x", pady=4)

    # Supplier is required for every import — each Excel is assumed to be
    # one supplier's price list, tagging every row lets the stocklist be
    # browsed by supplier later. Preferably pick an ALREADY REGISTERED
    # supplier from the dropdown (avoids "Acme Corp" vs "acme corp." ending
    # up as two different suppliers by accident), but if theirs isn't
    # registered yet, typing a name in the field next to it registers them
    # automatically on import instead of forcing a trip to Suppliers first.
    tk.Label(row2, text="Supplier:", bg="#EEF2F7").pack(side="left")
    import_supplier_menu = tk.OptionMenu(row2, import_supplier_var, "")
    import_supplier_menu.config(width=18)
    import_supplier_menu.pack(side="left", padx=(4, 8))

    def refresh_import_supplier_menu():
        menu = import_supplier_menu["menu"]
        menu.delete(0, "end")
        import_supplier_map.clear()

        suppliers = get_all_suppliers()
        for s in suppliers:
            import_supplier_map[s["name"]] = s["id"]
            menu.add_command(label=s["name"], command=lambda n=s["name"]: import_supplier_var.set(n))

        if import_supplier_var.get() not in import_supplier_map:
            import_supplier_var.set(suppliers[0]["name"] if suppliers else "")

    tk.Label(row2, text="or new supplier:", bg="#EEF2F7").pack(side="left")
    new_supplier_entry = tk.Entry(row2, width=16)
    new_supplier_entry.pack(side="left", padx=(4, 12))

    tk.Label(row2, text="Sheet:", bg="#EEF2F7").pack(side="left")
    sheet_menu = tk.OptionMenu(row2, sheet_var, "")
    sheet_menu.config(width=18)
    sheet_menu.pack(side="left", padx=(4, 12))

    tk.Label(row2, text="Group column:", bg="#EEF2F7").pack(side="left")
    group_menu = tk.OptionMenu(row2, group_var, "")
    group_menu.config(width=22)
    group_menu.pack(side="left", padx=(4, 12))

    tk.Label(row2, text="Price column:", bg="#EEF2F7").pack(side="left")
    price_menu = tk.OptionMenu(row2, price_var, "")
    price_menu.config(width=22)
    price_menu.pack(side="left", padx=(4, 12))

    tk.Label(row2, text="Name column:", bg="#EEF2F7").pack(side="left")
    name_menu = tk.OptionMenu(row2, name_var, "")
    name_menu.config(width=22)
    name_menu.pack(side="left", padx=(4, 0))

    info_label = tk.Label(top_frame, text="Total imported rows: 0", bg="#EEF2F7", fg="#0F2A4A", font=FONTS["f10b"])
    info_label.pack(anchor="w", pady=(8, 2))

    preview_frame = tk.LabelFrame(win, text="Browse Stock", font=FONTS["f12b"], bg="#EEF2F7", padx=10, pady=10)
    preview_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # Three ways to look at the same stock data:
    #   "group"    — every group's Min/Max/Avg combined across ALL suppliers
    #                (the original, only view this section used to have).
    #   "supplier" — drill down: pick a supplier -> see their groups (with
    #                the same Min/Max/Avg, just scoped to that supplier) ->
    #                pick a group -> see the individual items in it.
    #   "item"     — a flat, searchable table of every single item row,
    #                with its group/price/supplier/source as columns.
    nav_state = {"mode": "group", "supplier_id": None, "supplier_name": None, "group": None}

    mode_row = tk.Frame(preview_frame, bg="#EEF2F7")
    mode_row.pack(fill="x", pady=(0, 8))

    content_frame = tk.Frame(preview_frame, bg="#ffffff")
    content_frame.pack(fill="both", expand=True)

    def set_mode(mode):
        nav_state["mode"] = mode
        nav_state["supplier_id"] = None
        nav_state["supplier_name"] = None
        nav_state["group"] = None
        render_stocklist_view()

    def render_mode_buttons():
        for w in mode_row.winfo_children():
            w.destroy()

        muted = shade_color("#EEF2F7", -0.05)
        for key, label in (("group", "By Group"), ("supplier", "By Supplier"), ("item", "By Item")):
            active = nav_state["mode"] == key
            make_button(
                mode_row, text=label,
                bg="#0F2A4A" if active else muted,
                fg="white" if active else "#5B7A99",
                command=lambda k=key: set_mode(k)
            ).pack(side="left", padx=(0, 6))

    def clear_content():
        for w in content_frame.winfo_children():
            w.destroy()

    def make_stats_tree(parent):
        tree_frame = tk.Frame(parent, bg="#ffffff")
        tree_frame.pack(fill="both", expand=True)

        columns = ("group", "count", "min", "max", "avg")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=14)

        tree.heading("group", text="Group")
        tree.heading("count", text="Count")
        tree.heading("min", text="Min Price")
        tree.heading("max", text="Max Price")
        tree.heading("avg", text="Avg Price")

        tree.column("group", width=320, anchor="w")
        tree.column("count", width=100, anchor="center")
        tree.column("min", width=150, anchor="e")
        tree.column("max", width=150, anchor="e")
        tree.column("avg", width=150, anchor="e")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        xscroll.pack(side="bottom", fill="x")
        return tree

    def make_scroll_list(parent):
        outer = tk.Frame(parent, bg="#ffffff")
        outer.pack(fill="both", expand=True)

        canvas = tk.Canvas(outer, bg="#ffffff", highlightthickness=0)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        container = tk.Frame(canvas, bg="#ffffff")

        container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        window_id = canvas.create_window((0, 0), window=container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        bind_canvas_stretch(canvas, window_id)
        bind_mousewheel_scroll(canvas)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return container

    def render_group_view():
        clear_content()
        tree = make_stats_tree(content_frame)

        stats = get_stock_group_stats()
        total_rows = get_stock_overall_count()

        for row in stats:
            tree.insert(
                "", "end",
                values=(
                    row["item_group"], row["item_count"],
                    f"{row['min_price']:.2f}", f"{row['max_price']:.2f}", f"{row['avg_price']:.2f}",
                )
            )

        summary_text = "No stock data loaded." if not stats else f"Groups: {len(stats)} | Total rows: {total_rows}"
        tk.Label(content_frame, text=summary_text, bg="#ffffff", fg="#5B7A99", font=FONTS["f10i"]).pack(anchor="w", pady=(8, 0))
        info_label.config(text=f"Total imported rows: {total_rows}")

    def render_supplier_row(container, text, command):
        row = tk.Frame(container, bg="#EAF0F7", relief="flat", borderwidth=0,
                        highlightthickness=1, highlightbackground="#C8D4E2",
                        padx=10, pady=8, cursor="hand2")
        row.pack(fill="x", padx=4, pady=3)
        lbl = tk.Label(row, text=text, font=FONTS["f10"], bg="#EAF0F7", anchor="w", cursor="hand2")
        lbl.pack(fill="x")
        row.bind("<Button-1>", lambda e: command())
        lbl.bind("<Button-1>", lambda e: command())

    def render_supplier_view():
        clear_content()

        if nav_state["supplier_id"] is None:
            suppliers = get_suppliers_with_stock_db()
            if not suppliers:
                tk.Label(content_frame, text="Δεν υπάρχουν suppliers με stock data ακόμα — κάνε import ένα Excel πρώτα.",
                          bg="#ffffff", fg="#5B7A99", font=FONTS["f10i"]).pack(anchor="w", padx=8, pady=8)
                info_label.config(text=f"Total imported rows: {get_stock_overall_count()}")
                return

            container = make_scroll_list(content_frame)
            for s in suppliers:
                def go(supplier_id=s["id"], supplier_name=s["name"]):
                    nav_state["supplier_id"] = supplier_id
                    nav_state["supplier_name"] = supplier_name
                    nav_state["group"] = None
                    render_stocklist_view()
                render_supplier_row(container, f"{s['name']}   —   {s['item_count']} items", go)

            info_label.config(text=f"Total imported rows: {get_stock_overall_count()}")

        elif nav_state["group"] is None:
            back_row = tk.Frame(content_frame, bg="#ffffff")
            back_row.pack(fill="x", pady=(0, 6))

            def back_to_suppliers():
                nav_state["supplier_id"] = None
                render_stocklist_view()

            make_button(back_row, text="← All Suppliers", bg="#5B7A99", fg="white", command=back_to_suppliers).pack(side="left")
            tk.Label(back_row, text=f"  {nav_state['supplier_name']}", font=FONTS["f12b"], bg="#ffffff", fg="#0F2A4A").pack(side="left", padx=(8, 0))

            stats = get_stock_group_stats(supplier_id=nav_state["supplier_id"])
            if not stats:
                tk.Label(content_frame, text="Δεν υπάρχουν groups για αυτόν τον supplier.", bg="#ffffff", fg="#5B7A99", font=FONTS["f10i"]).pack(anchor="w", padx=8, pady=8)
                return

            container = make_scroll_list(content_frame)
            for row in stats:
                text = (
                    f"{row['item_group']}   •   {row['item_count']} items   •   "
                    f"Min €{row['min_price']:.2f}   Max €{row['max_price']:.2f}   Avg €{row['avg_price']:.2f}"
                )

                def go(group_name=row["item_group"]):
                    nav_state["group"] = group_name
                    render_stocklist_view()

                render_supplier_row(container, text, go)

        else:
            back_row = tk.Frame(content_frame, bg="#ffffff")
            back_row.pack(fill="x", pady=(0, 6))

            def back_to_groups():
                nav_state["group"] = None
                render_stocklist_view()

            make_button(back_row, text="← Groups", bg="#5B7A99", fg="white", command=back_to_groups).pack(side="left")
            tk.Label(back_row, text=f"  {nav_state['supplier_name']} / {nav_state['group']}", font=FONTS["f12b"], bg="#ffffff", fg="#0F2A4A").pack(side="left", padx=(8, 0))

            group_stats = get_stock_group_stats(supplier_id=nav_state["supplier_id"])
            matching = next((g for g in group_stats if g["item_group"] == nav_state["group"]), None)
            if matching:
                summary = (
                    f"Min €{matching['min_price']:.2f}   Max €{matching['max_price']:.2f}   "
                    f"Average €{matching['avg_price']:.2f}   ({matching['item_count']} items) — stats already calculated above, per group"
                )
                tk.Label(content_frame, text=summary, font=FONTS["f11b"], bg="#ffffff", fg="#0F2A4A", wraplength=900, justify="left").pack(anchor="w", pady=(0, 6))

            items = get_stock_items_for_group_db(nav_state["group"], supplier_id=nav_state["supplier_id"])
            container = make_scroll_list(content_frame)
            for it in items:
                name_display = it["item_name"] or "(no name)"
                box = tk.Frame(container, bg="#EAF0F7", relief="flat", borderwidth=0,
                                highlightthickness=1, highlightbackground="#C8D4E2", padx=10, pady=6)
                box.pack(fill="x", padx=4, pady=2)
                tk.Label(box, text=f"{name_display}   —   €{it['price']:.2f}", font=FONTS["f10"], bg="#EAF0F7", anchor="w").pack(fill="x")

    def render_item_view():
        clear_content()

        search_row = tk.Frame(content_frame, bg="#ffffff")
        search_row.pack(fill="x", pady=(0, 6))
        tk.Label(search_row, text="Search item or group:", bg="#ffffff").pack(side="left")
        item_search_entry = tk.Entry(search_row, width=30)
        item_search_entry.pack(side="left", padx=(4, 0))

        tree_frame = tk.Frame(content_frame, bg="#ffffff")
        tree_frame.pack(fill="both", expand=True, pady=(6, 0))

        columns = ("item", "group", "price", "supplier", "source", "imported")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=14)
        tree.heading("item", text="Item")
        tree.heading("group", text="Group")
        tree.heading("price", text="Price")
        tree.heading("supplier", text="Supplier")
        tree.heading("source", text="Source File")
        tree.heading("imported", text="Imported At")

        tree.column("item", width=220, anchor="w")
        tree.column("group", width=160, anchor="w")
        tree.column("price", width=90, anchor="e")
        tree.column("supplier", width=140, anchor="w")
        tree.column("source", width=180, anchor="w")
        tree.column("imported", width=130, anchor="w")

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        xscroll.pack(side="bottom", fill="x")

        def refresh_items(*args):
            for i in tree.get_children():
                tree.delete(i)

            rows = get_all_stock_items_flat_db(search=item_search_entry.get().strip())
            for r in rows:
                tree.insert(
                    "", "end",
                    values=(
                        r["item_name"] or "-", r["item_group"], f"{r['price']:.2f}",
                        r["supplier_name"] or "-", r["source_file"] or "-", r["imported_at"]
                    )
                )
            info_label.config(text=f"Showing {len(rows)} items")

        item_search_entry.bind("<KeyRelease>", refresh_items)
        refresh_items()

    def render_stocklist_view():
        render_mode_buttons()
        if nav_state["mode"] == "group":
            render_group_view()
        elif nav_state["mode"] == "supplier":
            render_supplier_view()
        else:
            render_item_view()

    def on_sheet_change(*args):
        current_sheet = sheet_var.get().strip()
        if not current_sheet or current_sheet not in excel_data["sheets"]:
            return

        df = excel_data["sheets"][current_sheet]
        cols = [str(c).strip() for c in df.columns.tolist()]

        update_option_menu(group_menu, group_var, cols, default_empty=False)
        update_option_menu(price_menu, price_var, cols, default_empty=False)
        update_option_menu(name_menu, name_var, cols, default_empty=True)

        detected_group, detected_price, detected_name = detect_columns(cols)

        if detected_group in cols:
            group_var.set(detected_group)
        if detected_price in cols:
            price_var.set(detected_price)
        if detected_name in cols or detected_name == "":
            name_var.set(detected_name)

        info_label.config(text=f"Loaded sheet rows: {len(df)}")

    def choose_excel_file():
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            xls = pd.ExcelFile(path)
            cleaned = {}

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(path, sheet_name=sheet_name, header=0)
                df = df.copy()
                df.columns = [str(c).strip() for c in df.columns]
                cleaned[sheet_name] = df

            if not cleaned:
                messagebox.showwarning("Warning", "Το Excel δεν περιέχει sheets.")
                return

            excel_data["sheets"] = cleaned
            selected_file["path"] = path

            sheet_names = list(cleaned.keys())
            update_option_menu(sheet_menu, sheet_var, sheet_names, default_empty=False)
            file_label.config(text=path)

            on_sheet_change()

        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file:\n{e}")

    def import_selected_sheet():
        if not selected_file["path"]:
            messagebox.showwarning("Warning", "Διάλεξε πρώτα Excel αρχείο.")
            return

        # A typed name in "or new supplier" wins over the dropdown — lets
        # someone register a not-yet-known supplier right here instead of
        # having to go set them up in Suppliers first.
        typed_supplier_name = new_supplier_entry.get().strip()
        if typed_supplier_name:
            error = validate_supplier_fields(typed_supplier_name, "", "")
            if error:
                messagebox.showwarning("Προσοχή", error)
                return

            existing = next(
                (s for s in get_all_suppliers() if s["name"].strip().lower() == typed_supplier_name.lower()),
                None
            )
            if existing:
                supplier_id = existing["id"]
                supplier_name = existing["name"]
            else:
                supplier_id = add_supplier_db(typed_supplier_name, "", "")
                supplier_name = typed_supplier_name
                refresh_import_supplier_menu()
                import_supplier_var.set(supplier_name)
        else:
            supplier_name = import_supplier_var.get().strip()
            supplier_id = import_supplier_map.get(supplier_name)
            if not supplier_id:
                messagebox.showwarning(
                    "Warning",
                    "Επίλεξε Supplier από τη λίστα, ή γράψε το όνομα ενός νέου supplier."
                )
                return

        current_sheet = sheet_var.get().strip()
        group_col = group_var.get().strip()
        price_col = price_var.get().strip()
        name_col = name_var.get().strip()

        if not current_sheet:
            messagebox.showwarning("Warning", "Επίλεξε sheet.")
            return

        if current_sheet not in excel_data["sheets"]:
            messagebox.showwarning("Warning", "Το επιλεγμένο sheet δεν βρέθηκε.")
            return

        if not group_col or not price_col:
            messagebox.showwarning("Warning", "Επίλεξε Group column και Price column.")
            return

        df = excel_data["sheets"][current_sheet].copy()

        if group_col not in df.columns or price_col not in df.columns:
            messagebox.showwarning("Warning", "Οι επιλεγμένες στήλες δεν υπάρχουν στο sheet.")
            return

        df[group_col] = df[group_col].astype(str).str.strip()
        df[group_col] = df[group_col].replace({"nan": "", "None": "", "NaN": ""})

        df[price_col] = df[price_col].apply(clean_price_to_float)

        if name_col and name_col in df.columns:
            df[name_col] = (
                df[name_col]
                .astype(str)
                .str.strip()
                .replace({"nan": "", "None": "", "NaN": ""})
            )
        else:
            name_col = None

        df = df[df[group_col] != ""]
        df = df[df[price_col].notna()]
        df = df[df[price_col] > 0]

        if df.empty:
            messagebox.showwarning(
                "Warning",
                "Δεν βρέθηκαν έγκυρες γραμμές μετά το καθάρισμα των δεδομένων.\n"
                "Έλεγξε αν το Group column είναι το σωστό group field και το Price column η τιμή σε EUR."
            )
            return

        imported_at = datetime.now().strftime("%d/%m/%Y %H:%M")
        source_file = os.path.basename(selected_file["path"])

        rows = []
        for _, r in df.iterrows():
            item_group = str(r[group_col]).strip()
            item_name = str(r[name_col]).strip() if name_col else ""
            price = float(r[price_col])

            rows.append((
                source_file,
                current_sheet,
                item_group,
                item_name,
                price,
                imported_at,
                supplier_id
            ))

        if not messagebox.askyesno(
            "Confirm Import",
            f"Βρέθηκαν {len(rows)} έγκυρες γραμμές για τον supplier '{supplier_name}'.\n"
            f"Θα γίνει αντικατάσταση του υπάρχοντος stocklist ΜΟΝΟ για αυτόν τον supplier "
            "(τα δεδομένα άλλων suppliers δεν επηρεάζονται).\nΣυνέχεια;"
        ):
            return

        clear_stock_items_for_supplier_db(supplier_id)
        add_stock_rows_db(rows)
        render_stocklist_view()
        new_supplier_entry.delete(0, tk.END)

        success_msg = f"Έγινε import {len(rows)} γραμμών από το sheet '{current_sheet}' για τον supplier '{supplier_name}'."
        messagebox.showinfo("Success", success_msg)

    btn_row = tk.Frame(top_frame, bg="#EEF2F7")
    btn_row.pack(fill="x", pady=(8, 0))

    make_button(btn_row, text="Choose Excel", bg="#0F2A4A", fg="white", width=14, command=choose_excel_file).pack(side="left", padx=(0, 8))
    make_button(btn_row, text="Import to Stocklist", bg="#639922", fg="white", width=16, command=import_selected_sheet).pack(side="left", padx=(0, 8))
    make_button(btn_row, text="Refresh", bg="#5B7A99", fg="white", width=12, command=render_stocklist_view).pack(side="left")

    sheet_var.trace_add("write", on_sheet_change)

    refresh_import_supplier_menu()
    render_stocklist_view()


# ================= SUPPLIERS WINDOW =================
def open_suppliers():
    win = open_section_window("Suppliers", 900, 620)

    form_frame = tk.LabelFrame(win, text="Add Supplier", font=FONTS["f12b"], bg="#EEF2F7", padx=14, pady=12)
    form_frame.pack(fill="x", padx=12, pady=12)

    tk.Label(form_frame, text="Name:", bg="#EEF2F7").grid(row=0, column=0, sticky="w", padx=(0, 4), pady=6)
    name_entry = tk.Entry(form_frame, width=22)
    name_entry.grid(row=0, column=1, padx=(0, 16), pady=6)

    tk.Label(form_frame, text="Tel:", bg="#EEF2F7").grid(row=0, column=2, sticky="w", padx=(0, 4), pady=6)
    tel_entry = tk.Entry(form_frame, width=16)
    tel_entry.grid(row=0, column=3, padx=(0, 16), pady=6)

    tk.Label(form_frame, text="Contract:", bg="#EEF2F7").grid(row=0, column=4, sticky="w", padx=(0, 4), pady=6)
    contact_entry = tk.Entry(form_frame, width=18)
    contact_entry.grid(row=0, column=5, padx=0, pady=6)

    suppliers_header = tk.Frame(win, bg="#EEF2F7")
    suppliers_header.pack(fill="x", padx=14, pady=(4, 0))

    tk.Label(suppliers_header, text="Suppliers", font=FONTS["f13b"], bg="#EEF2F7").pack(side="left")

    tk.Label(suppliers_header, text="Sort by:", bg="#EEF2F7", font=FONTS["f10"]).pack(side="left", padx=(20, 4))
    supplier_sort_var = tk.StringVar(value="Date registered")
    supplier_sort_menu = tk.OptionMenu(
        suppliers_header, supplier_sort_var, "Date registered", "Name (A-Z)",
        command=lambda *_: refresh_all()
    )
    supplier_sort_menu.config(width=16)
    supplier_sort_menu.pack(side="left")

    def current_sort_key():
        return "name" if supplier_sort_var.get() == "Name (A-Z)" else "date"

    suppliers_outer = tk.Frame(win, bg="#EEF2F7")
    suppliers_outer.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    s_canvas = tk.Canvas(suppliers_outer, bg="#ffffff", highlightthickness=0)
    s_scroll = tk.Scrollbar(suppliers_outer, orient="vertical", command=s_canvas.yview)
    suppliers_container = tk.Frame(s_canvas, bg="#ffffff")

    suppliers_container.bind("<Configure>", lambda e: s_canvas.configure(scrollregion=s_canvas.bbox("all")))
    suppliers_window_id = s_canvas.create_window((0, 0), window=suppliers_container, anchor="nw")
    s_canvas.configure(yscrollcommand=s_scroll.set)
    bind_canvas_stretch(s_canvas, suppliers_window_id)
    bind_mousewheel_scroll(s_canvas)
    s_canvas.pack(side="left", fill="both", expand=True)
    s_scroll.pack(side="right", fill="y")

    def build_attachment_widget(parent, supplier, key, label):
        cell = tk.Frame(parent, bg="#ffffff")
        cell.pack(side="left", padx=6)

        path = supplier["attachments"].get(key)
        filename = os.path.basename(path) if path else "— none —"

        tk.Label(cell, text=f"{label}:", font=FONTS["f9"], bg="#ffffff").pack(side="left")
        tk.Label(cell, text=filename, font=FONTS["f9i"], bg="#ffffff", fg="#3D5A78" if path else "#7A93AC", width=14, anchor="w").pack(side="left", padx=(2, 4))

        def browse():
            chosen = filedialog.askopenfilename(title=f"Select {label} file", filetypes=ATTACHMENT_FILETYPES)
            if chosen:
                set_supplier_attachment_db(supplier["id"], key, chosen)
                refresh_all()

        def open_attachment():
            if path:
                open_file(path)
            else:
                messagebox.showinfo(label, "Δεν έχει επισυναφθεί αρχείο ακόμα.")

        make_button(cell, text="Browse", width=6, bg="#5B7A99", fg="white", command=browse).pack(side="left", padx=1)
        make_button(cell, text="Open", width=5, bg="#3D5A78", fg="white", command=open_attachment).pack(side="left", padx=1)

    def add_supplier_row(supplier):
        row_outer = tk.Frame(suppliers_container, bg="#ffffff", relief="flat", borderwidth=0, highlightthickness=1, highlightbackground="#C8D4E2", pady=6, padx=6)
        row_outer.pack(fill="x", pady=4, padx=2)

        info_row = tk.Frame(row_outer, bg="#ffffff")
        info_row.pack(fill="x")

        tk.Label(info_row, text=supplier["name"], font=FONTS["f11b"], bg="#ffffff", anchor="w", width=20).pack(side="left")
        tk.Label(info_row, text=f"Tel: {supplier['tel'] or '-'}", font=FONTS["f10"], bg="#ffffff", anchor="w", width=18).pack(side="left")
        tk.Label(info_row, text=f"Contract: {supplier['contact'] or '-'}", font=FONTS["f10"], bg="#ffffff", anchor="w", width=22).pack(side="left")

        def edit_supplier_action():
            new_name = simpledialog.askstring("Edit Supplier", "Name:", initialvalue=supplier["name"])
            if new_name is None:
                return
            new_name = new_name.strip()

            new_tel = simpledialog.askstring("Edit Supplier", "Tel:", initialvalue=supplier["tel"])
            new_tel = (new_tel or "").strip()

            new_contact = simpledialog.askstring("Edit Supplier", "Contract:", initialvalue=supplier["contact"])
            new_contact = (new_contact or "").strip()

            error = validate_supplier_fields(new_name, new_tel, new_contact)
            if error:
                messagebox.showwarning("Προσοχή", error)
                return

            update_supplier_db(supplier["id"], new_name, new_tel, new_contact)
            refresh_all()

        def delete_supplier_action():
            if not messagebox.askyesno("Delete", f"Διαγραφή προμηθευτή '{supplier['name']}';"):
                return
            delete_supplier_db(supplier["id"])
            refresh_all()

        make_button(info_row, text="Edit", width=5, bg="#0F2A4A", fg="white", command=edit_supplier_action).pack(side="left", padx=2)
        make_icon_button(info_row, text="🗑️", command=delete_supplier_action).pack(side="left", padx=2)

        attach_row = tk.Frame(row_outer, bg="#ffffff")
        attach_row.pack(fill="x", pady=(4, 0))
        tk.Label(attach_row, text="Attachments:", font=FONTS["f9i"], bg="#ffffff", fg="#5B7A99").pack(side="left", padx=(0, 6))

        for key, label in ATTACHMENT_TYPES:
            build_attachment_widget(attach_row, supplier, key, label)

    def refresh_all():
        for widget in suppliers_container.winfo_children():
            widget.destroy()

        for supplier in get_all_suppliers(order_by=current_sort_key()):
            add_supplier_row(supplier)

    def add_supplier_action():
        name = name_entry.get().strip()
        tel = tel_entry.get().strip()
        contact = contact_entry.get().strip()

        error = validate_supplier_fields(name, tel, contact)
        if error:
            messagebox.showwarning("Προσοχή", error)
            return

        add_supplier_db(name, tel, contact)
        name_entry.delete(0, tk.END)
        tel_entry.delete(0, tk.END)
        contact_entry.delete(0, tk.END)
        refresh_all()

    make_button(form_frame, text="Add Supplier", bg="#639922", fg="white", command=add_supplier_action).grid(row=0, column=6, padx=10)

    refresh_all()


init_db()

root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
root.title("Dashboard MSC")
center_window(root, 620, 480)
root.configure(bg="#EEF2F7")

# Fonts are shared, scalable Font objects rather than plain ("Arial", N)
# tuples — this must run once an actual Tk root exists, and before any
# window creates widgets that reference FONTS[...].
setup_fonts()
bind_font_scaling(root, 620, 480)

header = tk.Label(root, text="Dashboard MSC", font=FONTS["f22b"], bg="#EEF2F7", fg="#0F2A4A")
header.pack(pady=(30, 4))

subtitle = tk.Label(root, text="Tasks · Suppliers · Members · Meetings · Stocklist",
                     font=FONTS["f10i"], bg="#EEF2F7", fg="#7A93AC")
subtitle.pack(pady=(0, 4))

divider = tk.Frame(root, bg="#C8D4E2", height=1)
divider.pack(fill="x", padx=60, pady=(10, 25))

btn_frame = tk.Frame(root, bg="#EEF2F7")
btn_frame.pack(pady=5)

make_button(btn_frame, text="Tasks", width=18, height=2, bg="#0F2A4A", fg="white", command=open_tasks).grid(row=0, column=0, padx=14, pady=10)
make_button(btn_frame, text="Suppliers", width=18, height=2, bg="#0F2A4A", fg="white", command=open_suppliers).grid(row=0, column=1, padx=14, pady=10)
make_button(btn_frame, text="Members", width=18, height=2, bg="#639922", fg="white", command=open_members).grid(row=1, column=0, padx=14, pady=10)
make_button(btn_frame, text="Meetings", width=18, height=2, bg="#0F2A4A", fg="white", command=open_meetings).grid(row=1, column=1, padx=14, pady=10)
make_button(btn_frame, text="Stocklist", width=18, height=2, bg="#3D5A78", fg="white", command=open_stocklist).grid(row=2, column=0, columnspan=2, padx=14, pady=10)

footer = tk.Label(root, text="Local SQLite storage enabled", font=FONTS["f10i"], bg="#EEF2F7", fg="#5B7A99")
footer.pack(pady=(20, 0))

root.mainloop()