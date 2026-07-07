import os
import sys
import sqlite3
import subprocess
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog, ttk
from datetime import datetime
import pandas as pd

CURRENCY_SYMBOLS = {"EUR": "€", "USD": "$", "JPY": "¥"}
ATTACHMENT_TYPES = [
    ("contract", "Contract"),
    ("email", "Email"),
    ("quote", "Price Quote"),
]
ATTACHMENT_FILETYPES = [
    ("PDF files", "*.pdf"),
    ("Word files", "*.doc *.docx"),
    ("Excel files", "*.xls *.xlsx"),
    ("Images", "*.jpg *.jpeg *.png"),
    ("All files", "*.*"),
]

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "dashboard.db")


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

    conn.commit()
    conn.close()


# ================= EMAILS =================
def get_all_emails():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM emails ORDER BY id DESC")
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def add_email_db(sender, subject, body, expanded=0):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO emails (sender, subject, body, expanded) VALUES (?, ?, ?, ?)",
        (sender, subject, body, expanded)
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


# ================= SUPPLIERS =================
def get_all_suppliers():
    conn = get_conn()
    cur = conn.cursor()
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
def get_all_members():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT * FROM members ORDER BY id DESC")
    rows = [dict(row) for row in cur.fetchall()]
    conn.close()
    return rows


def add_member_db(first_name, last_name, job_title, note, created_date):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO members (first_name, last_name, job_title, note, created_date) VALUES (?, ?, ?, ?, ?)",
        (first_name, last_name, job_title, note, created_date)
    )
    conn.commit()
    conn.close()


def delete_member_db(member_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM members WHERE id=?", (member_id,))
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


def add_stock_rows_db(rows):
    conn = get_conn()
    cur = conn.cursor()
    cur.executemany(
        """
        INSERT INTO stock_items (source_file, sheet_name, item_group, item_name, price, imported_at)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        rows
    )
    conn.commit()
    conn.close()


def get_stock_group_stats():
    conn = get_conn()
    cur = conn.cursor()
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


# ================= HELPERS =================
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


# ================= TASKS WINDOW =================
def open_tasks():
    win = tk.Toplevel(root)
    win.title("Tasks")
    win.geometry("850x550")
    win.configure(bg="#f5f5f5")

    main_frame = tk.Frame(win, bg="#f5f5f5")
    main_frame.pack(fill="both", expand=True)

    left_frame = tk.Frame(main_frame, bg="#e0e0e0", width=400)
    left_frame.pack(side="left", fill="both", expand=True)

    right_frame = tk.Frame(main_frame, bg="#ffffff", width=400)
    right_frame.pack(side="right", fill="both", expand=True)

    tk.Label(left_frame, text="Emails", font=("Arial", 14, "bold"), bg="#e0e0e0").pack(pady=(10, 5))

    new_email_btn_frame = tk.Frame(left_frame, bg="#e0e0e0")
    new_email_btn_frame.pack(pady=(0, 5))

    email_canvas = tk.Canvas(left_frame, bg="#e0e0e0", highlightthickness=0)
    email_scrollbar = tk.Scrollbar(left_frame, orient="vertical", command=email_canvas.yview)
    emails_container = tk.Frame(email_canvas, bg="#e0e0e0")

    emails_container.bind("<Configure>", lambda e: email_canvas.configure(scrollregion=email_canvas.bbox("all")))
    email_canvas.create_window((0, 0), window=emails_container, anchor="nw")
    email_canvas.configure(yscrollcommand=email_scrollbar.set)

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
        card = tk.Frame(emails_container, bg="#cfe8ff", relief="raised", borderwidth=1, pady=4, padx=6)
        card.pack(fill="x", pady=4, padx=2)

        def refresh_card():
            for widget in card.winfo_children():
                widget.destroy()

            header_row = tk.Frame(card, bg="#cfe8ff")
            header_row.pack(fill="x")

            arrow = "▼" if email_dict["expanded"] else "▶"
            subject_lbl = tk.Label(
                header_row,
                text=f"{arrow} {email_dict['subject']}",
                font=("Arial", 11, "bold"),
                bg="#cfe8ff",
                anchor="w",
                wraplength=280,
                cursor="hand2",
                justify="left"
            )
            subject_lbl.pack(side="left", fill="x", expand=True)

            def toggle_expand(event=None):
                email_dict["expanded"] = 0 if email_dict["expanded"] else 1
                set_email_expanded_db(email_dict["id"], email_dict["expanded"])
                refresh_card()

            subject_lbl.bind("<Button-1>", toggle_expand)
            subject_lbl.bind("<ButtonPress-1>", on_drag_start(lambda: email_dict["subject"]), add="+")
            subject_lbl.bind("<ButtonRelease-1>", on_drag_release, add="+")

            btns_row = tk.Frame(header_row, bg="#cfe8ff")
            btns_row.pack(side="right")

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

            tk.Button(btns_row, text="Edit", width=5, bg="#2196F3", fg="white", command=edit_email_action).pack(side="left", padx=2)
            tk.Button(btns_row, text="Delete", width=6, bg="#f44336", fg="white", command=delete_email_action).pack(side="left", padx=2)

            if email_dict["expanded"]:
                tk.Label(card, text=f"Από: {email_dict['sender']}", font=("Arial", 9), bg="#cfe8ff", anchor="w").pack(fill="x", pady=(4, 2))

                body_lbl = tk.Label(
                    card,
                    text=email_dict["body"],
                    font=("Arial", 10),
                    bg="#d9f2d9",
                    anchor="w",
                    justify="left",
                    wraplength=320,
                    cursor="hand2",
                    relief="raised",
                    padx=8,
                    pady=6
                )
                body_lbl.pack(fill="x", pady=(0, 4))

                body_lbl.bind("<ButtonPress-1>", on_drag_start(lambda: email_dict["body"]))
                body_lbl.bind("<ButtonRelease-1>", on_drag_release)

        refresh_card()

    def open_new_email_form():
        form = tk.Toplevel(win)
        form.title("Νέο Email")
        form.geometry("380x320")
        form.configure(bg="#f5f5f5")
        form.grab_set()

        tk.Label(form, text="Από (email):", font=("Arial", 10), bg="#f5f5f5", anchor="w").pack(fill="x", padx=15, pady=(15, 2))
        from_entry = tk.Entry(form, font=("Arial", 11))
        from_entry.insert(0, "user@example.com")
        from_entry.pack(fill="x", padx=15)

        tk.Label(form, text="Subject:", font=("Arial", 10), bg="#f5f5f5", anchor="w").pack(fill="x", padx=15, pady=(15, 2))
        subject_entry = tk.Entry(form, font=("Arial", 11))
        subject_entry.pack(fill="x", padx=15)

        tk.Label(form, text="Θέμα / Περιεχόμενο:", font=("Arial", 10), bg="#f5f5f5", anchor="w").pack(fill="x", padx=15, pady=(15, 2))
        body_text = tk.Text(form, font=("Arial", 11), height=6, wrap="word")
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

        btn_frame = tk.Frame(form, bg="#f5f5f5")
        btn_frame.pack(pady=15)

        tk.Button(btn_frame, text="Add", bg="#4CAF50", fg="white", width=10, command=submit_email).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", bg="#f44336", fg="white", width=10, command=form.destroy).pack(side="left", padx=5)

    tk.Button(new_email_btn_frame, text="+ Νέο Email", bg="#4CAF50", fg="white", command=open_new_email_form).pack()

    for existing_email in get_all_emails():
        render_email_card(existing_email)

    tk.Label(right_frame, text="To-Do List", font=("Arial", 14, "bold"), bg="#ffffff").pack(pady=(10, 5))

    entry_frame = tk.Frame(right_frame, bg="#ffffff")
    entry_frame.pack(pady=5, fill="x", padx=10)

    task_entry = tk.Entry(entry_frame, font=("Arial", 11))
    task_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

    task_canvas = tk.Canvas(right_frame, bg="#ffffff", highlightthickness=0)
    task_scrollbar = tk.Scrollbar(right_frame, orient="vertical", command=task_canvas.yview)
    tasks_container = tk.Frame(task_canvas, bg="#ffffff")

    tasks_container.bind("<Configure>", lambda e: task_canvas.configure(scrollregion=task_canvas.bbox("all")))
    task_canvas.create_window((0, 0), window=tasks_container, anchor="nw")
    task_canvas.configure(yscrollcommand=task_scrollbar.set)

    task_canvas.pack(side="left", fill="both", expand=True, padx=10)
    task_scrollbar.pack(side="right", fill="y")

    def render_task_row(task_dict):
        row_outer = tk.Frame(tasks_container, bg="#ffffff", relief="groove", borderwidth=1, pady=4, padx=4)
        row_outer.pack(fill="x", pady=4, padx=2)

        def refresh_row():
            for widget in row_outer.winfo_children():
                widget.destroy()

            top_row = tk.Frame(row_outer, bg="#ffffff")
            top_row.pack(fill="x")

            done_var = tk.BooleanVar(value=bool(task_dict["done"]))

            def toggle_done():
                task_dict["done"] = 1 if done_var.get() else 0
                set_task_done_db(task_dict["id"], task_dict["done"])
                refresh_row()

            tk.Checkbutton(top_row, variable=done_var, bg="#ffffff", command=toggle_done).pack(side="left")

            task_font = ("Arial", 11, "overstrike") if task_dict["done"] else ("Arial", 11)
            task_color = "#999999" if task_dict["done"] else "#000000"

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

            tk.Button(top_row, text="Edit", width=5, bg="#2196F3", fg="white", command=edit_task_action).pack(side="left", padx=2)
            tk.Button(top_row, text="Delete", width=6, bg="#f44336", fg="white", command=delete_task_action).pack(side="left", padx=2)

            bottom_row = tk.Frame(row_outer, bg="#ffffff")
            bottom_row.pack(fill="x", pady=(2, 0))
            tk.Label(bottom_row, text=f"📅 Καταχωρήθηκε: {task_dict['date']}", font=("Arial", 8, "italic"), fg="#777777", bg="#ffffff").pack(side="left")

        refresh_row()

    def add_task_row(text):
        date_str = datetime.now().strftime("%d/%m/%Y %H:%M")
        task_id = add_task_db(text, date_str, 0)
        render_task_row({
            "id": task_id,
            "text": text,
            "date": date_str,
            "done": 0
        })

    def add_task_action(event=None):
        text = task_entry.get().strip()
        if not text:
            messagebox.showwarning("Προσοχή", "Γράψε ένα task πρώτα.")
            return
        add_task_row(text)
        task_entry.delete(0, tk.END)

    tk.Button(entry_frame, text="Add Task", bg="#4CAF50", fg="white", command=add_task_action).pack(side="left")
    task_entry.bind("<Return>", add_task_action)

    for existing_task in get_all_tasks():
        render_task_row(existing_task)


# ================= MEMBERS WINDOW =================
def open_members():
    win = tk.Toplevel(root)
    win.title("Members")
    win.geometry("850x600")
    win.configure(bg="#f5f5f5")

    form_frame = tk.LabelFrame(win, text="New Employee", font=("Arial", 12, "bold"), bg="#f5f5f5", padx=10, pady=10)
    form_frame.pack(fill="x", padx=10, pady=10)

    tk.Label(form_frame, text="Name:", bg="#f5f5f5").grid(row=0, column=0, sticky="w")
    first_name_entry = tk.Entry(form_frame, width=20)
    first_name_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(form_frame, text="Surname:", bg="#f5f5f5").grid(row=0, column=2, sticky="w")
    last_name_entry = tk.Entry(form_frame, width=20)
    last_name_entry.grid(row=0, column=3, padx=5, pady=5)

    tk.Label(form_frame, text="Job Title:", bg="#f5f5f5").grid(row=1, column=0, sticky="w")
    job_title_entry = tk.Entry(form_frame, width=20)
    job_title_entry.grid(row=1, column=1, padx=5, pady=5)

    current_date = datetime.now().strftime("%d/%m/%Y")
    tk.Label(form_frame, text="Date:", bg="#f5f5f5").grid(row=1, column=2, sticky="w")
    tk.Label(form_frame, text=current_date, bg="#ffffff", relief="sunken", width=18, anchor="w").grid(row=1, column=3, padx=5, pady=5)

    tk.Label(form_frame, text="Note:", bg="#f5f5f5").grid(row=2, column=0, sticky="nw")
    note_text = tk.Text(form_frame, width=50, height=5)
    note_text.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="w")

    list_outer = tk.Frame(win, bg="#f5f5f5")
    list_outer.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    canvas = tk.Canvas(list_outer, bg="#ffffff", highlightthickness=0)
    scrollbar = tk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
    members_container = tk.Frame(canvas, bg="#ffffff")

    members_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=members_container, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def refresh_members():
        for widget in members_container.winfo_children():
            widget.destroy()

        members = get_all_members()

        if not members:
            tk.Label(members_container, text="Δεν υπάρχουν καταχωρημένοι εργαζόμενοι.", bg="#ffffff", fg="#777777", font=("Arial", 10)).pack(anchor="w", padx=8, pady=8)
            return

        for member in members:
            box = tk.Frame(members_container, bg="#eefaf0", relief="groove", borderwidth=1, padx=8, pady=6)
            box.pack(fill="x", pady=4, padx=4)

            tk.Label(box, text=f"{member['first_name']} {member['last_name']}", font=("Arial", 11, "bold"), bg="#eefaf0", anchor="w").pack(fill="x")
            tk.Label(box, text=f"Θέση εργασίας: {member['job_title']}", font=("Arial", 10), bg="#eefaf0", anchor="w").pack(fill="x", pady=(3, 0))
            tk.Label(box, text=f"Ημερομηνία: {member['created_date']}", font=("Arial", 9, "italic"), bg="#eefaf0", fg="#666666", anchor="w").pack(fill="x", pady=(3, 0))
            tk.Label(box, text=f"Note: {member['note'] or '-'}", font=("Arial", 10), bg="#eefaf0", justify="left", anchor="w", wraplength=650).pack(fill="x", pady=(4, 0))

            def delete_action(member_id=member["id"]):
                delete_member_db(member_id)
                refresh_members()

            tk.Button(box, text="Delete", bg="#f44336", fg="white", width=8, command=delete_action).pack(anchor="e", pady=(5, 0))

    def add_member_action():
        first_name = first_name_entry.get().strip()
        last_name = last_name_entry.get().strip()
        job_title = job_title_entry.get().strip()
        note = note_text.get("1.0", "end").strip()
        created_date = datetime.now().strftime("%d/%m/%Y")

        if not first_name or not last_name or not job_title:
            messagebox.showwarning("Προσοχή", "Συμπλήρωσε Name, Surname και Job Title.")
            return

        add_member_db(first_name, last_name, job_title, note, created_date)

        first_name_entry.delete(0, tk.END)
        last_name_entry.delete(0, tk.END)
        job_title_entry.delete(0, tk.END)
        note_text.delete("1.0", tk.END)

        refresh_members()

    tk.Button(form_frame, text="Add Member", bg="#4CAF50", fg="white", command=add_member_action).grid(row=3, column=3, pady=8, sticky="e")

    refresh_members()


# ================= MEETINGS WINDOW =================
def open_meetings():
    win = tk.Toplevel(root)
    win.title("Meetings")
    win.geometry("900x600")
    win.configure(bg="#f5f5f5")

    left = tk.Frame(win, bg="#f5f5f5", width=260)
    left.pack(side="left", fill="y", padx=10, pady=10)

    right = tk.Frame(win, bg="#ffffff")
    right.pack(side="right", fill="both", expand=True, padx=10, pady=10)

    tk.Label(left, text="Add Meeting", font=("Arial", 14, "bold"), bg="#f5f5f5").pack(anchor="w", pady=(0, 10))

    tk.Label(left, text="Date (YYYY-MM-DD):", bg="#f5f5f5", font=("Arial", 10)).pack(anchor="w")
    date_entry = tk.Entry(left, width=24, font=("Arial", 11))
    date_entry.pack(anchor="w", pady=(2, 10))
    date_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

    tk.Label(left, text="Title:", bg="#f5f5f5", font=("Arial", 10)).pack(anchor="w")
    title_entry = tk.Entry(left, width=24, font=("Arial", 11))
    title_entry.pack(anchor="w", pady=(2, 10))

    tk.Label(left, text="Note:", bg="#f5f5f5", font=("Arial", 10)).pack(anchor="w")
    note_text = tk.Text(left, width=28, height=8, font=("Arial", 10))
    note_text.pack(anchor="w", pady=(2, 10))

    tk.Label(right, text="Meetings", font=("Arial", 14, "bold"), bg="#ffffff").pack(anchor="w", padx=10, pady=(10, 8))

    top_filter = tk.Frame(right, bg="#ffffff")
    top_filter.pack(fill="x", padx=10, pady=(0, 8))

    tk.Label(top_filter, text="View date:", bg="#ffffff", font=("Arial", 10)).pack(side="left")
    filter_entry = tk.Entry(top_filter, width=18, font=("Arial", 11))
    filter_entry.pack(side="left", padx=(6, 8))
    filter_entry.insert(0, datetime.now().strftime("%Y-%m-%d"))

    selected_date_label = tk.Label(right, text="", font=("Arial", 11, "bold"), bg="#ffffff", fg="#1565c0")
    selected_date_label.pack(anchor="w", padx=10, pady=(0, 5))

    list_frame = tk.Frame(right, bg="#ffffff")
    list_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    canvas = tk.Canvas(list_frame, bg="#ffffff", highlightthickness=0)
    scrollbar = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    meetings_container = tk.Frame(canvas, bg="#ffffff")

    meetings_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=meetings_container, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def refresh_meetings_for_date(selected_date):
        for widget in meetings_container.winfo_children():
            widget.destroy()

        try:
            datetime.strptime(selected_date, "%Y-%m-%d")
        except ValueError:
            selected_date_label.config(text="Λάθος μορφή ημερομηνίας. Βάλε YYYY-MM-DD")
            return

        selected_date_label.config(text=f"Selected date: {selected_date}")

        meetings = get_meetings_by_date(selected_date)

        if not meetings:
            tk.Label(meetings_container, text="Δεν υπάρχουν meetings για αυτή την ημερομηνία.", bg="#ffffff", fg="#777777", font=("Arial", 10)).pack(anchor="w", pady=6)
            return

        for meeting in meetings:
            box = tk.Frame(meetings_container, bg="#eef6ff", relief="groove", borderwidth=1, padx=8, pady=6)
            box.pack(fill="x", pady=4)

            tk.Label(box, text=meeting["title"], font=("Arial", 11, "bold"), bg="#eef6ff", anchor="w").pack(fill="x")
            tk.Label(box, text=meeting["note"] or "-", font=("Arial", 10), bg="#eef6ff", justify="left", anchor="w", wraplength=420).pack(fill="x", pady=(4, 2))
            tk.Label(box, text=f"Καταχώριση: {meeting['created_at']}", font=("Arial", 8, "italic"), bg="#eef6ff", fg="#666666").pack(anchor="w")

            def delete_action(meeting_id=meeting["id"], current_date=selected_date):
                delete_meeting_db(meeting_id)
                refresh_meetings_for_date(current_date)

            tk.Button(box, text="Delete", bg="#f44336", fg="white", width=8, command=delete_action).pack(anchor="e", pady=(5, 0))

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

        refresh_meetings_for_date(meeting_date)

    def view_selected_date():
        selected_date = filter_entry.get().strip()
        refresh_meetings_for_date(selected_date)

    tk.Button(left, text="Add Meeting", bg="#4CAF50", fg="white", width=16, command=add_meeting_action).pack(anchor="w", pady=8)
    tk.Button(top_filter, text="View", bg="#2196F3", fg="white", width=10, command=view_selected_date).pack(side="left")

    filter_entry.bind("<Return>", lambda e: view_selected_date())

    refresh_meetings_for_date(filter_entry.get().strip())


# ================= STOCKLIST WINDOW =================
def open_stocklist():
    win = tk.Toplevel(root)
    win.title("Stocklist")
    win.geometry("1100x700")
    win.configure(bg="#f5f5f5")

    selected_file = {"path": None}
    excel_data = {"sheets": {}, "current_sheet": None}

    top_frame = tk.LabelFrame(win, text="Import Excel Stocklist", font=("Arial", 12, "bold"), bg="#f5f5f5", padx=10, pady=10)
    top_frame.pack(fill="x", padx=10, pady=10)

    row1 = tk.Frame(top_frame, bg="#f5f5f5")
    row1.pack(fill="x", pady=4)

    file_label = tk.Label(row1, text="No file selected", bg="#ffffff", relief="sunken", anchor="w", width=70)
    file_label.pack(side="left", padx=(0, 8), fill="x", expand=True)

    sheet_var = tk.StringVar(value="")
    group_var = tk.StringVar(value="")
    price_var = tk.StringVar(value="")
    name_var = tk.StringVar(value="")

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
            "group", "groups", "category", "product group", "item group",
            "family", "type", "brand", "department", "classification"
        ]
        price_candidates = [
            "price", "unit price", "cost", "sale price", "stock price",
            "amount", "value", "net price"
        ]
        name_candidates = [
            "item", "item name", "name", "description", "product", "product name"
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

    row2 = tk.Frame(top_frame, bg="#f5f5f5")
    row2.pack(fill="x", pady=4)

    tk.Label(row2, text="Sheet:", bg="#f5f5f5").pack(side="left")
    sheet_menu = tk.OptionMenu(row2, sheet_var, "")
    sheet_menu.config(width=18)
    sheet_menu.pack(side="left", padx=(4, 12))

    tk.Label(row2, text="Group column:", bg="#f5f5f5").pack(side="left")
    group_menu = tk.OptionMenu(row2, group_var, "")
    group_menu.config(width=18)
    group_menu.pack(side="left", padx=(4, 12))

    tk.Label(row2, text="Price column:", bg="#f5f5f5").pack(side="left")
    price_menu = tk.OptionMenu(row2, price_var, "")
    price_menu.config(width=18)
    price_menu.pack(side="left", padx=(4, 12))

    tk.Label(row2, text="Name column:", bg="#f5f5f5").pack(side="left")
    name_menu = tk.OptionMenu(row2, name_var, "")
    name_menu.config(width=18)
    name_menu.pack(side="left", padx=(4, 0))

    info_label = tk.Label(top_frame, text="Total imported rows: 0", bg="#f5f5f5", fg="#333333", font=("Arial", 10, "bold"))
    info_label.pack(anchor="w", pady=(8, 2))

    preview_frame = tk.LabelFrame(win, text="Group Statistics", font=("Arial", 12, "bold"), bg="#f5f5f5", padx=10, pady=10)
    preview_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    tree_frame = tk.Frame(preview_frame, bg="#ffffff")
    tree_frame.pack(fill="both", expand=True)

    columns = ("group", "count", "min", "max", "avg")
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=18)

    tree.heading("group", text="Group")
    tree.heading("count", text="Count")
    tree.heading("min", text="Min Price")
    tree.heading("max", text="Max Price")
    tree.heading("avg", text="Avg Price")

    tree.column("group", width=300, anchor="w")
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

    summary_label = tk.Label(preview_frame, text="No stock data loaded.", bg="#f5f5f5", fg="#555555", font=("Arial", 10, "italic"))
    summary_label.pack(anchor="w", pady=(8, 0))

    def refresh_stats_view():
        for item in tree.get_children():
            tree.delete(item)

        stats = get_stock_group_stats()
        total_rows = get_stock_overall_count()

        if not stats:
            summary_label.config(text="No stock data loaded.")
            info_label.config(text=f"Total imported rows: {total_rows}")
            return

        for row in stats:
            tree.insert(
                "",
                "end",
                values=(
                    row["item_group"],
                    row["item_count"],
                    f"{row['min_price']:.2f}",
                    f"{row['max_price']:.2f}",
                    f"{row['avg_price']:.2f}",
                )
            )

        summary_label.config(text=f"Groups: {len(stats)} | Total rows: {total_rows}")
        info_label.config(text=f"Total imported rows: {total_rows}")

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

    def choose_excel_file():
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return

        try:
            sheets = pd.read_excel(path, sheet_name=None)
            if not sheets:
                messagebox.showwarning("Warning", "Το Excel δεν περιέχει sheets.")
                return

            cleaned = {}
            for sheet_name, df in sheets.items():
                df = df.copy()
                df.columns = [str(c).strip() for c in df.columns]
                cleaned[sheet_name] = df

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
        df[price_col] = pd.to_numeric(df[price_col], errors="coerce")

        if name_col and name_col in df.columns:
            df[name_col] = df[name_col].astype(str).str.strip()
        else:
            name_col = None

        df = df[df[group_col].notna()]
        df = df[df[group_col] != ""]
        df = df[df[price_col].notna()]

        if df.empty:
            messagebox.showwarning("Warning", "Δεν βρέθηκαν έγκυρες γραμμές μετά το καθάρισμα των δεδομένων.")
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
                imported_at
            ))

        if not messagebox.askyesno("Confirm Import", f"Θα γίνει αντικατάσταση του υπάρχοντος stocklist με {len(rows)} γραμμές. Συνέχεια;"):
            return

        clear_stock_items_db()
        add_stock_rows_db(rows)
        refresh_stats_view()

        messagebox.showinfo("Success", f"Έγινε import {len(rows)} γραμμών από το sheet '{current_sheet}'.")

    btn_row = tk.Frame(top_frame, bg="#f5f5f5")
    btn_row.pack(fill="x", pady=(8, 0))

    tk.Button(btn_row, text="Choose Excel", bg="#2196F3", fg="white", width=14, command=choose_excel_file).pack(side="left", padx=(0, 8))
    tk.Button(btn_row, text="Import to Stocklist", bg="#4CAF50", fg="white", width=16, command=import_selected_sheet).pack(side="left", padx=(0, 8))
    tk.Button(btn_row, text="Refresh Stats", bg="#FF9800", fg="white", width=12, command=refresh_stats_view).pack(side="left")

    sheet_var.trace_add("write", on_sheet_change)

    refresh_stats_view()


# ================= SUPPLIERS WINDOW =================
def open_suppliers():
    win = tk.Toplevel(root)
    win.title("Suppliers")
    win.geometry("980x700")
    win.configure(bg="#f5f5f5")

    supplier_map = {}

    form_frame = tk.LabelFrame(win, text="Add Supplier", font=("Arial", 12, "bold"), bg="#f5f5f5", padx=10, pady=10)
    form_frame.pack(fill="x", padx=10, pady=10)

    tk.Label(form_frame, text="Name:", bg="#f5f5f5").grid(row=0, column=0, sticky="w")
    name_entry = tk.Entry(form_frame, width=22)
    name_entry.grid(row=0, column=1, padx=5, pady=3)

    tk.Label(form_frame, text="Tel:", bg="#f5f5f5").grid(row=0, column=2, sticky="w")
    tel_entry = tk.Entry(form_frame, width=16)
    tel_entry.grid(row=0, column=3, padx=5, pady=3)

    tk.Label(form_frame, text="Contact:", bg="#f5f5f5").grid(row=0, column=4, sticky="w")
    contact_entry = tk.Entry(form_frame, width=18)
    contact_entry.grid(row=0, column=5, padx=5, pady=3)

    tk.Label(win, text="Suppliers", font=("Arial", 13, "bold"), bg="#f5f5f5").pack(anchor="w", padx=12)

    suppliers_outer = tk.Frame(win, bg="#f5f5f5", height=230)
    suppliers_outer.pack(fill="x", padx=10, pady=(0, 10))
    suppliers_outer.pack_propagate(False)

    s_canvas = tk.Canvas(suppliers_outer, bg="#ffffff", highlightthickness=0)
    s_scroll = tk.Scrollbar(suppliers_outer, orient="vertical", command=s_canvas.yview)
    suppliers_container = tk.Frame(s_canvas, bg="#ffffff")

    suppliers_container.bind("<Configure>", lambda e: s_canvas.configure(scrollregion=s_canvas.bbox("all")))
    s_canvas.create_window((0, 0), window=suppliers_container, anchor="nw")
    s_canvas.configure(yscrollcommand=s_scroll.set)
    s_canvas.pack(side="left", fill="both", expand=True)
    s_scroll.pack(side="right", fill="y")

    quote_frame = tk.LabelFrame(win, text="Price Quotes", font=("Arial", 12, "bold"), bg="#f5f5f5", padx=10, pady=10)
    quote_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    control_row = tk.Frame(quote_frame, bg="#f5f5f5")
    control_row.pack(fill="x")

    tk.Label(control_row, text="Supplier:", bg="#f5f5f5").pack(side="left")
    selected_supplier = tk.StringVar(value="")
    supplier_menu = tk.OptionMenu(control_row, selected_supplier, "")
    supplier_menu.config(width=20)
    supplier_menu.pack(side="left", padx=(4, 12))

    tk.Label(control_row, text="Description:", bg="#f5f5f5").pack(side="left")
    desc_entry = tk.Entry(control_row, width=18)
    desc_entry.pack(side="left", padx=(4, 12))

    tk.Label(control_row, text="Qty:", bg="#f5f5f5").pack(side="left")
    qty_entry = tk.Entry(control_row, width=6)
    qty_entry.pack(side="left", padx=(4, 12))

    tk.Label(control_row, text="Price:", bg="#f5f5f5").pack(side="left")
    price_entry = tk.Entry(control_row, width=10)
    price_entry.pack(side="left", padx=(4, 4))

    currency_var = tk.StringVar(value="EUR")
    tk.OptionMenu(control_row, currency_var, "EUR", "USD", "JPY").pack(side="left", padx=(0, 12))

    items_outer = tk.Frame(quote_frame, bg="#f5f5f5")
    items_outer.pack(fill="both", expand=True, pady=(10, 0))

    i_canvas = tk.Canvas(items_outer, bg="#ffffff", highlightthickness=0)
    i_scroll = tk.Scrollbar(items_outer, orient="vertical", command=i_canvas.yview)
    items_container = tk.Frame(i_canvas, bg="#ffffff")

    items_container.bind("<Configure>", lambda e: i_canvas.configure(scrollregion=i_canvas.bbox("all")))
    i_canvas.create_window((0, 0), window=items_container, anchor="nw")
    i_canvas.configure(yscrollcommand=i_scroll.set)
    i_canvas.pack(side="left", fill="both", expand=True)
    i_scroll.pack(side="right", fill="y")

    summary_label = tk.Label(quote_frame, text="Min: -    Max: -    Average: -", font=("Arial", 11, "bold"), bg="#f5f5f5", anchor="w")
    summary_label.pack(fill="x", pady=(8, 0))

    def get_supplier_by_selected():
        supplier_id = supplier_map.get(selected_supplier.get())
        if not supplier_id:
            return None
        for s in get_all_suppliers():
            if s["id"] == supplier_id:
                return s
        return None

    def build_attachment_widget(parent, supplier, key, label):
        cell = tk.Frame(parent, bg="#ffffff")
        cell.pack(side="left", padx=6)

        path = supplier["attachments"].get(key)
        filename = os.path.basename(path) if path else "— none —"

        tk.Label(cell, text=f"{label}:", font=("Arial", 9), bg="#ffffff").pack(side="left")
        tk.Label(cell, text=filename, font=("Arial", 9, "italic"), bg="#ffffff", fg="#2e7d32" if path else "#999999", width=14, anchor="w").pack(side="left", padx=(2, 4))

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

        tk.Button(cell, text="Browse", width=6, bg="#607d8b", fg="white", command=browse).pack(side="left", padx=1)
        tk.Button(cell, text="Open", width=5, bg="#795548", fg="white", command=open_attachment).pack(side="left", padx=1)

    def add_supplier_row(supplier):
        row_outer = tk.Frame(suppliers_container, bg="#ffffff", relief="groove", borderwidth=1, pady=6, padx=6)
        row_outer.pack(fill="x", pady=4, padx=2)

        info_row = tk.Frame(row_outer, bg="#ffffff")
        info_row.pack(fill="x")

        tk.Label(info_row, text=supplier["name"], font=("Arial", 11, "bold"), bg="#ffffff", anchor="w", width=20).pack(side="left")
        tk.Label(info_row, text=f"Tel: {supplier['tel'] or '-'}", font=("Arial", 10), bg="#ffffff", anchor="w", width=18).pack(side="left")
        tk.Label(info_row, text=f"Contact: {supplier['contact'] or '-'}", font=("Arial", 10), bg="#ffffff", anchor="w", width=22).pack(side="left")

        def edit_supplier_action():
            new_name = simpledialog.askstring("Edit Supplier", "Name:", initialvalue=supplier["name"])
            new_tel = simpledialog.askstring("Edit Supplier", "Tel:", initialvalue=supplier["tel"])
            new_contact = simpledialog.askstring("Edit Supplier", "Contact:", initialvalue=supplier["contact"])

            if new_name:
                update_supplier_db(supplier["id"], new_name, new_tel or "", new_contact or "")
                refresh_all()

        def delete_supplier_action():
            if not messagebox.askyesno("Delete", f"Διαγραφή προμηθευτή '{supplier['name']}';"):
                return
            delete_supplier_db(supplier["id"])
            refresh_all()

        tk.Button(info_row, text="Edit", width=5, bg="#2196F3", fg="white", command=edit_supplier_action).pack(side="left", padx=2)
        tk.Button(info_row, text="Delete", width=6, bg="#f44336", fg="white", command=delete_supplier_action).pack(side="left", padx=2)

        attach_row = tk.Frame(row_outer, bg="#ffffff")
        attach_row.pack(fill="x", pady=(4, 0))
        tk.Label(attach_row, text="Attachments:", font=("Arial", 9, "italic"), bg="#ffffff", fg="#555").pack(side="left", padx=(0, 6))

        for key, label in ATTACHMENT_TYPES:
            build_attachment_widget(attach_row, supplier, key, label)

    def add_item_row(item):
        symbol = CURRENCY_SYMBOLS.get(item["currency"], item["currency"])
        total = item["qty"] * item["price"]

        row = tk.Frame(items_container, bg="#ffffff", relief="groove", borderwidth=1, padx=6, pady=4)
        row.pack(fill="x", padx=2, pady=2)

        tk.Label(row, text=item["description"], bg="#ffffff", anchor="w", width=22).pack(side="left")
        tk.Label(row, text=f"Qty: {item['qty']}", bg="#ffffff", anchor="w", width=8).pack(side="left")
        tk.Label(row, text=f"Price: {symbol}{item['price']:.2f}", bg="#ffffff", anchor="w", width=16).pack(side="left")
        tk.Label(row, text=f"Total: {symbol}{total:.2f}", bg="#ffffff", anchor="w", width=16, font=("Arial", 10, "bold")).pack(side="left")

        def delete_item_action():
            delete_supplier_item_db(item["id"])
            refresh_all()

        tk.Button(row, text="Delete", width=6, bg="#f44336", fg="white", command=delete_item_action).pack(side="left", padx=4)

    def update_summary(supplier):
        if not supplier or not supplier["items"]:
            summary_label.config(text="Min: -    Max: -    Average: -")
            return

        by_currency = {}
        for item in supplier["items"]:
            by_currency.setdefault(item["currency"], []).append(item["price"])

        parts = []
        for currency, prices in by_currency.items():
            symbol = CURRENCY_SYMBOLS.get(currency, currency)
            min_p = min(prices)
            max_p = max(prices)
            avg_p = sum(prices) / len(prices)
            parts.append(f"[{currency}] Min: {symbol}{min_p:.2f}   Max: {symbol}{max_p:.2f}   Average: {symbol}{avg_p:.2f}")

        summary_label.config(text="    ".join(parts))

    def refresh_supplier_menu():
        menu = supplier_menu["menu"]
        menu.delete(0, "end")
        supplier_map.clear()

        suppliers = get_all_suppliers()
        names = []

        for s in suppliers:
            display_name = f"{s['name']} (ID:{s['id']})"
            supplier_map[display_name] = s["id"]
            names.append(display_name)

        def pick(name):
            selected_supplier.set(name)
            refresh_items_view()

        for name in names:
            menu.add_command(label=name, command=lambda n=name: pick(n))

        if names and selected_supplier.get() not in names:
            selected_supplier.set(names[0])
        elif not names:
            selected_supplier.set("")

    def refresh_items_view():
        for widget in items_container.winfo_children():
            widget.destroy()

        supplier = get_supplier_by_selected()
        if supplier is None:
            tk.Label(items_container, text="Επίλεξε ή πρόσθεσε έναν προμηθευτή πρώτα.", bg="#ffffff", fg="#777").pack(anchor="w", padx=6, pady=6)
            summary_label.config(text="Min: -    Max: -    Average: -")
            return

        for item in supplier["items"]:
            add_item_row(item)

        update_summary(supplier)

    def refresh_all():
        for widget in suppliers_container.winfo_children():
            widget.destroy()

        for supplier in get_all_suppliers():
            add_supplier_row(supplier)

        refresh_supplier_menu()
        refresh_items_view()

    def add_supplier_action():
        name = name_entry.get().strip()
        tel = tel_entry.get().strip()
        contact = contact_entry.get().strip()

        if not name:
            messagebox.showwarning("Προσοχή", "Το πεδίο Name είναι υποχρεωτικό.")
            return

        add_supplier_db(name, tel, contact)
        name_entry.delete(0, tk.END)
        tel_entry.delete(0, tk.END)
        contact_entry.delete(0, tk.END)
        refresh_all()

    def add_item_action():
        supplier = get_supplier_by_selected()
        if supplier is None:
            messagebox.showwarning("Προσοχή", "Πρόσθεσε πρώτα έναν προμηθευτή.")
            return

        desc = desc_entry.get().strip()
        qty_raw = qty_entry.get().strip()
        price_raw = price_entry.get().strip()
        currency = currency_var.get().strip()

        if not desc or not qty_raw or not price_raw:
            messagebox.showwarning("Προσοχή", "Συμπλήρωσε Description, Qty και Price.")
            return

        try:
            qty = int(qty_raw)
            price = float(price_raw)
        except ValueError:
            messagebox.showwarning("Προσοχή", "Το Qty πρέπει να είναι ακέραιος και το Price αριθμός.")
            return

        add_supplier_item_db(supplier["id"], desc, qty, price, currency)
        desc_entry.delete(0, tk.END)
        qty_entry.delete(0, tk.END)
        price_entry.delete(0, tk.END)
        refresh_all()

    tk.Button(form_frame, text="Add Supplier", bg="#4CAF50", fg="white", command=add_supplier_action).grid(row=0, column=6, padx=10)
    tk.Button(control_row, text="Add Quote Item", bg="#4CAF50", fg="white", command=add_item_action).pack(side="left")

    selected_supplier.trace_add("write", lambda *args: refresh_items_view())

    refresh_all()


# ================= MAIN =================
init_db()

root = tk.Tk()
root.title("Dashboard MSC")
root.geometry("560x380")
root.configure(bg="#f5f5f5")

header = tk.Label(root, text="Dashboard MSC", font=("Arial", 18, "bold"), bg="#f5f5f5", fg="#222")
header.pack(pady=(25, 20))

btn_frame = tk.Frame(root, bg="#f5f5f5")
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="Tasks", width=18, height=2, bg="#2196F3", fg="white", command=open_tasks).grid(row=0, column=0, padx=10, pady=8)
tk.Button(btn_frame, text="Suppliers", width=18, height=2, bg="#FF9800", fg="white", command=open_suppliers).grid(row=0, column=1, padx=10, pady=8)
tk.Button(btn_frame, text="Members", width=18, height=2, bg="#4CAF50", fg="white", command=open_members).grid(row=1, column=0, padx=10, pady=8)
tk.Button(btn_frame, text="Meetings", width=18, height=2, bg="#9C27B0", fg="white", command=open_meetings).grid(row=1, column=1, padx=10, pady=8)
tk.Button(btn_frame, text="Stocklist", width=18, height=2, bg="#795548", fg="white", command=open_stocklist).grid(row=2, column=0, columnspan=2, padx=10, pady=8)

footer = tk.Label(root, text="Local SQLite storage enabled", font=("Arial", 10, "italic"), bg="#f5f5f5", fg="#666")
footer.pack(pady=(15, 0))

root.mainloop().\.venv\Scripts\python.exe dashboard.py