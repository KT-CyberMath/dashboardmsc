import os
import sys
import sqlite3
import subprocess
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
from datetime import datetime

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


# ================= DATABASE =================
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
    cur.execute("""
        UPDATE emails
        SET sender = ?, subject = ?, body = ?, expanded = ?
        WHERE id = ?
    """, (sender, subject, body, expanded, email_id))
    conn.commit()
    conn.close()


def delete_email_db(email_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM emails WHERE id = ?", (email_id,))
    conn.commit()
    conn.close()


def set_email_expanded_db(email_id, expanded):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE emails SET expanded = ? WHERE id = ?", (expanded, email_id))
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
    cur.execute(
        "INSERT INTO tasks (text, date, done) VALUES (?, ?, ?)",
        (text, date, done)
    )
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def update_task_db(task_id, text):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE tasks SET text = ? WHERE id = ?", (text, task_id))
    conn.commit()
    conn.close()


def set_task_done_db(task_id, done):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE tasks SET done = ? WHERE id = ?", (done, task_id))
    conn.commit()
    conn.close()


def delete_task_db(task_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM tasks WHERE id = ?", (task_id,))
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
            "SELECT type_key, file_path FROM supplier_attachments WHERE supplier_id = ?",
            (supplier["id"],)
        )
        attachments = {r["type_key"]: r["file_path"] for r in cur.fetchall()}

        cur.execute(
            "SELECT * FROM supplier_items WHERE supplier_id = ? ORDER BY id DESC",
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
    cur.execute("""
        UPDATE suppliers
        SET name = ?, tel = ?, contact = ?
        WHERE id = ?
    """, (name, tel, contact, supplier_id))
    conn.commit()
    conn.close()


def delete_supplier_db(supplier_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM suppliers WHERE id = ?", (supplier_id,))
    conn.commit()
    conn.close()


def set_supplier_attachment_db(supplier_id, type_key, file_path):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        UPDATE supplier_attachments
        SET file_path = ?
        WHERE supplier_id = ? AND type_key = ?
    """, (file_path, supplier_id, type_key))
    conn.commit()
    conn.close()


def add_supplier_item_db(supplier_id, description, qty, price, currency):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO supplier_items (supplier_id, description, qty, price, currency)
        VALUES (?, ?, ?, ?, ?)
    """, (supplier_id, description, qty, price, currency))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def delete_supplier_item_db(item_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM supplier_items WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()


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

            tk.Button(btns_row, text="Edit", width=5, bg="#2196F3", fg="white",
                      command=edit_email_action).pack(side="left", padx=2)
            tk.Button(btns_row, text="Delete", width=6, bg="#f44336", fg="white",
                      command=delete_email_action).pack(side="left", padx=2)

            if email_dict["expanded"]:
                tk.Label(card, text=f"Από: {email_dict['sender']}", font=("Arial", 9),
                         bg="#cfe8ff", anchor="w").pack(fill="x", pady=(4, 2))

                body_lbl = tk.Label(card, text=email_dict["body"], font=("Arial", 10),
                                    bg="#d9f2d9", anchor="w", justify="left",
                                    wraplength=320, cursor="hand2", relief="raised", padx=8, pady=6)
                body_lbl.pack(fill="x", pady=(0, 4))

                body_lbl.bind("<ButtonPress-1>", on_drag_start(lambda: email_dict["body"]))
                body_lbl.bind("<ButtonRelease-1>", on_drag_release)

                tk.Label(card, text="↳ Σύρε το Θέμα ή το Κείμενο στη λίστα Tasks δεξιά →",
                         font=("Arial", 8, "italic"), bg="#cfe8ff", fg="#555").pack(pady=(0, 2))

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

        tk.Button(btn_frame, text="Add", bg="#4CAF50", fg="white", width=10,
                  command=submit_email).pack(side="left", padx=5)
        tk.Button(btn_frame, text="Cancel", bg="#f44336", fg="white", width=10,
                  command=form.destroy).pack(side="left", padx=5)

    tk.Button(new_email_btn_frame, text="+ Νέο Email", bg="#4CAF50", fg="white",
              command=open_new_email_form).pack()

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

            tk.Label(top_row, text=task_dict["text"], font=task_font, fg=task_color,
                     bg="#ffffff", anchor="w", justify="left", wraplength=220).pack(side="left", fill="x", expand=True)

            def edit_task_action():
                new_text = simpledialog.askstring("Edit Task", "Επεξεργασία task:", initialvalue=task_dict["text"])
                if new_text:
                    task_dict["text"] = new_text
                    update_task_db(task_dict["id"], new_text)
                    refresh_row()

            def delete_task_action():
                delete_task_db(task_dict["id"])
                row_outer.destroy()

            tk.Button(top_row, text="Edit", width=5, bg="#2196F3", fg="white",
                      command=edit_task_action).pack(side="left", padx=2)
            tk.Button(top_row, text="Delete", width=6, bg="#f44336", fg="white",
                      command=delete_task_action).pack(side="left", padx=2)

            bottom_row = tk.Frame(row_outer, bg="#ffffff")
            bottom_row.pack(fill="x", pady=(2, 0))
            tk.Label(bottom_row, text=f"📅 Καταχωρήθηκε: {task_dict['date']}",
                     font=("Arial", 8, "italic"), fg="#777777", bg="#ffffff").pack(side="left")

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


# ================= SUPPLIERS WINDOW =================
def open_suppliers():
    win = tk.Toplevel(root)
    win.title("Suppliers")
    win.geometry("980x700")
    win.configure(bg="#f5f5f5")

    supplier_map = {}

    form_frame = tk.LabelFrame(win, text="Add Supplier", font=("Arial", 12, "bold"),
                               bg="#f5f5f5", padx=10, pady=10)
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

    quote_frame = tk.LabelFrame(win, text="Price Quotes", font=("Arial", 12, "bold"),
                                bg="#f5f5f5", padx=10, pady=10)
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

    summary_label = tk.Label(quote_frame, text="Min: -    Max: -    Average: -",
                             font=("Arial", 11, "bold"), bg="#f5f5f5", anchor="w")
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
        tk.Label(cell, text=filename, font=("Arial", 9, "italic"), bg="#ffffff",
                 fg="#2e7d32" if path else "#999999", width=14, anchor="w").pack(side="left", padx=(2, 4))

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

        tk.Label(info_row, text=supplier["name"], font=("Arial", 11, "bold"),
                 bg="#ffffff", anchor="w", width=20).pack(side="left")
        tk.Label(info_row, text=f"Tel: {supplier['tel'] or '-'}", font=("Arial", 10),
                 bg="#ffffff", anchor="w", width=18).pack(side="left")
        tk.Label(info_row, text=f"Contact: {supplier['contact'] or '-'}", font=("Arial", 10),
                 bg="#ffffff", anchor="w", width=22).pack(side="left")

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

        tk.Button(info_row, text="Edit", width=5, bg="#2196F3", fg="white",
                  command=edit_supplier_action).pack(side="left", padx=2)
        tk.Button(info_row, text="Delete", width=6, bg="#f44336", fg="white",
                  command=delete_supplier_action).pack(side="left", padx=2)

        attach_row = tk.Frame(row_outer, bg="#ffffff")
        attach_row.pack(fill="x", pady=(4, 0))
        tk.Label(attach_row, text="Attachments:", font=("Arial", 9, "italic"),
                 bg="#ffffff", fg="#555").pack(side="left", padx=(0, 6))

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
        tk.Label(row, text=f"Total: {symbol}{total:.2f}", bg="#ffffff", anchor="w", width=16,
                 font=("Arial", 10, "bold")).pack(side="left")

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
            tk.Label(items_container, text="Επίλεξε ή πρόσθεσε έναν προμηθευτή πρώτα.",
                     bg="#ffffff", fg="#777").pack(anchor="w", padx=6, pady=6)
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
        if not desc:
            messagebox.showwarning("Προσοχή", "Το πεδίο Description είναι υποχρεωτικό.")
            return

        try:
            qty = int(qty_entry.get().strip())
            if qty <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Προσοχή", "Το Qty πρέπει να είναι θετικός ακέραιος.")
            return

        try:
            price = float(price_entry.get().strip())
            if price < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Προσοχή", "Το Price πρέπει να είναι αριθμός.")
            return

        add_supplier_item_db(supplier["id"], desc, qty, price, currency_var.get())
        desc_entry.delete(0, tk.END)
        qty_entry.delete(0, tk.END)
        price_entry.delete(0, tk.END)
        refresh_all()

    tk.Button(form_frame, text="Add Supplier", bg="#4CAF50", fg="white",
              command=add_supplier_action).grid(row=0, column=6, padx=10)

    tk.Button(control_row, text="Add Item", bg="#4CAF50", fg="white",
              command=add_item_action).pack(side="left")

    refresh_all()


# ================= PLACEHOLDERS =================
def open_meetings():
    messagebox.showinfo("Meetings", "Άνοιξε η σελίδα Meetings")


def open_members():
    messagebox.showinfo("Members", "Άνοιξε η σελίδα Members")


def open_stocklist():
    messagebox.showinfo("Stocklist", "Άνοιξε η σελίδα Stocklist")


# ================= MAIN =================
init_db()

root = tk.Tk()
root.title("Dashboard")
root.geometry("400x400")
root.configure(bg="#f0f0f0")

title = tk.Label(root, text="Dashboard", font=("Arial", 20, "bold"), bg="#f0f0f0")
title.pack(pady=20)

buttons_frame = tk.Frame(root, bg="#f0f0f0")
buttons_frame.pack(expand=True)

btn_style = {"width": 20, "height": 2, "font": ("Arial", 12), "bg": "#4CAF50", "fg": "white"}

tk.Button(buttons_frame, text="Tasks", command=open_tasks, **btn_style).pack(pady=8)
tk.Button(buttons_frame, text="Meetings", command=open_meetings, **btn_style).pack(pady=8)
tk.Button(buttons_frame, text="Suppliers", command=open_suppliers, **btn_style).pack(pady=8)
tk.Button(buttons_frame, text="Members", command=open_members, **btn_style).pack(pady=8)
tk.Button(buttons_frame, text="Stocklist", command=open_stocklist, **btn_style).pack(pady=8)

root.mainloop()