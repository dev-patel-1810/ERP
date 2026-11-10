from datetime import datetime
from functools import wraps
from urllib.parse import quote
from flask import (
    Flask,
    flash,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
)
import io
import pandas as pd
import sqlite3
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = "super_secret_local_inventory_key"
DB_NAME = "inventory.db"


def init_db():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE,
            password TEXT,
            role TEXT
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS units (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unit_name TEXT UNIQUE
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS item_units (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER UNIQUE,
            unit_id INTEGER,
            FOREIGN KEY(item_id) REFERENCES master_items(id) ON DELETE CASCADE,
            FOREIGN KEY(unit_id) REFERENCES units(id) ON DELETE SET NULL
        )
    """)
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS master_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_name TEXT UNIQUE,
        item_category TEXT DEFAULT 'Brass Item',
        min_quantity REAL DEFAULT 0,
        min_casting_quantity REAL DEFAULT 0,
        min_machined_quantity REAL DEFAULT 0,
        min_assembly_quantity REAL DEFAULT 0,
        min_purchase_quantity REAL DEFAULT 0
    )
  """)
  # Add new minimum-quantity columns to existing databases.
  cursor.execute("PRAGMA table_info(master_items)")
  existing_columns = {row[1] for row in cursor.fetchall()}

  new_columns = {
    "min_quantity": "REAL DEFAULT 0",
    "min_casting_quantity": "REAL DEFAULT 0",
    "min_machined_quantity": "REAL DEFAULT 0",
    "min_assembly_quantity": "REAL DEFAULT 0",
    "min_purchase_quantity": "REAL DEFAULT 0",
  }

  for column_name, column_definition in new_columns.items():
    if column_name not in existing_columns:
        cursor.execute(
            f"ALTER TABLE master_items ADD COLUMN {column_name} {column_definition}"
        )

  conn.commit()
  try:
    cursor.execute("ALTER TABLE master_items ADD COLUMN min_quantity REAL DEFAULT 0")
  except sqlite3.OperationalError:
    pass
  cursor.execute("""
    CREATE TABLE IF NOT EXISTS finished_goods (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        box_pack_id INTEGER,
        box_pack_name TEXT,
        pcs INTEGER DEFAULT 0,
        FOREIGN KEY(box_pack_id) REFERENCES box_pack_items(id),
        UNIQUE(box_pack_id)
    )
""")

  cursor.execute("""
    CREATE TABLE IF NOT EXISTS finished_goods_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT,
        box_pack_id INTEGER,
        box_pack_name TEXT,
        pcs INTEGER,
        FOREIGN KEY(box_pack_id) REFERENCES box_pack_items(id)
    )
""")

  cursor.execute("""
    CREATE TABLE IF NOT EXISTS finished_goods_assembly_deductions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        finished_goods_log_id INTEGER,
        item_id INTEGER,
        item_name TEXT,
        pcs INTEGER,
        FOREIGN KEY(finished_goods_log_id) REFERENCES finished_goods_log(id),
        FOREIGN KEY(item_id) REFERENCES master_items(id)
    )
""")
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS vendors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vendor_name TEXT UNIQUE,
            contact_info TEXT
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS vendor_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vendor_id INTEGER,
            item_id INTEGER,
            FOREIGN KEY(vendor_id) REFERENCES vendors(id) ON DELETE CASCADE,
            FOREIGN KEY(item_id) REFERENCES master_items(id) ON DELETE CASCADE
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS box_pack_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            box_pack_name TEXT UNIQUE
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS box_pack_components (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            box_pack_id INTEGER,
            item_id INTEGER,
            FOREIGN KEY(box_pack_id) REFERENCES box_pack_items(id) ON DELETE CASCADE,
            FOREIGN KEY(item_id) REFERENCES master_items(id) ON DELETE CASCADE
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS vendor_registers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vendor_id INTEGER,
            register_type TEXT,
            FOREIGN KEY(vendor_id) REFERENCES vendors(id) ON DELETE CASCADE
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS brass_in (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bag_number INTEGER,
            date TEXT,
            vendor TEXT,
            weight REAL,
            remaining_weight REAL,
            status TEXT DEFAULT 'available'
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS brass_out (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            vendor TEXT,
            bag_number INTEGER,
            weight REAL,
            brass_in_id INTEGER
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS casting_in (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            vendor TEXT,
            item TEXT,
            pieces INTEGER,
            weight REAL
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS machining_register (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            item TEXT,
            direction TEXT,
            pieces INTEGER,
            weight REAL
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS plating_challans (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            challan_number TEXT,
            date TEXT,
            vendor TEXT,
            direction TEXT,
            status TEXT DEFAULT 'open'
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS plating_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            challan_id INTEGER,
            item TEXT,
            pieces INTEGER,
            weight REAL,
            bag_number INTEGER,
            remaining_pieces INTEGER,
            grinding_vendor TEXT,
            FOREIGN KEY(challan_id) REFERENCES plating_challans(id) ON DELETE CASCADE
        )
    """)
  try:
    cursor.execute("ALTER TABLE plating_items ADD COLUMN grinding_vendor TEXT")
  except sqlite3.OperationalError:
    pass

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS miscellaneous_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            option_type TEXT,
            vendor TEXT,
            item TEXT,
            pieces INTEGER,
            weight REAL
        )
    """)
  try:
    cursor.execute("ALTER TABLE miscellaneous_entries ADD COLUMN quantity_unit TEXT DEFAULT ''")
  except sqlite3.OperationalError:
    pass
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS rejection_register (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            item TEXT,
            rejection_type TEXT,
            vendor TEXT,
            pieces INTEGER,
            weight REAL
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            username TEXT,
            action TEXT,
            details TEXT
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS sale_register (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            party TEXT,
            bill_no TEXT,
            box_pack_id INTEGER,
            box_pack_name TEXT,
            pcs INTEGER,
            no_of_cartons INTEGER,
            FOREIGN KEY(box_pack_id) REFERENCES box_pack_items(id)
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS sale_register_cartons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sale_id INTEGER,
            carton_number INTEGER,
            weight REAL,
            FOREIGN KEY(sale_id) REFERENCES sale_register(id) ON DELETE CASCADE
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS sale_register_deductions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sale_id INTEGER,
            item_id INTEGER,
            item_name TEXT,
            pcs INTEGER,
            FOREIGN KEY(sale_id) REFERENCES sale_register(id) ON DELETE CASCADE,
            FOREIGN KEY(item_id) REFERENCES master_items(id)
        )
    """)
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS opening_stock (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item TEXT,
            stage TEXT,
            pieces INTEGER DEFAULT 0,
            weight REAL DEFAULT 0,
            date TEXT,
            UNIQUE(item, stage)
        )
    """)

  cursor.execute("""
        CREATE TABLE IF NOT EXISTS rate_list (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_id INTEGER NOT NULL,
            vendor_id INTEGER NOT NULL,
            rate REAL NOT NULL,
            FOREIGN KEY(item_id) REFERENCES master_items(id) ON DELETE CASCADE,
            FOREIGN KEY(vendor_id) REFERENCES vendors(id) ON DELETE CASCADE
        )
    """)

  # Payment tracking: add a "payment_made" flag to every table that feeds a
  # Vendor Ledger entry, so each ledger row can be ticked once payment has
  # been made to that vendor. Runs after all CREATE TABLE statements above
  # so it works both on a fresh DB and on an existing one being upgraded.
  for _pm_table in (
      "brass_in",
      "brass_out",
      "casting_in",
      "miscellaneous_entries",
      "plating_items",
  ):
    try:
      cursor.execute(
          f"ALTER TABLE {_pm_table} ADD COLUMN payment_made INTEGER DEFAULT 0"
      )
    except sqlite3.OperationalError:
      pass

  cursor.execute("SELECT * FROM users WHERE username = 'admin'")
  if not cursor.fetchone():
    cursor.execute(
        "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
        ("admin", "admin123", "admin"),
    )
  cursor.execute("SELECT * FROM users WHERE username = 'employee'")
  if not cursor.fetchone():
    cursor.execute(
        "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
        ("employee", "emp123", "employee"),
    )

  cursor.execute("SELECT COUNT(*) FROM units")
  if cursor.fetchone()[0] == 0:
    default_units = [("KG",), ("PCS",), ("Dozen",), ("Grams",), ("Boxes",)]
    cursor.executemany("INSERT INTO units (unit_name) VALUES (?)", default_units)

  try:
    cursor.execute("ALTER TABLE brass_out ADD COLUMN give_to TEXT DEFAULT ''")
    cursor.execute("ALTER TABLE brass_out RENAME COLUMN give_to TO given_to")
    cursor.connection.commit()
  except:
    pass
  conn.commit()
  conn.close()


init_db()


def log_action(username, action, details, conn=None):
    _close = False
    if conn is None:
        conn = sqlite3.connect(DB_NAME)
        _close = True
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO audit_logs (timestamp, username, action, details) VALUES (?, ?, ?, ?)",
        (datetime.now().isoformat(), username, action, details),
    )
    if _close:
        conn.commit()
        conn.close()


def payment_checkbox_html(table, row_id, paid, ledger_sub_type, ledger_vendor=""):
  """Renders a self-submitting 'Payment Made' tick-box for one Vendor Ledger
  entry. Ticking/unticking it POSTs to /registers (sub_action=toggle_payment)
  and redirects straight back to the same ledger view."""
  checked = "checked" if paid else ""
  return f"""<form method="POST" style="display:inline;">
        <input type="hidden" name="sub_action" value="toggle_payment">
        <input type="hidden" name="table" value="{table}">
        <input type="hidden" name="row_id" value="{row_id}">
        <input type="hidden" name="ledger_sub_type" value="{ledger_sub_type}">
        <input type="hidden" name="ledger_vendor" value="{ledger_vendor}">
        <input type="checkbox" name="payment_made" {checked} title="Payment Made" onchange="this.form.submit()">
    </form>"""


def is_admin():
  return session.get("role") == "admin"


def admin_delete_cell(delete_url, confirm_msg="Delete this entry? This cannot be undone."):
  """Renders a Delete button/link for admins only; empty string otherwise.

  Intended to be dropped straight into a <td> in any register table.
  """
  if not is_admin():
    return ""
  return (
      f'<a href="{delete_url}"'
      f" onclick=\"return confirm('{confirm_msg}')\">"
      '<button class="danger" style="padding:2px 10px;font-size:12px;">Delete</button></a>'
  )


def login_required(f):
  @wraps(f)
  def decorated_function(*args, **kwargs):
    if "user" not in session:
      return redirect(url_for("login"))
    return f(*args, **kwargs)

  return decorated_function


def admin_required(f):
  @wraps(f)
  def decorated_function(*args, **kwargs):
    if "user" not in session or session.get("role") != "admin":
      return redirect(url_for("add_items"))
    return f(*args, **kwargs)

  return decorated_function


HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Registers System</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body { font-family: Arial, sans-serif; background: #f4f4f9; margin: 0; padding: 20px; }
        .container { max-width: 1050px; margin: auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h2, h3, h4 { color: #333; }
        table { width: 100%%; border-collapse: collapse; margin-top: 15px; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; font-size: 14px; }
        th { background-color: #007bff; color: white; }
        input, select { padding: 8px; margin: 5px 0 15px 0; width: 100%%; box-sizing: border-box; font-size: 16px; }
        button { background: #007bff; color: white; border: none; padding: 10px 15px; cursor: pointer; border-radius: 4px; font-size: 15px; }
        button:hover { background: #0056b3; }
        .danger { background: #dc3545; }
        .danger:hover { background: #c82333; }
        .nav { margin-bottom: 20px; padding-bottom: 10px; border-bottom: 1px solid #eee; display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }
        .nav a { margin-right: 15px; text-decoration: none; color: #007bff; font-weight: bold; }
        .card { background: #f9f9f9; padding: 15px; border-radius: 6px; margin-bottom: 20px; border: 1px solid #e1e1e1; }
        .info-box { background: #e7f3fe; border-left: 6px solid #2196F3; padding: 10px; margin-bottom: 15px; font-size: 14px; }
        .checkbox-group { max-height: 150px; overflow-y: auto; background: white; padding: 10px; border: 1px solid #ccc; margin-bottom: 15px; }
        .ledger-container { display: flex; gap: 20px; flex-wrap: wrap; }
        .ledger-side { flex: 1; min-width: 450px; background: #fff; border: 1px solid #ddd; padding: 15px; border-radius: 6px; }
        .item-records-horizontal { display: flex; gap: 20px; overflow-x: auto; padding-bottom: 15px; align-items: flex-start; }
        .item-stage-box { min-width: 320px; flex: 0 0 auto; background: #fff; border: 1px solid #ddd; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .item-stage-box h4 { margin-top: 0; border-bottom: 2px solid #007bff; padding-bottom: 5px; color: #007bff; }

        /* ---- Mobile responsiveness ---- */
        .table-scroll { width: 100%%; overflow-x: auto; -webkit-overflow-scrolling: touch; margin-bottom: 20px; }
        .table-scroll table { margin-bottom: 0; }

        @media (max-width: 768px) {
            body { padding: 8px; }
            .container { padding: 12px; border-radius: 4px; }
            h2 { font-size: 20px; }
            h3 { font-size: 17px; }
            h4 { font-size: 15px; }
            .nav a { padding: 6px 10px; background: #eef4ff; border-radius: 4px; margin-right: 0; font-size: 14px; }
            table:not(.table-scroll table) { display: block; overflow-x: auto; white-space: nowrap; -webkit-overflow-scrolling: touch; }
            th, td { font-size: 13px; padding: 8px; }
            .card button[type="submit"] { width: 100%%; padding: 12px; margin-top: 6px; }
            .ledger-side { min-width: 90vw; }
            .item-stage-box { min-width: 85vw; }
            .ledger-container, .item-records-horizontal { gap: 12px; }
        }
    </style>
</head>
<body>
<div class="container">
    {% block content %}{% endblock %}
</div>

<div id="alert-panel" style="
    position:fixed;
    bottom:20px;
    right:20px;
    width:300px;
    max-height:60vh;
    overflow-y:auto;
    background:#dc3545;
    color:white;
    border-radius:8px;
    box-shadow:0 4px 16px rgba(0,0,0,0.4);
    z-index:99999;
    display:none;
    padding:14px 14px 10px 14px;
    font-size:14px;
    font-family:Arial,sans-serif;">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
    <b style="font-size:15px;">&#9888; Low Stock Alerts</b>
    <span onclick="document.getElementById('alert-panel').style.display='none';"
          style="cursor:pointer;font-size:22px;line-height:1;color:white;padding:0 4px;">&times;</span>
  </div>
  <div id="alert-list"></div>
</div>

<script>
(function() {
  setTimeout(function() {
    fetch('/api/low-stock-alerts')
      .then(function(r) { if (!r.ok) throw new Error('not ok'); return r.json(); })
      .then(function(alerts) {
        // Update badge on bell button
        var badge = document.getElementById('alert-badge');
        if (badge && alerts && alerts.length > 0) {
          badge.textContent = alerts.length;
          badge.style.display = 'inline-flex';
        }

        if (!alerts || alerts.length === 0) return;

        var list = document.getElementById('alert-list');
        list.innerHTML = '';
        alerts.forEach(function(a) {
          var link = a.category === 'Brass Item'
            ? '/registers?type=Brass+In'
            : '/registers?type=Miscellaneous&prefill_option=purchase&prefill_item=' + encodeURIComponent(a.name);
          var div = document.createElement('div');
          div.style.cssText = 'border-top:1px solid rgba(255,255,255,0.25);padding:7px 0;';
          div.innerHTML = '<b>' + a.name + '</b><br>'
            + '<span style="font-size:12px;">Stock: ' + a.qty + ' pcs &nbsp;|&nbsp; Min: ' + a.min_qty + ' pcs</span><br>'
            + '<a href="' + link + '" style="color:#fff;font-size:12px;text-decoration:underline;">Add Purchase &rarr;</a>';
          list.appendChild(div);
        });

        // Auto-show the panel on page load
        document.getElementById('alert-panel').style.display = 'block';
      })
      .catch(function(e) { console.log('Alert fetch error:', e); });
  }, 300);
})();
document.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON') {
        var fields = Array.from(document.querySelectorAll('input, select, textarea, button'));
        var idx = fields.indexOf(e.target);
        if (idx !== -1 && idx + 1 < fields.length) {
            e.preventDefault();
            fields[idx + 1].focus();
        }
    }
});
</script>

</body>
</html>
"""


def get_nav_links():
  links = '<a href="/add-items">Add Items</a>'
  links += '<a href="/registers">Registers</a>'
  links += '<a href="/audit-logs">Audit Logs</a>'
  if session.get("role") == "admin":
    links += '<a href="/manage-users">Manage User</a>'
  links += '<a href="/logout">Log Out</a>'
  links += '''<span 
      onclick="var p=document.getElementById(\'alert-panel\'); p.style.display=(p.style.display===\'none\'?\'block\':\'none\');"
      id="alert-bell-btn"
      title="View Low Stock Alerts"
      style="cursor:pointer; background:#dc3545; color:white; border-radius:6px;
             padding:6px 12px; font-size:13px; font-weight:bold; margin-left:8px;
             display:inline-flex; align-items:center; gap:5px; user-select:none;">
    &#9888; Alerts
    <span id="alert-badge" style="display:none; background:white; color:#dc3545;
          border-radius:50%; width:18px; height:18px; font-size:11px; font-weight:bold;
          display:none; align-items:center; justify-content:center;">0</span>
  </span>'''
  return links

@app.route("/login", methods=["GET", "POST"])
def login():
  if request.method == "POST":
    username = request.form["username"]
    password = request.form["password"]
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT * FROM users WHERE username = ? AND password = ?",
        (username, password),
    )
    user = cursor.fetchone()
    conn.close()
    if user:
      session["user"] = user[1]
      session["role"] = user[3]
      log_action(user[1], "LOGIN", "User logged into the system")
      return redirect(url_for("add_items"))
    else:
      return render_template_string(
          HTML_TEMPLATE
          + """
            <h2>Login</h2>
            <p style="color:red;">Invalid credentials.</p>
            <form method="POST">
                Username: <input type="text" name="username" required>
                Password: <input type="password" name="password" required>
                <button type="submit">Login</button>
            </form>
        """
      )
  return render_template_string(
      HTML_TEMPLATE
      + """
        <h2>Login</h2>
        <form method="POST">
            Username: <input type="text" name="username" required>
            Password: <input type="password" name="password" required>
            <button type="submit">Login</button>
        </form>
    """
  )


@app.route("/logout")
def logout():
  if "user" in session:
    log_action(session["user"], "LOGOUT", "User logged out")
  session.clear()
  return redirect(url_for("login"))


@app.route("/")
@login_required
def index():
  return redirect(url_for("add_items"))


@app.route("/add-items", methods=["GET", "POST"])
@login_required
def add_items():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  selected_type = request.values.get("type", "")

  if request.method == "POST":
    target_type = request.form.get("target_type")
    if target_type == "Items":
        item_name = request.form.get("item_name", "").strip()
        item_category = request.form.get(
            "item_category", "Brass Item"
        ).strip()

        min_quantity = float(request.form.get("min_quantity") or 0)

        min_casting_quantity = 0
        min_machined_quantity = 0
        min_assembly_quantity = 0
        min_purchase_quantity = 0

        if item_category == "Brass Item":
            min_casting_quantity = float(
                request.form.get("min_casting_quantity") or 0
            )
            min_machined_quantity = float(
                request.form.get("min_machined_quantity") or 0
            )
            min_assembly_quantity = float(
                request.form.get("min_assembly_quantity") or 0
            )

        elif item_category == "Miscellaneous Item":
            min_purchase_quantity = float(
                request.form.get("min_purchase_quantity") or 0
            )

        if item_name:
            try:
                cursor.execute(
                    """
                    INSERT INTO master_items (
                        item_name,
                        item_category,
                        min_quantity,
                        min_casting_quantity,
                        min_machined_quantity,
                        min_assembly_quantity,
                        min_purchase_quantity
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        item_name,
                        item_category,
                        min_quantity,
                        min_casting_quantity,
                        min_machined_quantity,
                        min_assembly_quantity,
                        min_purchase_quantity,
                    ),
                )

                conn.commit()

                log_action(
                    session["user"],
                    "CREATE ITEM",
                    f"Added item '{item_name}' ({item_category})",
                )

            except sqlite3.IntegrityError:
                pass

        return redirect(url_for("add_items", type="Items"))
    elif target_type == "Vendors":
      vendor_name = request.form.get("vendor_name", "").strip()
      contact_info = request.form.get("contact_info", "").strip()
      selected_items = request.form.getlist("item_ids")
      selected_registers = request.form.getlist("register_types")
      if vendor_name:
        try:
          cursor.execute(
              "INSERT INTO vendors (vendor_name, contact_info) VALUES (?, ?)",
              (vendor_name, contact_info),
          )
          vendor_id = cursor.lastrowid
          for item_id in selected_items:
            cursor.execute(
                "INSERT INTO vendor_items (vendor_id, item_id) VALUES (?, ?)",
                (vendor_id, item_id),
            )
          for reg_type in selected_registers:
            cursor.execute(
                "INSERT INTO vendor_registers (vendor_id, register_type) VALUES"
                " (?, ?)",
                (vendor_id, reg_type),
            )
          conn.commit()
          log_action(
              session["user"], "CREATE VENDOR", f"Added vendor '{vendor_name}'"
          )
        except sqlite3.IntegrityError:
          pass
      return redirect(url_for("add_items", type="Vendors"))

    elif target_type == "Box Pack Items":
      box_pack_name = request.form.get("box_pack_name", "").strip()
      selected_items = request.form.getlist("item_ids")
      if box_pack_name and selected_items:
        try:
          cursor.execute(
              "INSERT INTO box_pack_items (box_pack_name) VALUES (?)",
              (box_pack_name,),
          )
          box_pack_id = cursor.lastrowid
          for item_id in selected_items:
            cursor.execute(
                "INSERT INTO box_pack_components (box_pack_id, item_id) VALUES"
                " (?, ?)",
                (box_pack_id, item_id),
            )
          conn.commit()
          log_action(
              session["user"],
              "CREATE BOX PACK ITEM",
              f"Added box pack '{box_pack_name}' with {len(selected_items)}"
              " component item(s)",
          )
        except sqlite3.IntegrityError:
          pass
      return redirect(url_for("add_items", type="Box Pack Items"))

    elif target_type == "unit of measurement":
        unit_name = request.form.get("unit_name", "").strip()
        if unit_name:
            try:
                cursor.execute(
                    "INSERT INTO units (unit_name) VALUES (?)", (unit_name,)
                )
                conn.commit()
                log_action(
                    session["user"],
                    "CREATE UNIT",
                    f"Added measurement unit '{unit_name}'",
                )
            except sqlite3.IntegrityError:
                pass
        return redirect(url_for("add_items", type="unit of measurement"))

    elif target_type == "edit_unit":
        unit_id = request.form.get("unit_id")
        unit_name = request.form.get("unit_name", "").strip()
        if unit_id and unit_name:
            try:
                cursor.execute(
                    "UPDATE units SET unit_name = ? WHERE id = ?",
                    (unit_name, unit_id),
                )
                conn.commit()
                log_action(
                    session["user"], "UPDATE UNIT",
                    f"Renamed unit id={unit_id} to '{unit_name}'",
                )
            except sqlite3.IntegrityError:
                pass
        return redirect(url_for("add_items", type="unit of measurement"))

    elif target_type == "assign_unit_items":
        unit_id = request.form.get("unit_id")
        selected_items = request.form.getlist("item_ids")

        if unit_id:
            try:
                # Remove all existing assignments for this unit.
                cursor.execute(
                    "UPDATE item_units SET unit_id = NULL WHERE unit_id = ?",
                    (unit_id,),
                )

                # Assign the selected items to this unit.
                for item_id in selected_items:
                    cursor.execute(
                        """
                        INSERT INTO item_units (item_id, unit_id)
                        VALUES (?, ?)
                        ON CONFLICT(item_id)
                        DO UPDATE SET unit_id = excluded.unit_id
                        """,
                        (item_id, unit_id),
                    )

                conn.commit()

                log_action(
                    session["user"],
                    "ASSIGN ITEM UNIT",
                    f"unit_id={unit_id}, item_ids={','.join(selected_items)}",
                )

            except sqlite3.IntegrityError:
                conn.rollback()

        return redirect(
            url_for("add_items", type="unit of measurement")
        )
    elif target_type == "edit_item":
        item_id = request.form.get("item_edit_id")
        item_name = request.form.get("item_name", "").strip()
        item_category = request.form.get(
            "item_category", "Brass Item"
        ).strip()

        min_quantity = float(
            request.form.get("min_quantity") or 0
        )

        min_casting_quantity = 0
        min_machined_quantity = 0
        min_assembly_quantity = 0
        min_purchase_quantity = 0

        if item_category == "Brass Item":
            min_casting_quantity = float(
                request.form.get("min_casting_quantity") or 0
            )
            min_machined_quantity = float(
                request.form.get("min_machined_quantity") or 0
            )
            min_assembly_quantity = float(
                request.form.get("min_assembly_quantity") or 0
            )

        elif item_category == "Miscellaneous Item":
            min_purchase_quantity = float(
                request.form.get("min_purchase_quantity") or 0
            )

        if item_id and item_name:
            try:
                cursor.execute(
                    """
                    UPDATE master_items
                    SET
                        item_name = ?,
                        item_category = ?,
                        min_quantity = ?,
                        min_casting_quantity = ?,
                        min_machined_quantity = ?,
                        min_assembly_quantity = ?,
                        min_purchase_quantity = ?
                    WHERE id = ?
                    """,
                    (
                        item_name,
                        item_category,
                        min_quantity,
                        min_casting_quantity,
                        min_machined_quantity,
                        min_assembly_quantity,
                        min_purchase_quantity,
                        item_id,
                    ),
                )

                conn.commit()

                log_action(
                    session["user"],
                    "UPDATE ITEM",
                    f"Edited item id={item_id} -> '{item_name}' ({item_category})",
                )

            except sqlite3.IntegrityError:
                pass

        return redirect(url_for("add_items", type="Items"))
    elif target_type == "Rate List":
      import json
      if session.get("role") != "admin":
        flash("Only admins can edit the rate list.")
        return redirect(url_for("add_items", type="Rate List"))
      rate_data_raw = request.form.get("rate_data", "[]")
      try:
        rows = json.loads(rate_data_raw)
      except (json.JSONDecodeError, ValueError):
        rows = []
      for row in rows:
        action = row.get("action")
        rate_id = row.get("id")
        item_id = row.get("item_id")
        vendor_id = row.get("vendor_id")
        rate = row.get("rate")
        if action == "insert" and item_id and vendor_id and rate is not None:
          cursor.execute(
              "INSERT INTO rate_list (item_id, vendor_id, rate) VALUES (?, ?, ?)",
              (item_id, vendor_id, float(rate)),
          )
          log_action(session["user"], "RATE LIST INSERT",
           f"item_id={item_id}, vendor_id={vendor_id}, rate={rate}", conn=conn)
        elif action == "update" and rate_id and item_id and vendor_id and rate is not None:
          cursor.execute(
              "UPDATE rate_list SET item_id=?, vendor_id=?, rate=? WHERE id=?",
              (item_id, vendor_id, float(rate), rate_id),
          )
          log_action(session["user"], "RATE LIST UPDATE",
           f"id={rate_id}, item_id={item_id}, vendor_id={vendor_id}, rate={rate}", conn=conn)
        elif action == "delete" and rate_id:
          cursor.execute("DELETE FROM rate_list WHERE id=?", (rate_id,))
          log_action(session["user"], "RATE LIST DELETE", f"id={rate_id}", conn=conn)
      conn.commit()
      return redirect(url_for("add_items", type="Rate List"))

  cursor.execute("SELECT id, item_name FROM master_items")
  master_items = cursor.fetchall()

  # Fetch all items and vendors for Rate List dropdowns (always needed)
  cursor.execute("SELECT id, item_name FROM master_items ORDER BY item_name")
  all_items_for_rate = cursor.fetchall()
  cursor.execute("SELECT id, vendor_name FROM vendors ORDER BY vendor_name")
  all_vendors_for_rate = cursor.fetchall()

  # Fetch existing rate_list rows for display/edit
  rate_list_rows = []
  if selected_type == "Rate List":
    cursor.execute("""
        SELECT rl.id, rl.item_id, m.item_name, rl.vendor_id, v.vendor_name, rl.rate
        FROM rate_list rl
        JOIN master_items m ON rl.item_id = m.id
        JOIN vendors v ON rl.vendor_id = v.id
        ORDER BY m.item_name, v.vendor_name
    """)
    rate_list_rows = cursor.fetchall()

  import json as _json
  rate_list_js_data = _json.dumps([
      {"id": r[0], "item_id": r[1], "item_name": r[2],
       "vendor_id": r[3], "vendor_name": r[4], "rate": r[5]}
      for r in rate_list_rows
  ])
  items_js_data = _json.dumps([{"id": i[0], "name": i[1]} for i in all_items_for_rate])
  vendors_js_data = _json.dumps([{"id": v[0], "name": v[1]} for v in all_vendors_for_rate])

  items_html, vendors_html, units_html, box_packs_html = "", "", "", ""
  units_js = "[]"
  item_unit_js = "[]"

  form_fields = ""
  display_table = ""

  if selected_type == "Items":
    cursor.execute("""
        SELECT
            id,
            item_name,
            item_category,
            COALESCE(min_quantity, 0),
            COALESCE(min_casting_quantity, 0),
            COALESCE(min_machined_quantity, 0),
            COALESCE(min_assembly_quantity, 0),
            COALESCE(min_purchase_quantity, 0)
        FROM master_items
        ORDER BY item_name
    """)

    items = cursor.fetchall()
    is_admin_user = session.get("role") == "admin"

    # ------------------------------------------------------------
    # Build existing items table rows
    # ------------------------------------------------------------

    items_html = ""

    for i in items:
        item_id = i[0]
        item_name = i[1] or ""
        item_category = i[2] or "Brass Item"

        min_quantity = i[3] or 0
        min_casting_quantity = i[4] or 0
        min_machined_quantity = i[5] or 0
        min_assembly_quantity = i[6] or 0
        min_purchase_quantity = i[7] or 0

        # Safely create JavaScript string values.
        import json as _json

        js_item_name = _json.dumps(item_name)
        js_item_category = _json.dumps(item_category)

        edit_button = ""

        if is_admin_user:
            edit_button = f"""
                <button
                    type="button"
                    onclick='startEditItem(
                        {item_id},
                        {js_item_name},
                        {js_item_category},
                        {min_quantity},
                        {min_casting_quantity},
                        {min_machined_quantity},
                        {min_assembly_quantity},
                        {min_purchase_quantity}
                    )'
                    id="edit-btn-{item_id}"
                    style="
                        background:#ffc107;
                        color:#000;
                        border:none;
                        border-radius:4px;
                        padding:5px 10px;
                        font-size:12px;
                        margin-left:4px;
                        cursor:pointer;
                    ">
                    Edit
                </button>
            """

        items_html += f"""
            <tr id="item-row-{item_id}">
                <td id="item-name-{item_id}">
                    {item_name}
                </td>

                <td id="item-cat-{item_id}">
                    {item_category}
                </td>

                <td id="item-minqty-{item_id}">
                    {int(min_quantity) if min_quantity else 0}
                </td>

                <td>
                    {admin_delete_cell(
                        f'/delete-item/{item_id}?type=Items',
                        'Delete this item?'
                    )}

                    {edit_button}
                </td>
            </tr>
        """

    # ------------------------------------------------------------
    # Item Edit Modal
    # ------------------------------------------------------------

    item_edit_modal_html = ""

    if is_admin_user:
        item_edit_modal_html = """
        <!-- =====================================================
             ITEM EDIT MODAL
             ===================================================== -->

        <div id="item-edit-modal"
             style="
                display:none;
                position:fixed;
                inset:0;
                width:100%;
                height:100%;
                background:rgba(0,0,0,0.55);
                z-index:10000;
                align-items:center;
                justify-content:center;
                padding:20px;
                box-sizing:border-box;
             ">

            <div style="
                background:#fff;
                width:500px;
                max-width:95%;
                max-height:90vh;
                overflow-y:auto;
                padding:25px;
                border-radius:10px;
                box-shadow:0 6px 25px rgba(0,0,0,0.3);
                box-sizing:border-box;
             ">

                <!-- Modal Header -->

                <div style="
                    display:flex;
                    justify-content:space-between;
                    align-items:center;
                    margin-bottom:20px;
                    padding-bottom:12px;
                    border-bottom:1px solid #ddd;
                ">

                    <h3 style="
                        margin:0;
                        color:#333;
                    ">
                        Edit Item
                    </h3>

                    <button
                        type="button"
                        onclick="closeItemEdit()"
                        style="
                            background:#dc3545;
                            color:white;
                            border:none;
                            border-radius:50%;
                            width:30px;
                            height:30px;
                            font-size:18px;
                            line-height:30px;
                            padding:0;
                            cursor:pointer;
                        ">
                        &times;
                    </button>

                </div>


                <!-- =================================================
                     EDIT FORM
                     ================================================= -->

                <form method="POST" id="item-edit-form">

                    <input
                        type="hidden"
                        name="target_type"
                        value="edit_item">

                    <input
                        type="hidden"
                        name="item_edit_id"
                        id="item_edit_id">


                    <!-- Item Name -->

                    <label>
                        <b>Item Name:</b>
                    </label>

                    <input
                        type="text"
                        name="item_name"
                        id="item_edit_name"
                        required
                        autocomplete="off"
                        style="
                            width:100%;
                            box-sizing:border-box;
                            margin-bottom:15px;
                        ">


                    <!-- Item Category -->

                    <label>
                        <b>Item Category:</b>
                    </label>

                    <select
                        name="item_category"
                        id="item_edit_category"
                        required
                        onchange="updateEditItemMinimumFields()"
                        style="
                            width:100%;
                            box-sizing:border-box;
                            margin-bottom:15px;
                        ">

                        <option value="Brass Item">
                            Brass Item
                        </option>

                        <option value="Miscellaneous Item">
                            Miscellaneous Item
                        </option>

                    </select>


                    <!-- Total Minimum Quantity -->

                    <label>
                        <b>
                            Min Quantity Required (Total Pcs):
                        </b>
                    </label>

                    <input
                        type="number"
                        step="1"
                        min="0"
                        name="min_quantity"
                        id="item_edit_min_quantity"
                        placeholder="e.g. 5000 (0 = no total alert)"
                        style="
                            width:100%;
                            box-sizing:border-box;
                            margin-bottom:15px;
                        ">


                    <!-- =================================================
                         BRASS ITEM FIELDS
                         ================================================= -->

                    <div id="edit-brass-min-fields">

                        <div style="
                            background:#fff8e1;
                            border:1px solid #ffe082;
                            border-radius:6px;
                            padding:12px;
                            margin-bottom:15px;
                        ">

                            <div style="
                                font-weight:bold;
                                margin-bottom:12px;
                                color:#795548;
                            ">
                                Brass Item Minimum Quantities
                            </div>


                            <label>
                                <b>
                                    Min Casting Quantity Required (Pcs):
                                </b>
                            </label>

                            <input
                                type="number"
                                step="1"
                                min="0"
                                name="min_casting_quantity"
                                id="item_edit_min_casting_quantity"
                                placeholder="e.g. 2000 (0 = no alert)"
                                style="
                                    width:100%;
                                    box-sizing:border-box;
                                    margin-bottom:12px;
                                ">


                            <label>
                                <b>
                                    Min Machined Quantity Required (Pcs):
                                </b>
                            </label>

                            <input
                                type="number"
                                step="1"
                                min="0"
                                name="min_machined_quantity"
                                id="item_edit_min_machined_quantity"
                                placeholder="e.g. 1500 (0 = no alert)"
                                style="
                                    width:100%;
                                    box-sizing:border-box;
                                    margin-bottom:12px;
                                ">


                            <label>
                                <b>
                                    Min Assembly Quantity Required (Pcs):
                                </b>
                            </label>

                            <input
                                type="number"
                                step="1"
                                min="0"
                                name="min_assembly_quantity"
                                id="item_edit_min_assembly_quantity"
                                placeholder="e.g. 2000 (0 = no alert)"
                                style="
                                    width:100%;
                                    box-sizing:border-box;
                                ">

                        </div>

                    </div>


                    <!-- =================================================
                         MISCELLANEOUS ITEM FIELD
                         ================================================= -->

                    <div
                        id="edit-misc-min-fields"
                        style="display:none;">

                        <div style="
                            background:#e8f5e9;
                            border:1px solid #a5d6a7;
                            border-radius:6px;
                            padding:12px;
                            margin-bottom:15px;
                        ">

                            <div style="
                                font-weight:bold;
                                margin-bottom:12px;
                                color:#2e7d32;
                            ">
                                Miscellaneous Item Minimum Quantity
                            </div>

                            <label>
                                <b>
                                    Min Purchased Quantity Required (Pcs):
                                </b>
                            </label>

                            <input
                                type="number"
                                step="1"
                                min="0"
                                name="min_purchase_quantity"
                                id="item_edit_min_purchase_quantity"
                                placeholder="e.g. 500 (0 = no alert)"
                                style="
                                    width:100%;
                                    box-sizing:border-box;
                                ">

                        </div>

                    </div>


                    <!-- =================================================
                         BUTTONS
                         ================================================= -->

                    <div style="
                        display:flex;
                        justify-content:flex-end;
                        gap:10px;
                        margin-top:20px;
                        padding-top:15px;
                        border-top:1px solid #ddd;
                    ">

                        <button
                            type="button"
                            onclick="closeItemEdit()"
                            style="
                                background:#6c757d;
                                color:white;
                                border:none;
                                padding:8px 16px;
                                border-radius:5px;
                                cursor:pointer;
                            ">
                            Cancel
                        </button>

                        <button
                            type="button"
                            onclick="submitItemEdit()"
                            style="
                                background:#28a745;
                                color:white;
                                border:none;
                                padding:8px 18px;
                                border-radius:5px;
                                cursor:pointer;
                                font-weight:bold;
                            ">
                            &#10003; Save Changes
                        </button>

                    </div>

                </form>

            </div>
        </div>


        <!-- =====================================================
             ITEM EDIT JAVASCRIPT
             ===================================================== -->

        <script>

        function startEditItem(
            id,
            itemName,
            itemCategory,
            minQuantity,
            minCastingQuantity,
            minMachinedQuantity,
            minAssemblyQuantity,
            minPurchaseQuantity
        ) {

            document.getElementById("item_edit_id").value = id;

            document.getElementById("item_edit_name").value =
                itemName || "";

            document.getElementById("item_edit_category").value =
                itemCategory || "Brass Item";

            document.getElementById("item_edit_min_quantity").value =
                minQuantity || 0;

            document.getElementById("item_edit_min_casting_quantity").value =
                minCastingQuantity || 0;

            document.getElementById("item_edit_min_machined_quantity").value =
                minMachinedQuantity || 0;

            document.getElementById("item_edit_min_assembly_quantity").value =
                minAssemblyQuantity || 0;

            document.getElementById("item_edit_min_purchase_quantity").value =
                minPurchaseQuantity || 0;

            updateEditItemMinimumFields();

            document.getElementById("item-edit-modal").style.display =
                "flex";
        }


        function closeItemEdit() {

            var modal =
                document.getElementById("item-edit-modal");

            if (modal) {
                modal.style.display = "none";
            }
        }


        function updateEditItemMinimumFields() {

            var category =
                document.getElementById(
                    "item_edit_category"
                ).value;

            var brassFields =
                document.getElementById(
                    "edit-brass-min-fields"
                );

            var miscFields =
                document.getElementById(
                    "edit-misc-min-fields"
                );

            if (category === "Brass Item") {

                brassFields.style.display = "block";
                miscFields.style.display = "none";

            } else {

                brassFields.style.display = "none";
                miscFields.style.display = "block";
            }
        }


        function submitItemEdit() {

            var itemId =
                document.getElementById(
                    "item_edit_id"
                ).value;

            var itemName =
                document.getElementById(
                    "item_edit_name"
                ).value.trim();

            if (!itemId) {

                alert("Item ID is missing.");

                return;
            }

            if (!itemName) {

                alert("Please enter an Item Name.");

                document.getElementById(
                    "item_edit_name"
                ).focus();

                return;
            }

            document.getElementById(
                "item-edit-form"
            ).submit();
        }


        /* Close when clicking outside the modal */

        document.getElementById(
            "item-edit-modal"
        ).addEventListener(
            "click",
            function(event) {

                if (
                    event.target ===
                    document.getElementById(
                        "item-edit-modal"
                    )
                ) {

                    closeItemEdit();
                }
            }
        );


        /* ESC closes the modal */

        document.addEventListener(
            "keydown",
            function(event) {

                if (event.key === "Escape") {

                    var modal =
                        document.getElementById(
                            "item-edit-modal"
                        );

                    if (
                        modal &&
                        modal.style.display === "flex"
                    ) {

                        closeItemEdit();
                    }
                }
            }
        );


        updateEditItemMinimumFields();

        </script>
        """

    # ------------------------------------------------------------
    # IMPORTANT:
    # Build the Items table here.
    # ------------------------------------------------------------

    display_table = f"""
        <h3>Existing Items</h3>

        <table id="items-table">
            <thead>
                <tr>
                    <th>Item Name</th>
                    <th>Category</th>
                    <th>Min Quantity</th>
                    <th>Action</th>
                </tr>
            </thead>

            <tbody>
                {items_html}
            </tbody>
        </table>

        {item_edit_modal_html}
    """

  elif selected_type == "Vendors":
    cursor.execute("SELECT id, vendor_name, contact_info FROM vendors")
    vendors_raw = cursor.fetchall()
    vendors_list = []
    for v in vendors_raw:
      v_id, v_name, v_contact = v[0], v[1], v[2]
      cursor.execute(
          """SELECT m.item_name FROM master_items m 
             JOIN vendor_items vi ON m.id = vi.item_id 
             WHERE vi.vendor_id = ?""",
          (v_id,),
      )
      deals = ", ".join([row[0] for row in cursor.fetchall()])
      cursor.execute(
          "SELECT register_type FROM vendor_registers WHERE vendor_id = ?",
          (v_id,),
      )
      regs = ", ".join([row[0] for row in cursor.fetchall()])
      vendors_list.append(
          (v_id, v_name, v_contact, deals or "No items", regs or "None")
      )
    vendors_html = "".join([
        f"<tr><td>{v[1]}</td><td>{v[2]}</td><td>{v[3]}</td><td>{v[4]}</td><td>"
        f"<a href='/edit-vendor/{v[0]}'><button style='background:#ffc107;"
        " color:black; margin-right:5px;'>Edit</button></a>"
        f"{admin_delete_cell(f'/delete-vendor/{v[0]}?type=Vendors', 'Delete this vendor?')}"
        f"</td></tr>"
        for v in vendors_list
    ])

  elif selected_type == "unit of measurement":
    # Fetch every unit and the items currently assigned to it.
    cursor.execute("""
        SELECT
            u.id,
            u.unit_name,
            GROUP_CONCAT(m.item_name, ', ') AS item_names
        FROM units u
        LEFT JOIN item_units iu ON u.id = iu.unit_id
        LEFT JOIN master_items m ON iu.item_id = m.id
        GROUP BY u.id, u.unit_name
        ORDER BY u.unit_name
    """)
    units = cursor.fetchall()

    # Fetch all items and their current unit assignments.
    cursor.execute("""
        SELECT
            m.id,
            m.item_name,
            COALESCE(iu.unit_id, '') AS unit_id
        FROM master_items m
        LEFT JOIN item_units iu ON m.id = iu.item_id
        ORDER BY m.item_name
    """)
    item_unit_rows = cursor.fetchall()

    import json as _json2

    units_js = _json2.dumps([
        {"id": u[0], "name": u[1]}
        for u in units
    ])

    item_unit_js = _json2.dumps([
        {
            "id": r[0],
            "name": r[1],
            "unit_id": r[2]
        }
        for r in item_unit_rows
    ])

    is_admin_user = session.get("role") == "admin"

    units_html = "".join([
        f"""
        <tr>
            <td><b>{u[1]}</b></td>
            <td>{u[2] or '<span style="color:#888;">No items assigned</span>'}</td>
            <td>
                {admin_delete_cell(
                    f'/delete-unit/{u[0]}?type=unit of measurement',
                    'Delete this unit?'
                )}
                {
                    f'''
                    <button
                        type="button"
                        onclick="startEditUnit({u[0]})"
                        style="background:#ffc107;color:#000;
                               padding:2px 10px;font-size:12px;margin-left:4px;">
                        Edit
                    </button>

                    <button
                        type="button"
                        onclick="startEditUnitItems({u[0]})"
                        style="background:#17a2b8;color:#fff;
                               padding:2px 10px;font-size:12px;margin-left:4px;">
                        Edit Items
                    </button>
                    '''
                    if is_admin_user else ""
                }
            </td>
        </tr>
        """
        for u in units
    ])

  elif selected_type == "Box Pack Items":
    cursor.execute("SELECT id, box_pack_name FROM box_pack_items")
    box_packs_raw = cursor.fetchall()
    box_packs_list = []
    for bp_id, bp_name in box_packs_raw:
      cursor.execute(
          """SELECT m.item_name FROM master_items m
             JOIN box_pack_components bc ON m.id = bc.item_id
             WHERE bc.box_pack_id = ?""",
          (bp_id,),
      )
      comps = ", ".join([row[0] for row in cursor.fetchall()])
      box_packs_list.append((bp_id, bp_name, comps or "No items selected"))
    def bp_row(bp):
        import json as _json
        edit_btn = ""
        if session.get("role") == "admin":
            edit_btn = f"""
                <button type="button"
                    onclick="startEditBoxPack({bp[0]}, {_json.dumps(bp[1])}, {_json.dumps(bp[2])})"
                    style="background:#ffc107;color:#000;border:none;border-radius:4px;
                        padding:5px 10px;font-size:12px;margin-left:4px;cursor:pointer;">
                    Edit
                </button>"""
        return (f"<tr><td>{bp[1]}</td><td>{bp[2]}</td><td>"
                f"{admin_delete_cell(f'/delete-box-pack/{bp[0]}?type=Box Pack Items', 'Delete this box pack?')}"
                f"{edit_btn}</td></tr>")

    box_packs_html = "".join([bp_row(bp) for bp in box_packs_list])

  conn.close()

  checkboxes = "".join([
      f'<label><input type="checkbox" name="item_ids" value="{mi[0]}">'
      f" {mi[1]}</label><br>"
      for mi in master_items
  ])
  register_checkboxes = """
        <label><input type="checkbox" name="register_types" value="Brass In"> Brass In (Scrap)</label><br>
        <label><input type="checkbox" name="register_types" value="Brass Out"> Brass Out</label><br>
        <label><input type="checkbox" name="register_types" value="Casting In"> Casting In</label><br>
        <label><input type="checkbox" name="register_types" value="Machining"> Machining</label><br>
        <label><input type="checkbox" name="register_types" value="Grinding"> Grinding</label><br>
        <label><input type="checkbox" name="register_types" value="Plating"> Plating</label><br>
        <label><input type="checkbox" name="register_types" value="Miscellaneous"> Miscellaneous</label><br>
        <label><input type="checkbox" name="register_types" value="Rejection"> Rejection</label>
    """
    

  if selected_type == "Items":
    form_fields = """   
        <input type="hidden" name="target_type" value="Items">

        <label><b>Item Name:</b></label>
        <input type="text" name="item_name" required>

        <label><b>Item Category:</b></label>
        <select name="item_category" id="item_category_select" required
                onchange="updateItemMinimumFields()">
            <option value="Brass Item">Brass Item</option>
            <option value="Miscellaneous Item">Miscellaneous Item</option>
        </select>

        <label><b>Min Quantity Required (Total Pcs):</b></label>
        <input type="number" step="1" min="0"
               name="min_quantity"
               placeholder="e.g. 5000 (0 = no total alert)">

        <div id="brass-min-fields">

            <label><b>Min Casting Quantity Required (Pcs):</b></label>
            <input type="number" step="1" min="0"
                   name="min_casting_quantity"
                   placeholder="e.g. 2000 (0 = no alert)">

            <label><b>Min Machined Quantity Required (Pcs):</b></label>
            <input type="number" step="1" min="0"
                   name="min_machined_quantity"
                   placeholder="e.g. 1500 (0 = no alert)">

            <label><b>Min Assembly Quantity Required (Pcs):</b></label>
            <input type="number" step="1" min="0"
                   name="min_assembly_quantity"
                   placeholder="e.g. 2000 (0 = no alert)">

        </div>

        <div id="misc-min-fields" style="display:none;">

            <label><b>Min Purchased Quantity Required (Pcs):</b></label>
            <input type="number" step="1" min="0"
                   name="min_purchase_quantity"
                   placeholder="e.g. 500 (0 = no alert)">

        </div>

        <button type="submit">Save Item</button>

        <script>
        function updateItemMinimumFields() {
            var category =
                document.getElementById('item_category_select').value;

            var brassFields =
                document.getElementById('brass-min-fields');

            var miscFields =
                document.getElementById('misc-min-fields');

            if (category === 'Brass Item') {
                brassFields.style.display = 'block';
                miscFields.style.display = 'none';
            } else {
                brassFields.style.display = 'none';
                miscFields.style.display = 'block';
            }
        }

        updateItemMinimumFields();
        </script>
    """
  elif selected_type == "Vendors":
    form_fields = f"""
            <input type="hidden" name="target_type" value="Vendors">
            <label><b>Vendor Name:</b></label>
            <input type="text" name="vendor_name" required>
            <label><b>Contact Info:</b></label>
            <input type="text" name="contact_info">
            <label><b>Select Registers:</b></label>
            <div class="checkbox-group">{register_checkboxes}</div>
            <label><b>Select Master Items:</b></label>
            <div class="checkbox-group">{checkboxes}</div>
            <button type="submit">Save Vendor</button>
        """
    display_table = f"""
            <h3>Existing Vendors</h3>
            <table><tr><th>Vendor</th><th>Contact</th><th>Items</th><th>Registers</th><th>Action</th></tr>{vendors_html}</table>
        """

  elif selected_type == "unit of measurement":
    form_fields = """
            <input type="hidden" name="target_type" value="unit of measurement">

            <label><b>Unit Name:</b></label>
            <input type="text" name="unit_name" required>

            <button type="submit">Save Unit</button>
        """

    display_table = f"""
        <h3>Units of Measurement</h3>

        <div class="info-box">
            Each unit shows the items currently using it.
            For example: <b>Bottle → Dara</b> or
            <b>Packet → Taper Gutka</b>.
        </div>

        <table id="units-table">
            <tr>
                <th>Unit Name</th>
                <th>Items Using This Unit</th>
                <th>Action</th>
            </tr>
            {units_html}
        </table>


        <!-- Edit Unit Modal -->
        <div id="unit-edit-modal"
             style="display:none;position:fixed;top:0;left:0;width:100%;
                    height:100%;background:rgba(0,0,0,0.5);z-index:10000;
                    align-items:center;justify-content:center;">

            <div style="background:white;padding:25px;border-radius:8px;
                        min-width:300px;
                        box-shadow:0 4px 20px rgba(0,0,0,0.3);">

                <h3 style="margin-top:0;">Edit Unit</h3>

                <form method="POST" id="unit-edit-form">

                    <input type="hidden"
                           name="target_type"
                           value="edit_unit">

                    <input type="hidden"
                           name="unit_id"
                           id="edit_unit_id">

                    <label><b>Unit Name:</b></label>

                    <input type="text"
                           name="unit_name"
                           id="edit_unit_name"
                           required
                           style="margin-bottom:15px;">

                    <div style="display:flex;gap:10px;">

                        <button type="button"
                                onclick="submitUnitEdit()"
                                style="background:#28a745;">
                            &#10003; Done
                        </button>

                        <button type="button"
                                onclick="closeUnitEdit()"
                                style="background:#6c757d;">
                            Cancel
                        </button>

                    </div>

                </form>
            </div>
        </div>


        <!-- Edit Items Assigned To Unit Modal -->
        <div id="unit-items-modal"
             style="display:none;position:fixed;top:0;left:0;width:100%;
                    height:100%;background:rgba(0,0,0,0.5);z-index:10000;
                    align-items:center;justify-content:center;">

            <div style="background:white;padding:25px;border-radius:8px;
                        width:420px;max-width:90%;
                        box-shadow:0 4px 20px rgba(0,0,0,0.3);">

                <h3 style="margin-top:0;">
                    Assign Items
                </h3>

                <p>
                    <b>Unit:</b>
                    <span id="edit_items_unit_name"></span>
                </p>

                <form method="POST" id="unit-items-form">

                    <input type="hidden"
                           name="target_type"
                           value="assign_unit_items">

                    <input type="hidden"
                           name="unit_id"
                           id="edit_items_unit_id">

                    <div id="unit-items-checkboxes"
                         style="max-height:300px;
                                overflow-y:auto;
                                border:1px solid #ddd;
                                padding:10px;
                                border-radius:5px;">
                    </div>

                    <div style="display:flex;gap:10px;margin-top:15px;">

                        <button type="button"
                                onclick="submitUnitItems()"
                                style="background:#28a745;">
                            &#10003; Done
                        </button>

                        <button type="button"
                                onclick="closeUnitItems()"
                                style="background:#6c757d;">
                            Cancel
                        </button>

                    </div>

                </form>
            </div>
        </div>


        <script>
        var ALL_UNITS = {units_js};
        var ALL_ITEM_UNITS = {item_unit_js};


        function startEditUnit(id) {{
            var unit = ALL_UNITS.find(function(u) {{
                return String(u.id) === String(id);
            }});

            if (!unit) return;

            document.getElementById('edit_unit_id').value = unit.id;
            document.getElementById('edit_unit_name').value = unit.name;

            document.getElementById('unit-edit-modal').style.display = 'flex';
        }}


        function closeUnitEdit() {{
            document.getElementById('unit-edit-modal').style.display = 'none';
        }}


        function submitUnitEdit() {{
            document.getElementById('unit-edit-form').submit();
        }}


        function startEditUnitItems(unitId) {{
            var unit = ALL_UNITS.find(function(u) {{
                return String(u.id) === String(unitId);
            }});

            if (!unit) return;

            document.getElementById('edit_items_unit_id').value = unit.id;
            document.getElementById('edit_items_unit_name').textContent = unit.name;

            var container =
                document.getElementById('unit-items-checkboxes');

            container.innerHTML = '';

            if (ALL_ITEM_UNITS.length === 0) {{
                container.innerHTML =
                    '<span style="color:#888;">No items available.</span>';
            }} else {{

                ALL_ITEM_UNITS.forEach(function(item) {{

                    var label = document.createElement('label');

                    label.style.display = 'block';
                    label.style.padding = '5px 0';

                    var checkbox = document.createElement('input');

                    checkbox.type = 'checkbox';
                    checkbox.name = 'item_ids';
                    checkbox.value = item.id;

                    if (String(item.unit_id) === String(unitId)) {{
                        checkbox.checked = true;
                    }}

                    label.appendChild(checkbox);

                    label.appendChild(
                        document.createTextNode(' ' + item.name)
                    );

                    container.appendChild(label);
                }});
            }}

            document.getElementById('unit-items-modal').style.display = 'flex';
        }}


        function closeUnitItems() {{
            document.getElementById('unit-items-modal').style.display = 'none';
        }}


        function submitUnitItems() {{
            document.getElementById('unit-items-form').submit();
        }}
        </script>
    """
  elif selected_type == "Box Pack Items":
    checkboxes_with_data = "".join([
        f'<label><input type="checkbox" name="item_ids" value="{mi[0]}" data-name="{mi[1]}"> {mi[1]}</label><br>'
        for mi in master_items
    ])
    form_fields = f"""
            <input type="hidden" name="target_type" value="Box Pack Items">
            <label><b>Box Pack Name:</b></label>
            <input type="text" name="box_pack_name" placeholder="e.g. Push Cock BP" required>
            <label><b>Select Items (all items this box pack is assembled from):</b></label>
            <div class="checkbox-group">{checkboxes}</div>
            <button type="submit">Save Box Pack Item</button>
        """
    display_table = f"""
            <h3>Existing Box Pack Items</h3>
            <table><tr><th>Box Pack Name</th><th>Items Used</th><th>Action</th></tr>{box_packs_html}</table>
            <div id="edit-bp-modal" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;
                 background:rgba(0,0,0,0.5);z-index:9999;">
              <div style="background:#fff;margin:10% auto;padding:24px;max-width:420px;border-radius:8px;">
                <h3>Edit Box Pack</h3>
                <form method="POST" id="edit-bp-form">
                  <input type="hidden" name="target_type" value="Box Pack Items">
                  <input type="hidden" name="edit_bp_id" id="edit-bp-id">
                  <label><b>Box Pack Name:</b></label><br>
                  <input type="text" name="box_pack_name" id="edit-bp-name" style="width:100%;padding:6px;margin-bottom:10px;"><br>
                  <label><b>Select Items:</b></label>
                  <div class="checkbox-group" id="edit-bp-items">{checkboxes_with_data}</div><br>
                  <button type="submit" style="background:#28a745;color:#fff;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;">Save</button>
                  <button type="button" onclick="document.getElementById('edit-bp-modal').style.display='none'"
                          style="margin-left:8px;padding:8px 16px;cursor:pointer;">Cancel</button>
                </form>
              </div>
            </div>
            <script>
            function startEditBoxPack(id, name, comps) {{
                document.getElementById('edit-bp-id').value = id;
                document.getElementById('edit-bp-name').value = name;
                document.getElementById('edit-bp-modal').style.display = 'block';
                var compList = comps.split(', ');
                document.querySelectorAll('#edit-bp-items input[type=checkbox]').forEach(function(cb) {{
                    cb.checked = compList.indexOf(cb.dataset.name) !== -1;
                }});
                document.getElementById('edit-bp-form').action = '/edit-box-pack/' + id;
            }}
            </script>
        """

  elif selected_type == "Rate List":
    is_admin_user = session.get("role") == "admin"
    readonly_note = "" if is_admin_user else '<div class="info-box">&#128274; Only admins can edit the rate list. You are viewing in read-only mode.</div>'
    form_fields = f"""
            <input type="hidden" name="target_type" value="Rate List">
            <input type="hidden" name="rate_data" id="rate_data_input" value="[]">
            {readonly_note}
            <h3 style="margin-top:0">Rate List</h3>
            <div style="overflow-x:auto;">
            <table id="rate-list-table" style="min-width:600px;">
                <thead>
                    <tr>
                        <th style="width:35%">Item</th>
                        <th style="width:35%">Vendor</th>
                        <th style="width:15%">Rate (&#8377;)</th>
                        {'<th style="width:15%">Action</th>' if is_admin_user else ''}
                    </tr>
                </thead>
                <tbody id="rate-list-tbody">
                </tbody>
            </table>
            </div>
            {'<div style="margin-top:12px;display:flex;gap:10px;flex-wrap:wrap;"><button type="button" id="add-rate-row-btn" style="background:#28a745;">+ Add Row</button><button type="button" id="save-rate-btn" style="background:#007bff;">&#128190; Save Rate List</button></div>' if is_admin_user else ''}

            <script>
            (function() {{
                var ALL_ITEMS   = {items_js_data};
                var ALL_VENDORS = {vendors_js_data};
                var DB_ROWS     = {rate_list_js_data};
                var IS_ADMIN    = {'true' if is_admin_user else 'false'};

                var rowState = {{}};
                var rowCounter = 0;

                function itemName(id) {{
                    var f = ALL_ITEMS.find(function(i){{ return i.id == id; }});
                    return f ? f.name : id;
                }}
                function vendorName(id) {{
                    var f = ALL_VENDORS.find(function(v){{ return v.id == id; }});
                    return f ? f.name : id;
                }}

                function buildItemSelect(selectedId) {{
                    var opts = '<option value="">-- Select Item --</option>';
                    ALL_ITEMS.forEach(function(i) {{
                        opts += '<option value="' + i.id + '"' + (i.id == selectedId ? ' selected' : '') + '>' + i.name + '</option>';
                    }});
                    return '<select style="width:100%;margin:0;" class="rl-item-sel">' + opts + '</select>';
                }}
                function buildVendorSelect(selectedId) {{
                    var opts = '<option value="">-- Select Vendor --</option>';
                    ALL_VENDORS.forEach(function(v) {{
                        opts += '<option value="' + v.id + '"' + (v.id == selectedId ? ' selected' : '') + '>' + v.name + '</option>';
                    }});
                    return '<select style="width:100%;margin:0;" class="rl-vendor-sel">' + opts + '</select>';
                }}

                function renderRows() {{
                    var tbody = document.getElementById('rate-list-tbody');
                    tbody.innerHTML = '';

                    Object.keys(rowState).forEach(function(key) {{
                        var s = rowState[key];

                        if (s.action === 'delete') return;

                        var tr = document.createElement('tr');
                        tr.setAttribute('data-key', key);

                        if (s.action === 'insert') {{
                            tr.style.background = '#f0fff0';
                        }} else if (s.action === 'update') {{
                            tr.style.background = '#fff8e1';
                        }}

                        if (IS_ADMIN && s._editing) {{

                            tr.innerHTML =
                                '<td>' + buildItemSelect(s.item_id) + '</td>' +
                                '<td>' + buildVendorSelect(s.vendor_id) + '</td>' +
                                '<td><input type="number" class="rl-rate-inp" step="0.01" min="0" style="width:100%;margin:0;" value="' +
                                    (s.rate !== null ? s.rate : '') + '"></td>' +
                                '<td>' +
                                    '<button type="button" class="rl-done-btn" style="padding:6px 10px;font-size:13px;background:#28a745;margin-right:4px;">✓ Done</button>' +
                                    '<button type="button" class="danger rl-del-btn" style="padding:6px 10px;font-size:13px;">🗑 Delete</button>' +
                                '</td>';

                            tr.querySelector('.rl-item-sel').addEventListener('change', function() {{
                                rowState[key].item_id = parseInt(this.value) || null;
                                if (rowState[key].action !== 'insert') {{
                                    rowState[key].action = 'update';
                                }}
                            }});

                            tr.querySelector('.rl-vendor-sel').addEventListener('change', function() {{
                                rowState[key].vendor_id = parseInt(this.value) || null;
                                if (rowState[key].action !== 'insert') {{
                                    rowState[key].action = 'update';
                                }}
                            }});

                            tr.querySelector('.rl-rate-inp').addEventListener('input', function() {{
                                rowState[key].rate =
                                    this.value !== '' ? parseFloat(this.value) : null;
                                if (rowState[key].action !== 'insert') {{
                                    rowState[key].action = 'update';
                                }}
                            }});

                            tr.querySelector('.rl-done-btn').addEventListener('click', function() {{
                            var row = rowState[key];

                            if (row.action === 'insert') {{
                                if (!row.item_id) {{
                                    alert('Please select an Item.');
                                    return;
                                }}
                                if (!row.vendor_id) {{
                                    alert('Please select a Vendor.');
                                    return;
                                }}
                                if (row.rate === null || row.rate === '' || isNaN(row.rate)) {{
                                    alert('Please enter a Rate.');
                                    return;
                                }}
                            }}

                            row._editing = false;
                            if (row.id === null) {{
                                row.action = 'insert';
                            }}

                            renderRows();
                        }});

                            tr.querySelector('.rl-del-btn').addEventListener('click', function() {{
                                if (rowState[key].action === 'insert') {{
                                    delete rowState[key];
                                }} else {{
                                    rowState[key].action = 'delete';
                                    rowState[key]._editing = false;
                                }}
                                renderRows();
                            }});

                        }} else {{

                            tr.innerHTML =
                                '<td>' + itemName(s.item_id) + '</td>' +
                                '<td>' + vendorName(s.vendor_id) + '</td>' +
                                '<td>₹ ' + (s.rate !== null ? Number(s.rate).toFixed(2) : '-') + '</td>' +
                                (IS_ADMIN
                                    ? '<td>' +
                                        '<button type="button" class="rl-edit-btn" style="padding:6px 10px;font-size:13px;background:#ffc107;color:#000;margin-right:4px;">✎ Edit</button>' +
                                        '<button type="button" class="danger rl-del-btn" style="padding:6px 10px;font-size:13px;">🗑 Delete</button>' +
                                    '</td>'
                                    : '');

                            if (IS_ADMIN) {{
                                tr.querySelector('.rl-edit-btn').addEventListener('click', function() {{
                                    rowState[key]._editing = true;
                                    rowState[key].action = 'update';
                                    renderRows();
                                }});

                                tr.querySelector('.rl-del-btn').addEventListener('click', function() {{
                                    rowState[key].action = 'delete';
                                    renderRows();
                                }});
                            }}
                        }}

                        tbody.appendChild(tr);
                    }});
                }}

                DB_ROWS.forEach(function(r) {{
                    var key = 'db_' + r.id;
                    rowState[key] = {{
                        id: r.id,
                        item_id: r.item_id,
                        vendor_id: r.vendor_id,
                        rate: r.rate,
                        action: null
                    }};
                }});
                renderRows();

                if (IS_ADMIN) {{
                    document.getElementById('add-rate-row-btn').addEventListener('click', function() {{
                        var key = 'new_' + (++rowCounter);
                        rowState[key] = {{
                            id: null,
                            item_id: null,
                            vendor_id: null,
                            rate: null,
                            action: 'insert',
                            _editing: true
                        }};
                        renderRows();
                        var tbody = document.getElementById('rate-list-tbody');
                        if (tbody.lastElementChild) tbody.lastElementChild.scrollIntoView({{behavior:'smooth'}});
                    }});

                    document.getElementById('save-rate-btn').addEventListener('click', function() {{
                        doSave();
                    }});
                }}

                function getDirtyRows() {{
                    return Object.keys(rowState)
                        .map(function(k){{ return rowState[k]; }})
                        .filter(function(r){{ return r.action !== null; }});
                }}

                function buildSummaryLines(dirty) {{
                    return dirty.map(function(r) {{
                        var iName = itemName(r.item_id)   || '(no item)';
                        var vName = vendorName(r.vendor_id) || '(no vendor)';
                        var rVal  = r.rate !== null ? '₹' + Number(r.rate).toFixed(2) : '(no rate)';
                        var tag   = r.action === 'insert' ? '[NEW]'
                                  : r.action === 'update' ? '[EDITED]'
                                  : '[DELETED]';
                        return tag + '  Item: ' + iName + '  |  Vendor: ' + vName + '  |  Rate: ' + rVal;
                    }}).join('\\n');
                }}

                function doSave() {{
                    var dirty = getDirtyRows();
                    if (dirty.length === 0) {{
                        alert('No changes to save.');
                        return;
                    }}
                    var summary = buildSummaryLines(dirty);
                    var confirmed = confirm('Save the following changes?\\n\\n' + summary + '\\n\\nClick OK to save, Cancel to go back.');
                    if (!confirmed) return;

                    var payload = dirty.map(function(r) {{
                        return {{
                            action: r.action,
                            id: r.id,
                            item_id: r.item_id,
                            vendor_id: r.vendor_id,
                            rate: r.rate
                        }};
                    }});
                    document.getElementById('rate_data_input').value = JSON.stringify(payload);
                    Object.keys(rowState).forEach(function(k){{ rowState[k].action = null; }});
                    document.getElementById('rate-list-form').submit();
                }}

                function hasDirty() {{ return getDirtyRows().length > 0; }}

                window.addEventListener('beforeunload', function(e) {{
                    if (IS_ADMIN && hasDirty()) {{
                        e.preventDefault();
                        e.returnValue = '';
                    }}
                }});

                document.querySelectorAll('.nav a').forEach(function(link) {{
                    link.addEventListener('click', function(e) {{
                        if (!IS_ADMIN || !hasDirty()) return;
                        e.preventDefault();
                        var dirty = getDirtyRows();
                        var summary = buildSummaryLines(dirty);
                        var msg = 'You have unsaved changes:\\n\\n' + summary + '\\n\\nDo you want to save before leaving?\\n(OK = Save then go, Cancel = Leave without saving)';
                        var wantSave = confirm(msg);
                        if (wantSave) {{
                            var payload = dirty.map(function(r) {{
                                return {{action:r.action, id:r.id, item_id:r.item_id, vendor_id:r.vendor_id, rate:r.rate}};
                            }});
                            document.getElementById('rate_data_input').value = JSON.stringify(payload);
                            Object.keys(rowState).forEach(function(k){{ rowState[k].action = null; }});
                            document.getElementById('rate-list-form').submit();
                        }} else {{
                            Object.keys(rowState).forEach(function(k){{ rowState[k].action = null; }});
                            window.location.href = e.currentTarget.href;
                        }}
                    }});
                }});

                var catDropdown = document.querySelector('select[name="type"]');
                if (catDropdown) {{
                    catDropdown.addEventListener('change', function(e) {{
                        if (!IS_ADMIN || !hasDirty()) return;
                        var dirty = getDirtyRows();
                        var summary = buildSummaryLines(dirty);
                        var msg = 'You have unsaved changes:\\n\\n' + summary + '\\n\\nDo you want to save before switching tabs?\\n(OK = Save, Cancel = Switch without saving)';
                        var wantSave = confirm(msg);
                        if (wantSave) {{
                            var payload = dirty.map(function(r) {{
                                return {{action:r.action, id:r.id, item_id:r.item_id, vendor_id:r.vendor_id, rate:r.rate}};
                            }});
                            document.getElementById('rate_data_input').value = JSON.stringify(payload);
                            Object.keys(rowState).forEach(function(k){{ rowState[k].action = null; }});
                            document.getElementById('rate-list-form').submit();
                        }}
                        catDropdown.value = 'Rate List';
                    }});
                }}
            }})();
            </script>
        """
    display_table = ""  # everything is inline in the form for Rate List

  # For Rate List the outer wrapper must have an id so JS can reference it
  if selected_type == "Rate List":
    form_wrapper = f'<div class="card"><form method="POST" id="rate-list-form">{form_fields}</form></div>'
  else:
    form_wrapper = f'<div class="card"><form method="POST">{form_fields}</form></div>' if selected_type else ''

  page = (
      HTML_TEMPLATE
      + f"""
    <div class="nav">
        Logged in as: <b>{{{{ session['user'] }}}}</b> ({{{{ session['role'] }}}})<br><br>
        {get_nav_links()}
    </div>
    <h2>Add Items & Management</h2>
    <div class="card">
        <form method="GET" action="/add-items">
            <label><b>Select Category:</b></label>
            <select name="type" onchange="this.form.submit()">
                <option value="">-- Choose Option --</option>
                <option value="Items" {"selected" if selected_type == "Items" else ""}>1. Items</option>
                <option value="Vendors" {"selected" if selected_type == "Vendors" else ""}>2. Vendors</option>
                <option value="unit of measurement" {"selected" if selected_type == "unit of measurement" else ""}>3. unit of measurement</option>
                <option value="Box Pack Items" {"selected" if selected_type == "Box Pack Items" else ""}>4. Box Pack Items</option>
                <option value="Rate List" {"selected" if selected_type == "Rate List" else ""}>5. Rate List</option>
            </select>
        </form>
    </div>
    {form_wrapper}
    {display_table}
    """
  )
  return render_template_string(page)

@app.route("/edit-vendor/<int:v_id>", methods=["GET", "POST"])
@login_required
def edit_vendor(v_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()

  if request.method == "POST":
    vendor_name = request.form.get("vendor_name", "").strip()
    contact_info = request.form.get("contact_info", "").strip()
    selected_items = request.form.getlist("item_ids")
    selected_registers = request.form.getlist("register_types")

    if vendor_name:
      try:
        cursor.execute(
            "UPDATE vendors SET vendor_name = ?, contact_info = ? WHERE id = ?",
            (vendor_name, contact_info, v_id),
        )
        cursor.execute("DELETE FROM vendor_items WHERE vendor_id = ?", (v_id,))
        for item_id in selected_items:
          cursor.execute(
              "INSERT INTO vendor_items (vendor_id, item_id) VALUES (?, ?)",
              (v_id, item_id),
          )
        cursor.execute(
            "DELETE FROM vendor_registers WHERE vendor_id = ?", (v_id,)
        )
        for reg_type in selected_registers:
          cursor.execute(
              "INSERT INTO vendor_registers (vendor_id, register_type) VALUES (?, ?)",
              (v_id, reg_type),
          )
        conn.commit()
        log_action(
            session["user"], "UPDATE VENDOR", f"Edited vendor '{vendor_name}'"
        )
      except sqlite3.IntegrityError:
        pass

    conn.close()
    return redirect(url_for("add_items", type="Vendors"))

  cursor.execute(
      "SELECT id, vendor_name, contact_info FROM vendors WHERE id = ?", (v_id,)
  )
  vendor = cursor.fetchone()
  if not vendor:
    conn.close()
    return redirect(url_for("add_items", type="Vendors"))

  cursor.execute(
      "SELECT item_id FROM vendor_items WHERE vendor_id = ?", (v_id,)
  )
  assigned_items = [row[0] for row in cursor.fetchall()]

  cursor.execute(
      "SELECT register_type FROM vendor_registers WHERE vendor_id = ?", (v_id,)
  )
  assigned_registers = [row[0] for row in cursor.fetchall()]

  cursor.execute("SELECT id, item_name FROM master_items")
  master_items = cursor.fetchall()

  conn.close()

  checkboxes = "".join([
      f'<label><input type="checkbox" name="item_ids" value="{mi[0]}"'
      f' {"checked" if mi[0] in assigned_items else ""}> {mi[1]}</label><br>'
      for mi in master_items
  ])

  reg_types_list = [
      "Brass In", "Brass Out", "Casting In", "Machining",
      "Grinding", "Plating", "Miscellaneous", "Rejection",
  ]
  register_checkboxes = "".join([
      f'<label><input type="checkbox" name="register_types" value="{rt}"'
      f' {"checked" if rt in assigned_registers else ""}> {rt}</label><br>'
      for rt in reg_types_list
  ])

  page = (
      HTML_TEMPLATE
      + f"""
    <div class="nav">
        Logged in as: <b>{{{{ session['user'] }}}}</b> ({{{{ session['role'] }}}})<br><br>
        {get_nav_links()}
    </div>
    <h2>Edit Vendor: {vendor[1]}</h2>
    <div class="card">
        <form method="POST">
            <label><b>Vendor Name:</b></label>
            <input type="text" name="vendor_name" value="{vendor[1]}" required>
            
            <label><b>Contact Info:</b></label>
            <input type="text" name="contact_info" value="{vendor[2] or ''}">
            
            <label><b>Select Registers:</b></label>
            <div class="checkbox-group">{register_checkboxes}</div>
            
            <label><b>Select Master Items:</b></label>
            <div class="checkbox-group">{checkboxes}</div>
            
            <button type="submit" style="background:#28a745;">Update Vendor</button>
            <a href="/add-items?type=Vendors"><button type="button" class="danger" style="margin-left: 10px;">Cancel</button></a>
        </form>
    </div>
    """
  )
  return render_template_string(page)

@app.route("/delete-item/<int:item_id>")
@admin_required
def delete_item(item_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute("SELECT item_name FROM master_items WHERE id = ?", (item_id,))
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM master_items WHERE id = ?", (item_id,))
    conn.commit()
    log_action(session["user"], "DELETE ITEM", f"Removed item '{row[0]}'")
  conn.close()
  return redirect(url_for("add_items", type="Items"))

@app.route("/edit-item/<int:item_id>", methods=["POST"])
@admin_required
def edit_item(item_id):
    item_name     = request.form.get("item_name", "").strip()
    item_category = request.form.get("item_category", "Brass Item").strip()
    min_quantity  = float(request.form.get("min_quantity") or 0)
    if item_name:
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        try:
            # Get old name first
            cursor.execute("SELECT item_name FROM master_items WHERE id=?", (item_id,))
            old_row = cursor.fetchone()
            old_name = old_row[0] if old_row else None

            cursor.execute(
                "UPDATE master_items SET item_name=?, item_category=?, min_quantity=? WHERE id=?",
                (item_name, item_category, min_quantity, item_id)
            )

            # Cascade to all transaction tables
            if old_name and old_name != item_name:
                for table, col in [
                    ("casting_in", "item"),
                    ("machining_register", "item"),
                    ("rejection_register", "item"),
                    ("plating_out", "item"),
                    ("plating_in", "item"),
                    ("grinder_register", "item"),
                    ("opening_stock", "item"),
                    ("fg_deductions", "item"),
                ]:
                    cursor.execute(
                        f"UPDATE {table} SET {col}=? WHERE {col}=?",
                        (item_name, old_name)
                    )

            conn.commit()
            log_action(session["user"], "UPDATE ITEM", f"Renamed '{old_name}' → '{item_name}'")
        except sqlite3.IntegrityError:
            pass
        conn.close()
    return redirect(url_for("add_items", type="Items"))

@app.route("/delete-vendor/<int:v_id>")
@admin_required
def delete_vendor(v_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute("SELECT vendor_name FROM vendors WHERE id = ?", (v_id,))
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM vendors WHERE id = ?", (v_id,))
    conn.commit()
    log_action(session["user"], "DELETE VENDOR", f"Removed vendor '{row[0]}'")
  conn.close()
  return redirect(url_for("add_items", type="Vendors"))


@app.route("/delete-box-pack/<int:bp_id>")
@admin_required
def delete_box_pack(bp_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT box_pack_name FROM box_pack_items WHERE id = ?", (bp_id,)
  )
  row = cursor.fetchone()
  if row:
    cursor.execute(
        "DELETE FROM box_pack_components WHERE box_pack_id = ?", (bp_id,)
    )
    cursor.execute("DELETE FROM box_pack_items WHERE id = ?", (bp_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE BOX PACK ITEM",
        f"Removed box pack '{row[0]}'",
    )
  conn.close()
  return redirect(url_for("add_items", type="Box Pack Items"))

@app.route("/edit-box-pack/<int:bp_id>", methods=["POST"])
@admin_required
def edit_box_pack(bp_id):
    bp_name = request.form.get("box_pack_name", "").strip()
    item_ids = request.form.getlist("item_ids")
    if not bp_name or not item_ids:
        return redirect(url_for("add_items", type="Box Pack Items"))
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("UPDATE box_pack_items SET box_pack_name=? WHERE id=?", (bp_name, bp_id))
    cursor.execute("DELETE FROM box_pack_components WHERE box_pack_id=?", (bp_id,))
    for iid in item_ids:
        cursor.execute("INSERT INTO box_pack_components (box_pack_id, item_id) VALUES (?,?)", (bp_id, int(iid)))
    conn.commit()
    log_action(session["user"], "EDIT BOX PACK ITEM", f"Updated '{bp_name}'")
    conn.close()
    return redirect(url_for("add_items", type="Box Pack Items"))

@app.route("/delete-sale-register/<int:sale_id>")
@admin_required
def delete_sale_register(sale_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT box_pack_name, party, bill_no FROM sale_register WHERE id = ?",
      (sale_id,),
  )
  row = cursor.fetchone()
  # Restore pcs back to finished_goods
  cursor.execute(
      "SELECT item_name, pcs FROM sale_register_deductions WHERE sale_id = ?",
      (sale_id,)
    )
    # item_name here is now the box_pack_name, item_id is box_pack_id
  cursor.execute(
      "SELECT item_id, pcs FROM sale_register_deductions WHERE sale_id = ?",
      (sale_id,)
  )
  ded_row = cursor.fetchone()
  if ded_row:
      bp_id_to_restore, pcs_to_restore = ded_row
      cursor.execute(
          "UPDATE finished_goods SET pcs = pcs + ? WHERE box_pack_id = ?",
          (pcs_to_restore, bp_id_to_restore)
      )
  if row:
    box_pack_name, party, bill_no = row

    # Get how many pcs were sold to restore to finished goods
    cursor.execute(
        "SELECT pcs FROM sale_register WHERE id = ?", (sale_id,)
    )
    sale_pcs_row = cursor.fetchone()
    sold_pcs = sale_pcs_row[0] if sale_pcs_row else 0

    # Also get box_pack_id to restore finished goods
    cursor.execute(
        "SELECT box_pack_id FROM sale_register WHERE id = ?", (sale_id,)
    )
    bp_id_row = cursor.fetchone()
    bp_id = bp_id_row[0] if bp_id_row else None

    # Restore pcs back to finished goods
    if bp_id and sold_pcs:
        cursor.execute(
            "UPDATE finished_goods SET pcs = pcs + ? WHERE box_pack_id = ?",
            (sold_pcs, bp_id)
        )

    cursor.execute(
        "DELETE FROM sale_register_cartons WHERE sale_id = ?", (sale_id,)
    )
    cursor.execute(
        "DELETE FROM sale_register_deductions WHERE sale_id = ?", (sale_id,)
    )
    cursor.execute("DELETE FROM sale_register WHERE id = ?", (sale_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE SALE REGISTER ENTRY",
        f"Removed sale entry of box pack '{box_pack_name}' to party '{party}'"
        f" (Bill No: {bill_no}). {sold_pcs} pcs restored to Finished Goods stock.",
    )
  conn.close()
  return redirect(url_for("registers", type="Sale Register"))

@app.route("/delete-finished-goods-log/<int:log_id>")
@admin_required
def delete_finished_goods_log(log_id):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "SELECT box_pack_id, box_pack_name, pcs FROM finished_goods_log WHERE id = ?", (log_id,)
    )
    row = cursor.fetchone()
    if row:
        bp_id, bp_name, pcs = row
        # Reverse finished goods stock
        cursor.execute(
            "UPDATE finished_goods SET pcs = pcs - ? WHERE box_pack_id = ?",
            (pcs, bp_id)
        )
        # Reverse assembly deductions (delete them — inventory calc will naturally re-include them)
        cursor.execute(
            "DELETE FROM finished_goods_assembly_deductions WHERE finished_goods_log_id = ?",
            (log_id,)
        )
        cursor.execute("DELETE FROM finished_goods_log WHERE id = ?", (log_id,))
        conn.commit()
        log_action(
            session["user"],
            "DELETE FINISHED GOODS ENTRY",
            f"Removed log entry for '{bp_name}' ({pcs} pcs). Assembly deductions reversed."
        )
    conn.close()
    return redirect(url_for("registers", type="Finished Goods"))

@app.route("/delete-unit/<int:u_id>")
@admin_required
def delete_unit(u_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute("SELECT unit_name FROM units WHERE id = ?", (u_id,))
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM units WHERE id = ?", (u_id,))
    conn.commit()
    log_action(session["user"], "DELETE UNIT", f"Removed unit '{row[0]}'")
  conn.close()
  return redirect(url_for("add_items", type="unit of measurement"))

# @app.route("/assign-item-unit", methods=["POST"])
# @admin_required
# def assign_item_unit():
#     item_id = request.form.get("item_id")
#     unit_id = request.form.get("unit_id") or None
#     if item_id:
#         conn = sqlite3.connect(DB_NAME)
#         cursor = conn.cursor()
#         if unit_id:
#             cursor.execute("""
#                 INSERT INTO item_units (item_id, unit_id) VALUES (?, ?)
#                 ON CONFLICT(item_id) DO UPDATE SET unit_id = excluded.unit_id
#             """, (item_id, unit_id))
#         else:
#             cursor.execute("DELETE FROM item_units WHERE item_id = ?", (item_id,))
#         conn.commit()
#         log_action(session["user"], "ASSIGN ITEM UNIT",
#                    f"item_id={item_id} assigned unit_id={unit_id}")
#         conn.close()
#     return redirect(url_for("add_items", type="unit of measurement"))

@app.route("/delete-opening-stock/<int:os_id>")
@admin_required
def delete_opening_stock(os_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT item, stage FROM opening_stock WHERE id = ?", (os_id,)
  )
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM opening_stock WHERE id = ?", (os_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE OPENING STOCK",
        f"Removed opening stock for '{row[0]}' ({row[1]})",
    )
  conn.close()
  return redirect(url_for("registers", type="Opening Stock"))

@app.route("/edit-opening-stock/<int:os_id>", methods=["POST"])
@admin_required
def edit_opening_stock(os_id):
    date = request.form.get("date")
    item = request.form.get("item")
    stage = request.form.get("stage")
    pieces = request.form.get("pieces") or 0
    weight = request.form.get("weight") or 0
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE opening_stock SET date=?, item=?, stage=?, pieces=?, weight=? WHERE id=?",
        (date, item, stage, pieces, weight, os_id)
    )
    conn.commit()
    log_action(session["user"], "EDIT OPENING STOCK", f"Updated id={os_id} item={item} stage={stage}")
    conn.close()
    return redirect(url_for("registers", type="Opening Stock"))

@app.route("/delete-opening-scrap-bag/<int:bag_id>")
@admin_required
def delete_opening_scrap_bag(bag_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT bag_number FROM brass_in WHERE id = ? AND vendor ="
      " 'Opening Balance'",
      (bag_id,),
  )
  row = cursor.fetchone()
  if row:
    cursor.execute(
        "DELETE FROM brass_in WHERE id = ? AND vendor = 'Opening Balance'",
        (bag_id,),
    )
    conn.commit()
    log_action(
        session["user"],
        "DELETE OPENING SCRAP STOCK",
        f"Removed opening scrap balance Bag #{row[0]}",
    )
  conn.close()
  return redirect(url_for("registers", type="Opening Stock"))

@app.route("/edit-brass-in/<int:row_id>", methods=["POST"])
@admin_required
def edit_brass_in(row_id):
    date   = request.form.get("date")
    vendor = request.form.get("vendor")
    bag    = int(request.form.get("bag_number") or 0)
    weight = float(request.form.get("weight") or 0)
    conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
    cursor.execute(
        "UPDATE brass_in SET date=?, vendor=?, bag_number=?, weight=?, remaining_weight=? WHERE id=?",
        (date, vendor, bag, weight, weight, row_id))
    conn.commit(); conn.close()
    log_action(session["user"], "UPDATE BRASS IN", f"id={row_id}")
    return redirect(url_for("registers", type="Brass In"))

@app.route("/delete-brass-in/<int:bag_id>")
@admin_required
def delete_brass_in(bag_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT bag_number, vendor FROM brass_in WHERE id = ?", (bag_id,)
  )
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM brass_in WHERE id = ?", (bag_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE BRASS IN",
        f"Removed Brass In Bag #{row[0]} for vendor '{row[1]}'",
    )
  conn.close()
  return redirect(
      url_for("registers", type="Brass In", vendor=request.args.get("vendor", ""))
  )


@app.route("/delete-brass-out/<int:out_id>")
@admin_required
def delete_brass_out(out_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT bag_number, vendor, weight, brass_in_id FROM brass_out WHERE"
      " id = ?",
      (out_id,),
  )
  row = cursor.fetchone()
  if row:
    bag_num, vendor, out_weight, brass_in_id = row
    cursor.execute("DELETE FROM brass_out WHERE id = ?", (out_id,))
    # Restore the linked scrap bag back to available, if it still exists.
    if brass_in_id:
      cursor.execute(
          "SELECT weight FROM brass_in WHERE id = ?", (brass_in_id,)
      )
      bin_row = cursor.fetchone()
      if bin_row:
        cursor.execute(
            "UPDATE brass_in SET status = 'available', remaining_weight = ?"
            " WHERE id = ?",
            (bin_row[0], brass_in_id),
        )
    conn.commit()
    log_action(
        session["user"],
        "DELETE BRASS OUT",
        f"Removed Brass Out entry for Bag #{bag_num}, vendor '{vendor}'"
        f" ({out_weight:.3f} KG). Linked scrap bag restored to available.",
    )
  conn.close()
  return redirect(
      url_for("registers", type="Brass Out", vendor=request.args.get("vendor", ""))
  )

@app.route("/edit-casting-in/<int:row_id>", methods=["POST"])
@admin_required
def edit_casting_in(row_id):
    date   = request.form.get("date")
    vendor = request.form.get("vendor")
    item   = request.form.get("item")
    pieces = int(request.form.get("pieces") or 0)
    weight = float(request.form.get("weight") or 0)
    conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
    cursor.execute(
        "UPDATE casting_in SET date=?, vendor=?, item=?, pieces=?, weight=? WHERE id=?",
        (date, vendor, item, pieces, weight, row_id))
    conn.commit(); conn.close()
    log_action(session["user"], "UPDATE CASTING IN", f"id={row_id}")
    return redirect(url_for("registers", type="Casting In"))

@app.route("/delete-casting-in/<int:c_id>")
@admin_required
def delete_casting_in(c_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT item, vendor, pieces, weight FROM casting_in WHERE id = ?",
      (c_id,),
  )
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM casting_in WHERE id = ?", (c_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE CASTING IN",
        f"Removed Casting In entry for item '{row[0]}', vendor '{row[1]}'"
        f" ({row[2]} pcs / {row[3]:.3f} KG)",
    )
  conn.close()
  return redirect(
      url_for("registers", type="Casting In", vendor=request.args.get("vendor", ""))
  )

@app.route("/edit-machining/<int:row_id>", methods=["POST"])
@admin_required
def edit_machining(row_id):
    date   = request.form.get("date")
    item   = request.form.get("item")
    pieces = int(request.form.get("pieces") or 0)
    weight = float(request.form.get("weight") or 0)
    conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
    cursor.execute(
        "UPDATE machining_register SET date=?, item=?, pieces=?, weight=? WHERE id=?",
        (date, item, pieces, weight, row_id))
    conn.commit(); conn.close()
    log_action(session["user"], "UPDATE MACHINING", f"id={row_id}")
    return redirect(url_for("registers", type="Machining"))

@app.route("/delete-machining/<int:m_id>")
@admin_required
def delete_machining(m_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT item, direction, pieces, weight FROM machining_register WHERE"
      " id = ?",
      (m_id,),
  )
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM machining_register WHERE id = ?", (m_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE MACHINING ENTRY",
        f"Removed Machining entry for item '{row[0]}' ({row[1]})",
    )
  conn.close()
  return redirect(url_for("registers", type="Machining"))

@app.route("/edit-plating-inward/<int:challan_id>", methods=["POST"])
@admin_required
def edit_plating_inward(challan_id):
    pieces = int(request.form.get("pieces") or 0)
    weight = float(request.form.get("weight") or 0)
    conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
    cursor.execute(
        "UPDATE plating_items SET pieces=?, weight=? WHERE challan_id=?",
        (pieces, weight, challan_id))
    conn.commit(); conn.close()
    log_action(session["user"], "UPDATE PLATING INWARD", f"challan_id={challan_id}")
    return redirect(url_for("registers", type="Plating", direction="Inward"))

@app.route("/delete-plating-challan/<int:challan_id>")
@admin_required
def delete_plating_challan(challan_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT challan_number, date, vendor, direction FROM plating_challans"
      " WHERE id = ?",
      (challan_id,),
  )
  row = cursor.fetchone()
  if row:
    challan_no, c_date, vendor, direction = row
    cursor.execute(
        "DELETE FROM plating_items WHERE challan_id = ?", (challan_id,)
    )
    cursor.execute("DELETE FROM plating_challans WHERE id = ?", (challan_id,))
    conn.commit()
    label = f"Challan #{challan_no}" if challan_no else "Inward receipt"
    log_action(
        session["user"],
        "DELETE PLATING ENTRY",
        f"Removed Plating {direction} {label} dated {c_date} for vendor"
        f" '{vendor}'",
    )
  conn.close()
  return redirect(
      url_for(
          "registers",
          type="Plating",
          direction=request.args.get("direction", "Outward"),
      )
  )

@app.route("/edit-miscellaneous/<int:row_id>", methods=["POST"])
@admin_required
def edit_miscellaneous(row_id):
    date        = request.form.get("date")
    option_type = request.form.get("option_type")
    vendor      = request.form.get("vendor", "")
    item        = request.form.get("item")
    pieces      = int(request.form.get("pieces") or 0)
    weight      = float(request.form.get("weight") or 0)
    unit        = request.form.get("quantity_unit", "")
    conn = sqlite3.connect(DB_NAME); cursor = conn.cursor()
    cursor.execute(
        "UPDATE miscellaneous_entries SET date=?,option_type=?,vendor=?,item=?,pieces=?,weight=?,quantity_unit=? WHERE id=?",
        (date, option_type, vendor, item, pieces, weight, unit, row_id))
    conn.commit(); conn.close()
    log_action(session["user"], "UPDATE MISC ENTRY", f"id={row_id}")
    return redirect(url_for("registers", type="Miscellaneous"))

@app.route("/delete-miscellaneous/<int:m_id>")
@admin_required
def delete_miscellaneous(m_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT option_type, item, vendor, pieces, weight FROM"
      " miscellaneous_entries WHERE id = ?",
      (m_id,),
  )
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM miscellaneous_entries WHERE id = ?", (m_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE MISCELLANEOUS ENTRY",
        f"Removed Miscellaneous '{row[0]}' entry for item '{row[1]}'"
        f" ({row[3]} pcs / {row[4]:.3f} KG)",
    )
  conn.close()
  return redirect(url_for("registers", type="Miscellaneous"))


@app.route("/delete-rejection/<int:r_id>")
@admin_required
def delete_rejection(r_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT item, rejection_type, vendor, pieces, weight FROM"
      " rejection_register WHERE id = ?",
      (r_id,),
  )
  row = cursor.fetchone()
  if row:
    cursor.execute("DELETE FROM rejection_register WHERE id = ?", (r_id,))
    conn.commit()
    log_action(
        session["user"],
        "DELETE REJECTION ENTRY",
        f"Removed '{row[1]}' rejection entry (item '{row[0]}', vendor"
        f" '{row[2]}', {row[3]} pcs / {row[4]:.3f} KG). Note: any stock"
        " already deducted by this rejection is NOT automatically restored.",
    )
  conn.close()
  return redirect(
      url_for(
          "registers",
          type="Rejection",
          sub_rej_type=request.args.get("sub_rej_type", "inventory"),
      )
  )


_THIN_SIDE = Side(style="thin", color="B7B7B7")
_CELL_BORDER = Border(
    left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE
)
_TITLE_FILL = PatternFill("solid", fgColor="FFFFFF")
_TITLE_FONT = Font(bold=True, size=11)
_HEADER_FILL = PatternFill("solid", fgColor="FFFFFF")
_HEADER_FONT = Font(bold=True)
_TOTAL_FILL = PatternFill("solid", fgColor="FFFFFF")
_TOTAL_FONT = Font(bold=True)
_EMPTY_FONT = Font(italic=True)
_CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
_COLUMN_WIDTHS = {
    "date": 12,
    "vendor": 16,
    "pieces": 9,
    "weight": 11,
    "challan_number": 15,
    "grinding_vendor": 16,
    "pieces_received": 16,
}
_GAP_COLUMN_WIDTH = 3


def _column_number_format(col_name):
  lowered = col_name.lower()
  if "weight" in lowered:
    return "0.000"
  if "piece" in lowered:
    return "0"
  return None


def _sanitize_sheet_name(name, used_names):
  """Make a string safe as an Excel sheet name and unique within the workbook."""
  invalid_chars = ["\\", "/", "*", "[", "]", ":", "?"]
  cleaned = name
  for ch in invalid_chars:
    cleaned = cleaned.replace(ch, "")
  cleaned = cleaned.strip() or "Item"
  cleaned = cleaned[:31]

  candidate = cleaned
  suffix = 2
  while candidate in used_names:
    trimmed = cleaned[: 31 - len(f"_{suffix}")]
    candidate = f"{trimmed}_{suffix}"
    suffix += 1
  used_names.add(candidate)
  return candidate

def _extract_tables_from_html(html):
  try:
    import io
    return [df.fillna("") for df in pd.read_html(io.StringIO(html))]
  except ValueError:
    return []

def _fmt_paid(val):
  return "Yes" if val else "No"


def _get_vendor_ledger_data(cursor, ledger_sub_type, vendor_filter=""):
  """Pulls Vendor Ledger data straight from the DB so exports get clean
  values: real 'Yes'/'No' for payment status, no NaN, no scraped HTML."""
  results = []

  if ledger_sub_type == "brass_out_casting_in":
    cast_date_from = request.values.get("cast_date_from", "")
    cast_date_to   = request.values.get("cast_date_to", "")

    cursor.execute(
        """SELECT DISTINCT v.vendor_name FROM vendors v
           JOIN vendor_registers vr ON v.id = vr.vendor_id
           WHERE vr.register_type IN ('Brass Out', 'Casting In')
           ORDER BY v.vendor_name"""
    )
    vendor_names = [r[0] for r in cursor.fetchall()]
    if vendor_filter:
      vendor_names = [v for v in vendor_names if v == vendor_filter]

    for v_name in vendor_names:
      # ── Helper: cumulative brass balance up to (but not including) a date ──
      # brass_balance on any day = cumulative brass_out − cumulative casting_accounted_return
      # We compute opening balance = balance strictly BEFORE cast_date_from.
      opening_balance = None
      if cast_date_from:
        cursor.execute(
            "SELECT COALESCE(SUM(weight),0) FROM brass_out WHERE given_to=? AND date<?",
            (v_name, cast_date_from),
        )
        pre_bo = cursor.fetchone()[0] or 0.0
        cursor.execute(
            "SELECT COALESCE(SUM(weight),0) FROM casting_in WHERE vendor=? AND date<?",
            (v_name, cast_date_from),
        )
        pre_ci = cursor.fetchone()[0] or 0.0
        cursor.execute(
            "SELECT COALESCE(SUM(weight),0) FROM rejection_register"
            " WHERE vendor=? AND rejection_type='Vendor Rejection' AND date<?",
            (v_name, cast_date_from),
        )
        pre_rej = cursor.fetchone()[0] or 0.0
        pre_adj = max(0.0, pre_ci - pre_rej)
        pre_burn = pre_adj * 0.05
        pre_accounted = pre_adj + pre_burn
        opening_balance = pre_bo - pre_accounted

      # ── Brass Out rows (filtered) ──
      bo_query = "SELECT id, date, bag_number, weight, payment_made FROM brass_out WHERE given_to=?"
      bo_params = [v_name]
      if cast_date_from:
        bo_query += " AND date>=?"; bo_params.append(cast_date_from)
      if cast_date_to:
        bo_query += " AND date<=?"; bo_params.append(cast_date_to)
      bo_query += " ORDER BY date ASC"
      cursor.execute(bo_query, bo_params)
      bo_raw = cursor.fetchall()   # (id, date, bag_number, weight, payment_made)

      # Build running brass_balance per row (cumulative within the filtered window,
      # starting from opening_balance if a date filter is active).
      # brass_balance[i] = opening + sum(brass_out up to row i) - sum(casting_accounted up to same date)
      # We store it per-row in the export data but NOT shown in the on-screen table.
      bo_rows = []
      # opening balance marker row for brass-out side
      if opening_balance is not None:
        label = f"Brass Balance Before {cast_date_from}"
        bo_rows.append([label, "", f"{opening_balance:.3f}", ""])

      for _id, d, b, w, p in bo_raw:
        bo_rows.append([d, f"Bag #{b}", f"{w:.3f}", _fmt_paid(p)])

      total_bo_weight_period = sum(r[3] for r in bo_raw)
      total_bo_weight_all = (opening_balance or 0.0) + total_bo_weight_period if cast_date_from else total_bo_weight_period

      # ── Casting In rows (filtered) ──
      ci_query = "SELECT date, item, pieces, weight, payment_made FROM casting_in WHERE vendor=?"
      ci_params = [v_name]
      if cast_date_from:
        ci_query += " AND date>=?"; ci_params.append(cast_date_from)
      if cast_date_to:
        ci_query += " AND date<=?"; ci_params.append(cast_date_to)
      ci_query += " ORDER BY date ASC"
      cursor.execute(ci_query, ci_params)
      ci_raw = cursor.fetchall()

      ci_rows = [
          [d, i, str(pcs), f"{w:.3f}", _fmt_paid(p)]
          for d, i, pcs, w, p in ci_raw
      ]
      total_ci_pieces = sum(r[2] for r in ci_raw)
      total_ci_weight  = sum(r[3] for r in ci_raw)

      # Vendor rejection — filter by same date window
      rej_query = ("SELECT COALESCE(SUM(weight),0) FROM rejection_register"
                   " WHERE vendor=? AND rejection_type='Vendor Rejection'")
      rej_params = [v_name]
      if cast_date_from:
        rej_query += " AND date>=?"; rej_params.append(cast_date_from)
      if cast_date_to:
        rej_query += " AND date<=?"; rej_params.append(cast_date_to)
      cursor.execute(rej_query, rej_params)
      vendor_rejection_weight = cursor.fetchone()[0] or 0.0

      adjusted_casting_weight  = max(0.0, total_ci_weight - vendor_rejection_weight)
      burning_loss             = adjusted_casting_weight * 0.05
      total_accounted_return   = adjusted_casting_weight + burning_loss

      # Net balance uses ALL brass out (including pre-period opening) vs period casting
      if cast_date_from:
        # full period net: opening_balance (already net) + period brass_out - period accounted
        net_balance = opening_balance + total_bo_weight_period - total_accounted_return
      else:
        cursor.execute(
            "SELECT COALESCE(SUM(weight),0) FROM brass_out WHERE given_to=?", (v_name,)
        )
        all_bo = cursor.fetchone()[0] or 0.0
        net_balance = all_bo - total_accounted_return

      # Append totals to brass-out side
      if bo_raw:
        bo_rows.append(["Total Brass Out (Period)", "", f"{total_bo_weight_period:.3f}", ""])

      # Append totals to casting-in side
      if ci_raw:
        ci_rows.append(["Total", "", str(total_ci_pieces), f"{total_ci_weight:.3f}", ""])
        ci_rows.append(["Less: Vendor Rejection", "", "", f"-{vendor_rejection_weight:.3f}", ""])
        ci_rows.append(["Adjusted Subtotal", "", "", f"{adjusted_casting_weight:.3f}", ""])
        ci_rows.append(["Burning Loss (5%)", "", "", f"{burning_loss:.3f}", ""])
        ci_rows.append(["Total Accounted Return", "", "", f"{total_accounted_return:.3f}", ""])

      results.append({
          "vendor": v_name,
          "cast_date_from": cast_date_from,
          "cast_date_to":   cast_date_to,
          "opening_balance": opening_balance,
          "total_bo_weight_period": total_bo_weight_period,
          "tables": [
              {"title": "Brass Out (Scrap Given)",
               "columns": ["Date", "Bag Number", "Weight (KG)", "Paid"],
               "rows": bo_rows},
              {"title": "Casting In",
               "columns": ["Date", "Item", "Pieces", "Weight (KG)", "Paid"],
               "rows": ci_rows},
          ],
          "summary": f"Net Balance: {net_balance:.3f} KG "
                     f"({'Due From Vendor' if net_balance > 0 else 'Advance With Vendor' if net_balance < 0 else 'Fully Settled'})",
      })

  elif ledger_sub_type == "brass_in":
    cursor.execute(
        """SELECT DISTINCT v.vendor_name FROM vendors v
           JOIN vendor_registers vr ON v.id = vr.vendor_id
           WHERE vr.register_type = 'Brass In'
           ORDER BY v.vendor_name"""
    )
    vendor_names = [r[0] for r in cursor.fetchall()]
    if vendor_filter:
      vendor_names = [v for v in vendor_names if v == vendor_filter]

    for v_name in vendor_names:
      cursor.execute(
          "SELECT date, bag_number, weight, remaining_weight, status,"
          " payment_made FROM brass_in WHERE vendor = ? ORDER BY date ASC",
          (v_name,),
      )
      raw = cursor.fetchall()
      rows = [
          [d, f"Bag #{b}", f"{w:.3f}", f"{rem:.3f}", st.capitalize(), _fmt_paid(p)]
          for d, b, w, rem, st, p in raw
      ]
      total_in_weight = sum(r[2] for r in raw)
      total_remaining = sum(r[3] for r in raw if r[4] == "available")

      cursor.execute(
          "SELECT COALESCE(SUM(weight), 0) FROM rejection_register WHERE"
          " vendor = ? AND rejection_type = 'Brass In Vendor Rejection'",
          (v_name,),
      )
      brass_in_vendor_rejection = cursor.fetchone()[0] or 0.0

      if raw:
        rows.append(["Total Scrap Supplied In", "", f"{total_in_weight:.3f}", "", "", ""])
        rows.append(["Total Rejected", "", f"-{brass_in_vendor_rejection:.3f}", "", "", ""])
        rows.append(["Currently Available Stock", "", "", f"{total_remaining:.3f}", "", ""])

      results.append({
          "vendor": v_name,
          "tables": [{
              "title": "Brass In (Scrap Supplied)",
              "columns": ["Date", "Bag Number", "Original Weight (KG)",
                          "Remaining Weight (KG)", "Status", "Paid"],
              "rows": rows,
          }],
      })

  elif ledger_sub_type == "miscellaneous":
    cursor.execute(
        """SELECT DISTINCT v.vendor_name FROM vendors v
           JOIN vendor_registers vr ON v.id = vr.vendor_id
           WHERE vr.register_type = 'Miscellaneous'
           ORDER BY v.vendor_name"""
    )
    vendor_names = [r[0] for r in cursor.fetchall()]
    if vendor_filter:
      vendor_names = [v for v in vendor_names if v == vendor_filter]

    for v_name in vendor_names:
      cursor.execute(
          "SELECT date, item, pieces, weight, payment_made FROM"
          " miscellaneous_entries WHERE vendor = ? AND option_type ="
          " 'purchase' ORDER BY date ASC",
          (v_name,),
      )
      raw = cursor.fetchall()
      rows = [
          [d, i, str(pcs), f"{w:.3f}", _fmt_paid(p)] for d, i, pcs, w, p in raw
      ]
      total_pcs = sum(r[2] for r in raw)
      total_wt = sum(r[3] for r in raw)
      if raw:
        rows.append(["Total", "", str(total_pcs), f"{total_wt:.3f}", ""])

      results.append({
          "vendor": v_name,
          "tables": [{
              "title": "Miscellaneous Purchases",
              "columns": ["Date", "Item", "Pieces", "Weight (KG)", "Paid"],
              "rows": rows,
          }],
      })

  elif ledger_sub_type in ("grinding", "plating"):
    reg_type_name = "Grinding" if ledger_sub_type == "grinding" else "Plating"
    date_from = request.values.get(
        "grind_date_from" if ledger_sub_type == "grinding" else "plate_date_from", ""
    )
    date_to = request.values.get(
        "grind_date_to" if ledger_sub_type == "grinding" else "plate_date_to", ""
    )

    cursor.execute(
        """SELECT DISTINCT v.vendor_name FROM vendors v
           JOIN vendor_registers vr ON v.id = vr.vendor_id
           WHERE vr.register_type = ?
           ORDER BY v.vendor_name""",
        (reg_type_name,),
    )
    vendor_names = [r[0] for r in cursor.fetchall()]
    if vendor_filter:
      vendor_names = [v for v in vendor_names if v == vendor_filter]

    vendor_col = "pi.grinding_vendor" if ledger_sub_type == "grinding" else "pc.vendor"

    for v_name in vendor_names:
      query = f"""SELECT pc.date, pc.challan_number, pi.item, pi.pieces,
                         pi.payment_made,
                         COALESCE((
                             SELECT rl.rate FROM rate_list rl
                             JOIN master_items mi ON mi.id = rl.item_id
                             JOIN vendors vnd ON vnd.id = rl.vendor_id
                             WHERE mi.item_name = pi.item AND vnd.vendor_name = ?
                             LIMIT 1
                         ), 0) as rate
                  FROM plating_items pi
                  JOIN plating_challans pc ON pi.challan_id = pc.id
                  WHERE {vendor_col} = ? AND pc.direction = 'Outward'"""
      params = [v_name, v_name]
      if date_from:
        query += " AND pc.date >= ?"
        params.append(date_from)
      if date_to:
        query += " AND pc.date <= ?"
        params.append(date_to)
      query += " ORDER BY pc.date ASC"
      cursor.execute(query, params)
      raw = cursor.fetchall()

      rows = []
      total_pcs = 0
      total_cost = 0.0
      for d, chno, item, pcs, paid, rate in raw:
        rate = float(rate)
        cost = pcs * rate
        total_pcs += pcs
        total_cost += cost
        rows.append([
            d, chno or "", item, str(pcs),
            f"{rate:.2f}" if rate > 0 else "",
            f"{cost:.2f}" if rate > 0 else "",
            _fmt_paid(paid),
        ])
      if raw:
        rows.append(["Total", "", "", str(total_pcs), "", f"{total_cost:.2f}", ""])

      results.append({
          "vendor": v_name,
          "tables": [{
              "title": f"{reg_type_name} Entries",
              "columns": ["Date", "Challan No", "Item", "Pieces",
                          "Rate", "Total Cost", "Paid"],
              "rows": rows,
          }],
      })

  return results

def _get_low_stock_alerts(cursor):
    """
    Returns a list of low-stock alerts.

    Alerts include:
      1. Overall item quantity below min_quantity.
      2. Brass item casting quantity below min_casting_quantity.
      3. Brass item machined quantity below min_machined_quantity.
      4. Brass item assembly quantity below min_assembly_quantity.
      5. Miscellaneous purchased quantity below min_purchase_quantity.
    """

    cursor.execute("""
        SELECT
            id,
            item_name,
            item_category,
            COALESCE(min_quantity, 0),
            COALESCE(min_casting_quantity, 0),
            COALESCE(min_machined_quantity, 0),
            COALESCE(min_assembly_quantity, 0),
            COALESCE(min_purchase_quantity, 0)
        FROM master_items
        WHERE
            COALESCE(min_quantity, 0) > 0
            OR COALESCE(min_casting_quantity, 0) > 0
            OR COALESCE(min_machined_quantity, 0) > 0
            OR COALESCE(min_assembly_quantity, 0) > 0
            OR COALESCE(min_purchase_quantity, 0) > 0
    """)

    items = cursor.fetchall()

    # ---------------------------------------------------------
    # Opening stock
    # ---------------------------------------------------------

    cursor.execute("""
        SELECT item, stage, pieces, weight
        FROM opening_stock
    """)

    opening_map = {}

    for op_item, op_stage, op_pcs, op_wt in cursor.fetchall():
        opening_map[(op_item, op_stage)] = (
            op_pcs or 0,
            op_wt or 0.0
        )

    alerts = []

    for (
        item_id,
        item_name,
        item_category,
        min_quantity,
        min_casting_quantity,
        min_machined_quantity,
        min_assembly_quantity,
        min_purchase_quantity,
    ) in items:

        # =====================================================
        # BRASS ITEMS
        # =====================================================

        if item_category == "Brass Item":

            # -------------------------------------------------
            # Opening stock
            # -------------------------------------------------

            op_casted_pcs, _ = opening_map.get(
                (item_name, "casted"),
                (0, 0.0)
            )

            op_mach_pcs, _ = opening_map.get(
                (item_name, "machined"),
                (0, 0.0)
            )

            op_assembly_pcs, _ = opening_map.get(
                (item_name, "assembly"),
                (0, 0.0)
            )

            # -------------------------------------------------
            # Casting
            # -------------------------------------------------

            cursor.execute(
                """
                SELECT
                    COALESCE(SUM(pieces), 0),
                    COALESCE(SUM(weight), 0)
                FROM casting_in
                WHERE item = ?
                """,
                (item_name,)
            )

            casting_pcs, casting_wt = cursor.fetchone()

            casting_pcs = casting_pcs or 0

            # -------------------------------------------------
            # Machining
            # -------------------------------------------------

            cursor.execute(
                """
                SELECT COALESCE(SUM(pieces), 0)
                FROM machining_register
                WHERE item = ?
                """,
                (item_name,)
            )

            m_pcs = cursor.fetchone()[0] or 0

            # -------------------------------------------------
            # Rejections
            # -------------------------------------------------

            cursor.execute(
                """
                SELECT
                    rejection_type,
                    COALESCE(SUM(pieces), 0)
                FROM rejection_register
                WHERE
                    item = ?
                    AND rejection_type != 'Vendor Rejection'
                GROUP BY rejection_type
                """,
                (item_name,)
            )

            rej_data = {
                "Casted Rejection": 0,
                "Machined Rejection": 0,
                "Grinder Rejection": 0,
                "Plated Rejection": 0,
            }

            for r_type, r_pcs in cursor.fetchall():

                if r_type in rej_data:
                    rej_data[r_type] = r_pcs or 0

            casted_reject_pcs = (
                rej_data["Casted Rejection"]
                + rej_data["Machined Rejection"]
            )

            machined_reject_pcs = (
                rej_data["Grinder Rejection"]
            )

            assembly_reject_pcs = (
                rej_data["Plated Rejection"]
            )

            # -------------------------------------------------
            # Plating
            # -------------------------------------------------

            cursor.execute(
                """
                SELECT
                    COALESCE(SUM(pi.pieces), 0),
                    COALESCE(
                        SUM(
                            pi.pieces - pi.remaining_pieces
                        ),
                        0
                    )
                FROM plating_items pi
                JOIN plating_challans pc
                    ON pi.challan_id = pc.id
                WHERE
                    pi.item = ?
                    AND pc.direction = 'Outward'
                """,
                (item_name,)
            )

            plate_out_pcs, plate_in_pcs_old = cursor.fetchone()

            plate_out_pcs = plate_out_pcs or 0
            plate_in_pcs_old = plate_in_pcs_old or 0

            cursor.execute(
                """
                SELECT COALESCE(SUM(pi.pieces), 0)
                FROM plating_items pi
                JOIN plating_challans pc
                    ON pi.challan_id = pc.id
                WHERE
                    pi.item = ?
                    AND pc.direction = 'Inward'
                """,
                (item_name,)
            )

            plate_in_pcs_new = cursor.fetchone()[0] or 0

            plate_in_pcs = (
                plate_in_pcs_old
                + plate_in_pcs_new
            )

            # -------------------------------------------------
            # Finished goods deductions
            # -------------------------------------------------

            cursor.execute(
                """
                SELECT COALESCE(SUM(fgd.pcs), 0)
                FROM finished_goods_assembly_deductions fgd
                JOIN finished_goods_log fgl
                    ON fgd.finished_goods_log_id = fgl.id
                WHERE fgd.item_name = ?
                """,
                (item_name,)
            )

            fg_assembly_deduct_pcs = (
                cursor.fetchone()[0] or 0
            )

            # -------------------------------------------------
            # Calculate current quantities
            # -------------------------------------------------

            net_casted_pcs = (
                casting_pcs
                - m_pcs
                - casted_reject_pcs
                + op_casted_pcs
            )

            net_mach_pcs = (
                m_pcs
                - plate_out_pcs
                - machined_reject_pcs
                + op_mach_pcs
            )

            net_assembly_pcs = (
                plate_in_pcs
                - assembly_reject_pcs
                + op_assembly_pcs
                - fg_assembly_deduct_pcs
            )

            # Don't allow negative stock to make alerts confusing.
            actual_casting = max(0, net_casted_pcs)
            actual_machined = max(0, net_mach_pcs)
            actual_assembly = max(0, net_assembly_pcs)

            # -------------------------------------------------
            # Overall total quantity
            # -------------------------------------------------

            total_qty = (
                actual_casting
                + actual_machined
                + actual_assembly
            )

            # -------------------------------------------------
            # Overall item alert
            # -------------------------------------------------

            if min_quantity > 0 and total_qty < min_quantity:

                alerts.append({
                    "name": item_name,
                    "qty": int(total_qty),
                    "min_qty": int(min_quantity),
                    "category": item_category,
                    "stage": "Total",
                    "message": (
                        f"Low items of {item_name} in Total"
                    ),
                })

            # -------------------------------------------------
            # Casting alert
            # -------------------------------------------------

            if (
                min_casting_quantity > 0
                and actual_casting < min_casting_quantity
            ):

                alerts.append({
                    "name": item_name,
                    "qty": int(actual_casting),
                    "min_qty": int(min_casting_quantity),
                    "category": item_category,
                    "stage": "Casting",
                    "message": (
                        f"Low items of {item_name} in Casting"
                    ),
                })

            # -------------------------------------------------
            # Machined alert
            # -------------------------------------------------

            if (
                min_machined_quantity > 0
                and actual_machined < min_machined_quantity
            ):

                alerts.append({
                    "name": item_name,
                    "qty": int(actual_machined),
                    "min_qty": int(min_machined_quantity),
                    "category": item_category,
                    "stage": "Machined",
                    "message": (
                        f"Low items of {item_name} in Machined"
                    ),
                })

            # -------------------------------------------------
            # Assembly alert
            # -------------------------------------------------

            if (
                min_assembly_quantity > 0
                and actual_assembly < min_assembly_quantity
            ):

                alerts.append({
                    "name": item_name,
                    "qty": int(actual_assembly),
                    "min_qty": int(min_assembly_quantity),
                    "category": item_category,
                    "stage": "Assembly",
                    "message": (
                        f"Low items of {item_name} in Assembly"
                    ),
                })

        # =====================================================
        # MISCELLANEOUS ITEMS
        # =====================================================

        else:

            op_misc_pcs, _ = opening_map.get(
                (item_name, "misc"),
                (0, 0.0)
            )

            # -------------------------------------------------
            # Purchased
            # -------------------------------------------------

            cursor.execute(
                """
                SELECT
                    COALESCE(
                        SUM(
                            CASE
                                WHEN option_type = 'purchase'
                                THEN pieces
                                ELSE 0
                            END
                        ),
                        0
                    ),
                    COALESCE(
                        SUM(
                            CASE
                                WHEN option_type = 'in assembly'
                                THEN pieces
                                ELSE 0
                            END
                        ),
                        0
                    )
                FROM miscellaneous_entries
                WHERE item = ?
                """,
                (item_name,)
            )

            misc_pur_pcs, misc_asm_pcs = cursor.fetchone()

            misc_pur_pcs = misc_pur_pcs or 0
            misc_asm_pcs = misc_asm_pcs or 0

            # -------------------------------------------------
            # Rejection
            # -------------------------------------------------

            cursor.execute(
                """
                SELECT COALESCE(SUM(pieces), 0)
                FROM rejection_register
                WHERE
                    item = ?
                    AND rejection_type =
                        'Misc Vendor Rejection'
                """,
                (item_name,)
            )

            misc_reject_pcs = (
                cursor.fetchone()[0] or 0
            )

            # -------------------------------------------------
            # Current purchased stock
            # -------------------------------------------------

            purchased_qty = max(
                0,
                misc_pur_pcs
                - misc_asm_pcs
                - misc_reject_pcs
                + op_misc_pcs
            )

            total_qty = purchased_qty

            # -------------------------------------------------
            # Overall miscellaneous alert
            # -------------------------------------------------

            if min_quantity > 0 and total_qty < min_quantity:

                alerts.append({
                    "name": item_name,
                    "qty": int(total_qty),
                    "min_qty": int(min_quantity),
                    "category": item_category,
                    "stage": "Total",
                    "message": (
                        f"Low items of {item_name} in Total"
                    ),
                })

            # -------------------------------------------------
            # Purchased alert
            # -------------------------------------------------

            if (
                min_purchase_quantity > 0
                and purchased_qty < min_purchase_quantity
            ):

                alerts.append({
                    "name": item_name,
                    "qty": int(purchased_qty),
                    "min_qty": int(min_purchase_quantity),
                    "category": item_category,
                    "stage": "Purchased",
                    "message": (
                        f"Low items of {item_name} in Purchased"
                    ),
                })

    return alerts
def _vendor_export_buttons(ledger_sub_type, vendor_name, extra_params=""):
  """Small Excel/PDF export buttons for one vendor's ledger card."""
  base = (
      f"type=Vendor+Ledgers&ledger_sub_type={ledger_sub_type}"
      f"&vendor={quote(vendor_name)}"
  )
  if extra_params:
    base += f"&{extra_params}"
  return f"""<div style="margin-bottom:10px;">
        <a href="/export-register-excel?{base}">
            <button style="background:#217346;padding:4px 10px;font-size:12px;">⬇ Excel</button></a>
        <a href="/export-register-pdf?{base}">
            <button style="background:#c0392b;padding:4px 10px;font-size:12px;margin-left:6px;">⬇ PDF</button></a>
    </div>"""


def _render_full_page_html(view_func):
  """Calls an existing view function and returns its raw HTML string,
  reusing whatever `request.values` the current request already has."""
  result = view_func()
  if isinstance(result, tuple):
    result = result[0]
  return result

@app.route("/export-item-records")
@login_required
def export_item_records():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT item_name, item_category FROM master_items ORDER BY item_name ASC"
  )
  items = cursor.fetchall()

  output = io.BytesIO()
  used_sheet_names = set()

  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    if not items:
      pd.DataFrame([["No items found."]]).to_excel(
          writer, sheet_name="Items", index=False, header=False
      )

    for item_name, item_category in items:
      sheet_name = _sanitize_sheet_name(item_name, used_sheet_names)
      current_col = 0
      ws = writer.book.create_sheet(title=sheet_name)

      def write_block(title, df, sum_cols=None):
        nonlocal current_col
        sum_cols = sum_cols or []
        title_col_excel = current_col + 1  # openpyxl columns are 1-indexed

        if df.empty:
          empty_cell = ws.cell(
              row=2, column=title_col_excel, value="No records."
          )
          empty_cell.font = _EMPTY_FONT
          empty_cell.border = _CELL_BORDER
          ws.column_dimensions[get_column_letter(title_col_excel)].width = 20
          block_width = 1
        else:
          df.to_excel(
              writer,
              sheet_name=sheet_name,
              startrow=1,
              startcol=current_col,
              index=False,
          )
          columns = list(df.columns)
          ncols = len(columns)
          block_width = ncols

          totals_row = ["" for _ in columns]
          totals_row[0] = "Total"
          for col_name in sum_cols:
            if col_name in columns:
              totals_row[columns.index(col_name)] = df[col_name].sum()
          total_row_excel = 1 + len(df) + 2  # 1-indexed: header + data + next
          pd.DataFrame([totals_row]).to_excel(
              writer,
              sheet_name=sheet_name,
              startrow=total_row_excel - 1,
              startcol=current_col,
              index=False,
              header=False,
          )

          # Header row styling + number formats + column widths
          for i, col_name in enumerate(columns):
            col_excel = title_col_excel + i
            header_cell = ws.cell(row=2, column=col_excel)
            header_cell.font = _HEADER_FONT
            header_cell.fill = _HEADER_FILL
            header_cell.border = _CELL_BORDER
            header_cell.alignment = _CENTER_ALIGN

            number_format = _column_number_format(col_name)
            for data_row in range(3, 3 + len(df)):
              cell = ws.cell(row=data_row, column=col_excel)
              cell.border = _CELL_BORDER
              if number_format:
                cell.number_format = number_format

            total_cell = ws.cell(row=total_row_excel, column=col_excel)
            total_cell.font = _TOTAL_FONT
            total_cell.fill = _TOTAL_FILL
            total_cell.border = _CELL_BORDER
            if number_format and total_cell.value != "":
              total_cell.number_format = number_format

            ws.column_dimensions[get_column_letter(col_excel)].width = (
                _COLUMN_WIDTHS.get(col_name, 12)
            )

        # Title bar spanning the block's width
        if block_width > 1:
          ws.merge_cells(
              start_row=1,
              start_column=title_col_excel,
              end_row=1,
              end_column=title_col_excel + block_width - 1,
          )
        title_cell = ws.cell(row=1, column=title_col_excel)
        title_cell.value = title
        title_cell.font = _TITLE_FONT
        title_cell.fill = _TITLE_FILL
        title_cell.alignment = _CENTER_ALIGN
        for i in range(block_width):
          ws.cell(row=1, column=title_col_excel + i).border = _CELL_BORDER
          ws.cell(row=1, column=title_col_excel + i).fill = _TITLE_FILL

        gap_col_excel = title_col_excel + block_width
        ws.column_dimensions[get_column_letter(gap_col_excel)].width = (
            _GAP_COLUMN_WIDTH
        )
        current_col += block_width + 1  # section width + 1 gap column

      if item_category == "Brass Item":
        df_casting = pd.read_sql_query(
            "SELECT date, vendor, pieces, weight FROM casting_in WHERE item = ?"
            " ORDER BY id DESC",
            conn,
            params=(item_name,),
        )
        write_block("Casting", df_casting, sum_cols=["pieces", "weight"])

        df_machining = pd.read_sql_query(
            "SELECT date, pieces, weight FROM machining_register WHERE item = ?"
            " ORDER BY id DESC",
            conn,
            params=(item_name,),
        )
        write_block("Machining", df_machining, sum_cols=["pieces", "weight"])

        df_grind = pd.read_sql_query(
            """SELECT c.date, c.challan_number, i.grinding_vendor, i.pieces, i.weight 
               FROM plating_challans c 
               JOIN plating_items i ON c.id = i.challan_id 
               WHERE i.item = ? AND i.grinding_vendor IS NOT NULL AND i.grinding_vendor != '' 
               ORDER BY c.id DESC""",
            conn,
            params=(item_name,),
        )
        write_block("Grinding", df_grind, sum_cols=["pieces", "weight"])

        df_plating = pd.read_sql_query(
            """SELECT c.date, COALESCE(c.challan_number, '(Inward Receipt)') as challan_number,
                      c.vendor,
                      CASE WHEN c.direction = 'Inward' THEN i.pieces
                           ELSE (i.pieces - i.remaining_pieces) END as pieces_received
               FROM plating_challans c 
               JOIN plating_items i ON c.id = i.challan_id 
               WHERE i.item = ?
                 AND ((c.direction = 'Outward' AND (i.pieces - i.remaining_pieces) > 0)
                      OR (c.direction = 'Inward' AND i.pieces > 0))
               ORDER BY c.date DESC, c.id DESC""",
            conn,
            params=(item_name,),
        )
        write_block("Plating Inward", df_plating, sum_cols=["pieces_received"])
      else:
        df_misc_pur = pd.read_sql_query(
            "SELECT date, vendor, pieces, weight FROM miscellaneous_entries WHERE"
            " item = ? AND option_type = 'purchase' ORDER BY id DESC",
            conn,
            params=(item_name,),
        )
        write_block("Miscellaneous - Purchase", df_misc_pur, sum_cols=["pieces", "weight"])

        df_misc_asm = pd.read_sql_query(
            "SELECT date, pieces, weight FROM miscellaneous_entries WHERE item = ?"
            " AND option_type = 'in assembly' ORDER BY id DESC",
            conn,
            params=(item_name,),
        )
        write_block("Miscellaneous - In Assembly", df_misc_asm, sum_cols=["pieces", "weight"])

      df_sale = pd.read_sql_query(
          """SELECT sr.date, sr.bill_no, sd.pcs
             FROM sale_register_deductions sd
             JOIN sale_register sr ON sd.sale_id = sr.id
             WHERE sd.item_name = ? ORDER BY sr.id DESC""",
          conn,
          params=(item_name,),
      )
      write_block("Sale", df_sale, sum_cols=["pcs"])

    if items and "Sheet" in writer.book.sheetnames:
      writer.book.remove(writer.book["Sheet"])

  conn.close()
  output.seek(0)

  timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
  safe_filename = f"All_Item_Records_{timestamp_str}.xlsx"
  log_action(
      session["user"],
      "EXPORT EXCEL",
      "Exported item records spreadsheet for all items",
  )

  return send_file(
      output,
      as_attachment=True,
      download_name=safe_filename,
      mimetype=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

@app.route("/api/low-stock-alerts")
@login_required
def low_stock_alerts():
    import json
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    alerts = _get_low_stock_alerts(cursor)
    conn.close()
    return app.response_class(
        response=json.dumps(alerts),
        mimetype="application/json"
    )

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet

@app.route("/registers", methods=["GET", "POST"])
@login_required
def registers():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  reg_type = request.values.get("type", "")
  today_str = datetime.now().strftime("%Y-%m-%d")
  prefill_option = request.values.get("prefill_option", "")
  prefill_item = request.values.get("prefill_item", "")

  if request.method == "POST":
    sub_action = request.form.get("sub_action")

    if sub_action == "save_brass_in":
      batch = session.get("brass_in_batch", [])
      for item in batch:
        cursor.execute(
            "INSERT INTO brass_in (bag_number, date, vendor, weight,"
            " remaining_weight, status) VALUES (?, ?, ?, ?, ?, 'available')",
            (
                item["bag_number"],
                item["date"],
                item["vendor"],
                item["weight"],
                item["weight"],
            ),
        )
      conn.commit()
      log_action(
          session["user"],
          "CREATE BRASS IN SCRAP",
          f"Saved scrap batch of {len(batch)} bags.",
      )
      session.pop("brass_in_batch", None)
      return redirect(url_for("registers", type="Brass In"))

    elif sub_action == "add_bag_brass_in":
      vendor = request.form.get("vendor")
      date = request.form.get("date") or today_str
      bag_number = int(request.form.get("bag_number") or 1)
      weight = float(request.form.get("weight") or 0)
      batch = session.get("brass_in_batch", [])
      batch.append({
          "bag_number": bag_number,
          "date": date,
          "vendor": vendor,
          "weight": weight,
      })
      session["brass_in_batch"] = batch
      return redirect(url_for("registers", type="Brass In"))

    elif sub_action == "save_brass_out":
      date = request.form.get("date") or today_str
      vendor = request.form.get("vendor")
      bag_in_id_raw = request.form.get("bag_in_id")
      bag_in_id = int(bag_in_id_raw) if bag_in_id_raw else 0
      out_weight = float(request.form.get("weight") or 0)
      given_to = request.form.get("given_to", "").strip()

      cursor.execute(
          "SELECT bag_number FROM brass_in WHERE id = ?",
          (bag_in_id,),
      )
      b_row = cursor.fetchone()
      if b_row:
        bag_num = b_row[0]

        cursor.execute(
            "INSERT INTO brass_out (date, vendor, bag_number, weight, brass_in_id, given_to)"
            " VALUES (?, ?, ?, ?, ?, ?)",
            (date, vendor, bag_num, out_weight, bag_in_id, given_to),
        )

        cursor.execute(
            "UPDATE brass_in SET remaining_weight = 0, status = 'out' WHERE id ="
            " ?",
            (bag_in_id,),
        )
        conn.commit()
        log_action(
            session["user"],
            "CREATE BRASS OUT",
            f"Processed scrap bag #{bag_num} out completely with weight"
            f" {out_weight:.3f} KG.",
        )
      return redirect(
          url_for("registers", type="Brass Out", date=date, vendor=vendor)
      )

    elif sub_action == "save_casting_in":
      date = request.form.get("date") or today_str
      vendor = request.form.get("vendor")
      item = request.form.get("item")
      pieces = int(request.form.get("pieces") or 0)
      quantity_unit = request.form.get("quantity_unit", "").strip()
      weight = float(request.form.get("weight") or 0)
      if not vendor or not item:
          flash("Vendor and Item are required.", "error")
          return redirect(url_for("registers", type="Casting In"))
      cursor.execute(
          "INSERT INTO casting_in (date, vendor, item, pieces, weight) VALUES"
          " (?, ?, ?, ?, ?)",
          (date, vendor, item, pieces, weight),
      )
      conn.commit()

    elif sub_action == "save_machining":
      date = request.form.get("date") or today_str
      item = request.form.get("item")
      direction = "Machined"
      pieces = int(request.form.get("pieces") or 0)
      weight = float(request.form.get("weight") or 0)

      if item:
        cursor.execute(
            "SELECT COALESCE(SUM(pieces), 0) FROM casting_in WHERE item = ?",
            (item,),
        )
        total_casting_pieces = cursor.fetchone()[0]

        cursor.execute(
            "SELECT COALESCE(SUM(pieces), 0) FROM machining_register WHERE item = ?",
            (item,),
        )
        mach_total = cursor.fetchone()[0]

        available_for_machining = total_casting_pieces - mach_total

        if pieces > available_for_machining:
          conn.close()
          return redirect(url_for("registers", type="Machining"))

      cursor.execute(
          "INSERT INTO machining_register (date, item, direction, pieces,"
          " weight) VALUES (?, ?, ?, ?, ?)",
          (date, item, direction, pieces, weight),
      )
      conn.commit()
      log_action(
          session["user"],
          "CREATE MACHINING ENTRY",
          f"Added machining entry for item '{item}' ({direction})",
      )
      return redirect(url_for("registers", type="Machining"))

    elif sub_action == "add_plating_item_batch":
      challan_number = request.form.get("challan_number")
      challan_date = request.form.get("date") or today_str
      vendor = request.form.get("vendor")

      session["plating_out_meta"] = {
          "challan_number": challan_number,
          "date": challan_date,
          "vendor": vendor,
      }

      batch = session.get("plating_out_batch", [])
      batch.append({
          "item": request.form.get("item"),
          "pieces": int(request.form.get("pieces") or 0),
          "weight": float(request.form.get("weight") or 0.0),
          "bag_number": int(request.form.get("bag_number") or 1),
          "grinding_vendor": request.form.get("grinding_vendor", ""),
      })
      session["plating_out_batch"] = batch
      return redirect(
          url_for("registers", type="Plating", direction="Outward")
      )

    elif sub_action == "save_outward_challan":
      meta = session.get("plating_out_meta", {})
      challan_no = meta.get("challan_number")
      date = meta.get("date") or today_str
      vendor = meta.get("vendor")
      batch = session.get("plating_out_batch", [])

      if challan_no and batch:
        cursor.execute(
            "INSERT INTO plating_challans (challan_number, date, vendor,"
            " direction, status) VALUES (?, ?, ?, 'Outward', 'open')",
            (challan_no, date, vendor),
        )
        challan_id = cursor.lastrowid
        for item in batch:
          cursor.execute(
              "INSERT INTO plating_items (challan_id, item, pieces, weight,"
              " bag_number, remaining_pieces, grinding_vendor) VALUES (?, ?, ?, ?, ?, ?, ?)",
              (
                  challan_id,
                  item["item"],
                  item["pieces"],
                  item["weight"],
                  item["bag_number"],
                  item["pieces"],
                  item.get("grinding_vendor", ""),
              ),
          )
        conn.commit()
        log_action(
            session["user"],
            "CREATE PLATING OUTWARD",
            f"Created plating challan #{challan_no} for vendor '{vendor}'",
        )
        session.pop("plating_out_batch", None)
        session.pop("plating_out_meta", None)
      return redirect(url_for("registers", type="Plating", direction="Outward"))

    elif sub_action == "save_inward":
      date = request.form.get("date") or today_str
      vendor = request.form.get("vendor")
      item = request.form.get("item")
      in_pieces = int(request.form.get("pieces") or 0)
      in_weight = float(request.form.get("weight") or 0.0)

      if vendor and item and (in_pieces > 0 or in_weight > 0):
        cursor.execute(
            "INSERT INTO plating_challans (challan_number, date, vendor,"
            " direction, status) VALUES (NULL, ?, ?, 'Inward', 'received')",
            (date, vendor),
        )
        challan_id = cursor.lastrowid
        cursor.execute(
            "INSERT INTO plating_items (challan_id, item, pieces, weight,"
            " bag_number, remaining_pieces, grinding_vendor) VALUES (?, ?, ?,"
            " ?, NULL, 0, NULL)",
            (challan_id, item, in_pieces, in_weight),
        )
        conn.commit()
        log_action(
            session["user"],
            "CREATE PLATING INWARD",
            f"Received {in_pieces} pcs / {in_weight} KG of '{item}' from"
            f" plating vendor '{vendor}'",
        )
      return redirect(url_for("registers", type="Plating", direction="Inward"))

    elif sub_action == "save_miscellaneous":
        date = request.form.get("date") or today_str
        option_type = request.form.get("option_type")
        vendor = (
            request.form.get("vendor") if option_type == "purchase" else ""
        )
        item = request.form.get("item")
        pieces = int(request.form.get("pieces") or 0)
        weight = float(request.form.get("weight") or 0.0)
        quantity_unit = request.form.get("quantity_unit", "").strip()

        cursor.execute(
            "INSERT INTO miscellaneous_entries (date, option_type, vendor, item,"
            " pieces, weight, quantity_unit) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (date, option_type, vendor, item, pieces, weight, quantity_unit),
        )
        conn.commit()
        log_action(
            session["user"],
            "CREATE MISCELLANEOUS ENTRY",
            f"Added miscellaneous entry for item '{item}' ({option_type})",
        )
        return redirect(url_for("registers", type="Miscellaneous"))
    elif sub_action == "save_rejection":
      date = request.form.get("date") or today_str
      sub_rej_type = request.form.get("sub_rej_type", "inventory")
      vendor = request.form.get("vendor", "")
      weight = float(request.form.get("weight") or 0.0)

      if sub_rej_type == "brass_in_vendor":
        # Rejection against a Brass In (scrap supplier) vendor: physically
        # deduct the rejected weight from that vendor's available scrap bags,
        # oldest bag first (FIFO), consuming bags fully before moving on.
        item = ""
        rejection_type = "Brass In Vendor Rejection"
        pieces = 0

        cursor.execute(
            "INSERT INTO rejection_register (date, item, rejection_type,"
            " vendor, pieces, weight) VALUES (?, ?, ?, ?, ?, ?)",
            (date, item, rejection_type, vendor, pieces, weight),
        )

        cursor.execute(
            "SELECT id, remaining_weight FROM brass_in WHERE vendor = ? AND"
            " status = 'available' AND remaining_weight > 0 ORDER BY date"
            " ASC, id ASC",
            (vendor,),
        )
        bags = cursor.fetchall()
        remaining_to_deduct = weight
        for bag_id, bag_remaining in bags:
          if remaining_to_deduct <= 0:
            break
          deduct_now = min(bag_remaining, remaining_to_deduct)
          new_remaining = round(bag_remaining - deduct_now, 3)
          new_status = "out" if new_remaining <= 0.0001 else "available"
          cursor.execute(
              "UPDATE brass_in SET remaining_weight = ?, status = ? WHERE id"
              " = ?",
              (max(0.0, new_remaining), new_status, bag_id),
          )
          remaining_to_deduct -= deduct_now
        conn.commit()

        detail = (
            f"Brass In vendor rejection of {weight:.3f} KG deducted from"
            f" vendor '{vendor}' scrap stock."
        )
        if remaining_to_deduct > 0.0001:
          detail += (
              f" WARNING: only {(weight - remaining_to_deduct):.3f} KG of"
              " available stock could be deducted; vendor stock fell short"
              f" by {remaining_to_deduct:.3f} KG."
          )
        log_action(session["user"], "CREATE REJECTION ENTRY", detail)

      elif sub_rej_type == "casting_vendor":
        # Rejection against a Casting In / Brass Out vendor: subtracted from
        # that vendor's casting return in the vendor ledger calculations
        # (stored as 'Vendor Rejection' for backward compatibility with
        # existing ledger queries).
        item = ""
        rejection_type = "Vendor Rejection"
        pieces = 0
        cursor.execute(
            "INSERT INTO rejection_register (date, item, rejection_type,"
            " vendor, pieces, weight) VALUES (?, ?, ?, ?, ?, ?)",
            (date, item, rejection_type, vendor, pieces, weight),
        )
        conn.commit()
        log_action(
            session["user"],
            "CREATE REJECTION ENTRY",
            f"Casting/Brass Out vendor rejection of {weight:.3f} KG recorded"
            f" for vendor '{vendor}'.",
        )

      elif sub_rej_type == "misc_vendor":
        # Rejection against a Miscellaneous item vendor: deducted from the
        # purchased Miscellaneous stock of that item.
        item = request.form.get("item")
        rejection_type = "Misc Vendor Rejection"
        pieces = int(request.form.get("pieces") or 0)
        cursor.execute(
            "INSERT INTO rejection_register (date, item, rejection_type,"
            " vendor, pieces, weight) VALUES (?, ?, ?, ?, ?, ?)",
            (date, item, rejection_type, vendor, pieces, weight),
        )
        conn.commit()
        log_action(
            session["user"],
            "CREATE REJECTION ENTRY",
            f"Misc vendor rejection: item '{item}' qty {pieces} pcs /"
            f" {weight:.3f} KG deducted from purchased stock, vendor"
            f" '{vendor}'.",
        )

      else:
        item = request.form.get("item")
        rejection_type = request.form.get("rejection_type")
        pieces = int(request.form.get("pieces") or 0)
        cursor.execute(
            "INSERT INTO rejection_register (date, item, rejection_type,"
            " vendor, pieces, weight) VALUES (?, ?, ?, ?, ?, ?)",
            (date, item, rejection_type, vendor, pieces, weight),
        )
        conn.commit()
        log_action(
            session["user"],
            "CREATE REJECTION ENTRY",
            f"Added rejection entry for vendor '{vendor}' ({sub_rej_type})",
        )

      return redirect(
          url_for("registers", type="Rejection", sub_rej_type=sub_rej_type)
      )

    elif sub_action == "save_finished_goods":
        date = request.form.get("date") or today_str
        box_pack_id_raw = request.form.get("box_pack_id")
        box_pack_id = int(box_pack_id_raw) if box_pack_id_raw else 0
        pcs = int(request.form.get("pcs") or 0)

        cursor.execute("SELECT box_pack_name FROM box_pack_items WHERE id = ?", (box_pack_id,))
        bp_row = cursor.fetchone()

        if bp_row and box_pack_id and pcs > 0:
            box_pack_name = bp_row[0]

            # 1. Upsert into finished_goods (add to existing pcs)
            cursor.execute("SELECT id, pcs FROM finished_goods WHERE box_pack_id = ?", (box_pack_id,))
            fg_row = cursor.fetchone()
            if fg_row:
                cursor.execute(
                    "UPDATE finished_goods SET pcs = pcs + ? WHERE box_pack_id = ?",
                    (pcs, box_pack_id)
                )
            else:
                cursor.execute(
                    "INSERT INTO finished_goods (box_pack_id, box_pack_name, pcs) VALUES (?, ?, ?)",
                    (box_pack_id, box_pack_name, pcs)
                )

            # 2. Log this addition
            cursor.execute(
                "INSERT INTO finished_goods_log (date, box_pack_id, box_pack_name, pcs) VALUES (?, ?, ?, ?)",
                (date, box_pack_id, box_pack_name, pcs)
            )
            log_id = cursor.lastrowid

            # 3. Deduct component items from assembly
            cursor.execute(
                """SELECT m.id, m.item_name FROM master_items m
                   JOIN box_pack_components bc ON m.id = bc.item_id
                   WHERE bc.box_pack_id = ?""",
                (box_pack_id,)
            )
            component_rows = cursor.fetchall()
            for comp_item_id, comp_item_name in component_rows:
                cursor.execute(
                    "INSERT INTO finished_goods_assembly_deductions"
                    " (finished_goods_log_id, item_id, item_name, pcs)"
                    " VALUES (?, ?, ?, ?)",
                    (log_id, comp_item_id, comp_item_name, pcs)
                )

            conn.commit()
            log_action(
                session["user"],
                "CREATE FINISHED GOODS ENTRY",
                f"Added {pcs} pcs of '{box_pack_name}' to Finished Goods. "
                f"Deducted {pcs} pcs from {len(component_rows)} assembly component(s)."
            )
        return redirect(url_for("registers", type="Finished Goods"))
    
    elif sub_action == "save_sale_register":
      date = request.form.get("date") or today_str
      party = request.form.get("party", "").strip()
      bill_no = request.form.get("bill_no", "").strip()
      box_pack_id_raw = request.form.get("box_pack_id")
      box_pack_id = int(box_pack_id_raw) if box_pack_id_raw else 0
      pcs = int(request.form.get("pcs") or 0)
      no_of_cartons = int(request.form.get("no_of_cartons") or 0)

      cursor.execute(
          "SELECT box_pack_name FROM box_pack_items WHERE id = ?",
          (box_pack_id,),
      )
      bp_row = cursor.fetchone()

      if bp_row and party and bill_no and box_pack_id and pcs > 0:
        box_pack_name = bp_row[0]

        cursor.execute(
            "INSERT INTO sale_register (date, party, bill_no, box_pack_id,"
            " box_pack_name, pcs, no_of_cartons) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (
                date,
                party,
                bill_no,
                box_pack_id,
                box_pack_name,
                pcs,
                no_of_cartons,
            ),
        )
        sale_id = cursor.lastrowid

        for i in range(1, no_of_cartons + 1):
          carton_weight = float(request.form.get(f"carton_weight_{i}") or 0.0)
          cursor.execute(
              "INSERT INTO sale_register_cartons (sale_id, carton_number,"
              " weight) VALUES (?, ?, ?)",
              (sale_id, i, carton_weight),
          )

        # Deduct from finished goods stock directly
                    # 3. Deduct component items from assembly
        cursor.execute(
            """SELECT m.id, m.item_name FROM master_items m
                JOIN box_pack_components bc ON m.id = bc.item_id
                WHERE bc.box_pack_id = ?""",
            (box_pack_id,)
        )
        component_rows = cursor.fetchall()
        for comp_item_id, comp_item_name in component_rows:
            cursor.execute(
                "INSERT INTO finished_goods_assembly_deductions"
                " (finished_goods_log_id, item_id, item_name, pcs)"
                " VALUES (?, ?, ?, ?)",
                (log_id, comp_item_id, comp_item_name, pcs)
            )

        conn.commit()
        log_action(
            session["user"],
            "CREATE FINISHED GOODS ENTRY",
            f"Added {pcs} pcs of '{box_pack_name}' to Finished Goods. "
            f"Deducted {pcs} pcs from {len(component_rows)} assembly component(s)."
        )
        return redirect(url_for("registers", type="Sale Register"))
    elif sub_action == "save_opening_stock":
      date = request.form.get("date") or today_str
      item = request.form.get("item")
      stage = request.form.get("stage")
      pieces = int(request.form.get("pieces") or 0)
      weight = float(request.form.get("weight") or 0.0)

      cursor.execute(
          """INSERT INTO opening_stock (item, stage, pieces, weight, date)
             VALUES (?, ?, ?, ?, ?)
             ON CONFLICT(item, stage) DO UPDATE SET
                 pieces = excluded.pieces,
                 weight = excluded.weight,
                 date = excluded.date""",
          (item, stage, pieces, weight, date),
      )
      conn.commit()
      log_action(
          session["user"],
          "SET OPENING STOCK",
          f"Set opening stock for '{item}' ({stage}) to {pieces} pcs /"
          f" {weight:.3f} KG as of {date}",
      )
      return redirect(url_for("registers", type="Opening Stock"))

    elif sub_action == "edit_opening_stock":
        os_id = request.form.get("os_id")
        date = request.form.get("date")
        item = request.form.get("item")
        stage = request.form.get("stage")
        pieces = request.form.get("pieces") or 0
        weight = request.form.get("weight") or 0
        with sqlite3.connect(DB_NAME) as conn:
            conn.execute(
                "UPDATE opening_stock SET date=?, item=?, stage=?, pieces=?, weight=? WHERE id=?",
                (date, item, stage, float(pieces), float(weight), int(os_id))
            )
            conn.commit()
        log_action(session["user"], "EDIT OPENING STOCK", f"Updated id={os_id} item={item} stage={stage}")
        return redirect(url_for("registers", type="Opening Stock"))

    elif sub_action == "save_opening_scrap":
      date = request.form.get("date") or today_str
      bag_number = int(request.form.get("bag_number") or 1)
      weight = float(request.form.get("weight") or 0.0)

      cursor.execute(
          "INSERT INTO brass_in (bag_number, date, vendor, weight,"
          " remaining_weight, status) VALUES (?, ?, 'Opening Balance', ?, ?,"
          " 'available')",
          (bag_number, date, weight, weight),
      )
      conn.commit()
      log_action(
          session["user"],
          "SET OPENING SCRAP STOCK",
          f"Added opening scrap balance of {weight:.3f} KG (Bag"
          f" #{bag_number}) as of {date}",
      )
      return redirect(url_for("registers", type="Opening Stock"))

    elif sub_action == "toggle_payment":
      # Generic "Payment Made" tick used across every Vendor Ledger entry
      # (Brass In, Brass Out, Casting In, Miscellaneous, Grinding/Plating
      # entries which live in plating_items).
      pm_table = request.form.get("table", "")
      row_id = int(request.form.get("row_id") or 0)
      ledger_sub_type = request.form.get("ledger_sub_type", "brass_out_casting_in")
      ledger_vendor = request.form.get("ledger_vendor", "")
      allowed_payment_tables = {
          "brass_in",
          "brass_out",
          "casting_in",
          "miscellaneous_entries",
          "plating_items",
      }
      if pm_table in allowed_payment_tables and row_id:
        cursor.execute(
            f"SELECT payment_made FROM {pm_table} WHERE id = ?", (row_id,)
        )
        row = cursor.fetchone()
        if row is not None:
          new_val = 0 if (row[0] or 0) else 1
          cursor.execute(
              f"UPDATE {pm_table} SET payment_made = ? WHERE id = ?",
              (new_val, row_id),
          )
          conn.commit()
          log_action(
              session["user"],
              "TOGGLE PAYMENT MADE",
              f"Set payment_made={new_val} on {pm_table} row #{row_id}",
          )
      return redirect(
          url_for(
              "registers",
              type="Vendor Ledgers",
              ledger_sub_type=ledger_sub_type,
              vendor=ledger_vendor,
          )
      )

  if reg_type and reg_type not in [
      "Vendor Ledgers",
      "Inventory",
      "Item Records",
  ]:
    cursor.execute(
        """SELECT v.vendor_name FROM vendors v 
           JOIN vendor_registers vr ON v.id = vr.vendor_id 
           WHERE vr.register_type = ?""",
        (reg_type,),
    )
    vendors = [v[0] for v in cursor.fetchall()]
  else:
    cursor.execute("SELECT vendor_name FROM vendors")
    vendors = [v[0] for v in cursor.fetchall()]

  cursor.execute("SELECT item_name FROM master_items")
  master_items = [m[0] for m in cursor.fetchall()]

  content_html = ""

  if reg_type == "Brass In":
    selected_vendor = request.values.get("vendor", "")
    prefill_option = request.values.get("prefill_option", "")
    prefill_item   = request.values.get("prefill_item", "")

    vendor_options_html = "".join([
        f'<option value="{v}" {"selected" if v == selected_vendor else ""}>{v}</option>'
        for v in vendors
    ])

    batch = session.get("brass_in_batch", [])
    next_bag_num = len(batch) + 1
    batch_rows = "".join([
        f"<tr><td>{b['date']}</td><td>{b['vendor']}</td><td>{b['bag_number']}</td><td>{b['weight']:.3f} KG</td></tr>"
        for b in batch
    ])

    content_html = f"""
        <h3>Brass In (Scrap) Register</h3>
        <div class="card">
            <form method="POST">
                <input type="hidden" name="sub_action" value="add_bag_brass_in">
                <label><b>Date:</b></label><input type="date" name="date" value="{today_str}" required>
                <label><b>Vendor:</b></label><select name="vendor" required><option value="">-- Select Vendor --</option>{vendor_options_html}</select>
                <label><b>Bag Number:</b></label><input type="number" name="bag_number" value="{next_bag_num}" required>
                <label><b>Scrap Weight (KG):</b></label><input type="number" step="0.001" name="weight" required>
                <button type="submit">Add Bag to Batch</button>
            </form>
        </div>
        <h4>Current Unsaved Scrap Batch</h4>
        <table><tr><th>Date</th><th>Vendor</th><th>Bag Number</th><th>Scrap Weight</th></tr>{batch_rows if batch_rows else '<tr><td colspan="4">No bags.</td></tr>'}</table>
        {f'<form method="POST"><input type="hidden" name="sub_action" value="save_brass_in"><button type="submit" style="background:#28a745;">Save Batch</button></form>' if batch else ''}
        <hr style="margin: 30px 0;">
        
        <div class="card">
            <form method="GET" action="/registers" id="brassInFilterForm">
                <input type="hidden" name="type" value="Brass In">
                <label><b>Filter Real-Time Ledger by Vendor:</b></label>
                <select name="vendor" onchange="document.getElementById('brassInFilterForm').submit()">
                    <option value="">-- Select Vendor to View Real-Time Ledger --</option>
                    {vendor_options_html}
                </select>
            </form>
        </div>
        """

    if selected_vendor:
      cursor.execute(
          "SELECT id, date, bag_number, weight, remaining_weight, status FROM"
          " brass_in WHERE vendor = ? ORDER BY date DESC, id DESC",
          (selected_vendor,),
      )
      v_in_rows = cursor.fetchall()
      total_in_weight = sum([r[3] for r in v_in_rows])
      v_in_html = "".join([
          f"<tr><td>{r[1]}</td><td>Bag #{r[2]}</td><td>{r[3]:.3f} KG</td><td>{r[4]:.3f} KG</td><td>{r[5].capitalize()}</td>"
          f"<td>{admin_delete_cell(f'/delete-brass-in/{r[0]}?vendor={selected_vendor}')}</td></tr>"
          for r in v_in_rows
      ])
      content_html += f"""
            <div class="card" style="border: 2px solid #007bff;">
                <h4 style="color:#007bff; margin-top:0;">Real-Time Ledger for Vendor: {selected_vendor}</h4>
                <table>
                    <tr><th>Date</th><th>Bag Number</th><th>Original Scrap Weight</th><th>Remaining Scrap Weight</th><th>Status</th><th>Action</th></tr>
                    {v_in_html if v_in_html else '<tr><td colspan="6">No records for this vendor.</td></tr>'}
                    <tr style="background:#d4edda; font-weight:bold;">
                        <td colspan="2">Total Scrap Supplied In:</td>
                        <td colspan="4">{total_in_weight:.3f} KG</td>
                    </tr>
                </table>
            </div>
            """
    else:
      cursor.execute(
          "SELECT id, date, vendor, bag_number, weight, remaining_weight FROM"
          " brass_in WHERE status = 'available' AND remaining_weight > 0"
      )
      avail_rows = cursor.fetchall()
      avail_html = "".join([
          f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]:.3f} KG</td><td><b>{r[5]:.3f} KG</b></td>"
          f"<td>{admin_delete_cell(f'/delete-brass-in/{r[0]}')}</td></tr>"
          for r in avail_rows
      ])
      content_html += f"""
            <h4>Real-Time Available Scrap Bags (All Vendors)</h4>
            <table><tr><th>Date</th><th>Vendor</th><th>Bag Number</th><th>Original Scrap</th><th>Available Scrap</th><th>Action</th></tr>{avail_html if avail_html else '<tr><td colspan="6">No stock.</td></tr>'}</table>
            """

  elif reg_type == "Brass Out":
    selected_vendor = request.values.get("vendor", "")
    selected_bag_id = request.values.get("bag_id", "")
    selected_date = request.values.get("date", today_str)

    cursor.execute(
        "SELECT id, bag_number, date, remaining_weight, vendor FROM brass_in"
        " WHERE status = 'available' AND remaining_weight > 0"
    )
    avail_bags = cursor.fetchall()

    bag_options = "".join([
        f'<option value="{b[0]}" data-vendor="{b[4]}" data-weight="{b[3]}" {"selected" if str(b[0]) == selected_bag_id else ""}>Bag #{b[1]} | Vendor: {b[4]} | Available Scrap: {b[3]:.3f} KG</option>'
        for b in avail_bags
    ])

    selected_bag_data = None
    if selected_bag_id:
      cursor.execute(
          "SELECT id, bag_number, weight, vendor FROM brass_in WHERE id = ?",
          (selected_bag_id,),
      )
      selected_bag_data = cursor.fetchone()

    vendor_summary_html = ""
    if selected_vendor:
      cursor.execute(
          "SELECT SUM(weight) FROM brass_out WHERE vendor = ?",
          (selected_vendor,),
      )
      t_out = cursor.fetchone()[0] or 0.0

      cursor.execute(
          "SELECT SUM(weight) FROM casting_in WHERE vendor = ?",
          (selected_vendor,),
      )
      t_casting = cursor.fetchone()[0] or 0.0

      cursor.execute(
          "SELECT SUM(weight) FROM rejection_register WHERE vendor = ? AND"
          " rejection_type = 'Vendor Rejection'",
          (selected_vendor,),
      )
      t_rejection = cursor.fetchone()[0] or 0.0

      adjusted_casting_weight = max(0.0, t_casting - t_rejection)
      burn_loss = adjusted_casting_weight * 0.05
      total_ret = adjusted_casting_weight + burn_loss
      net_bal = t_out - total_ret

      if net_bal > 0:
        bal_str = (
            f"<span style='color:red;'><b>+{net_bal:.3f} KG</b> (Balance Due"
            " From Vendor)</span>"
        )
      elif net_bal < 0:
        bal_str = (
            f"<span style='color:green;'><b>{net_bal:.3f} KG</b> (Advance"
            " With Vendor)</span>"
        )
      else:
        bal_str = "<span style='color:blue;'><b>0.000 KG</b> (Fully Settled)</span>"

      vendor_summary_html = f"""
            <div class="info-box">
                <b>Real-Time Account Summary for Vendor: {selected_vendor}</b><br>
                • Total Scrap Taken Out: <b>{t_out:.3f} KG</b><br>
                • Total Casting In: <b>{t_casting:.3f} KG</b><br>
                • Less: Vendor Rejection: <b>-{t_rejection:.3f} KG</b><br>
                • Adjusted Casting Weight: <b>{adjusted_casting_weight:.3f} KG</b><br>
                • Burning Loss (5%): <b>{burn_loss:.3f} KG</b><br>
                • Advance / Balance Scrap: {bal_str}
            </div>
            """

    vendor_options_html = "".join([
        f'<option value="{v}" {"selected" if v == selected_vendor else ""}>{v}</option>'
        for v in vendors
    ])

    default_weight_val = (
        f"{selected_bag_data[2]:.3f}" if selected_bag_data else ""
    )

    content_html = f"""
        <h3>Brass Out Register</h3>
        <div class="card">
            <form method="POST" id="brassOutForm">
                <input type="hidden" name="sub_action" value="save_brass_out">
                <input type="hidden" name="bag_in_id" id="bagInIdInput" value="{selected_bag_id}">
                
                <label><b>Date:</b></label>
                <input type="date" name="date" value="{selected_date}" required>
                
                <label><b>Vendor:</b></label>
                <select name="vendor" id="vendorSelect" onchange="window.location.href='/registers?type=Brass Out&vendor=' + encodeURIComponent(this.value);" required>
                    <option value="">-- Choose Vendor --</option>
                    {vendor_options_html}
                </select>

                <label><b>Select Available Scrap Bag:</b></label>
                <select name="bag_id_display" id="bagSelect" onchange="onBagChange(this)" required>
                    <option value="">-- Choose Bag --</option>
                    {bag_options}
                </select>

                <label><b>Weight (KG) [Fully Editable for Outbound Vendor]:</b></label>
                <input type="number" step="0.001" name="weight" id="weightInput" value="{default_weight_val}" required>
                <label><b>Given To (Casting Vendor):</b></label>
                <input type="text" name="given_to" placeholder="Vendor name" required>

                <button type="submit" style="background:#28a745;">Confirm & Save</button>
            </form>
        </div>
        """

    if selected_vendor:
      content_html += f"""
        <div class="card">
            {vendor_summary_html}
        </div>
        """

    content_html += """
        <script>
        function onBagChange(selectElement) {
            var selectedOption = selectElement.options[selectElement.selectedIndex];
            var weight = selectedOption.getAttribute('data-weight');
            var bagId = selectElement.value;
            
            var weightInput = document.getElementById('weightInput');
            if (weight) {
                weightInput.value = weight;
            }
            
            var bagInIdInput = document.getElementById('bagInIdInput');
            if (bagInIdInput) {
                bagInIdInput.value = bagId;
            }
        }
        </script>
        """

    cursor.execute(
        "SELECT id, date, vendor, bag_number, weight FROM brass_out ORDER BY"
        " id DESC"
    )
    out_rows = cursor.fetchall()
    out_html = "".join([
        f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>Bag #{r[3]}</td><td>{r[4]:.3f} KG</td>"
        f"<td>{admin_delete_cell(f'/delete-brass-out/{r[0]}?vendor={selected_vendor}', 'Delete this Brass Out entry? The linked scrap bag will be restored to available.')}</td></tr>"
        for r in out_rows
    ])
    content_html += f"""
        <h4>All Saved Brass Out Records</h4>
        <table><tr><th>Date</th><th>Vendor</th><th>Bag Number</th><th>Weight Out</th><th>Action</th></tr>{out_html if out_html else '<tr><td colspan="5">No records.</td></tr>'}</table>
        """

  elif reg_type == "Casting In":
    selected_vendor = request.values.get("vendor", "")
    selected_date = request.values.get("date", today_str)

    vendor_summary_html = ""
    if selected_vendor:
      cursor.execute(
          "SELECT SUM(weight) FROM brass_out WHERE vendor = ?",
          (selected_vendor,),
      )
      t_out = cursor.fetchone()[0] or 0.0

      cursor.execute(
          "SELECT SUM(weight) FROM casting_in WHERE vendor = ?",
          (selected_vendor,),
      )
      t_casting = cursor.fetchone()[0] or 0.0

      cursor.execute(
          "SELECT SUM(weight) FROM rejection_register WHERE vendor = ? AND"
          " rejection_type = 'Vendor Rejection'",
          (selected_vendor,),
      )
      t_rejection = cursor.fetchone()[0] or 0.0

      adjusted_casting_weight = max(0.0, t_casting - t_rejection)
      burn_loss = adjusted_casting_weight * 0.05
      total_ret = adjusted_casting_weight + burn_loss
      net_bal = t_out - total_ret

      if net_bal > 0:
        bal_str = (
            f"<span style='color:red;'><b>+{net_bal:.3f} KG</b> (Balance Due"
            " From Vendor)</span>"
        )
      elif net_bal < 0:
        bal_str = (
            f"<span style='color:green;'><b>{net_bal:.3f} KG</b> (Advance"
            " With Vendor)</span>"
        )
      else:
        bal_str = "<span style='color:blue;'><b>0.000 KG</b> (Fully Settled)</span>"

      vendor_summary_html = f"""
            <div class="info-box">
                <b>Real-Time Account Summary for Vendor: {selected_vendor}</b><br>
                • Total Scrap Taken Out: <b>{t_out:.3f} KG</b><br>
                • Total Casting In: <b>{t_casting:.3f} KG</b><br>
                • Less: Vendor Rejection: <b>-{t_rejection:.3f} KG</b><br>
                • Adjusted Casting Weight: <b>{adjusted_casting_weight:.3f} KG</b><br>
                • Burning Loss (5%): <b>{burn_loss:.3f} KG</b><br>
                • Advance / Balance Scrap: {bal_str}
            </div>
            """

    cursor.execute(
        """SELECT DISTINCT v.vendor_name FROM vendors v
           JOIN vendor_registers vr ON v.id = vr.vendor_id
           WHERE vr.register_type = 'Casting In'"""
    )
    casting_vendors = [v[0] for v in cursor.fetchall()]
    vendor_options_html = "".join([
        f'<option value="{v}" {"selected" if v == selected_vendor else ""}>{v}</option>'
        for v in casting_vendors
    ])
    if selected_vendor:
        cursor.execute(
            "SELECT DISTINCT item FROM casting_in WHERE vendor=? ORDER BY item",
            (selected_vendor,)
        )
        vendor_items = [r[0] for r in cursor.fetchall()]
    else:
        vendor_items = master_items
    
    item_options = "".join(
        [f'<option value="{i}">{i}</option>' for i in vendor_items]
    )

    content_html = f"""
        <h3>Casting In Register</h3>
        <div class="card">
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_casting_in">
                <input type="hidden" name="type" value="Casting In">
                <label><b>Date:</b></label><input type="date" name="date" value="{selected_date}" required>
                <label><b>Vendor:</b></label>
                <select name="vendor" id="vendorSelect" onchange="window.location.href='/registers?type=Casting In&vendor=' + encodeURIComponent(this.value);" required>
                    <option value="">-- Select Vendor --</option>
                    {vendor_options_html}
                </select>
                <label><b>Item:</b></label><select name="item" required><option value="">-- Select --</option>{item_options}</select>
                <label><b>Pieces:</b></label><input type="number" name="pieces" required>
                <label><b>Weight (KG):</b></label><input type="number" step="0.001" name="weight" required>
                <button type="submit" style="background:#28a745;">Save Casting In</button>
            </form>
        </div>
        """

    if selected_vendor:
      content_html += f"""
        <div class="card">
            {vendor_summary_html}
        </div>
        """

    if selected_vendor:
      cursor.execute(
          "SELECT id, date, item, pieces, weight FROM casting_in WHERE vendor = ?"
          " ORDER BY date DESC, id DESC",
          (selected_vendor,),
      )
      c_rows = cursor.fetchall()
      total_c_weight = sum([r[4] for r in c_rows])
      total_c_pieces = sum([r[3] for r in c_rows])
      c_html = "".join([
          f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]:.3f} KG</td>"
          f"<td>{admin_delete_cell(f'/delete-casting-in/{r[0]}?vendor={selected_vendor}')}</td></tr>"
          for r in c_rows
      ])
      content_html += f"""
            <div class="card" style="border: 2px solid #28a745;">
                <h4 style="color:#28a745; margin-top:0;">Real-Time Ledger for Vendor: {selected_vendor}</h4>
                <table>
                    <tr><th>Date</th><th>Item</th><th>Pieces</th><th>Weight</th><th>Action</th></tr>
                    {c_html if c_html else '<tr><td colspan="5">No records for this vendor.</td></tr>'}
                    <tr style="background:#d4edda; font-weight:bold;">
                        <td colspan="2">Total Casting In:</td>
                        <td>{total_c_pieces} Pcs</td>
                        <td>{total_c_weight:.3f} KG</td>
                        <td></td>
                    </tr>
                </table>
            </div>
            """
    else:
      cursor.execute(
          "SELECT id, date, vendor, item, pieces, weight FROM casting_in ORDER"
          " BY id DESC"
      )
      c_rows = cursor.fetchall()
      c_html = "".join([
          f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]}</td><td>{r[5]:.3f} KG</td>"
          f"<td>{admin_delete_cell(f'/delete-casting-in/{r[0]}')}</td></tr>"
          for r in c_rows
      ])
      content_html += f"""
        <h4>All Saved Casting In Records</h4>
        <table><tr><th>Date</th><th>Vendor</th><th>Item</th><th>Pieces</th><th>Weight</th><th>Action</th></tr>{c_html if c_html else '<tr><td colspan="6">No records.</td></tr>'}</table>
        """

  elif reg_type == "Machining":
    cursor.execute("SELECT DISTINCT item FROM casting_in ORDER BY item ASC")
    casting_items = set(row[0] for row in cursor.fetchall())

    cursor.execute("SELECT DISTINCT item FROM opening_stock WHERE stage='casted' ORDER BY item ASC")
    op_items = set(row[0] for row in cursor.fetchall())

    all_items = sorted(casting_items | op_items)

    machining_item_options = "".join(
        [f'<option value="{i}">{i}</option>' for i in all_items]
    )

    item_limits_data = {}
    for item_name in all_items:
        cursor.execute("SELECT COALESCE(SUM(pieces), 0) FROM casting_in WHERE item = ?", (item_name,))
        tot_cast = cursor.fetchone()[0]

        cursor.execute("SELECT COALESCE(SUM(pieces), 0) FROM machining_register WHERE item = ?", (item_name,))
        mach_total = cursor.fetchone()[0]

        cursor.execute("SELECT COALESCE(pieces, 0) FROM opening_stock WHERE item = ? AND stage='casted'", (item_name,))
        op_p = cursor.fetchone()[0]

        item_limits_data[item_name] = {
            "casting_avail": tot_cast - mach_total + op_p
        }

    content_html = f"""
        <h3>Machining Register</h3>
        <div class="card">
            <form method="POST" action="/registers" onsubmit="return validateMachiningForm()">
                <input type="hidden" name="sub_action" value="save_machining">
                <input type="hidden" name="type" value="Machining">
                
                <label><b>1. Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>
                
                <label><b>2. Item (Received from Casting In):</b></label>
                <select name="item" id="machiningItemSelect" onchange="updateAvailabilityInfo()" required>
                    <option value="">-- Select Casting Item --</option>
                    {machining_item_options}
                </select>
                
                <div id="availabilityNote" class="info-box" style="display:none; margin-top:5px; margin-bottom:15px;"></div>
                
                <label><b>3. Direction:</b></label>
                <input type="text" name="direction" value="Machined" readonly style="background-color: #e9ecef;">
                
                <label><b>4. Pieces:</b></label>
                <input type="number" name="pieces" id="piecesInput" placeholder="Enter pieces (optional if weight entered)" oninput="checkPiecesLimit()">
                
                <label><b>5. Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" id="weightInput" placeholder="Enter weight (optional if pieces entered)">
                
                <div id="errorMsg" style="color: red; font-weight: bold; margin-bottom: 15px; display: none;">
                    Please enter either Pieces or Weight (at least one is required).
                </div>
                <div id="limitErrorMsg" style="color: red; font-weight: bold; margin-bottom: 15px; display: none;">
                    Entered pieces exceed available stock limits for this action!
                </div>
                
                <button type="submit" style="background:#28a745;" id="saveMachiningBtn">Save Machining Entry</button>
            </form>
        </div>

        <script>
        var itemLimits = {item_limits_data};

        function updateAvailabilityInfo() {{
            var itemSelect = document.getElementById('machiningItemSelect');
            var noteBox = document.getElementById('availabilityNote');
            
            var selectedItem = itemSelect.value;
            
            if (selectedItem && itemLimits[selectedItem]) {{
                noteBox.style.display = 'block';
                var castAvail = itemLimits[selectedItem].casting_avail;
                noteBox.innerHTML = '<b>Stock Status for ' + selectedItem + ':</b> Available in Casting (for Machining): ' + castAvail + ' pcs';
            }} else {{
                noteBox.style.display = 'none';
            }}
            checkPiecesLimit();
        }}

        function checkPiecesLimit() {{
            var itemSelect = document.getElementById('machiningItemSelect');
            var piecesInput = document.getElementById('piecesInput');
            var limitErrorMsg = document.getElementById('limitErrorMsg');
            var saveBtn = document.getElementById('saveMachiningBtn');
            
            var selectedItem = itemSelect.value;
            var enteredPieces = parseInt(piecesInput.value) || 0;
            
            if (selectedItem && itemLimits[selectedItem]) {{
                if (enteredPieces > itemLimits[selectedItem].casting_avail) {{
                    limitErrorMsg.style.display = 'block';
                    limitErrorMsg.innerHTML = 'Error: Cannot machine ' + enteredPieces + ' pcs. Only ' + itemLimits[selectedItem].casting_avail + ' pcs available in Casting In!';
                    saveBtn.disabled = true;
                    saveBtn.style.background = '#ccc';
                    return;
                }}
            }}
            limitErrorMsg.style.display = 'none';
            saveBtn.disabled = false;
            saveBtn.style.background = '#28a745';
        }}

        function validateMachiningForm() {{
            var pieces = document.getElementById('piecesInput').value.trim();
            var weight = document.getElementById('weightInput').value.trim();
            var errorMsg = document.getElementById('errorMsg');
            
            if (pieces === "" && weight === "") {{
                errorMsg.style.display = 'block';
                return false;
            }}
            errorMsg.style.display = 'none';
            return true;
        }}
        </script>
        """

    cursor.execute(
        "SELECT id, date, item, direction, pieces, weight FROM"
        " machining_register ORDER BY id DESC"
    )
    mach_rows = cursor.fetchall()
    mach_html = ""
    for r in mach_rows:
      pcs_val = "-"
      try:
        if r[5] is not None and str(r[5]).strip() != "":
            val_int = int(float(str(r[5])))
            q_unit = r[7] if len(r) > 7 else ""
            if val_int > 0:
                pcs_val = f"{val_int} {q_unit}".strip()
      except (ValueError, TypeError):
        pass

      wt_val = "-"
      try:
        if r[5] is not None and str(r[5]).strip() != "":
          val_float = float(r[5])
          if val_float > 0:
            wt_val = f"{val_float:.3f} KG"
      except (ValueError, TypeError):
        pass

      mach_html += f"<tr><td>{r[1]}</td><td>{r[2]}</td><td><b>{r[3]}</b></td><td>{pcs_val}</td><td>{wt_val}</td><td>{admin_delete_cell(f'/delete-machining/{r[0]}')}</td></tr>"

    content_html += f"""
        <h4>All Saved Machining Register Entries</h4>
        <table>
            <tr><th>Date</th><th>Item</th><th>Direction</th><th>Pieces</th><th>Weight</th><th>Action</th></tr>
            {mach_html if mach_html else '<tr><td colspan="6">No machining records found.</td></tr>'}
        </table>
        """

  elif reg_type == "Plating":
    direction = request.values.get("direction", "Outward")
    cursor.execute(
        """SELECT DISTINCT v.vendor_name FROM vendors v 
           JOIN vendor_registers vr ON v.id = vr.vendor_id 
           WHERE vr.register_type = 'Plating'"""
    )
    plating_vendors = [v[0] for v in cursor.fetchall()]
    vendor_options_html = "".join([
        f'<option value="{v}">{v}</option>' for v in plating_vendors
    ])

    cursor.execute(
        """SELECT DISTINCT v.vendor_name FROM vendors v 
           JOIN vendor_registers vr ON v.id = vr.vendor_id 
           WHERE vr.register_type = 'Grinding'"""
    )
    grinding_vendors = [v[0] for v in cursor.fetchall()]
    grinding_vendor_options = "".join([
        f'<option value="{v}">{v}</option>' for v in grinding_vendors
    ])

    # Build item list from master_items (not just machining_register) so
    # items that only have Opening Stock at the 'machined' stage still show
    # up here, and compute real pending machined qty for each item so the
    # user can see/select correct available quantity when sending Outward.
    cursor.execute("SELECT item_name FROM master_items ORDER BY item_name ASC")
    all_master_items = [row[0] for row in cursor.fetchall()]

    cursor.execute(
        "SELECT item, pieces, weight FROM opening_stock WHERE stage ="
        " 'machined'"
    )
    opening_mach_map = {row[0]: (row[1], row[2]) for row in cursor.fetchall()}
    machined_pending = {}
    for item_name in all_master_items:
        cursor.execute(
            "SELECT COALESCE(SUM(pieces), 0), COALESCE(SUM(weight), 0) FROM machining_register WHERE item = ?",
            (item_name,),
        )
        m_pcs, m_wt = cursor.fetchone()

        cursor.execute(
            """SELECT COALESCE(SUM(pi.pieces), 0), COALESCE(SUM(pi.weight), 0)
            FROM plating_items pi
            JOIN plating_challans pc ON pi.challan_id = pc.id
            WHERE pi.item = ? AND pc.direction = 'Outward'""",
            (item_name,),
        )
        plate_out_pcs, plate_out_wt = cursor.fetchone()

        cursor.execute(
            "SELECT COALESCE(SUM(pieces), 0), COALESCE(SUM(weight), 0) FROM"
            " rejection_register WHERE item = ? AND rejection_type ="
            " 'Grinder Rejection'",
            (item_name,),
        )
        grind_rej_pcs, grind_rej_wt = cursor.fetchone()

        op_pcs, op_wt = opening_mach_map.get(item_name, (0, 0.0))

        net_mach_pcs = m_pcs - plate_out_pcs - grind_rej_pcs + op_pcs
        net_mach_wt = max(0.0, m_wt - plate_out_wt - grind_rej_wt + op_wt)
        
        # ONLY add to dictionary if there are pieces left to send out for plating
        if net_mach_pcs > 0:
            machined_pending[item_name] = (net_mach_pcs, net_mach_wt)

    # --- CHANGE THIS LINE ---
    # Instead of all master items, limit machining_items strictly to what is pending:
    machining_items = list(machined_pending.keys())

    machining_item_options = "".join([
        f'<option value="{i}" data-pcs="{machined_pending[i][0]}"'
        f' data-wt="{machined_pending[i][1]:.3f}">{i} (Available:'
        f' {machined_pending[i][0]} Pcs / {machined_pending[i][1]:.3f} KG)</option>'
        for i in machining_items
    ])

    pending_rows_html = "".join([
        f"<tr><td>{i}</td><td>{machined_pending[i][0]}</td><td>{machined_pending[i][1]:.3f} KG</td></tr>"
        for i in machining_items
        if machined_pending[i][0] > 0 or machined_pending[i][1] > 0
    ])

    content_html = f"""
        <h3>Plating Register</h3>
        <div class="card">
            <form method="GET" action="/registers">
                <input type="hidden" name="type" value="Plating">
                <label><b>Direction:</b></label>
                <select name="direction" onchange="this.form.submit()">
                    <option value="Outward" {"selected" if direction == "Outward" else ""}>Outward (Send to Plating)</option>
                    <option value="Inward" {"selected" if direction == "Inward" else ""}>Inward (Receive from Plating)</option>
                </select>
            </form>
        </div>
        """

    if direction == "Outward":
      batch = session.get("plating_out_batch", [])
      meta = session.get("plating_out_meta", {})
      curr_challan_no = meta.get("challan_number", "")
      curr_date = meta.get("date", today_str)
      curr_vendor = meta.get("vendor", "")

      vendor_options_out_html = "".join([
          f'<option value="{v}" {"selected" if v == curr_vendor else ""}>{v}</option>'
          for v in plating_vendors
      ])

      next_bag_num = len(batch) + 1
      batch_rows = "".join([
          f"<tr><td>{b['item']}</td><td>{b.get('grinding_vendor', 'N/A')}</td><td>{b['pieces']} Pcs</td><td>{b['weight']:.3f} KG</td><td>Bag #{b['bag_number']}</td></tr>"
          for b in batch
      ])

      content_html += f"""
            <div class="card">
                <h4>Create Plating Outward Challan</h4>
                <form method="POST">
                    <input type="hidden" name="sub_action" value="add_plating_item_batch">
                    <label><b>Challan Number:</b></label>
                    <input type="text" name="challan_number" value="{curr_challan_no}" placeholder="Enter Challan Number" required>
                    <label><b>Date:</b></label>
                    <input type="date" name="date" value="{curr_date}" required>
                    <label><b>Vendor:</b></label>
                    <select name="vendor" required><option value="">-- Select Vendor --</option>{vendor_options_out_html}</select>
                    
                    <hr>
                    <h4>Add Item to Challan Batch</h4>
                    <label><b>Item (Machined Stock, incl. Opening Stock):</b></label>
                    <select name="item" id="plating_out_item" required onchange="updatePlatingAvailable()"><option value="">-- Select Item --</option>{machining_item_options}</select>
                    <div id="plating_available_hint" style="margin:4px 0; font-size:0.9em; color:#555;"></div>
                    <label><b>Grinding Vendor:</b></label>
                    <select name="grinding_vendor"><option value="">-- Select Grinding Vendor --</option>{grinding_vendor_options}</select>
                    <label><b>Pieces:</b></label>
                    <input type="number" name="pieces" id="plating_out_pieces" required>
                    <label><b>Weight (KG):</b></label>
                    <input type="number" step="0.001" name="weight" required>
                    <label><b>No. of Bags:</b></label>
                    <input type="number" name="bag_number" value="{next_bag_num}" required>
                    <button type="submit">Add Item to Batch</button>
                </form>
                <script>
                function updatePlatingAvailable() {{
                    var sel = document.getElementById('plating_out_item');
                    var opt = sel.options[sel.selectedIndex];
                    var hint = document.getElementById('plating_available_hint');
                    if (!opt || !opt.value) {{ hint.innerHTML = ''; return; }}
                    var pcs = opt.getAttribute('data-pcs');
                    var wt = opt.getAttribute('data-wt');
                    hint.innerHTML = '<b>Available (pending) in Machined stage:</b> ' + pcs + ' Pcs / ' + wt + ' KG';
                    document.getElementById('plating_out_pieces').setAttribute('max', pcs);
                    document.getElementById('plating_out_weight').setAttribute('max', wt);
                }}
                </script>
            </div>

            <h4>Available Machined Stock (Pending for Plating Outward)</h4>
            <table><tr><th>Item</th><th>Pending Pieces</th><th>Pending Weight</th></tr>{pending_rows_html if pending_rows_html else '<tr><td colspan="3">No machined stock available.</td></tr>'}</table>

            <h4>Current Unsaved Challan Batch Items</h4>
            <table><tr><th>Item</th><th>Grinding Vendor</th><th>Pieces</th><th>Weight</th><th>Bag No</th></tr>{batch_rows if batch_rows else '<tr><td colspan="5">No items added yet.</td></tr>'}</table>
            {f'<form method="POST"><input type="hidden" name="sub_action" value="save_outward_challan"><button type="submit" style="background:#28a745;">Create Challan</button></form>' if batch else ''}
            """
    elif direction == "Inward":
      vendor_options_in_html = "".join([
          f'<option value="{v}">{v}</option>' for v in plating_vendors
      ])

      # Compute what is actually outstanding (sent Outward, not yet received
      # back) per vendor + item, so the Inward item dropdown only offers
      # items genuinely pending with that vendor -- not the full machined
      # item list.
      cursor.execute("""
            SELECT pc.vendor, pi.item,
                   COALESCE(SUM(pi.pieces), 0) AS out_pcs,
                   COALESCE(SUM(pi.weight), 0) AS out_wt,
                   COALESCE(SUM(pi.pieces - pi.remaining_pieces), 0) AS old_in_pcs,
                   COALESCE(SUM(
                       CASE WHEN pi.pieces > 0
                            THEN pi.weight * (pi.pieces - pi.remaining_pieces) * 1.0 / pi.pieces
                            ELSE 0 END
                   ), 0) AS old_in_wt
            FROM plating_challans pc
            JOIN plating_items pi ON pi.challan_id = pc.id
            WHERE pc.direction = 'Outward'
            GROUP BY pc.vendor, pi.item
        """)
      out_rows_for_pending = cursor.fetchall()

      cursor.execute("""
            SELECT pc.vendor, pi.item,
                   COALESCE(SUM(pi.pieces), 0) AS in_pcs,
                   COALESCE(SUM(pi.weight), 0) AS in_wt
            FROM plating_challans pc
            JOIN plating_items pi ON pi.challan_id = pc.id
            WHERE pc.direction = 'Inward'
            GROUP BY pc.vendor, pi.item
        """)
      new_in_map_early = {(r[0], r[1]): (r[2], r[3]) for r in cursor.fetchall()}

      pending_by_vendor = {}
      for v_name, i_name, out_pcs, out_wt, old_in_pcs, old_in_wt in out_rows_for_pending:
        new_in_pcs, new_in_wt = new_in_map_early.get((v_name, i_name), (0, 0.0))
        total_in_pcs = old_in_pcs + new_in_pcs
        total_in_wt = old_in_wt + new_in_wt
        short_pcs = out_pcs - total_in_pcs
        short_wt = max(0.0, out_wt - total_in_wt)
        if short_pcs > 0 or short_wt > 0:
          pending_by_vendor.setdefault(v_name, []).append(
              (i_name, short_pcs, short_wt)
          )

      selected_in_vendor = request.values.get("vendor", "")

      vendor_filter_options = "".join([
          f'<option value="{v}" {"selected" if v == selected_in_vendor else ""}>{v}</option>'
          for v in plating_vendors
      ])

      if selected_in_vendor:
        pending_items_for_vendor = sorted(
            pending_by_vendor.get(selected_in_vendor, []),
            key=lambda x: x[0],
        )
        inward_item_options = "".join([
            f'<option value="{i_name}">{i_name} (Pending: {s_pcs} Pcs /'
            f' {s_wt:.3f} KG)</option>'
            for i_name, s_pcs, s_wt in pending_items_for_vendor
        ])
        if not pending_items_for_vendor:
          inward_item_options = (
              '<option value="">-- No items pending with this vendor'
              ' --</option>'
          )
        vendor_selected_html = (
            f'<input type="hidden" name="vendor" value="{selected_in_vendor}">'
            f'<input type="text" value="{selected_in_vendor}" disabled'
            ' style="background:#eee;">'
        )
      else:
        inward_item_options = (
            '<option value="">-- Select a vendor above first --</option>'
        )
        vendor_selected_html = (
            f'<select name="vendor" required><option value="">-- Select'
            f' Vendor --</option>{vendor_options_in_html}</select>'
        )

      content_html += f"""
            <div class="card">
                <form method="GET" action="/registers">
                    <input type="hidden" name="type" value="Plating">
                    <input type="hidden" name="direction" value="Inward">
                    <label><b>Select Plating Vendor to Receive From:</b></label>
                    <select name="vendor" onchange="this.form.submit()">
                        <option value="">-- Select Vendor --</option>
                        {vendor_filter_options}
                    </select>
                </form>
            </div>

            <div class="card">
                <h4>Process Plating Inward Receipt</h4>
                <form method="POST">
                    <input type="hidden" name="sub_action" value="save_inward">
                    <label><b>Date:</b></label>
                    <input type="date" name="date" value="{today_str}" required>
                    <label><b>Plating Vendor:</b></label>
                    {vendor_selected_html}
                    <label><b>Item (Outstanding with this Vendor):</b></label>
                    <select name="item" required><option value="">-- Select Item --</option>{inward_item_options}</select>
                    <label><b>Pieces Received:</b></label>
                    <input type="number" name="pieces" required>
                    <label><b>Weight Received (KG):</b></label>
                    <input type="number" step="0.001" name="weight" required>
                    <button type="submit" style="background:#28a745;">Save Inward Receipt</button>
                </form>
            </div>
            """

      short_rows_html = ""
      for v_name in sorted(pending_by_vendor.keys()):
        for i_name, short_pcs, short_wt in sorted(
            pending_by_vendor[v_name], key=lambda x: x[0]
        ):
          row_style = "color:#dc3545; font-weight:bold;" if short_pcs > 0 else ""
          short_rows_html += (
              f"<tr><td>{v_name}</td><td>{i_name}</td>"
              f"<td>{short_pcs} Pcs</td><td>{short_wt:.3f} KG</td></tr>"
          )

      content_html += f"""
            <h4>Plating Short Quantity &mdash; by Vendor &amp; Item (Pending with Vendor)</h4>
            <table>
                <tr><th>Plating Vendor</th><th>Item</th><th>Short (Pending Pieces)</th><th>Short (Pending Weight)</th></tr>
                {short_rows_html if short_rows_html else '<tr><td colspan="4">No plating outward activity found.</td></tr>'}
            </table>
            """

    cursor.execute("""
            SELECT c.id, c.challan_number, c.date, c.vendor, i.item, i.grinding_vendor, i.pieces, i.weight 
            FROM plating_challans c 
            JOIN plating_items i ON c.id = i.challan_id 
            WHERE c.direction = 'Outward'
            ORDER BY c.id DESC
        """)
    all_outward_rows = cursor.fetchall()
    outward_table_html = "".join([
        f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]}</td><td>{r[5] if r[5] else 'N/A'}</td><td>{r[6]} Pcs</td><td>{r[7]:.3f} KG</td>"
        f"<td>{admin_delete_cell(f'/delete-plating-challan/{r[0]}?direction=Outward', 'Delete this entire outward challan (all its items)?')}</td></tr>"
        for r in all_outward_rows
    ])

    cursor.execute("""
            SELECT c.id, c.date, c.vendor, i.item, i.pieces, i.weight 
            FROM plating_challans c 
            JOIN plating_items i ON c.id = i.challan_id 
            WHERE c.direction = 'Inward'
            ORDER BY c.id DESC
        """)
    all_inward_rows = cursor.fetchall()
    inward_table_html = "".join([
        f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]} Pcs</td><td>{r[5]:.3f} KG</td>"
        f"<td>{admin_delete_cell(f'/delete-plating-challan/{r[0]}?direction=Inward', 'Delete this plating inward receipt?')}</td></tr>"
        for r in all_inward_rows
    ])

    content_html += f"""
        <h4>Plating Outward Challans Register</h4>
        <table>
            <tr><th>Challan No</th><th>Date</th><th>Vendor</th><th>Item</th><th>Grinding Vendor</th><th>Pcs Sent</th><th>Weight Sent</th><th>Action</th></tr>
            {outward_table_html if outward_table_html else '<tr><td colspan="8">No plating outward challans found.</td></tr>'}
        </table>

        <h4>Plating Inward Receipts Register</h4>
        <table>
            <tr><th>Date</th><th>Vendor</th><th>Item</th><th>Pcs Received</th><th>Weight Received</th><th>Action</th></tr>
            {inward_table_html if inward_table_html else '<tr><td colspan="6">No plating inward receipts found.</td></tr>'}
        </table>
        """

  elif reg_type == "Miscellaneous":
    selected_vendor = request.values.get("vendor", "")

    cursor.execute(
        """SELECT DISTINCT v.id, v.vendor_name FROM vendors v 
           JOIN vendor_registers vr ON v.id = vr.vendor_id 
           WHERE vr.register_type = 'Miscellaneous'"""
    )
    misc_vendors_raw = cursor.fetchall()
    misc_vendors = [v[1] for v in misc_vendors_raw]

    vendor_options_html = "".join([
        f'<option value="{v}" {"selected" if v == selected_vendor else ""}>{v}</option>'
        for v in misc_vendors
    ])

    vendor_items_map = {}
    for v_id, v_name in misc_vendors_raw:
      cursor.execute(
          """SELECT m.item_name FROM master_items m 
             JOIN vendor_items vi ON m.id = vi.item_id 
             WHERE vi.vendor_id = ?""",
          (v_id,),
      )
      vendor_items_map[v_name] = [row[0] for row in cursor.fetchall()]
    # After vendor_items_map is built, add:
    cursor.execute("""
        SELECT m.item_name, u.unit_name
        FROM item_units iu
        JOIN master_items m ON m.id = iu.item_id
        JOIN units u ON u.id = iu.unit_id
    """)
    item_unit_map = {row[0]: row[1] for row in cursor.fetchall()}

    cursor.execute("SELECT unit_name FROM units ORDER BY unit_name")
    all_units = [row[0] for row in cursor.fetchall()]
    content_html = f"""
        <h3>Miscellaneous Register</h3>
        <div class="card">
            <form method="POST" action="/registers" id="miscForm" onsubmit="return validateMiscForm()">
                <input type="hidden" name="sub_action" value="save_miscellaneous">
                <input type="hidden" name="type" value="Miscellaneous">
                
                <label><b>Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>
                
                <label><b>Option Type:</b></label>
                <select name="option_type" id="optionTypeSelect" onchange="toggleVendorField(this.value)" required>
                    <option value="">-- Choose Option Type --</option>
                    <option value="purchase">Purchase (Add to Inventory)</option>
                    <option value="in assembly">In Assembly (Deduct from Inventory)</option>
                </select>
                
                <div id="vendorFieldContainer" style="display:none;">
                    <label><b>Vendor (Assigned to Miscellaneous):</b></label>
                    <select name="vendor" id="vendorSelect" onchange="updateItemsList(this.value)">
                        <option value="">-- Select Vendor --</option>
                        {vendor_options_html}
                    </select>
                </div>
                
                <label><b>Item:</b></label>
                <select name="item" id="itemSelect" required>
                    <option value="">-- Select Item --</option>
                </select>
                
                <label><b>Quantity:</b></label>
                <div style="display:flex;gap:8px;align-items:center;">
                    <input type="number" name="pieces" id="miscPiecesInput"
                        placeholder="Enter quantity"
                        style="margin:0;padding:8px;font-size:16px;width:50%;box-sizing:border-box;">
                    <select name="quantity_unit" id="miscUnitSelect"
                            style="margin:0;padding:8px;font-size:16px;width:50%;box-sizing:border-box;">
                        <option value="">-- Unit --</option>
                    </select>
                </div>
                
                <label><b>Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" id="miscWeightInput" placeholder="Enter weight (optional if pieces entered)">
                
                <div id="miscErrorMsg" style="color: red; font-weight: bold; margin-bottom: 15px; display: none;">
                    Please enter either Pieces or Weight (at least one is required).
                </div>
                
                <button type="submit" style="background:#28a745;">Save Entry</button>
            </form>
        </div>

        <script>
        var vendorItemsMap = {vendor_items_map};
        var itemUnitMap = {item_unit_map};
        var allUnits = {all_units};
        var allMasterItems = {master_items};
        function updateUnitForItem(itemName) {{
            var unitSelect = document.getElementById('miscUnitSelect');
            unitSelect.innerHTML = '<option value="">-- Unit --</option>';
            allUnits.forEach(function(u) {{
                var opt = document.createElement('option');
                opt.value = u;
                opt.innerHTML = u;
                unitSelect.appendChild(opt);
            }});
            if (itemName && itemUnitMap[itemName]) {{
                unitSelect.value = itemUnitMap[itemName];
            }}
        }}

        // Hook item select onchange (for in-assembly mode where items load immediately)
        document.getElementById('itemSelect').addEventListener('change', function() {{
            updateUnitForItem(this.value);
        }});
        // Autofill from alert link
        (function() {{
            var prefillOption = "{prefill_option}";
            var prefillItem   = "{prefill_item}";
            if (prefillOption) {{
                var optSel = document.getElementById('optionTypeSelect');
                optSel.value = prefillOption;
                toggleVendorField(prefillOption);
                if (prefillItem) {{
                    // Wait for DOM to settle then set item
                    setTimeout(function() {{
                        var itemSel = document.getElementById('itemSelect');
                        // For 'in assembly', items are already populated
                        // For 'purchase', user still needs to pick vendor first — just note it
                        itemSel.value = prefillItem;
                        if (!itemSel.value) {{
                            // item not yet in list (vendor not selected), store for later
                            itemSel.setAttribute('data-prefill', prefillItem);
                        }}
                    }}, 50);
                }}
            }}
        }})();
        function toggleVendorField(optionType) {{
            var vendorContainer = document.getElementById('vendorFieldContainer');
            var vendorSelect = document.getElementById('vendorSelect');
            var itemSelect = document.getElementById('itemSelect');
            
            itemSelect.innerHTML = '<option value="">-- Select Item --</option>';

            if (optionType === 'purchase') {{
                vendorContainer.style.display = 'block';
                vendorSelect.required = true;
                vendorSelect.value = "";
            }} else if (optionType === 'in assembly') {{
                vendorContainer.style.display = 'none';
                vendorSelect.required = false;
                vendorSelect.value = "";
                
                allMasterItems.forEach(function(item) {{
                    var opt = document.createElement('option');
                    opt.value = item;
                    opt.innerHTML = item;
                    itemSelect.appendChild(opt);
                }});
            }} else {{
                vendorContainer.style.display = 'none';
                vendorSelect.required = false;
            }}
        }}

        function updateItemsList(vendorName) {{
            var itemSelect = document.getElementById('itemSelect');
            itemSelect.innerHTML = '<option value="">-- Select Item --</option>';
            var stored = itemSelect.getAttribute('data-prefill');
            if (vendorName && vendorItemsMap[vendorName]) {{
                vendorItemsMap[vendorName].forEach(function(item) {{
                    var opt = document.createElement('option');
                    opt.value = item;
                    opt.innerHTML = item;
                    itemSelect.appendChild(opt);
                }});
                if (stored) {{
                    itemSelect.value = stored;
                    itemSelect.removeAttribute('data-prefill');
                    updateUnitForItem(stored);   // ← ADD THIS
                }}
            }}
        }}
        function validateMiscForm() {{
            var pieces = document.getElementById('miscPiecesInput').value.trim();
            var weight = document.getElementById('miscWeightInput').value.trim();
            var errorMsg = document.getElementById('miscErrorMsg');
            
            if (pieces === "" && weight === "") {{
                errorMsg.style.display = 'block';
                return false;
            }}
            errorMsg.style.display = 'none';
            return true;
        }}
        </script>
        """

    cursor.execute(
    "SELECT id, date, option_type, vendor, item, pieces, weight, "
    "COALESCE(quantity_unit, '') FROM miscellaneous_entries ORDER BY id DESC"
    )
    misc_rows = cursor.fetchall()
    misc_html = ""
    for r in misc_rows:
      pcs_val = "-"
      try:
        if r[5] is not None and str(r[5]).strip() != "":
          val_int = int(float(str(r[5])))
          q_unit = r[7] if len(r) > 7 else ""
          if val_int > 0:
            pcs_val = f"{val_int} {q_unit}".strip()
      except (ValueError, TypeError):
        pass

      wt_val = "-"
      try:
        if r[6] is not None and str(r[6]).strip() != "":
          val_float = float(r[6])
          if val_float > 0:
            wt_val = f"{val_float:.3f} KG"
      except (ValueError, TypeError):
        pass

      misc_html += f"<tr><td>{r[1]}</td><td><b>{str(r[2]).capitalize()}</b></td><td>{r[3] if r[3] else 'N/A (Assembly)'}</td><td>{r[4]}</td><td>{pcs_val}</td><td>{wt_val}</td><td>{admin_delete_cell(f'/delete-miscellaneous/{r[0]}')}</td></tr>"

    content_html += f"""
        <h4>All Saved Miscellaneous Entries</h4>
        <table>
            <tr><th>Date</th><th>Option Type</th><th>Vendor</th><th>Item</th><th>Pieces</th><th>Weight</th><th>Action</th></tr>
            {misc_html if misc_html else '<tr><td colspan="7">No miscellaneous records.</td></tr>'}
        </table>
        """

  elif reg_type == "Rejection":
    sub_rej_type = request.values.get("sub_rej_type", "inventory")
    item_options = "".join(
        [f'<option value="{i}">{i}</option>' for i in master_items]
    )

    content_html = f"""
        <h3>Rejection Register</h3>
        <div class="card">
            <form method="GET" action="/registers">
                <input type="hidden" name="type" value="Rejection">
                <label><b>Select Rejection Category:</b></label>
                <select name="sub_rej_type" onchange="this.form.submit()">
                    <option value="inventory" {"selected" if sub_rej_type == "inventory" else ""}>1. Rejection for Inventory (Process Rejection)</option>
                    <option value="brass_in_vendor" {"selected" if sub_rej_type == "brass_in_vendor" else ""}>2. Rejection for Brass In Vendor (Scrap Supplier)</option>
                    <option value="casting_vendor" {"selected" if sub_rej_type == "casting_vendor" else ""}>3. Rejection for Casting In / Brass Out Vendor</option>
                    <option value="misc_vendor" {"selected" if sub_rej_type == "misc_vendor" else ""}>4. Rejection for Miscellaneous Vendor</option>
                </select>
            </form>
        </div>
        """

    if sub_rej_type == "inventory":
      cursor.execute("SELECT vendor_name FROM vendors")
      rejection_vendors = [v[0] for v in cursor.fetchall()]
      rej_vendor_options = "".join(
          [f'<option value="{v}">{v}</option>' for v in rejection_vendors]
      )
      content_html += f"""
        <div class="card">
            <h4>Rejection for Inventory Entry</h4>
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_rejection">
                <input type="hidden" name="type" value="Rejection">
                <input type="hidden" name="sub_rej_type" value="inventory">
                
                <label><b>1. Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>
                
                <label><b>2. Item:</b></label>
                <select name="item" required>
                    <option value="">-- Select Item --</option>
                    {item_options}
                </select>
                
                <label><b>3. Type of Rejection:</b></label>
                <select name="rejection_type" required>
                    <option value="">-- Select Rejection Level --</option>
                    <option value="Casted Rejection">Casted Rejection (deducts from Casted stock)</option>
                    <option value="Machined Rejection">Machined Rejection (deducts from Casted stock)</option>
                    <option value="Grinder Rejection">Grinder Rejection (deducts from Machined stock)</option>
                    <option value="Plated Rejection">Plated Rejection (deducts from Assembly stock)</option>
                </select>
                
                <label><b>4. Vendor:</b></label>
                <select name="vendor" required>
                    <option value="">-- Select Vendor --</option>
                    {rej_vendor_options}
                </select>
                
                <label><b>5. Pieces:</b></label>
                <input type="number" name="pieces" placeholder="Enter pieces" required>
                
                <label><b>6. Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" placeholder="Enter weight" required>
                
                <button type="submit" style="background:#28a745;">Save Rejection Entry</button>
            </form>
        </div>
        """

      cursor.execute(
          "SELECT id, date, item, rejection_type, vendor, pieces, weight FROM"
          " rejection_register WHERE rejection_type != 'Vendor Rejection'"
          " ORDER BY id DESC"
      )
      rej_rows = cursor.fetchall()
      rej_html = "".join([
          f"<tr><td>{r[1]}</td><td>{r[2]}</td><td><b>{r[3]}</b></td><td>{r[4]}</td><td>{r[5]} Pcs</td><td>{r[6]:.3f} KG</td>"
          f"<td>{admin_delete_cell(f'/delete-rejection/{r[0]}?sub_rej_type=inventory')}</td></tr>"
          for r in rej_rows
      ])

      content_html += f"""
        <h4>All Saved Rejection for Inventory Entries</h4>
        <table>
            <tr><th>Date</th><th>Item</th><th>Type of Rejection</th><th>Vendor</th><th>Pieces</th><th>Weight</th><th>Action</th></tr>
            {rej_html if rej_html else '<tr><td colspan="7">No rejection records found.</td></tr>'}
        </table>
        """

    elif sub_rej_type == "brass_in_vendor":
      # Only vendors registered against "Brass In" are eligible, and each
      # option shows that vendor's currently available scrap so it's clear
      # how much can be rejected against them.
      cursor.execute(
          """SELECT DISTINCT v.vendor_name FROM vendors v
             JOIN vendor_registers vr ON v.id = vr.vendor_id
             WHERE vr.register_type = 'Brass In'"""
      )
      brass_in_vendor_names = [row[0] for row in cursor.fetchall()]

      vendor_option_rows = []
      for v_name in brass_in_vendor_names:
        cursor.execute(
            "SELECT COALESCE(SUM(remaining_weight), 0) FROM brass_in WHERE"
            " vendor = ? AND status = 'available'",
            (v_name,),
        )
        avail = cursor.fetchone()[0] or 0.0
        vendor_option_rows.append(
            f'<option value="{v_name}">{v_name} (Available: {avail:.3f} KG)</option>'
        )
      brass_in_vendor_options = "".join(vendor_option_rows)

      content_html += f"""
        <div class="card">
            <h4>Rejection for Brass In Vendor</h4>
            <p style="color:#555;">Selecting a vendor here deducts the rejected weight directly from that vendor's available scrap stock (oldest bags first).</p>
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_rejection">
                <input type="hidden" name="type" value="Rejection">
                <input type="hidden" name="sub_rej_type" value="brass_in_vendor">

                <label><b>1. Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>

                <label><b>2. Vendor (Brass In):</b></label>
                <select name="vendor" required>
                    <option value="">-- Select Vendor --</option>
                    {brass_in_vendor_options if brass_in_vendor_options else ''}
                </select>

                <label><b>3. Rejection Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" placeholder="Enter weight in KG" required>

                <button type="submit" style="background:#28a745;">Save Brass In Vendor Rejection</button>
            </form>
            {'<p style="color:#dc3545;"><b>No vendors are registered under Brass In.</b> Assign a vendor to the Brass In register first.</p>' if not brass_in_vendor_names else ''}
        </div>
        """

      cursor.execute(
          "SELECT id, date, vendor, weight FROM rejection_register WHERE"
          " rejection_type = 'Brass In Vendor Rejection' ORDER BY id DESC"
      )
      brass_in_rej_rows = cursor.fetchall()
      brass_in_rej_html = "".join([
          f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]:.3f} KG</td>"
          f"<td>{admin_delete_cell(f'/delete-rejection/{r[0]}?sub_rej_type=brass_in_vendor', 'Delete this rejection entry? Scrap bag weights already deducted will NOT be automatically restored.')}</td></tr>"
          for r in brass_in_rej_rows
      ])

      content_html += f"""
        <h4>All Saved Brass In Vendor Rejection Entries</h4>
        <table>
            <tr><th>Date</th><th>Vendor</th><th>Rejection Weight</th><th>Action</th></tr>
            {brass_in_rej_html if brass_in_rej_html else '<tr><td colspan="4">No Brass In vendor rejection records found.</td></tr>'}
        </table>
        """

    elif sub_rej_type == "casting_vendor":
      # Only vendors registered against "Brass Out" or "Casting In" are
      # eligible. This is subtracted from that vendor's casting return in
      # the Vendor Ledgers view.
      cursor.execute(
          """SELECT DISTINCT v.vendor_name FROM vendors v
             JOIN vendor_registers vr ON v.id = vr.vendor_id
             WHERE vr.register_type IN ('Brass Out', 'Casting In')"""
      )
      casting_vendor_names = [row[0] for row in cursor.fetchall()]
      casting_vendor_options = "".join(
          [f'<option value="{v}">{v}</option>' for v in casting_vendor_names]
      )

      content_html += f"""
        <div class="card">
            <h4>Rejection for Casting In / Brass Out Vendor</h4>
            <p style="color:#555;">Selecting a vendor here deducts the rejected weight from that vendor's casting return in the Vendor Ledger.</p>
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_rejection">
                <input type="hidden" name="type" value="Rejection">
                <input type="hidden" name="sub_rej_type" value="casting_vendor">

                <label><b>1. Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>

                <label><b>2. Vendor (Casting In / Brass Out):</b></label>
                <select name="vendor" required>
                    <option value="">-- Select Vendor --</option>
                    {casting_vendor_options}
                </select>

                <label><b>3. Rejection Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" placeholder="Enter weight in KG" required>

                <button type="submit" style="background:#28a745;">Save Vendor Rejection Entry</button>
            </form>
            {'<p style="color:#dc3545;"><b>No vendors are registered under Brass Out / Casting In.</b> Assign a vendor to those registers first.</p>' if not casting_vendor_names else ''}
        </div>
        """

      cursor.execute(
          "SELECT id, date, vendor, weight FROM rejection_register WHERE"
          " rejection_type = 'Vendor Rejection' ORDER BY id DESC"
      )
      rej_vendor_rows = cursor.fetchall()
      rej_vendor_html = "".join([
          f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]:.3f} KG</td>"
          f"<td>{admin_delete_cell(f'/delete-rejection/{r[0]}?sub_rej_type=casting_vendor')}</td></tr>"
          for r in rej_vendor_rows
      ])

      content_html += f"""
        <h4>All Saved Casting In / Brass Out Vendor Rejection Entries</h4>
        <table>
            <tr><th>Date</th><th>Vendor</th><th>Rejection Weight</th><th>Action</th></tr>
            {rej_vendor_html if rej_vendor_html else '<tr><td colspan="4">No vendor rejection records found.</td></tr>'}
        </table>
        """

    else:
      # misc_vendor: only vendors registered against "Miscellaneous" are
      # eligible. Item, pieces, and weight are all captured and deducted
      # from that item's purchased Miscellaneous stock.
      cursor.execute(
          """SELECT DISTINCT v.vendor_name FROM vendors v
             JOIN vendor_registers vr ON v.id = vr.vendor_id
             WHERE vr.register_type = 'Miscellaneous'"""
      )
      misc_vendor_names = [row[0] for row in cursor.fetchall()]
      misc_vendor_options = "".join(
          [f'<option value="{v}">{v}</option>' for v in misc_vendor_names]
      )

      content_html += f"""
        <div class="card">
            <h4>Rejection for Miscellaneous Vendor</h4>
            <p style="color:#555;">Selecting an item here deducts the entered quantity and weight from that item's purchased Miscellaneous stock.</p>
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_rejection">
                <input type="hidden" name="type" value="Rejection">
                <input type="hidden" name="sub_rej_type" value="misc_vendor">

                <label><b>1. Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>

                <label><b>2. Vendor (Miscellaneous):</b></label>
                <select name="vendor" required>
                    <option value="">-- Select Vendor --</option>
                    {misc_vendor_options}
                </select>

                <label><b>3. Item:</b></label>
                <select name="item" required>
                    <option value="">-- Select Item --</option>
                    {item_options}
                </select>

                <label><b>4. Pieces:</b></label>
                <input type="number" name="pieces" placeholder="Enter pieces" required>

                <label><b>5. Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" placeholder="Enter weight" required>

                <button type="submit" style="background:#28a745;">Save Misc Vendor Rejection</button>
            </form>
            {'<p style="color:#dc3545;"><b>No vendors are registered under Miscellaneous.</b> Assign a vendor to the Miscellaneous register first.</p>' if not misc_vendor_names else ''}
        </div>
        """

      cursor.execute(
          "SELECT id, date, item, vendor, pieces, weight FROM rejection_register"
          " WHERE rejection_type = 'Misc Vendor Rejection' ORDER BY id DESC"
      )
      misc_rej_rows = cursor.fetchall()
      misc_rej_html = "".join([
          f"<tr><td>{r[1]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]} Pcs</td><td>{r[5]:.3f} KG</td>"
          f"<td>{admin_delete_cell(f'/delete-rejection/{r[0]}?sub_rej_type=misc_vendor')}</td></tr>"
          for r in misc_rej_rows
      ])

      content_html += f"""
        <h4>All Saved Misc Vendor Rejection Entries</h4>
        <table>
            <tr><th>Date</th><th>Item</th><th>Vendor</th><th>Pieces</th><th>Weight</th><th>Action</th></tr>
            {misc_rej_html if misc_rej_html else '<tr><td colspan="6">No misc vendor rejection records found.</td></tr>'}
        </table>
        """

  elif reg_type == "Sale Register":
    cursor.execute("SELECT id, box_pack_name FROM box_pack_items ORDER BY box_pack_name ASC")
    box_pack_choices = cursor.fetchall()
    box_pack_options = "".join(
        [f'<option value="{bp[0]}">{bp[1]}</option>' for bp in box_pack_choices]
    )

    content_html = f"""
        <h3>Sale Register</h3>
        <div class="info-box">
            Selling a Box Pack Item deducts the sold Pcs from the Assembly stock of
            every component item that makes up that Box Pack (each component is
            deducted by the same Pcs entered below, not divided among components).
        </div>
        <div class="card">
            <h4>New Sale Entry</h4>
            <form method="POST" action="/registers" id="saleRegisterForm">
                <input type="hidden" name="sub_action" value="save_sale_register">
                <input type="hidden" name="type" value="Sale Register">

                <label><b>1. Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>

                <label><b>2. Party:</b></label>
                <input type="text" name="party" placeholder="Enter party / customer name" required>

                <label><b>3. Bill No.:</b></label>
                <input type="text" name="bill_no" placeholder="Enter bill number" required>

                <label><b>4. Box Pack Items:</b></label>
                <select name="box_pack_id" required>
                    <option value="">-- Select Box Pack --</option>
                    {box_pack_options}
                </select>

                <label><b>5. Pcs:</b></label>
                <input type="number" name="pcs" placeholder="Enter pieces sold" required>

                <label><b>6. No. of Cartons:</b></label>
                <input type="number" name="no_of_cartons" id="noOfCartons" min="0"
                       placeholder="Enter number of cartons" required
                       onchange="renderCartonWeightInputs()" oninput="renderCartonWeightInputs()">

                <div id="cartonWeightsContainer"></div>

                <button type="submit" style="background:#28a745;">Save Sale Entry</button>
            </form>
        </div>

        <script>
        function renderCartonWeightInputs() {{
            const count = parseInt(document.getElementById('noOfCartons').value || '0', 10);
            const container = document.getElementById('cartonWeightsContainer');
            container.innerHTML = '';
            if (!count || count < 1) {{ return; }}
            for (let i = 1; i <= count; i++) {{
                const label = document.createElement('label');
                label.innerHTML = '<b>Weight of Carton ' + i + ' (KG):</b>';
                const input = document.createElement('input');
                input.type = 'number';
                input.step = '0.001';
                input.name = 'carton_weight_' + i;
                input.placeholder = 'Enter weight for carton ' + i;
                input.required = true;
                container.appendChild(label);
                container.appendChild(input);
            }}
        }}
        </script>
        """

    cursor.execute(
        "SELECT id, date, party, bill_no, box_pack_name, pcs, no_of_cartons"
        " FROM sale_register ORDER BY id DESC"
    )
    sale_rows = cursor.fetchall()
    sale_html = ""
    for s_id, s_date, s_party, s_bill_no, s_bp_name, s_pcs, s_cartons in sale_rows:
      cursor.execute(
          "SELECT carton_number, weight FROM sale_register_cartons"
          " WHERE sale_id = ? ORDER BY carton_number ASC",
          (s_id,),
      )
      carton_rows = cursor.fetchall()
      cartons_str = ", ".join(
          [f"#{c_num}: {c_wt:.3f} KG" for c_num, c_wt in carton_rows]
      ) or "No cartons recorded"

      sale_html += (
          f"<tr><td>{s_date}</td><td>{s_party}</td><td>{s_bill_no}</td>"
          f"<td><b>{s_bp_name}</b></td><td>{s_pcs}</td><td>{s_cartons}</td>"
          f"<td>{cartons_str}</td>"
          f"<td>{admin_delete_cell(f'/delete-sale-register/{s_id}', 'Delete this sale entry? Component assembly deductions will be reversed.')}</td></tr>"
      )

    content_html += f"""
        <h4>All Saved Sale Entries</h4>
        <table>
            <tr>
                <th>Date</th><th>Party</th><th>Bill No.</th>
                <th>Box Pack Items</th><th>Pcs</th><th>No. of Cartons</th>
                <th>Carton Weights</th><th>Action</th>
            </tr>
            {sale_html if sale_html else '<tr><td colspan="8">No sale records found.</td></tr>'}
        </table>
        """

  elif reg_type == "Finished Goods":
    cursor.execute("SELECT id, box_pack_name FROM box_pack_items ORDER BY box_pack_name ASC")
    box_pack_choices = cursor.fetchall()
    box_pack_options = "".join(
        [f'<option value="{bp[0]}">{bp[1]}</option>' for bp in box_pack_choices]
    )

    # Current finished goods stock
    cursor.execute(
        "SELECT fg.box_pack_id, fg.box_pack_name, fg.pcs FROM finished_goods fg ORDER BY fg.box_pack_name ASC"
    )
    fg_stock = cursor.fetchall()
    fg_stock_html = "".join([
        f"<tr><td><b>{row[1]}</b></td><td>{row[2]} Pcs</td></tr>"
        for row in fg_stock
    ]) or "<tr><td colspan='2'>No finished goods stock yet.</td></tr>"

    # Addition log
    cursor.execute(
        "SELECT fgl.id, fgl.date, fgl.box_pack_name, fgl.pcs FROM finished_goods_log fgl ORDER BY fgl.id DESC"
    )
    fg_log = cursor.fetchall()
    fg_log_html = "".join([
        f"<tr><td>{row[1]}</td><td>{row[2]}</td><td>{row[3]} Pcs</td>"
        f"<td>{admin_delete_cell(f'/delete-finished-goods-log/{row[0]}', 'Delete this entry? Assembly deductions will be reversed.')}</td></tr>"
        for row in fg_log
    ]) or "<tr><td colspan='4'>No entries yet.</td></tr>"

    content_html = f"""
        <h3>Finished Goods Register</h3>
        <div class="info-box">
            Adding items here deducts the equivalent component items from Assembly stock.
            Sale Register entries will deduct directly from Finished Goods stock.
        </div>
        <div class="card">
            <h4>Add to Finished Goods</h4>
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_finished_goods">
                <input type="hidden" name="type" value="Finished Goods">
                <label><b>Date:</b></label>
                <input type="date" name="date" value="{today_str}" required>
                <label><b>Box Pack Item:</b></label>
                <select name="box_pack_id" required>
                    <option value="">-- Select Box Pack --</option>
                    {box_pack_options}
                </select>
                <label><b>Pcs:</b></label>
                <input type="number" name="pcs" placeholder="Enter pieces" required>
                <button type="submit" style="background:#28a745;">Add to Finished Goods</button>
            </form>
        </div>

        <h4>Current Finished Goods Stock</h4>
        <table>
            <tr><th>Box Pack Item</th><th>Available Pcs</th></tr>
            {fg_stock_html}
        </table>

        <h4>Finished Goods Addition Log</h4>
        <table>
            <tr><th>Date</th><th>Box Pack Item</th><th>Pcs Added</th><th>Action</th></tr>
            {fg_log_html}
        </table>
    """

  elif reg_type == "Item Records":
    selected_item = request.values.get("item", "")
    item_options_html = "".join([
        f'<option value="{i}" {"selected" if i == selected_item else ""}>{i}</option>'
        for i in master_items
    ])

    content_html = f"""
        <h3>Item Records Register</h3>
        <div class="card" style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <form method="GET" action="/registers" style="margin:0;">
                <input type="hidden" name="type" value="Item Records">
                <label><b>Select Item:</b></label>
                <select name="item" onchange="this.form.submit()">
                    <option value="">-- Choose Item --</option>
                    {item_options_html}
                </select>
            </form>
            <a href="/export-item-records"><button style="background:#28a745;">Download All Items (Excel)</button></a>
        </div>
    """

    if selected_item:
      cursor.execute(
          "SELECT item_category FROM master_items WHERE item_name = ?",
          (selected_item,),
      )
      cat_row = cursor.fetchone()
      item_category = cat_row[0] if cat_row else "Brass Item"

      cursor.execute(
          "SELECT date, vendor, pieces, weight FROM casting_in WHERE item = ? ORDER BY id DESC",
          (selected_item,),
      )
      c_rows = cursor.fetchall()
      total_c_pieces = sum([r[2] for r in c_rows])
      total_c_weight = sum([r[3] for r in c_rows])
      c_html = "".join([
          f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]} Pcs</td><td>{r[3]:.3f} KG</td></tr>"
          for r in c_rows
      ])
      if c_rows:
        c_html += f"""
            <tr style="background:#d4edda; font-weight:bold;">
                <td colspan="2">Total:</td>
                <td>{total_c_pieces} Pcs</td>
                <td>{total_c_weight:.3f} KG</td>
            </tr>
        """

      cursor.execute(
          "SELECT date, pieces, weight FROM machining_register WHERE item = ? ORDER BY id DESC",
          (selected_item,),
      )
      m_rows = cursor.fetchall()
      total_m_pieces = sum([r[1] for r in m_rows if r[1] is not None])
      total_m_weight = sum([r[2] for r in m_rows if r[2] is not None])
      m_html = "".join([
          f"<tr><td>{r[0]}</td><td>{r[1] or '-'} Pcs</td><td>{r[2] or '-'} KG</td></tr>"
          for r in m_rows
      ])
      if m_rows:
        m_html += f"""
            <tr style="background:#d4edda; font-weight:bold;">
                <td>Total:</td>
                <td>{total_m_pieces} Pcs</td>
                <td>{total_m_weight:.3f} KG</td>
            </tr>
        """

      cursor.execute(
          """SELECT c.date, c.challan_number, i.grinding_vendor, i.pieces, i.weight 
             FROM plating_challans c 
             JOIN plating_items i ON c.id = i.challan_id 
             WHERE i.item = ? AND c.direction = 'Outward'
               AND i.grinding_vendor IS NOT NULL AND i.grinding_vendor != '' 
             ORDER BY c.id DESC""",
          (selected_item,),
      )
      g_rows = cursor.fetchall()
      total_g_pieces = sum([r[3] for r in g_rows])
      total_g_weight = sum([r[4] for r in g_rows])
      g_html = "".join([
          f"<tr><td>{r[0]}</td><td>Ch #{r[1]}</td><td>{r[2]}</td><td>{r[3]} Pcs</td><td>{r[4]:.3f} KG</td></tr>"
          for r in g_rows
      ])
      if g_rows:
        g_html += f"""
            <tr style="background:#d4edda; font-weight:bold;">
                <td colspan="3">Total:</td>
                <td>{total_g_pieces} Pcs</td>
                <td>{total_g_weight:.3f} KG</td>
            </tr>
        """

      cursor.execute(
          """SELECT c.date, COALESCE(c.challan_number, '(Inward Receipt)'), c.vendor,
                    CASE WHEN c.direction = 'Inward' THEN i.pieces
                         ELSE (i.pieces - i.remaining_pieces) END AS pcs_in
             FROM plating_challans c 
             JOIN plating_items i ON c.id = i.challan_id 
             WHERE i.item = ?
               AND ((c.direction = 'Outward' AND (i.pieces - i.remaining_pieces) > 0)
                    OR (c.direction = 'Inward' AND i.pieces > 0))
             ORDER BY c.date DESC, c.id DESC""",
          (selected_item,),
      )
      p_rows = cursor.fetchall()
      total_p_pieces = sum([r[3] for r in p_rows])
      p_html = "".join([
          f"<tr><td>{r[0]}</td><td>Ch #{r[1]}</td><td>{r[2]}</td><td>{r[3]} Pcs</td></tr>"
          for r in p_rows
      ])
      if p_rows:
        p_html += f"""
            <tr style="background:#d4edda; font-weight:bold;">
                <td colspan="3">Total:</td>
                <td>{total_p_pieces} Pcs</td>
            </tr>
        """

      cursor.execute(
          "SELECT date, vendor, pieces, weight FROM miscellaneous_entries WHERE item = ? AND option_type = 'purchase' ORDER BY id DESC",
          (selected_item,),
      )
      misc_pur_rows = cursor.fetchall()
      total_misc_pur_pieces = sum([r[2] for r in misc_pur_rows])
      total_misc_pur_weight = sum([r[3] for r in misc_pur_rows])
      misc_pur_html = "".join([
          f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]} Pcs</td><td>{r[3]:.3f} KG</td></tr>"
          for r in misc_pur_rows
      ])
      if misc_pur_rows:
        misc_pur_html += f"""
            <tr style="background:#d4edda; font-weight:bold;">
                <td colspan="2">Total:</td>
                <td>{total_misc_pur_pieces} Pcs</td>
                <td>{total_misc_pur_weight:.3f} KG</td>
            </tr>
        """

      cursor.execute(
          "SELECT date, pieces, weight FROM miscellaneous_entries WHERE item = ? AND option_type = 'in assembly' ORDER BY id DESC",
          (selected_item,),
      )
      misc_asm_rows = cursor.fetchall()
      total_misc_asm_pieces = sum([r[1] for r in misc_asm_rows])
      total_misc_asm_weight = sum([r[2] for r in misc_asm_rows])
      misc_asm_html = "".join([
          f"<tr><td>{r[0]}</td><td>{r[1]} Pcs</td><td>{r[2]:.3f} KG</td></tr>"
          for r in misc_asm_rows
      ])
      if misc_asm_rows:
        misc_asm_html += f"""
            <tr style="background:#d4edda; font-weight:bold;">
                <td>Total:</td>
                <td>{total_misc_asm_pieces} Pcs</td>
                <td>{total_misc_asm_weight:.3f} KG</td>
            </tr>
        """

      cursor.execute(
          """SELECT sr.date, sr.bill_no, sd.pcs
             FROM sale_register_deductions sd
             JOIN sale_register sr ON sd.sale_id = sr.id
             WHERE sd.item_name = ? ORDER BY sr.id DESC""",
          (selected_item,),
      )
      sale_rows = cursor.fetchall()
      total_sale_pieces = sum([r[2] for r in sale_rows])
      sale_html = "".join([
          f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]} Pcs</td></tr>"
          for r in sale_rows
      ])
      if sale_rows:
        sale_html += f"""
            <tr style="background:#d4edda; font-weight:bold;">
                <td colspan="2">Total:</td>
                <td>{total_sale_pieces} Pcs</td>
            </tr>
        """

      sections_html = ""
      if item_category == "Brass Item":
        sections_html = f"""
            <div class="item-records-horizontal">
                <div class="item-stage-box">
                    <h4>1. Casting</h4>
                    <table>
                        <tr><th>Date</th><th>Vendor</th><th>Pcs</th><th>Weight</th></tr>
                        {c_html if c_html else '<tr><td colspan="4">No casting records.</td></tr>'}
                    </table>
                </div>
                
                <div class="item-stage-box">
                    <h4>2. Machined</h4>
                    <table>
                        <tr><th>Date</th><th>Pcs</th><th>Weight</th></tr>
                        {m_html if m_html else '<tr><td colspan="3">No machined records.</td></tr>'}
                    </table>
                </div>
                
                <div class="item-stage-box">
                    <h4>3. Grinding</h4>
                    <table>
                        <tr><th>Date</th><th>Ch No.</th><th>Vendor</th><th>Pcs</th><th>Weight</th></tr>
                        {g_html if g_html else '<tr><td colspan="5">No grinding records.</td></tr>'}
                    </table>
                </div>

                <div class="item-stage-box">
                    <h4>4. Plating</h4>
                    <table>
                        <tr><th>Date</th><th>Ch No.</th><th>Vendor</th><th>Pcs</th></tr>
                        {p_html if p_html else '<tr><td colspan="4">No plating inward.</td></tr>'}
                    </table>
                </div>

                <div class="item-stage-box">
                    <h4>5. Sale</h4>
                    <table>
                        <tr><th>Date</th><th>Bill No.</th><th>Pcs</th></tr>
                        {sale_html if sale_html else '<tr><td colspan="3">No sale records.</td></tr>'}
                    </table>
                </div>
            </div>
        """
      else:
        cursor.execute(
            "SELECT date, vendor, pieces, weight FROM rejection_register"
            " WHERE item = ? AND rejection_type = 'Misc Vendor Rejection'"
            " ORDER BY id DESC",
            (selected_item,),
        )
        misc_rej_item_rows = cursor.fetchall()
        total_misc_rej_pieces = sum([r[2] for r in misc_rej_item_rows])
        total_misc_rej_weight = sum([r[3] for r in misc_rej_item_rows])
        misc_rej_item_html = "".join([
            f"<tr><td>{r[0]}</td><td>{r[1]}</td><td>{r[2]} Pcs</td><td>{r[3]:.3f} KG</td></tr>"
            for r in misc_rej_item_rows
        ])
        if misc_rej_item_rows:
          misc_rej_item_html += f"""
              <tr style="background:#f8d7da; font-weight:bold;">
                  <td colspan="2">Total:</td>
                  <td>{total_misc_rej_pieces} Pcs</td>
                  <td>{total_misc_rej_weight:.3f} KG</td>
              </tr>
          """

        sections_html = f"""
            <div class="item-records-horizontal">
                <div class="item-stage-box">
                    <h4>5. Miscellaneous - Purchase</h4>
                    <table>
                        <tr><th>Date</th><th>Vendor</th><th>Pcs</th><th>Weight</th></tr>
                        {misc_pur_html if misc_pur_html else '<tr><td colspan="4">No purchase records.</td></tr>'}
                    </table>
                </div>

                <div class="item-stage-box">
                    <h4>6. Miscellaneous - In Assembly</h4>
                    <table>
                        <tr><th>Date</th><th>Pcs</th><th>Weight</th></tr>
                        {misc_asm_html if misc_asm_html else '<tr><td colspan="3">No assembly records.</td></tr>'}
                    </table>
                </div>

                <div class="item-stage-box">
                    <h4>7. Misc Vendor Rejection</h4>
                    <table>
                        <tr><th>Date</th><th>Vendor</th><th>Pcs</th><th>Weight</th></tr>
                        {misc_rej_item_html if misc_rej_item_html else '<tr><td colspan="4">No rejection records.</td></tr>'}
                    </table>
                </div>

                <div class="item-stage-box">
                    <h4>8. Sale</h4>
                    <table>
                        <tr><th>Date</th><th>Bill No.</th><th>Pcs</th></tr>
                        {sale_html if sale_html else '<tr><td colspan="3">No sale records.</td></tr>'}
                    </table>
                </div>
            </div>
        """

      content_html += f"""
        <div class="card" style="border: 2px solid #007bff;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                <h4 style="color:#007bff; margin:0;">Records for Item: {selected_item} ({item_category})</h4>
                <a href="/export-item-records"><button style="background:#28a745;">Download All Items (Excel)</button></a>
            </div>
            {sections_html}
        </div>
      """

  elif reg_type == "Opening Stock":
    stage_labels = {
        "casted": "Casted",
        "machined": "Machined",
        "assembly": "Assembly",
        "plating_short": "Plating Short (with plating vendor)",
        "misc": "Miscellaneous Stock",
    }

    item_options_html = "".join(
        [f'<option value="{i}">{i}</option>' for i in master_items]
    )
    stage_options_html = "".join([
        f'<option value="{key}">{label}</option>'
        for key, label in stage_labels.items()
    ])

    cursor.execute(
        "SELECT id, item, stage, pieces, weight, date FROM opening_stock"
        " ORDER BY item ASC, stage ASC"
    )
    os_rows = cursor.fetchall()
    def os_row_html(r):
        edit_btn = ""
        if session.get("role") == "admin":
            edit_btn = f"""<button type="button"
                onclick="startEditOS({r[0]}, '{r[5]}', '{r[1]}', '{r[2]}', {r[3] or 0}, {r[4] or 0})"
                style="background:#ffc107;color:#000;border:none;border-radius:4px;
                       padding:5px 10px;font-size:12px;margin-left:4px;cursor:pointer;">Edit</button>"""
        return (f"<tr><td>{r[5]}</td><td>{r[1]}</td><td>{stage_labels.get(r[2], r[2])}</td>"
                f"<td>{r[3]}</td><td>{r[4]:.3f} KG</td>"
                f"<td>{admin_delete_cell(f'/delete-opening-stock/{r[0]}', 'Delete this opening stock entry?')}{edit_btn}</td></tr>")
    os_rows_html = "".join([os_row_html(r) for r in os_rows])


    cursor.execute(
        "SELECT id, bag_number, weight, date FROM brass_in WHERE vendor ="
        " 'Opening Balance' AND status = 'available' ORDER BY date ASC"
    )
    os_scrap_rows = cursor.fetchall()
    os_scrap_html = "".join([
        f"<tr><td>{r[3]}</td><td>Bag #{r[1]}</td><td>{r[2]:.3f} KG</td>"
        f"<td>{admin_delete_cell(f'/delete-opening-scrap-bag/{r[0]}', 'Delete this opening scrap balance?')}</td></tr>"
        for r in os_scrap_rows
    ])

    content_html = f"""
        <h3>Opening Stock (Starting Balances)</h3>
        <div class="info-box">
            Use this to set your starting stock when you begin entries in continuation of a previous
            period, instead of re-entering all historical transactions. Set one opening balance per
            item &amp; stage (and one for Scrap Brass). Saving again for the same item &amp; stage
            updates that balance. These balances are automatically added into the Inventory Register.
        </div>

        <div class="card">
            <h4 style="margin-top:0;">Opening Balance &mdash; Item Stock</h4>
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_opening_stock">
                <input type="hidden" name="type" value="Opening Stock">

                <label><b>Date (as of):</b></label>
                <input type="date" name="date" value="{today_str}" required>

                <label><b>Item:</b></label>
                <select name="item" required>
                    <option value="">-- Select Item --</option>
                    {item_options_html}
                </select>

                <label><b>Stage:</b></label>
                <select name="stage" required>
                    <option value="">-- Select Stage --</option>
                    {stage_options_html}
                </select>

                <label><b>Pieces:</b></label>
                <input type="number" name="pieces" placeholder="Enter opening pieces">

                <label><b>Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" placeholder="Enter opening weight">

                <button type="submit" style="background:#28a745;">Save Opening Balance</button>
            </form>
        </div>

        <div class="card">
            <h4 style="margin-top:0;">Opening Balance &mdash; Scrap Brass</h4>
            <form method="POST" action="/registers">
                <input type="hidden" name="sub_action" value="save_opening_scrap">
                <input type="hidden" name="type" value="Opening Stock">

                <label><b>Date (as of):</b></label>
                <input type="date" name="date" value="{today_str}" required>

                <label><b>Bag Number:</b></label>
                <input type="number" name="bag_number" value="1" required>

                <label><b>Weight (KG):</b></label>
                <input type="number" step="0.001" name="weight" placeholder="Enter opening scrap weight" required>

                <button type="submit" style="background:#28a745;">Save Opening Scrap Balance</button>
            </form>
        </div>

        <h4>Current Opening Balances &mdash; Item Stock</h4>
        <table>
            <tr><th>Date</th><th>Item</th><th>Stage</th><th>Pieces</th><th>Weight</th><th>Action</th></tr>
            {os_rows_html if os_rows_html else '<tr><td colspan="6">No opening stock entries yet.</td></tr>'}
        </table>

        <h4>Current Opening Balances &mdash; Scrap Brass</h4>
        <table>
            <tr><th>Date</th><th>Bag Number</th><th>Weight</th><th>Action</th></tr>
            {os_scrap_html if os_scrap_html else '<tr><td colspan="4">No opening scrap balance yet.</td></tr>'}
        </table>
                <div id="edit-os-modal" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;
             background:rgba(0,0,0,0.5);z-index:9999;">
          <div style="background:#fff;margin:8% auto;padding:24px;max-width:420px;border-radius:8px;">
            <h3>Edit Opening Stock</h3>
            <form method="POST" action="/registers" id="edit-os-form">
              <input type="hidden" name="sub_action" value="edit_opening_stock">
              <input type="hidden" name="type" value="Opening Stock">
              <input type="hidden" name="os_id" id="edit-os-id">
              <label><b>Date:</b></label>
              <input type="date" name="date" id="edit-os-date" required><br><br>
              <label><b>Item:</b></label>
              <select name="item" id="edit-os-item" required>
                <option value="">-- Select Item --</option>
                {item_options_html}
              </select>
              <label><b>Stage:</b></label>
              <select name="stage" id="edit-os-stage" required>
                <option value="">-- Select Stage --</option>
                {stage_options_html}
              </select>
              <label><b>Pieces:</b></label>
              <input type="number" name="pieces" id="edit-os-pieces">
              <label><b>Weight (KG):</b></label>
              <input type="number" step="0.001" name="weight" id="edit-os-weight">
              <button type="submit" style="background:#28a745;color:#fff;border:none;padding:8px 16px;border-radius:4px;cursor:pointer;margin-top:10px;">Save</button>
              <button type="button" onclick="document.getElementById('edit-os-modal').style.display='none'"
                      style="margin-left:8px;padding:8px 16px;cursor:pointer;">Cancel</button>
            </form>
          </div>
        </div>
        <script>
        function startEditOS(id, date, item, stage, pieces, weight) {{
            document.getElementById('edit-os-id').value = id;
            document.getElementById('edit-os-date').value = date;
            document.getElementById('edit-os-item').value = item;
            document.getElementById('edit-os-stage').value = stage;
            document.getElementById('edit-os-pieces').value = pieces;
            document.getElementById('edit-os-weight').value = weight;
            document.getElementById('edit-os-modal').style.display = 'block';
        }}
        </script>
        """

  elif reg_type == "Inventory":
    inventory_rows = []

    cursor.execute(
        "SELECT SUM(remaining_weight) FROM brass_in WHERE status = 'available'"
        " AND remaining_weight > 0"
    )
    scrap_res = cursor.fetchone()
    total_scrap = scrap_res[0] if scrap_res and scrap_res[0] is not None else 0.0

    inventory_rows.append({
        "item": "Scrap (Brass)",
        "weight": f"{total_scrap:.3f} KG" if total_scrap > 0 else "",
        "pieces": "",
    })

    cursor.execute("SELECT item_name FROM master_items ORDER BY item_name ASC")
    all_master_items = [row[0] for row in cursor.fetchall()]

    cursor.execute("SELECT item, stage, pieces, weight FROM opening_stock")
    opening_map = {}
    for op_item, op_stage, op_pcs, op_wt in cursor.fetchall():
      opening_map[(op_item, op_stage)] = (op_pcs, op_wt)

    cursor.execute("""
        SELECT m.item_name, u.unit_name
        FROM item_units iu
        JOIN master_items m ON m.id = iu.item_id
        JOIN units u ON u.id = iu.unit_id
    """)
    inv_item_unit_map = {row[0]: row[1] for row in cursor.fetchall()}

    for item_name in all_master_items:
      op_casted_pcs, op_casted_wt = opening_map.get(
          (item_name, "casted"), (0, 0.0)
      )
      op_mach_pcs, op_mach_wt = opening_map.get(
          (item_name, "machined"), (0, 0.0)
      )
      op_assembly_pcs, op_assembly_wt = opening_map.get(
          (item_name, "assembly"), (0, 0.0)
      )
      op_plate_short_pcs, op_plate_short_wt = opening_map.get(
          (item_name, "plating_short"), (0, 0.0)
      )
      op_misc_pcs, op_misc_wt = opening_map.get((item_name, "misc"), (0, 0.0))

      cursor.execute(
          """
          SELECT rejection_type, COALESCE(SUM(pieces), 0), COALESCE(SUM(weight), 0)
          FROM rejection_register
          WHERE item = ? AND rejection_type != 'Vendor Rejection'
          GROUP BY rejection_type
      """,
          (item_name,),
      )

      rej_data = {
          "Casted Rejection": {"pcs": 0, "wt": 0.0},
          "Machined Rejection": {"pcs": 0, "wt": 0.0},
          "Grinder Rejection": {"pcs": 0, "wt": 0.0},
          "Plated Rejection": {"pcs": 0, "wt": 0.0},
      }
      for r_type, r_pcs, r_wt in cursor.fetchall():
        if r_type in rej_data:
          rej_data[r_type]["pcs"] = r_pcs
          rej_data[r_type]["wt"] = r_wt

      casted_reject_pcs = (
          rej_data["Casted Rejection"]["pcs"]
          + rej_data["Machined Rejection"]["pcs"]
      )
      casted_reject_wt = (
          rej_data["Casted Rejection"]["wt"]
          + rej_data["Machined Rejection"]["wt"]
      )
      machined_reject_pcs = rej_data["Grinder Rejection"]["pcs"]
      machined_reject_wt = rej_data["Grinder Rejection"]["wt"]
      assembly_reject_pcs = rej_data["Plated Rejection"]["pcs"]
      assembly_reject_wt = rej_data["Plated Rejection"]["wt"]

      cursor.execute(
          "SELECT COALESCE(SUM(pieces), 0), COALESCE(SUM(weight), 0) FROM"
          " casting_in WHERE item = ?",
          (item_name,),
      )
      casting_pcs, casting_wt = cursor.fetchone()

      cursor.execute(
          "SELECT COALESCE(SUM(pieces), 0), COALESCE(SUM(weight), 0) FROM"
          " machining_register WHERE item = ?",
          (item_name,),
      )
      m_pcs, m_wt = cursor.fetchone()

      cursor.execute(
          """SELECT
             COALESCE(SUM(pi.pieces), 0),
             COALESCE(SUM(pi.weight), 0),
             COALESCE(SUM(pi.pieces - pi.remaining_pieces), 0),
             COALESCE(SUM(
                 CASE WHEN pi.pieces > 0
                      THEN pi.weight * (pi.pieces - pi.remaining_pieces) * 1.0 / pi.pieces
                      ELSE 0 END
             ), 0)
             FROM plating_items pi
             JOIN plating_challans pc ON pi.challan_id = pc.id
             WHERE pi.item = ? AND pc.direction = 'Outward'""",
          (item_name,),
      )
      plate_out_pcs, plate_out_wt, plate_in_pcs_old, plate_in_wt_old = (
          cursor.fetchone()
      )

      cursor.execute(
          """SELECT COALESCE(SUM(pi.pieces), 0), COALESCE(SUM(pi.weight), 0)
             FROM plating_items pi
             JOIN plating_challans pc ON pi.challan_id = pc.id
             WHERE pi.item = ? AND pc.direction = 'Inward'""",
          (item_name,),
      )
      plate_in_pcs_new, plate_in_wt_new = cursor.fetchone()

      plate_in_pcs = plate_in_pcs_old + plate_in_pcs_new
      plate_in_wt = plate_in_wt_old + plate_in_wt_new
      plate_short_pcs = plate_out_pcs - plate_in_pcs + op_plate_short_pcs
      plate_short_wt = max(
          0.0, plate_out_wt - plate_in_wt + op_plate_short_wt
      )

      cursor.execute(
          """SELECT
             COALESCE(SUM(CASE WHEN option_type = 'purchase' THEN pieces ELSE 0 END), 0),
             COALESCE(SUM(CASE WHEN option_type = 'purchase' THEN weight ELSE 0 END), 0),
             COALESCE(SUM(CASE WHEN option_type = 'in assembly' THEN pieces ELSE 0 END), 0),
             COALESCE(SUM(CASE WHEN option_type = 'in assembly' THEN weight ELSE 0 END), 0)
             FROM miscellaneous_entries WHERE item = ?""",
          (item_name,),
      )
      misc_pur_pcs, misc_pur_wt, misc_asm_pcs, misc_asm_wt = cursor.fetchone()

      cursor.execute(
          "SELECT COALESCE(SUM(pieces), 0), COALESCE(SUM(weight), 0) FROM"
          " rejection_register WHERE item = ? AND rejection_type = 'Misc"
          " Vendor Rejection'",
          (item_name,),
      )
      misc_reject_pcs, misc_reject_wt = cursor.fetchone()

      cursor.execute(
          "SELECT COALESCE(SUM(pcs), 0) FROM sale_register_deductions"
          " WHERE item_name = ?",
          (item_name,),
      )
      sale_deduct_pcs = cursor.fetchone()[0]

      net_misc_pcs = (
          misc_pur_pcs - misc_asm_pcs - misc_reject_pcs + op_misc_pcs
      )
      net_misc_wt = max(
          0.0, misc_pur_wt - misc_asm_wt - misc_reject_wt + op_misc_wt
      )

      net_casted_pcs = casting_pcs - m_pcs - casted_reject_pcs + op_casted_pcs
      net_casted_wt = max(
          0.0, casting_wt - m_wt - casted_reject_wt + op_casted_wt
      )

      net_mach_pcs = (
          m_pcs - plate_out_pcs - machined_reject_pcs + op_mach_pcs
      )
      net_mach_wt = max(
          0.0, m_wt - plate_out_wt - machined_reject_wt + op_mach_wt
      )

      # Deductions from assembly now include finished goods additions
      cursor.execute(
        """SELECT COALESCE(SUM(fgd.pcs), 0) FROM finished_goods_assembly_deductions fgd
        JOIN finished_goods_log fgl ON fgd.finished_goods_log_id = fgl.id
        WHERE fgd.item_name = ?""",
        (item_name,)
      )
      fg_assembly_deduct_pcs = cursor.fetchone()[0]

      net_assembly_pcs = (
          plate_in_pcs - assembly_reject_pcs + op_assembly_pcs - fg_assembly_deduct_pcs
      )
      net_assembly_wt = max(
          0.0, plate_in_wt - assembly_reject_wt + op_assembly_wt
      )

      if net_casted_pcs > 0 or net_casted_wt > 0:
        inventory_rows.append({
            "item": f"{item_name} - Casted",
            "weight": f"{net_casted_wt:.3f} KG" if net_casted_wt > 0 else "",
            "pieces": f"{net_casted_pcs}" if net_casted_pcs > 0 else "",
        })

      if net_mach_pcs > 0 or net_mach_wt > 0:
        inventory_rows.append({
            "item": f"{item_name} - Machined",
            "weight": f"{net_mach_wt:.3f} KG" if net_mach_wt > 0 else "",
            "pieces": f"{net_mach_pcs}" if net_mach_pcs > 0 else "",
        })

      if net_assembly_pcs > 0 or net_assembly_wt > 0:
        inventory_rows.append({
            "item": f"{item_name} - Assembly",
            "weight": f"{net_assembly_wt:.3f} KG" if net_assembly_wt > 0 else "",
            "pieces": f"{net_assembly_pcs}" if net_assembly_pcs > 0 else "",
        })

      if plate_short_pcs > 0 or plate_short_wt > 0:
        inventory_rows.append({
            "item": f"{item_name} - Plating Short (with plating vendor)",
            "weight": f"{plate_short_wt:.3f} KG" if plate_short_wt > 0 else "",
            "pieces": f"{plate_short_pcs}" if plate_short_pcs > 0 else "",
        })

      if net_misc_pcs != 0 or net_misc_wt != 0.0:
        misc_label = "Miscellaneous Stock (Purchases / Assembly Net)"
        unit_label = inv_item_unit_map.get(item_name, "pcs")
        qty_display = f"{net_misc_pcs} {unit_label}".strip() if net_misc_pcs != 0 else ""
        inventory_rows.append({
            "item": f"{item_name} - {misc_label}",
            "weight": f"{net_misc_wt:.3f} KG" if net_misc_wt != 0 else "",
            "pieces": qty_display,
        })
    inventory_html = "".join([
        (
            f"<tr><td>{idx}</td><td><b>{row['item']}</b></td><td>{row['weight']}</td><td>{row['pieces']}</td></tr>"
        )
        for idx, row in enumerate(inventory_rows, 1)
    ])

    content_html = f"""
        <h3>Inventory Register (Casted &rarr; Machined &rarr; Assembly, with Plating Short)</h3>
        <div class="info-box">
            Casted = Casting In &minus; sent to Machining &minus; Casted/Machined rejections.<br>
            Machined = Machined &minus; sent Outward for Plating &minus; Grinder rejections.<br>
            Assembly = received back via Plating Inward &minus; Plated rejections &minus; Box Pack sales (Sale Register).<br>
            Plating Short = sent Outward for Plating but not yet received back (currently with the plating vendor).<br>
            Opening balances set under "Opening Stock" are automatically included in each stage above.
        </div>
        
        <table>
            <tr>
                <th style="width: 60px;">S. No</th>
                <th>Item & Stage</th>
                <th>Weight</th>
                <th>Quantity</th>
            </tr>
            {inventory_html if len(inventory_rows) > 0 else '<tr><td colspan="4">No inventory data available.</td></tr>'}
        </table>
        """

  elif reg_type == "Vendor Ledgers":
    ledger_sub_type = request.values.get(
        "ledger_sub_type", "brass_out_casting_in"
    )
    ledger_vendor_filter = request.values.get("vendor", "")
    content_html = f"""
        <h3>Unified Vendor Ledgers</h3>
        <div class="card">
            <form method="GET" action="/registers">
                <input type="hidden" name="type" value="Vendor Ledgers">
                <label><b>Select Ledger Type:</b></label>
                <select name="ledger_sub_type" onchange="this.form.submit()">
                    <option value="brass_out_casting_in" {"selected" if ledger_sub_type == "brass_out_casting_in" else ""}>1. Casting Vendor - Brass Out & Casting In (Two-Sided Ledger)</option>
                    <option value="brass_in" {"selected" if ledger_sub_type == "brass_in" else ""}>2. Scrap Vendor - Brass In (Supply History Ledger)</option>
                    <option value="miscellaneous" {"selected" if ledger_sub_type == "miscellaneous" else ""}>3. Miscellaneous Vendor - Purchase Entries</option>
                    <option value="grinding" {"selected" if ledger_sub_type == "grinding" else ""}>4. Grinding Vendor - Plating Out Entries</option>
                    <option value="plating" {"selected" if ledger_sub_type == "plating" else ""}>5. Plating Vendor - Plating Out Entries</option>
                </select>
            </form>
        </div>
        """

    if ledger_sub_type == "brass_out_casting_in":
      cast_date_from = request.values.get("cast_date_from", "")
      cast_date_to   = request.values.get("cast_date_to", "")

      cursor.execute(
          """SELECT DISTINCT v.vendor_name FROM vendors v
             JOIN vendor_registers vr ON v.id = vr.vendor_id
             WHERE vr.register_type IN ('Brass Out', 'Casting In')"""
      )
      ledger_vendors = [row[0] for row in cursor.fetchall()]

      if not ledger_vendors:
        content_html += (
            "<p>No vendors found assigned to Brass Out or Casting In.</p>"
        )
      else:
        content_html += f"""
          <div class="card">
            <form method="GET" action="/registers" style="display:flex;gap:15px;flex-wrap:wrap;align-items:flex-end;">
              <input type="hidden" name="type" value="Vendor Ledgers">
              <input type="hidden" name="ledger_sub_type" value="brass_out_casting_in">
              <div>
                <label><b>From Date:</b></label><br>
                <input type="date" name="cast_date_from" value="{cast_date_from}" style="width:auto;margin:0;">
              </div>
              <div>
                <label><b>To Date:</b></label><br>
                <input type="date" name="cast_date_to" value="{cast_date_to}" style="width:auto;margin:0;">
              </div>
              <div>
                <button type="submit" style="margin-bottom:0;">Apply Filter</button>
                <a href="/registers?type=Vendor+Ledgers&ledger_sub_type=brass_out_casting_in">
                  <button type="button" style="margin-bottom:0;background:#6c757d;">Clear</button>
                </a>
              </div>
            </form>
          </div>
        """

        for v_name in ledger_vendors:
          # ── Compute opening balance before cast_date_from ──
          opening_balance = None
          if cast_date_from:
            cursor.execute(
                "SELECT COALESCE(SUM(weight),0) FROM brass_out WHERE given_to=? AND date<?",
                (v_name, cast_date_from),
            )
            pre_bo = cursor.fetchone()[0] or 0.0
            cursor.execute(
                "SELECT COALESCE(SUM(weight),0) FROM casting_in WHERE vendor=? AND date<?",
                (v_name, cast_date_from),
            )
            pre_ci = cursor.fetchone()[0] or 0.0
            cursor.execute(
                "SELECT COALESCE(SUM(weight),0) FROM rejection_register"
                " WHERE vendor=? AND rejection_type='Vendor Rejection' AND date<?",
                (v_name, cast_date_from),
            )
            pre_rej = cursor.fetchone()[0] or 0.0
            pre_adj  = max(0.0, pre_ci - pre_rej)
            pre_burn = pre_adj * 0.05
            opening_balance = pre_bo - (pre_adj + pre_burn)

          # ── Brass Out (filtered) ──
          bo_query = ("SELECT id, date, weight, bag_number, payment_made, given_to FROM brass_out"
            " WHERE given_to=?")
          bo_params = [v_name]
          if cast_date_from:
            bo_query += " AND date>=?"; bo_params.append(cast_date_from)
          if cast_date_to:
            bo_query += " AND date<=?"; bo_params.append(cast_date_to)
          bo_query += " ORDER BY date ASC"
          cursor.execute(bo_query, bo_params)
          brass_out_rows = cursor.fetchall()   # (id, date, weight, bag_number, payment_made)
          total_brass_out_period = sum(r[2] for r in brass_out_rows)

          brass_out_html = ""
          # Opening balance marker row on the brass-out side
          if opening_balance is not None:
            ob_label = f"Brass Balance Before {cast_date_from}"
            ob_color = "#f8d7da" if opening_balance >= 0 else "#d4edda"
            brass_out_html += (
                f"<tr style='background:{ob_color};font-weight:bold;'>"
                f"<td colspan='2'>{ob_label}</td>"
                f"<td>{opening_balance:.3f} KG</td><td></td></tr>"
            )

          if brass_out_rows:
            for bo_id, b_date, b_weight, b_bag, b_paid, b_given_to in brass_out_rows:
              paid_cell = payment_checkbox_html(
                  "brass_out", bo_id, b_paid, ledger_sub_type, v_name
              )
              brass_out_html += (
                    f"<tr><td>{b_date}</td><td>Bag #{b_bag}</td>"
                    f"<td>{b_weight:.3f} KG</td>"
                    f"<td style='text-align:center;'>{paid_cell}</td></tr>"
                )
          if not brass_out_rows and opening_balance is None:
            brass_out_html = '<tr><td colspan="4">No records yet.</td></tr>'

          # ── Casting In (filtered) ──
          ci_query  = ("SELECT id, date, item, pieces, weight, payment_made FROM casting_in"
                       " WHERE vendor=?")
          ci_params = [v_name]
          if cast_date_from:
            ci_query += " AND date>=?"; ci_params.append(cast_date_from)
          if cast_date_to:
            ci_query += " AND date<=?"; ci_params.append(cast_date_to)
          ci_query += " ORDER BY date ASC"
          cursor.execute(ci_query, ci_params)
          casting_in_rows = cursor.fetchall()
          total_casting_weight = sum(r[4] for r in casting_in_rows)

          # Vendor rejection (same date window)
          rej_query  = ("SELECT COALESCE(SUM(weight),0) FROM rejection_register"
                        " WHERE vendor=? AND rejection_type='Vendor Rejection'")
          rej_params = [v_name]
          if cast_date_from:
            rej_query += " AND date>=?"; rej_params.append(cast_date_from)
          if cast_date_to:
            rej_query += " AND date<=?"; rej_params.append(cast_date_to)
          cursor.execute(rej_query, rej_params)
          vendor_rejection_weight = cursor.fetchone()[0] or 0.0

          adjusted_casting_weight = max(0.0, total_casting_weight - vendor_rejection_weight)
          burning_loss             = adjusted_casting_weight * 0.05
          total_accounted_return   = adjusted_casting_weight + burning_loss

          # Net balance
          if cast_date_from:
            net_balance = opening_balance + total_brass_out_period - total_accounted_return
            total_brass_out_display = total_brass_out_period  # period only in summary
          else:
            cursor.execute(
                "SELECT COALESCE(SUM(weight),0) FROM brass_out WHERE given_to=?", (v_name,)
            )
            total_brass_out_display = cursor.fetchone()[0] or 0.0
            net_balance = total_brass_out_display - total_accounted_return

          casting_in_html = ""
          if casting_in_rows:
            for ci_id, c_date, c_item, c_pcs, c_weight, c_paid in casting_in_rows:
              paid_cell = payment_checkbox_html(
                  "casting_in", ci_id, c_paid, ledger_sub_type, v_name
              )
              casting_in_html += (
                  f"<tr><td>{c_date}</td><td>{c_item}</td><td>{c_pcs} Pcs</td>"
                  f"<td>{c_weight:.3f} KG</td>"
                  f"<td style='text-align:center;'>{paid_cell}</td></tr>"
              )
          else:
            casting_in_html = '<tr><td colspan="5">No records yet.</td></tr>'

          if net_balance > 0:
            balance_text = (
                f"<span style='color:red;'><b>+{net_balance:.3f} KG</b> (Balance"
                " Due From Vendor)</span>"
            )
          elif net_balance < 0:
            balance_text = (
                f"<span style='color:green;'><b>{net_balance:.3f} KG</b>"
                " (Advance / Credit With Vendor)</span>"
            )
          else:
            balance_text = (
                "<span style='color:blue;'><b>0.000 KG</b> (Fully Settled)</span>"
            )

          date_filter_params = ""
          if cast_date_from:
            date_filter_params += f"&cast_date_from={cast_date_from}"
          if cast_date_to:
            date_filter_params += f"&cast_date_to={cast_date_to}"

          content_html += f"""
                <div class="card" style="border: 2px solid #007bff;">
                  <h4 style="color:#007bff; margin-top:0;">Vendor Ledger: {v_name}</h4>
                  {_vendor_export_buttons('brass_out_casting_in', v_name, date_filter_params.lstrip('&'))}
                  <div class="ledger-container">
                      <div class="ledger-side">
                          <h4 style="margin-top:0; background:#dc3545; color:white; padding:8px; border-radius:4px;">Brass Out (Scrap Given)</h4>
                          <table>
                              <tr><th>Date</th><th>Bag No</th><th>Weight</th><th>Paid</th></tr>
                              {brass_out_html}
                              <tr style="background:#f8d7da; font-weight:bold;">
                                  <td colspan="3">Total Brass Out (Period):</td>
                                  <td colspan="2">{total_brass_out_period:.3f} KG</td>
                              </tr>
                          </table>
                      </div>
                      <div class="ledger-side">
                          <h4 style="margin-top:0; background:#28a745; color:white; padding:8px; border-radius:4px;">Casting In</h4>
                          <table>
                              <tr><th>Date</th><th>Item</th><th>Pieces</th><th>Weight</th><th>Paid</th></tr>
                              {casting_in_html}
                              <tr style="background:#fff3cd;">
                                  <td colspan="3">Total Casting In Weight:</td>
                                  <td colspan="2">{total_casting_weight:.3f} KG</td>
                              </tr>
                              <tr style="background:#fff3cd;">
                                  <td colspan="3">Less: Vendor Rejection Deduction:</td>
                                  <td colspan="2">-{vendor_rejection_weight:.3f} KG</td>
                              </tr>
                              <tr style="background:#e9ecef;">
                                  <td colspan="3">Adjusted Subtotal Weight:</td>
                                  <td colspan="2">{adjusted_casting_weight:.3f} KG</td>
                              </tr>
                              <tr style="background:#e9ecef;">
                                  <td colspan="3">Burning Loss (5%):</td>
                                  <td colspan="2">{burning_loss:.3f} KG</td>
                              </tr>
                              <tr style="background:#d4edda; font-weight:bold;">
                                  <td colspan="3">Total Accounted Return:</td>
                                  <td colspan="2">{total_accounted_return:.3f} KG</td>
                              </tr>
                          </table>
                      </div>
                  </div>
                  <div style="margin-top:15px;padding:10px;background:#fff3cd;border:1px solid #ffeeba;border-radius:4px;">
                      <b>Net Ledger Summary:</b>
                      {f"Opening Balance ({cast_date_from}): {opening_balance:.3f} KG + " if opening_balance is not None else ""}
                      Period Brass Out ({total_brass_out_period:.3f} KG) &minus;
                      Total Accounted Return ({total_accounted_return:.3f} KG) = {balance_text}
                  </div>
              </div>
              """
    elif ledger_sub_type == "brass_in":
      cursor.execute(
          """SELECT DISTINCT v.vendor_name FROM vendors v
             JOIN vendor_registers vr ON v.id = vr.vendor_id
             WHERE vr.register_type = 'Brass In'"""
      )
      brass_in_vendors = [row[0] for row in cursor.fetchall()]
      if not brass_in_vendors:
        content_html += "<p>No vendors found.</p>"
      else:
        for v_name in brass_in_vendors:
          cursor.execute(
              "SELECT id, date, bag_number, weight, remaining_weight, status,"
              " payment_made FROM brass_in WHERE vendor = ? ORDER BY date"
              " ASC",
              (v_name,),
          )
          b_in_rows = cursor.fetchall()
          total_in_weight = sum([row[3] for row in b_in_rows])
          total_remaining_weight = sum(
              [row[4] for row in b_in_rows if row[5] == "available"]
          )
          rows_html = ""
          for bi_id, r_date, r_bag, r_wt, r_rem, r_status, r_paid in b_in_rows:
            paid_cell = payment_checkbox_html(
                "brass_in", bi_id, r_paid, ledger_sub_type, v_name
            )
            rows_html += f"<tr><td>{r_date}</td><td>Bag #{r_bag}</td><td>{r_wt:.3f} KG</td><td>{r_rem:.3f} KG</td><td>{r_status.capitalize()}</td><td style='text-align:center;'>{paid_cell}</td></tr>"

          cursor.execute(
              "SELECT COALESCE(SUM(weight), 0) FROM rejection_register WHERE"
              " vendor = ? AND rejection_type = 'Brass In Vendor Rejection'",
              (v_name,),
          )
          brass_in_vendor_rejection = cursor.fetchone()[0] or 0.0

          content_html += f"""
              <div class="card" style="border: 2px solid #28a745;">
                  <h4 style="color:#28a745; margin-top:0;">Vendor: {v_name}</h4>
                  {_vendor_export_buttons('brass_in', v_name)}
                  <table>
                      <tr><th>Date</th><th>Bag Number</th><th>Original Scrap Weight</th><th>Remaining Scrap Weight</th><th>Status</th><th>Paid</th></tr>
                      {rows_html if rows_html else '<tr><td colspan="6">No records.</td></tr>'}
                      <tr style="background:#d4edda; font-weight:bold;">
                          <td colspan="2">Total Scrap Supplied In:</td>
                          <td colspan="4">{total_in_weight:.3f} KG</td>
                      </tr>
                      <tr style="background:#fff3cd; font-weight:bold;">
                          <td colspan="2">Total Rejected (deducted from stock):</td>
                          <td colspan="4">-{brass_in_vendor_rejection:.3f} KG</td>
                      </tr>
                      <tr style="background:#e9ecef; font-weight:bold;">
                          <td colspan="2">Currently Available Scrap Stock:</td>
                          <td colspan="4">{total_remaining_weight:.3f} KG</td>
                      </tr>
                  </table>
              </div>
              """

    elif ledger_sub_type == "miscellaneous":
      cursor.execute(
          """SELECT DISTINCT v.vendor_name FROM vendors v
             JOIN vendor_registers vr ON v.id = vr.vendor_id
             WHERE vr.register_type = 'Miscellaneous'"""
      )
      misc_ledger_vendors = [row[0] for row in cursor.fetchall()]
      if not misc_ledger_vendors:
        content_html += (
            "<p>No vendors found assigned to Miscellaneous.</p>"
        )
      else:
        for v_name in misc_ledger_vendors:
          cursor.execute(
              "SELECT id, date, item, pieces, weight, payment_made FROM"
              " miscellaneous_entries WHERE vendor = ? AND option_type ="
              " 'purchase' ORDER BY date ASC",
              (v_name,),
          )
          misc_rows = cursor.fetchall()
          total_misc_pcs = sum([row[3] for row in misc_rows])
          total_misc_weight = sum([row[4] for row in misc_rows])
          rows_html = ""
          for m_id, m_date, m_item, m_pcs, m_wt, m_paid in misc_rows:
            paid_cell = payment_checkbox_html(
                "miscellaneous_entries", m_id, m_paid, ledger_sub_type, v_name
            )
            rows_html += f"<tr><td>{m_date}</td><td>{m_item}</td><td>{m_pcs} Pcs</td><td>{m_wt:.3f} KG</td><td style='text-align:center;'>{paid_cell}</td></tr>"

          content_html += f"""
              <div class="card" style="border: 2px solid #6f42c1;">
                  <h4 style="color:#6f42c1; margin-top:0;">Vendor: {v_name}</h4>
                  {_vendor_export_buttons('miscellaneous', v_name)}
                  <table>
                      <tr><th>Date</th><th>Item</th><th>Pieces</th><th>Weight</th><th>Paid</th></tr>
                      {rows_html if rows_html else '<tr><td colspan="5">No purchase entries yet.</td></tr>'}
                      <tr style="background:#e9ecef; font-weight:bold;">
                          <td>Total:</td>
                          <td></td>
                          <td>{total_misc_pcs} Pcs</td>
                          <td>{total_misc_weight:.3f} KG</td>
                          <td></td>
                      </tr>
                  </table>
              </div>
              """

    elif ledger_sub_type == "grinding":
          cursor.execute(
              """SELECT DISTINCT v.vendor_name FROM vendors v
                 JOIN vendor_registers vr ON v.id = vr.vendor_id
                 WHERE vr.register_type = 'Grinding'"""
          )
          grinding_ledger_vendors = [row[0] for row in cursor.fetchall()]
          
          grind_date_from = request.values.get("grind_date_from", "")
          grind_date_to = request.values.get("grind_date_to", "")
          
          if not grinding_ledger_vendors:
            content_html += "<p>No vendors found assigned to Grinding.</p>"
          else:
            content_html += f"""
              <div class="card">
                <form method="GET" action="/registers" style="display:flex;gap:15px;flex-wrap:wrap;align-items:flex-end;">
                  <input type="hidden" name="type" value="Vendor Ledgers">
                  <input type="hidden" name="ledger_sub_type" value="grinding">
                  <div>
                    <label><b>From Date:</b></label><br>
                    <input type="date" name="grind_date_from" value="{grind_date_from}" style="width:auto;margin:0;">
                  </div>
                  <div>
                    <label><b>To Date:</b></label><br>
                    <input type="date" name="grind_date_to" value="{grind_date_to}" style="width:auto;margin:0;">
                  </div>
                  <div>
                    <button type="submit" style="margin-bottom:0;">Apply Filter</button>
                    <a href="/registers?type=Vendor+Ledgers&ledger_sub_type=grinding">
                      <button type="button" style="margin-bottom:0;background:#6c757d;">Clear</button>
                    </a>
                  </div>
                </form>
              </div>
            """
            
            for v_name in grinding_ledger_vendors:
              query = """SELECT pi.id, pc.date, pc.challan_number, pi.item, pi.pieces,
                                pi.payment_made,
                                COALESCE((
                                    SELECT rl.rate FROM rate_list rl
                                    JOIN master_items mi ON mi.id = rl.item_id
                                    JOIN vendors vnd ON vnd.id = rl.vendor_id
                                    WHERE mi.item_name = pi.item AND vnd.vendor_name = ?
                                    LIMIT 1
                                ), 0) as rate
                         FROM plating_items pi
                         JOIN plating_challans pc ON pi.challan_id = pc.id
                         WHERE pi.grinding_vendor = ? AND pc.direction = 'Outward'"""
              params = [v_name, v_name]
              
              if grind_date_from:
                query += " AND pc.date >= ?"
                params.append(grind_date_from)
              if grind_date_to:
                query += " AND pc.date <= ?"
                params.append(grind_date_to)
              
              query += " ORDER BY pc.date ASC"
              cursor.execute(query, params)
              grind_rows = cursor.fetchall()
              total_grind_pcs = sum([row[4] for row in grind_rows])
              total_grind_cost = sum([row[4] * float(row[6]) for row in grind_rows])
              rows_html = ""
              for g_id, g_date, g_chno, g_item, g_pcs, g_paid, g_rate in grind_rows:
                g_rate = float(g_rate)
                g_total_cost = g_pcs * g_rate
                rate_display = f"&#8377; {g_rate:.2f}" if g_rate > 0 else "-"
                cost_display = f"&#8377; {g_total_cost:.2f}" if g_rate > 0 else "-"
                paid_cell = payment_checkbox_html(
                    "plating_items", g_id, g_paid, ledger_sub_type, v_name
                )
                rows_html += f"<tr><td>{g_date}</td><td>{g_chno}</td><td>{g_item}</td><td>{g_pcs} Pcs</td><td>{rate_display}</td><td>{cost_display}</td><td style='text-align:center;'>{paid_cell}</td></tr>"

              row_count = len(grind_rows)
              scroll_div_start = f'<div style="max-height:calc(10 * 41px);overflow-y:auto;">' if row_count > 10 else '<div>'
              
              content_html += f"""
                  <div class="card" style="border: 2px solid #fd7e14;">
                      <h4 style="color:#fd7e14; margin-top:0;">Vendor: {v_name}</h4>
                      {_vendor_export_buttons('grinding', v_name, f'grind_date_from={grind_date_from}&grind_date_to={grind_date_to}')}
                      {scroll_div_start}
                      <table>
                          <thead>
                              <tr><th>Date</th><th>Ch. No</th><th>Item</th><th>Pcs</th><th>Cost/Piece</th><th>Total Cost</th><th>Paid</th></tr>
                          </thead>
                          <tbody>
                              {rows_html if rows_html else '<tr><td colspan="7">No entries yet.</td></tr>'}
                          </tbody>
                          <tfoot>
                              <tr style="background:#e9ecef; font-weight:bold;">
                                  <td colspan="3">Total:</td>
                                  <td>{total_grind_pcs} Pcs</td>
                                  <td></td>
                                  <td>&#8377; {total_grind_cost:.2f}</td>
                                  <td></td>
                              </tr>
                          </tfoot>
                      </table>
                      </div>
                  </div>
                  """

    elif ledger_sub_type == "plating":
          cursor.execute(
              """SELECT DISTINCT v.vendor_name FROM vendors v
                 JOIN vendor_registers vr ON v.id = vr.vendor_id
                 WHERE vr.register_type = 'Plating'"""
          )
          plating_ledger_vendors = [row[0] for row in cursor.fetchall()]
          
          plate_date_from = request.values.get("plate_date_from", "")
          plate_date_to = request.values.get("plate_date_to", "")
          
          if not plating_ledger_vendors:
            content_html += "<p>No vendors found assigned to Plating.</p>"
          else:
            content_html += f"""
              <div class="card">
                <form method="GET" action="/registers" style="display:flex;gap:15px;flex-wrap:wrap;align-items:flex-end;">
                  <input type="hidden" name="type" value="Vendor Ledgers">
                  <input type="hidden" name="ledger_sub_type" value="plating">
                  <div>
                    <label><b>From Date:</b></label><br>
                    <input type="date" name="plate_date_from" value="{plate_date_from}" style="width:auto;margin:0;">
                  </div>
                  <div>
                    <label><b>To Date:</b></label><br>
                    <input type="date" name="plate_date_to" value="{plate_date_to}" style="width:auto;margin:0;">
                  </div>
                  <div>
                    <button type="submit" style="margin-bottom:0;">Apply Filter</button>
                    <a href="/registers?type=Vendor+Ledgers&ledger_sub_type=plating">
                      <button type="button" style="margin-bottom:0;background:#6c757d;">Clear</button>
                    </a>
                  </div>
                </form>
              </div>
            """
            
            for v_name in plating_ledger_vendors:
              query = """SELECT pi.id, pc.date, pc.challan_number, pi.item, pi.pieces,
                                pi.payment_made,
                                COALESCE((
                                    SELECT rl.rate FROM rate_list rl
                                    JOIN master_items mi ON mi.id = rl.item_id
                                    JOIN vendors vnd ON vnd.id = rl.vendor_id
                                    WHERE mi.item_name = pi.item AND vnd.vendor_name = ?
                                    LIMIT 1
                                ), 0) as rate
                         FROM plating_items pi
                         JOIN plating_challans pc ON pi.challan_id = pc.id
                         WHERE pc.vendor = ? AND pc.direction = 'Outward'"""
              params = [v_name, v_name]
              
              if plate_date_from:
                query += " AND pc.date >= ?"
                params.append(plate_date_from)
              if plate_date_to:
                query += " AND pc.date <= ?"
                params.append(plate_date_to)
              
              query += " ORDER BY pc.date ASC"
              cursor.execute(query, params)
              plate_rows = cursor.fetchall()
              total_plate_pcs = sum([row[4] for row in plate_rows])
              total_plate_cost = sum([row[4] * float(row[6]) for row in plate_rows])
              rows_html = ""
              for p_id, p_date, p_chno, p_item, p_pcs, p_paid, p_rate in plate_rows:
                p_rate = float(p_rate)
                p_total_cost = p_pcs * p_rate
                rate_display = f"&#8377; {p_rate:.2f}" if p_rate > 0 else "-"
                cost_display = f"&#8377; {p_total_cost:.2f}" if p_rate > 0 else "-"
                paid_cell = payment_checkbox_html(
                    "plating_items", p_id, p_paid, ledger_sub_type, v_name
                )
                rows_html += f"<tr><td>{p_date}</td><td>{p_chno}</td><td>{p_item}</td><td>{p_pcs} Pcs</td><td>{rate_display}</td><td>{cost_display}</td><td style='text-align:center;'>{paid_cell}</td></tr>"

              row_count = len(plate_rows)
              scroll_div_start = f'<div style="max-height:calc(10 * 41px);overflow-y:auto;">' if row_count > 10 else '<div>'
              
              content_html += f"""
                  <div class="card" style="border: 2px solid #20c997;">
                      <h4 style="color:#20c997; margin-top:0;">Vendor: {v_name}</h4>
                      {_vendor_export_buttons('plating', v_name, f'plate_date_from={plate_date_from}&plate_date_to={plate_date_to}')}
                      {scroll_div_start}
                      <table>
                          <thead>
                              <tr><th>Date</th><th>Ch. No</th><th>Item</th><th>Pcs</th><th>Cost/Piece</th><th>Total Cost</th><th>Paid</th></tr>
                          </thead>
                          <tbody>
                              {rows_html if rows_html else '<tr><td colspan="7">No entries yet.</td></tr>'}
                          </tbody>
                          <tfoot>
                              <tr style="background:#e9ecef; font-weight:bold;">
                                  <td colspan="3">Total:</td>
                                  <td>{total_plate_pcs} Pcs</td>
                                  <td></td>
                                  <td>&#8377; {total_plate_cost:.2f}</td>
                                  <td></td>
                              </tr>
                          </tfoot>
                      </table>
                      </div>
                  </div>
                  """

  conn.close()

  page = (
      HTML_TEMPLATE
      + f"""
    <div class="nav">
        Logged in as: <b>{{{{ session['user'] }}}}</b> ({{{{ session['role'] }}}})<br><br>
        {get_nav_links()}
    </div>
    <h2>Registers</h2>
    <div class="card">
        <form method="GET" action="/registers">
            <label><b>Select Register / Ledger:</b></label>
            <select name="type" onchange="this.form.submit()">
                <option value="">-- Choose Option --</option>
                <option value="Opening Stock" {"selected" if reg_type == "Opening Stock" else ""}>0. Opening Stock (Set Starting Balances)</option>
                <option value="Inventory" {"selected" if reg_type == "Inventory" else ""}>0b. Inventory (Current Availability)</option>
                <option value="Brass In" {"selected" if reg_type == "Brass In" else ""}>1. Brass In (Scrap)</option>
                <option value="Brass Out" {"selected" if reg_type == "Brass Out" else ""}>2. Brass Out</option>
                <option value="Casting In" {"selected" if reg_type == "Casting In" else ""}>3. Casting In</option>
                <option value="Machining" {"selected" if reg_type == "Machining" else ""}>4. Machining</option>
                <option value="Plating" {"selected" if reg_type == "Plating" else ""}>5. Plating</option>
                <option value="Miscellaneous" {"selected" if reg_type == "Miscellaneous" else ""}>6. Miscellaneous</option>
                <option value="Rejection" {"selected" if reg_type == "Rejection" else ""}>7. Rejection</option>
                <option value="Sale Register" {"selected" if reg_type == "Sale Register" else ""}>7b. Sale Register</option>
                <option value="Finished Goods" {"selected" if reg_type == "Finished Goods" else ""}>7c. Finished Goods</option>
                <option value="Vendor Ledgers" {"selected" if reg_type == "Vendor Ledgers" else ""}>8. Vendor Ledgers</option>
                <option value="Item Records" {"selected" if reg_type == "Item Records" else ""}>9. Item Records</option>
            </select>
        </form>
    </div>
    {f'''<div class="card" style="display:flex; gap:10px;">
            <a href="/export-register-excel?{request.query_string.decode()}">
                <button style="background:#217346;">⬇ Export to Excel</button>
            </a>
            <a href="/export-register-pdf?{request.query_string.decode()}">
                <button style="background:#c0392b;">⬇ Export to PDF</button>
            </a>
        </div>''' if reg_type else ''}
    {content_html}
    """
  )
  return render_template_string(page)

from reportlab.lib.styles import ParagraphStyle

@app.route("/export-register-pdf")
@login_required
def export_register_pdf():
  reg_type = request.values.get("type", "Register")

  output = io.BytesIO()
  doc = SimpleDocTemplate(
      output, pagesize=landscape(A4),
      leftMargin=1.2 * cm, rightMargin=1.2 * cm,
      topMargin=1.2 * cm, bottomMargin=1.2 * cm,
  )
  styles = getSampleStyleSheet()
  title_style = ParagraphStyle(
      "BigTitle", parent=styles["Title"], fontSize=26, spaceAfter=14
  )
  subtitle_style = ParagraphStyle(
      "BigSubtitle", parent=styles["Heading3"], fontSize=16, spaceAfter=6
  )
  elements = []

  def build_table(columns, rows):
    col_count = len(columns)
    col_width = doc.width / col_count
    data = [columns] + rows
    table = Table(data, colWidths=[col_width] * col_count, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#007bff")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 13),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f4f4f9")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return table

  if reg_type == "Vendor Ledgers":
    ledger_sub_type = request.values.get("ledger_sub_type", "brass_out_casting_in")
    vendor_filter = request.values.get("vendor", "")

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    vendor_data = _get_vendor_ledger_data(cursor, ledger_sub_type, vendor_filter)
    conn.close()

    if not vendor_data:
      elements.append(Paragraph("No data to export.", styles["Normal"]))

    for v in vendor_data:
      elements.append(Paragraph(v["vendor"], title_style))
      elements.append(Spacer(1, 10))
      for t in v["tables"]:
        if len(v["tables"]) > 1:
          elements.append(Paragraph(t["title"], subtitle_style))
        body_rows = t["rows"] if t["rows"] else [["No records."] + [""] * (len(t["columns"]) - 1)]
        elements.append(build_table(t["columns"], body_rows))
        elements.append(Spacer(1, 20))
      if v.get("summary"):
        elements.append(Paragraph(f"<b>{v['summary']}</b>", styles["Normal"]))
        elements.append(Spacer(1, 12))

    if vendor_filter:
      filename = f"{vendor_filter} - Vendor Ledgers.pdf"
    else:
      timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
      filename = f"Vendor_Ledgers_{ledger_sub_type}_{timestamp_str}.pdf"

    log_action(
        session["user"], "EXPORT PDF",
        f"Exported Vendor Ledgers ({ledger_sub_type}) for"
        f" {'vendor ' + vendor_filter if vendor_filter else 'all vendors'}",
    )

  else:
    page_html = _render_full_page_html(registers)
    dfs = _extract_tables_from_html(page_html)

    elements.append(Paragraph(f"{reg_type} Register", title_style))
    elements.append(Spacer(1, 14))

    if not dfs:
      elements.append(Paragraph("No data to export.", styles["Normal"]))
    else:
      for idx, df in enumerate(dfs, 1):
        if df.empty:
          continue
        if len(dfs) > 1:
          elements.append(Paragraph(f"Table {idx}", subtitle_style))
        columns = list(df.columns)
        rows = df.astype(str).values.tolist()
        elements.append(build_table(columns, rows))
        elements.append(Spacer(1, 20))

    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{reg_type.replace(' ', '_')}_{timestamp_str}.pdf"
    log_action(session["user"], "EXPORT PDF", f"Exported '{reg_type}' register to PDF")

  doc.build(elements)
  output.seek(0)
  return send_file(output, as_attachment=True, download_name=filename, mimetype="application/pdf")


@app.route("/export-register-excel")
@login_required
def export_register_excel():
  reg_type = request.values.get("type", "Register")

  if reg_type == "Vendor Ledgers":
    ledger_sub_type = request.values.get("ledger_sub_type", "brass_out_casting_in")
    vendor_filter = request.values.get("vendor", "")

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    vendor_data = _get_vendor_ledger_data(cursor, ledger_sub_type, vendor_filter)
    conn.close()

    output = io.BytesIO()
    used_names = set()

    # ── Bold fill colours for summary rows ──
    _BO_TOTAL_FILL   = PatternFill("solid", fgColor="FFFFFF")   # light red
    _BO_TOTAL_FILL   = PatternFill("solid", fgColor="FFFFFF")   # light red
    _ADJ_FILL        = PatternFill("solid", fgColor="FFFFFF")   # grey
    _ACCOUNTED_FILL  = PatternFill("solid", fgColor="FFFFFF")   # light green
    _BALANCE_FILL    = PatternFill("solid", fgColor="FFFFFF")   # cornsilk

    # Row labels that should always be bold in the casting-in side
    _CI_BOLD_LABELS  = {
        "Total", "Less: Vendor Rejection", "Adjusted Subtotal",
        "Burning Loss (5%)", "Total Accounted Return",
    }
    # Row labels that should always be bold in the brass-out side
    _BO_BOLD_LABELS  = {
        "Total Brass Out (Period)", "Total Scrap Out",
    }

    def _apply_cell(ws, row, col, value, bold=False, fill=None, number_format=None,
                    border=True, align_center=False):
      cell = ws.cell(row=row, column=col, value=value)
      if bold:
        cell.font = Font(bold=True)
      if fill:
        cell.fill = fill
      if number_format:
        cell.number_format = number_format
      if border:
        cell.border = _CELL_BORDER
      if align_center:
        cell.alignment = _CENTER_ALIGN
      return cell

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      if not vendor_data:
        pd.DataFrame([["No data to export."]]).to_excel(
            writer, sheet_name="Export", index=False, header=False
        )

      for v in vendor_data:
        sheet_name = _sanitize_sheet_name(v["vendor"], used_names)
        ws = writer.book.create_sheet(title=sheet_name)

        # ── For brass_out_casting_in: custom side-by-side layout ──
        if ledger_sub_type == "brass_out_casting_in" and len(v["tables"]) == 2:
          bo_table = v["tables"][0]   # Brass Out
          ci_table = v["tables"][1]   # Casting In

          # Columns to write:
          # Brass Out: Date | Bag Number | Weight (KG)   (skip Paid)
          # Gap: 2 empty columns
          # Casting In: Date | Item | Pieces | Weight (KG)   (skip Paid)
          BO_COLS = ["Date", "Bag Number", "Weight (KG)"]
          CI_COLS = ["Date", "Item", "Pieces", "Weight (KG)"]
          GAP     = 2          # empty columns between the two tables
          CI_START = len(BO_COLS) + GAP + 1   # 1-indexed column where CI starts

          # ── Row 1: Vendor name title bar ──
          total_width = len(BO_COLS) + GAP + len(CI_COLS)
          title_cell = ws.cell(row=1, column=1, value=v["vendor"])
          title_cell.font  = _TITLE_FONT
          title_cell.fill  = _TITLE_FILL
          title_cell.alignment = _CENTER_ALIGN
          ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_width)
          for c in range(1, total_width + 1):
            ws.cell(row=1, column=c).fill = _TITLE_FILL

          # Date filter info row if applicable
          info_row = 2
          date_info = ""
          if v.get("cast_date_from") or v.get("cast_date_to"):
            parts = []
            if v.get("cast_date_from"):
              parts.append(f"From: {v['cast_date_from']}")
            if v.get("cast_date_to"):
              parts.append(f"To: {v['cast_date_to']}")
            date_info = "  |  ".join(parts)
          if date_info:
            info_cell = ws.cell(row=info_row, column=1, value=f"Date Filter: {date_info}")
            info_cell.font = Font(italic=True, color="555555")
            ws.merge_cells(start_row=info_row, start_column=1,
                           end_row=info_row, end_column=total_width)
            data_start_row = 3
          else:
            data_start_row = 2

          # ── Section headers (Brass Out / Casting In) ──
          bo_hdr_cell = ws.cell(row=data_start_row, column=1,
                                value="Brass Out (Scrap Given)")
          bo_hdr_cell.font  = _TITLE_FONT
          bo_hdr_cell.fill  =     _CI_TOTAL_FILL   = PatternFill("solid", fgColor="FFFFFF")   # light yellow

          bo_hdr_cell.alignment = _CENTER_ALIGN
          ws.merge_cells(start_row=data_start_row, start_column=1,
                         end_row=data_start_row, end_column=len(BO_COLS))
          for c in range(1, len(BO_COLS) + 1):
            ws.cell(row=data_start_row, column=c).fill =     _CI_TOTAL_FILL   = PatternFill("solid", fgColor="FFFFFF")   # light yellow


          ci_hdr_cell = ws.cell(row=data_start_row, column=CI_START,
                                value="Casting In")
          ci_hdr_cell.font  = _TITLE_FONT
          ci_hdr_cell.fill  =     _CI_TOTAL_FILL   = PatternFill("solid", fgColor="FFFFFF")   # light yellow

          ci_hdr_cell.alignment = _CENTER_ALIGN
          ws.merge_cells(start_row=data_start_row, start_column=CI_START,
                         end_row=data_start_row, end_column=CI_START + len(CI_COLS) - 1)
          for c in range(CI_START, CI_START + len(CI_COLS)):
            ws.cell(row=data_start_row, column=c).fill =     _CI_TOTAL_FILL   = PatternFill("solid", fgColor="FFFFFF")   # light yellow


          # ── Column header row ──
          hdr_row = data_start_row + 1
          for ci, col_name in enumerate(BO_COLS):
            c = ws.cell(row=hdr_row, column=ci + 1, value=col_name)
            c.font  = _HEADER_FONT; c.fill  = _HEADER_FILL
            c.border = _CELL_BORDER; c.alignment = _CENTER_ALIGN
            ws.column_dimensions[get_column_letter(ci + 1)].width = (
                18 if "Date" in col_name else 16 if "Bag" in col_name else 14
            )
          for ci, col_name in enumerate(CI_COLS):
            col_excel = CI_START + ci
            c = ws.cell(row=hdr_row, column=col_excel, value=col_name)
            c.font  = _HEADER_FONT; c.fill  = _HEADER_FILL
            c.border = _CELL_BORDER; c.alignment = _CENTER_ALIGN
            ws.column_dimensions[get_column_letter(col_excel)].width = (
                18 if "Date" in col_name else 20 if "Item" in col_name else 12
            )
          # Gap columns width
          for g in range(len(BO_COLS) + 1, CI_START):
            ws.column_dimensions[get_column_letter(g)].width = 3

          # ── Build raw row lists (exclude Paid column; filter summary rows) ──
          # bo_data_rows: rows that are actual data (no summary labels) from bo_table
          # bo_summary_rows: rows that are summaries (e.g. "Total Brass Out (Period)")
          bo_all  = bo_table["rows"]   # already includes opening balance row + totals
          ci_all  = ci_table["rows"]

          # Separate data vs summary for writing
          # We'll write ALL rows but style summary rows differently
          # Pad the shorter side with empty rows so they align at the top
          max_rows = max(len(bo_all), len(ci_all))
          data_row_start = hdr_row + 1

          for r_idx in range(max_rows):
            excel_row = data_row_start + r_idx

            # ── Brass Out cell ──
            if r_idx < len(bo_all):
              bo_row = bo_all[r_idx]
              # bo_row: [Date/label, Bag Number, Weight, Paid]
              label_val = str(bo_row[0]) if bo_row[0] is not None else ""
              is_bo_summary = (label_val in _BO_BOLD_LABELS or
                               label_val.startswith("Brass Balance Before"))
              bo_fill = _BO_TOTAL_FILL if is_bo_summary else None
              if label_val.startswith("Brass Balance Before"):
                # Opening balance row: spans first 2 cols, weight in col 3
                ob_cell = ws.cell(row=excel_row, column=1, value=label_val)
                ob_cell.font   = Font(bold=True)
                ob_cell.fill   = _BALANCE_FILL
                ob_cell.border = _CELL_BORDER
                ws.merge_cells(start_row=excel_row, start_column=1,
                               end_row=excel_row, end_column=2)
                ws.cell(row=excel_row, column=2).fill   = _BALANCE_FILL
                ws.cell(row=excel_row, column=2).border = _CELL_BORDER
                wt_val = bo_row[2]
                try:
                  wt_float = float(str(wt_val).replace(" KG", "").replace(",", ""))
                except ValueError:
                  wt_float = wt_val
                wc = ws.cell(row=excel_row, column=3, value=wt_float)
                wc.font   = Font(bold=True); wc.fill = _BALANCE_FILL
                wc.border = _CELL_BORDER; wc.number_format = "0.000"
              else:
                # Date
                d_cell = ws.cell(row=excel_row, column=1, value=label_val)
                d_cell.border = _CELL_BORDER
                if is_bo_summary:
                  d_cell.font = Font(bold=True); d_cell.fill = bo_fill
                # Bag Number
                bag_val = str(bo_row[1]) if bo_row[1] is not None else ""
                bg_cell = ws.cell(row=excel_row, column=2, value=bag_val)
                bg_cell.border = _CELL_BORDER
                if is_bo_summary:
                  bg_cell.font = Font(bold=True); bg_cell.fill = bo_fill
                # Weight
                wt_val = bo_row[2] if len(bo_row) > 2 else ""
                try:
                  wt_float = float(str(wt_val).replace(" KG", "").replace(",", "")) if wt_val else None
                except ValueError:
                  wt_float = None
                wc = ws.cell(row=excel_row, column=3,
                             value=wt_float if wt_float is not None else "")
                wc.border = _CELL_BORDER
                if wt_float is not None:
                  wc.number_format = "0.000"
                if is_bo_summary:
                  wc.font = Font(bold=True); wc.fill = bo_fill
            else:
              for ci in range(1, len(BO_COLS) + 1):
                ws.cell(row=excel_row, column=ci).border = _CELL_BORDER

            # ── Casting In cells ──
            if r_idx < len(ci_all):
              ci_row = ci_all[r_idx]
              # ci_row: [Date/label, Item, Pieces, Weight, Paid]
              label_val = str(ci_row[0]) if ci_row[0] is not None else ""
              is_ci_summary = label_val in _CI_BOLD_LABELS
              if label_val in ("Total", ):
                ci_fill = _CI_TOTAL_FILL
              elif label_val in ("Less: Vendor Rejection",):
                ci_fill = _CI_TOTAL_FILL
              elif label_val in ("Adjusted Subtotal", "Burning Loss (5%)"):
                ci_fill = _ADJ_FILL
              elif label_val == "Total Accounted Return":
                ci_fill = _ACCOUNTED_FILL
              else:
                ci_fill = None

              # Date/label col
              dc = ws.cell(row=excel_row, column=CI_START, value=label_val)
              dc.border = _CELL_BORDER
              if is_ci_summary:
                dc.font = Font(bold=True)
                if ci_fill: dc.fill = ci_fill
              # Item col
              item_val = str(ci_row[1]) if len(ci_row) > 1 and ci_row[1] is not None else ""
              ic = ws.cell(row=excel_row, column=CI_START + 1, value=item_val)
              ic.border = _CELL_BORDER
              if is_ci_summary:
                ic.font = Font(bold=True)
                if ci_fill: ic.fill = ci_fill
              # Pieces col
              pcs_raw = str(ci_row[2]) if len(ci_row) > 2 and ci_row[2] is not None else ""
              try:
                pcs_val = int(pcs_raw.replace(" Pcs", "")) if pcs_raw else None
              except ValueError:
                pcs_val = None
              pc = ws.cell(row=excel_row, column=CI_START + 2,
                           value=pcs_val if pcs_val is not None else pcs_raw)
              pc.border = _CELL_BORDER
              if is_ci_summary:
                pc.font = Font(bold=True)
                if ci_fill: pc.fill = ci_fill
              # Weight col
              wt_raw = str(ci_row[3]) if len(ci_row) > 3 and ci_row[3] is not None else ""
              try:
                wt_float = float(wt_raw.replace(" KG","").replace(",","").lstrip("-")) if wt_raw else None
                if wt_raw.lstrip().startswith("-"):
                  wt_float = -abs(wt_float) if wt_float is not None else None
              except ValueError:
                wt_float = None
              wc = ws.cell(row=excel_row, column=CI_START + 3,
                           value=wt_float if wt_float is not None else wt_raw)
              wc.border = _CELL_BORDER
              if wt_float is not None:
                wc.number_format = "0.000"
              if is_ci_summary:
                wc.font = Font(bold=True)
                if ci_fill: wc.fill = ci_fill
            else:
              for ci in range(CI_START, CI_START + len(CI_COLS)):
                ws.cell(row=excel_row, column=ci).border = _CELL_BORDER

          # ── Summary row below both tables ──
          summary_row = data_row_start + max_rows + 1
          if v.get("summary"):
            sc = ws.cell(row=summary_row, column=1, value=v["summary"])
            sc.font = Font(bold=True, size=11)
            sc.fill = _BALANCE_FILL
            sc.border = _CELL_BORDER
            ws.merge_cells(start_row=summary_row, start_column=1,
                           end_row=summary_row, end_column=total_width)
            for c in range(1, total_width + 1):
              ws.cell(row=summary_row, column=c).fill   = _BALANCE_FILL
              ws.cell(row=summary_row, column=c).border = _CELL_BORDER

        else:
          # ── Generic layout for other ledger sub-types ──
          max_cols = max((len(t["columns"]) for t in v["tables"]), default=1)
          title_cell = ws.cell(row=1, column=1, value=v["vendor"])
          title_cell.font  = _TITLE_FONT
          title_cell.fill  = _TITLE_FILL
          ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
          for c in range(1, max_cols + 1):
            ws.cell(row=1, column=c).fill = _TITLE_FILL

          current_row = 2

          for t in v["tables"]:
            if len(v["tables"]) > 1:
              sub_title_cell = ws.cell(row=current_row + 1, column=1, value=t["title"])
              sub_title_cell.font = _HEADER_FONT
              current_row += 1

            df = pd.DataFrame(t["rows"], columns=t["columns"])
            df.to_excel(writer, sheet_name=sheet_name, startrow=current_row, index=False)

            for i in range(len(t["columns"])):
              header_cell = ws.cell(row=current_row + 1, column=i + 1)
              header_cell.font  = _HEADER_FONT
              header_cell.fill  = _HEADER_FILL
              ws.column_dimensions[get_column_letter(i + 1)].width = 16

            current_row += len(df) + 3

          if v.get("summary"):
            ws.cell(row=current_row, column=1, value=v["summary"]).font = _TOTAL_FONT
            current_row += 2

      if "Sheet" in writer.book.sheetnames:
        writer.book.remove(writer.book["Sheet"])

    output.seek(0)
    if vendor_filter:
      filename = f"{vendor_filter} - Vendor Ledgers.xlsx"
    else:
      timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
      filename = f"Vendor_Ledgers_{ledger_sub_type}_{timestamp_str}.xlsx"

    log_action(
        session["user"], "EXPORT EXCEL",
        f"Exported Vendor Ledgers ({ledger_sub_type}) for"
        f" {'vendor ' + vendor_filter if vendor_filter else 'all vendors'}",
    )
    return send_file(
        output, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

  # --- generic path for every other register type ---
  page_html = _render_full_page_html(registers)
  dfs = _extract_tables_from_html(page_html)

  output = io.BytesIO()
  used_names = set()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    if not dfs:
      pd.DataFrame([["No data to export."]]).to_excel(
          writer, sheet_name="Export", index=False, header=False
      )
    for idx, df in enumerate(dfs, 1):
      base_name = f"{reg_type}_{idx}" if len(dfs) > 1 else reg_type
      sheet_name = _sanitize_sheet_name(base_name, used_names)
      df.to_excel(writer, sheet_name=sheet_name, index=False)
  output.seek(0)

  timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
  filename = f"{reg_type.replace(' ', '_')}_{timestamp_str}.xlsx"
  log_action(session["user"], "EXPORT EXCEL", f"Exported '{reg_type}' register to Excel")

  return send_file(
      output, as_attachment=True, download_name=filename,
      mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

@app.route("/audit-logs")
@login_required
def audit_logs():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT timestamp, username, action, details FROM audit_logs ORDER BY id"
      " DESC"
  )
  logs = cursor.fetchall()
  conn.close()
  logs_html = "".join([
      f"<tr><td>{l[0]}</td><td><b>{l[1]}</b></td><td>{l[2]}</td><td>{l[3]}</td></tr>"
      for l in logs
  ])
  page = (
      HTML_TEMPLATE
      + f"""
    <div class="nav">
        Logged in as: <b>{{{{ session['user'] }}}}</b> ({{{{ session['role'] }}}})<br><br>
        {get_nav_links()}
    </div>
    <h2>System Audit Logs</h2>
    <table><tr><th>Timestamp</th><th>User</th><th>Action</th><th>Details</th></tr>{logs_html}</table>
    """
  )
  return render_template_string(page)


@app.route("/manage-users", methods=["GET", "POST"])
@admin_required
def manage_users():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  if request.method == "POST":
    new_user = request.form["new_username"]
    new_pass = request.form["new_password"]
    new_role = request.form["new_role"]
    try:
      cursor.execute(
          "INSERT INTO users (username, password, role) VALUES (?, ?, ?)",
          (new_user, new_pass, new_role),
      )
      conn.commit()
      log_action(session["user"], "CREATE USER", f"Created user '{new_user}'")
    except sqlite3.IntegrityError:
      pass
    return redirect(url_for("manage_users"))
  cursor.execute("SELECT id, username, password, role FROM users")
  users = cursor.fetchall()
  conn.close()
  users_html = ""
  for u in users:
    actions = f'<a href="/delete-user/{u[0]}"><button class="danger">Delete</button></a>'
    if u[1] == session["user"]:
      actions = "<i>(Current User)</i>"
    users_html += f"<tr><td>{u[1]}</td><td><code>{u[2]}</code></td><td>{u[3]}</td><td>{actions}</td></tr>"
  page = (
      HTML_TEMPLATE
      + f"""
    <div class="nav">
        Logged in as: <b>{{{{ session['user'] }}}}</b> ({{{{ session['role'] }}}})<br><br>
        {get_nav_links()}
    </div>
    <h2>Manage Users</h2>
    <div class="card">
        <form method="POST">
            <label><b>Username:</b></label><input type="text" name="new_username" required>
            <label><b>Password:</b></label><input type="password" name="new_password" required>
            <label><b>Role:</b></label>
            <select name="new_role">
                <option value="employee">Employee</option>
                <option value="admin">Admin</option>
            </select>
            <button type="submit">Create User</button>
        </form>
    </div>
    <table><tr><th>Username</th><th>Password</th><th>Role</th><th>Actions</th></tr>{users_html}</table>
    """
  )
  return render_template_string(page)


@app.route("/delete-user/<int:u_id>")
@admin_required
def delete_user(u_id):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute("SELECT username FROM users WHERE id = ?", (u_id,))
  row = cursor.fetchone()
  if row and row[1] != session["user"]:
    cursor.execute("DELETE FROM users WHERE id = ?", (u_id,))
    conn.commit()
    log_action(session["user"], "DELETE USER", f"Removed user '{row[0]}'")
  conn.close()
  return redirect(url_for("manage_users"))


if __name__ == "__main__":
  # host="0.0.0.0" makes the app reachable from other devices (e.g. a phone)
  # on the same local network, not just from this computer itself.
  app.run(host="0.0.0.0", debug=True, port=5000)