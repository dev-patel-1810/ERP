"""
gui_app.py  —  Desktop GUI for the Registers System
=====================================================
Place this file in the SAME folder as app.py and inventory.db.

Imports from app.py:
    DB_NAME, init_db, log_action,
    _get_vendor_ledger_data, _get_low_stock_alerts

All DB queries for GUI screens are written here directly
(Flask session/request objects cannot be used outside a request context).

Run:  python gui_app.py
Exe:  pyinstaller --onefile --windowed gui_app.py
      (add pandas, openpyxl, reportlab to hiddenimports in spec if needed)
"""

import threading
import webbrowser
from app import app as flask_app   # the Flask instance already created in app.py
import sys
import os
import io
from datetime import date, timedelta
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
# pyrefly: ignore [missing-import]
from tkcalendar import DateEntry
# ── import shared helpers from app.py ────────────────────────────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from app import (
    DB_NAME,
    init_db,
    log_action,
    _get_low_stock_alerts,
    _get_vendor_ledger_data,
)

# ── optional export libraries (bundled via PyInstaller spec) ─────────────────
try:
    import pandas as pd
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import cm
    from reportlab.platypus import (SimpleDocTemplate, Table,
                                    TableStyle, Paragraph, Spacer)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ═══════════════════════════════════════════════════════════════════════════════
#  PALETTE  (modernized — same roles/usage as before, refreshed colours)
# ═══════════════════════════════════════════════════════════════════════════════
C_BG       = "#eef1f8"      # app background (soft blue-grey instead of flat grey)
C_CARD     = "#ffffff"
C_TEXT     = "#26324a"      # near-navy instead of plain grey — crisper on white
C_MUTED    = "#6b7690"      # secondary/help text
C_PRIMARY  = "#2f6fed"      # refreshed blue accent
C_PRI_HOV  = "#1f56c4"
C_DANGER   = "#e5484d"
C_DAN_HOV  = "#c53a3e"
C_SUCCESS  = "#20a869"
C_SUC_HOV  = "#188352"
C_WARNING  = "#f5a623"
C_WAR_HOV  = "#d8890f"
C_GRAY     = "#6b7690"
C_GRA_HOV  = "#565f75"
C_INFO_BG  = "#eaf2ff"
C_INFO_BOR = "#2f6fed"
C_BORDER   = "#e1e5f0"
C_TH_BG    = "#2f6fed"
C_TH_FG    = "#ffffff"
C_ROW_ALT  = "#f6f8fc"
C_SEL      = "#d7e6ff"
C_GREEN_ROW= "#e1f5ec"
C_YEL_ROW  = "#fff4de"

# sidebar / shell accents
C_SIDEBAR       = "#1c2438"
C_SIDEBAR_TEXT  = "#c7cee0"
C_SIDEBAR_HOV   = "#2a3552"
C_SIDEBAR_ACT   = "#2f6fed"
C_TOPBAR        = "#ffffff"
C_SCROLL_TRACK  = "#e6e9f2"
C_SCROLL_THUMB  = "#b7c0d6"
C_SCROLL_THUMB_H= "#98a4c2"

# One place to pick the display font: Segoe UI looks native/modern on Windows
# (where this app actually runs) and falls back gracefully elsewhere.
_FONT_FAMILY = "Segoe UI"

def _calc_scale():
    import tkinter as _tk
    _r = _tk.Tk(); _r.withdraw()
    sw = _r.winfo_screenwidth()
    _r.destroy()
    if sw >= 1920: return 1.2
    if sw >= 1440: return 1.1
    if sw >= 1280: return 1.0
    return 0.9

_S = _calc_scale()
def _fs(n): return max(8, int(n * _S))

FN  = (_FONT_FAMILY, _fs(10))
FNB = (_FONT_FAMILY, _fs(10), "bold")
FH  = (_FONT_FAMILY, _fs(16), "bold")
FSH = (_FONT_FAMILY, _fs(12), "bold")
FSM = (_FONT_FAMILY, _fs(9))
# ── date display helper (COSMETIC ONLY) ──────────────────────────────────────
# Every date is still stored/queried/compared as "YYYY-MM-DD" everywhere in the
# app (today(), date_picker, SQL). This helper only changes how a date LOOKS
# when printed in a read-only list/ledger — it never feeds back into a form,
# a save(), or a query, so it can't affect any calculation or stored value.
def fmt_date_display(iso_str):
    if not iso_str:
        return iso_str
    try:
        return datetime.strptime(str(iso_str)[:10], "%Y-%m-%d").strftime("%d %b %Y")
    except (ValueError, TypeError):
        return iso_str

# ═══════════════════════════════════════════════════════════════════════════════
#  DB HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
_qry_conn = None
def _get_qry_conn():
    global _qry_conn
    if _qry_conn is None:
        _qry_conn = sqlite3.connect(DB_NAME, check_same_thread=False)
        _qry_conn.execute("PRAGMA journal_mode=WAL")
        _qry_conn.execute("PRAGMA synchronous=NORMAL")
    return _qry_conn

def _close_db():
    global _qry_conn
    if _qry_conn is not None:
        try:
            _qry_conn.execute("PRAGMA wal_checkpoint(TRUNCATE)")
            _qry_conn.commit()
        except: pass
        finally:
            _qry_conn.close()
            _qry_conn = None

def get_conn():
    """A fresh connection for code that manages its own open/close cycle
    (the export functions). Deliberately NOT the same connection qry() uses,
    so closing this one never breaks qry() elsewhere."""
    conn = sqlite3.connect(DB_NAME)
    conn.execute("PRAGMA busy_timeout=5000")
    return conn

def qry(sql, params=(), fetch=None, commit=False):
    conn = _get_qry_conn()
    cur  = conn.cursor()
    cur.execute(sql, params)
    lid  = cur.lastrowid
    res  = None
    if fetch == "one": res = cur.fetchone()
    if fetch == "all": res = cur.fetchall()
    if commit: conn.commit()
    if commit and fetch is None: return lid
    return res

def today():
    return datetime.now().strftime("%Y-%m-%d")

# ═══════════════════════════════════════════════════════════════════════════════
#  SMALL UI HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def _button_enter(event):
    """Pressing Enter activates the focused button."""
    widget = event.widget

    if isinstance(widget, (tk.Button, ttk.Button)):
        widget.invoke()
        return "break"

def btn(parent, text, cmd, bg=C_PRIMARY, hov=C_PRI_HOV, fg="white",
        padx=14, pady=7, **kw):
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                  activebackground=hov, activeforeground=fg,
                  relief="flat", padx=padx, pady=pady, font=FNB,
                  cursor="hand2", bd=0, highlightthickness=0, **kw)
    b.bind("<Enter>", lambda e: b.config(bg=hov))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    b.bind("<Return>", _button_enter)
    b.bind("<KP_Enter>", _button_enter)
    return b

def lbl(parent, text, **kw):
    return tk.Label(parent, text=text, bg=kw.pop("bg", C_CARD),
                    fg=kw.pop("fg", C_TEXT), font=kw.pop("font", FN), **kw)

def entry(parent, default="", width=18, **kw):
    e = tk.Entry(parent, font=FN, bg="white", fg=C_TEXT,
                 relief="solid", bd=1, width=width,
                 insertbackground=C_TEXT,
                 highlightthickness=1,
                 highlightbackground=C_BORDER,
                 highlightcolor=C_PRIMARY, **kw)
    if default: e.insert(0, default)
    e.bind("<Return>", lambda event: e.tk_focusNext().focus())   # ← ADD THIS
    return e

import calendar as _pycal

class DatePicker(tk.Frame):
    """Date field + calendar popup. Replaces tkcalendar.DateEntry.
    The popup is closed ONLY by: picking a day, Today, Esc, or clicking
    somewhere else in the window. It never looks at keyboard focus.
    Public API used by the app: get_date(), set_date(date), get()."""
    _FMTS = ("%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y",
             "%Y-%m-%d", "%d.%m.%Y", "%d %m %Y")
    _MONTHS = ["January","February","March","April","May","June","July",
               "August","September","October","November","December"]

    def __init__(self, parent, default=None, width=12, **kw):
        super().__init__(parent, bg=kw.pop("bg", C_CARD))
        self._date = date.today()
        self._view = self._date.replace(day=1)
        self._pop = None
        self._bound_top = None

        self.entry = tk.Entry(self, font=FN, bg="white", fg=C_TEXT, relief="solid",
                              bd=1, width=width, insertbackground=C_TEXT,
                              highlightthickness=1, highlightbackground=C_BORDER,
                              highlightcolor=C_PRIMARY, justify="center")
        self.entry.pack(side="left", fill="x", expand=True)
        self.btn = tk.Label(self, text="\u25BE", font=FNB, bg=C_PRIMARY, fg="white",
                            padx=8, cursor="hand2")
        self.btn.pack(side="left", fill="y", padx=(2, 0))
        self.btn.bind("<Button-1>", lambda e: self.toggle())
        self.entry.bind("<FocusOut>", lambda e: self._commit_text())
        self.entry.bind("<Return>", lambda e: (self._commit_text(), self.close()))
        self.entry.bind("<Escape>", lambda e: self.close())
        self.bind("<Destroy>", lambda e: self.close() if e.widget is self else None)
        if default:
            self.set_date(default)
        else:
            self._show()

    # ── value handling ──────────────────────────────────────────────────────
    def _show(self):
        self.entry.delete(0, "end")
        self.entry.insert(0, self._date.strftime("%d-%m-%Y"))

    def _parse(self, text):
        text = (text or "").strip()
        for f in self._FMTS:
            try:
                return datetime.strptime(text, f).date()
            except ValueError:
                pass
        return None

    def _commit_text(self):
        d = self._parse(self.entry.get())
        if d: self._date = d
        self._show()                     # invalid text snaps back to last good date

    def set_date(self, value):
        if isinstance(value, datetime): value = value.date()
        elif isinstance(value, str):
            value = self._parse(value)
            if value is None: raise ValueError("invalid date")
        self._date = value
        self._view = value.replace(day=1)
        self._show()
        if self._pop: self._draw()

    def get_date(self):
        self._commit_text()
        return self._date

    def get(self):
        return self.entry.get()

    # ── popup ───────────────────────────────────────────────────────────────
    def toggle(self):
        self.close() if self._pop else self.open()

    def open(self):
        if self._pop: return
        self._commit_text()
        self._view = self._date.replace(day=1)
        p = self._pop = tk.Toplevel(self)
        p.overrideredirect(True)
        p.configure(bg=C_BORDER)
        try:
            if self.winfo_toplevel().attributes("-topmost"):
                p.attributes("-topmost", True)
        except tk.TclError:
            pass
        self._body = tk.Frame(p, bg="white")
        self._body.pack(padx=1, pady=1)
        self._draw()
        p.update_idletasks()
        w, h = p.winfo_reqwidth(), p.winfo_reqheight()
        x = self.winfo_rootx()
        y = self.winfo_rooty() + self.winfo_height() + 2
        if y + h > self.winfo_screenheight():          # no room below -> open above
            y = max(0, self.winfo_rooty() - h - 2)
        x = min(x, max(0, self.winfo_screenwidth() - w))
        p.geometry(f"+{x}+{y}")
        p.lift()
        top = self.winfo_toplevel()
        self._bound_top = top
        self._click_id = top.bind("<ButtonPress-1>", self._outside_click, add="+")
        self._esc_id = top.bind("<Escape>", lambda e: self.close(), add="+")
        self._cfg_id = top.bind("<Configure>", self._on_top_configure, add="+")
        p.bind("<Escape>", lambda e: self.close())
        p.bind("<MouseWheel>", self._wheel)
        p.bind("<Button-4>", self._wheel); p.bind("<Button-5>", self._wheel)

    def close(self):
        top, self._bound_top = self._bound_top, None
        if top is not None:
            for seq, fid in (("<ButtonPress-1>", getattr(self, "_click_id", None)),
                             ("<Escape>", getattr(self, "_esc_id", None)),
                             ("<Configure>", getattr(self, "_cfg_id", None))):
                try:
                    cur = top.bind(seq)
                    if fid and cur:
                        lines = [l for l in cur.split("\n") if fid not in l]
                        top.bind(seq, "\n".join(lines))
                except tk.TclError:
                    pass
        if self._pop is not None:
            try: self._pop.destroy()
            except tk.TclError: pass
            self._pop = None

    def _outside_click(self, e):
        if self._pop is None: return
        if e.widget is self.btn: return          # the button toggles by itself
        self.close()

    def _on_top_configure(self, e):
        if e.widget is self._bound_top:          # window moved / resized
            self.close()

    def _wheel(self, e):
        up = getattr(e, "num", None) == 4 or getattr(e, "delta", 0) > 0
        self._step_month(-1 if up else 1)

    def _step_month(self, n):
        m = self._view.month - 1 + n
        self._view = self._view.replace(year=self._view.year + m // 12, month=m % 12 + 1)
        self._draw()

    def _step_year(self, n):
        self._view = self._view.replace(year=min(9998, max(1, self._view.year + n)))
        self._draw()

    def _pick(self, d):
        self._date = d
        self._show()
        self.close()
        self.entry.event_generate("<<DateSelected>>")

    def _draw(self):
        for w in self._body.winfo_children(): w.destroy()
        nav = dict(font=FNB, bg=C_PRIMARY, fg="white", cursor="hand2", padx=7)
        hdr = tk.Frame(self._body, bg=C_PRIMARY); hdr.pack(fill="x")
        for txt, cmd in (("\u00AB", lambda: self._step_year(-1)), ("\u2039", lambda: self._step_month(-1))):
            l = tk.Label(hdr, text=txt, **nav); l.pack(side="left"); l.bind("<Button-1>", lambda e, c=cmd: c())
        for txt, cmd in (("\u00BB", lambda: self._step_year(1)), ("\u203A", lambda: self._step_month(1))):
            l = tk.Label(hdr, text=txt, **nav); l.pack(side="right"); l.bind("<Button-1>", lambda e, c=cmd: c())
        tk.Label(hdr, text=f"{self._MONTHS[self._view.month-1]} {self._view.year}",
                 font=FNB, bg=C_PRIMARY, fg="white").pack(side="left", expand=True)

        grid = tk.Frame(self._body, bg="white"); grid.pack(padx=4, pady=4)
        for c, nm in enumerate(("Mo","Tu","We","Th","Fr","Sa","Su")):
            tk.Label(grid, text=nm, font=FNB, bg="white", fg=C_GRAY, width=4).grid(row=0, column=c)
        today_ = date.today()
        weeks = _pycal.Calendar(firstweekday=0).monthdatescalendar(self._view.year, self._view.month)
        for r, week in enumerate(weeks, start=1):
            for c, d in enumerate(week):
                inmonth = d.month == self._view.month
                sel = d == self._date
                bg = C_PRIMARY if sel else "white"
                fg = "white" if sel else (C_TEXT if inmonth else "#b4bccd")
                if d == today_ and not sel: fg = C_PRIMARY
                l = tk.Label(grid, text=str(d.day), font=FNB if d == today_ else FN, width=4,
                             bg=bg, fg=fg, cursor="hand2", pady=3,
                             highlightthickness=1 if d == today_ else 0,
                             highlightbackground=C_PRIMARY)
                l.grid(row=r, column=c, padx=1, pady=1)
                l.bind("<Button-1>", lambda e, dd=d: self._pick(dd))
                if not sel:
                    l.bind("<Enter>", lambda e, w=l: w.config(bg=C_SEL))
                    l.bind("<Leave>", lambda e, w=l, b=bg: w.config(bg=b))
        foot = tk.Frame(self._body, bg="white"); foot.pack(fill="x", padx=6, pady=(0, 6))
        t = tk.Label(foot, text="Today", font=FNB, fg=C_PRIMARY, bg="white", cursor="hand2")
        t.pack(side="left"); t.bind("<Button-1>", lambda e: self._pick(date.today()))
        x = tk.Label(foot, text="Close", font=FN, fg=C_GRAY, bg="white", cursor="hand2")
        x.pack(side="right"); x.bind("<Button-1>", lambda e: self.close())


def date_picker(parent, default="", width=12, **kw):
    """Date field with calendar popup. `default` is an ISO 'YYYY-MM-DD' string
    (from today() or the database); the field shows and accepts dd-mm-yyyy."""
    d = DatePicker(parent, width=width)
    if default:
        d.set_date(datetime.strptime(default, "%Y-%m-%d").date())
    return d
def combo(parent, values, default="", width=18, **kw):
    cb = ttk.Combobox(parent, values=values, font=FN,
                      state="readonly", width=width, **kw)
    if default and default in values: cb.set(default)
    elif values: cb.set(values[0])
    cb.bind("<Return>", lambda event: cb.tk_focusNext().focus())  # ← ADD THIS
    def _cb_wheel(e):
        w = cb.master
        while w is not None and not isinstance(w, ScrollFrame):
            w = w.master
        if w is not None:
            w._wheel(e)
        return "break"
    cb.bind("<MouseWheel>", _cb_wheel)
    cb.bind("<Button-4>", _cb_wheel)
    cb.bind("<Button-5>", _cb_wheel)
    return cb

def info_box(parent, text):
    f = tk.Frame(parent, bg=C_INFO_BG,
                 highlightbackground=C_INFO_BOR, highlightthickness=1)
    f.pack(fill="x", padx=4, pady=4)
    tk.Label(f, text="ℹ  " + text, bg=C_INFO_BG, fg=C_TEXT, font=FSM,
             anchor="w", justify="left", wraplength=850).pack(
        fill="x", padx=10, pady=8)

def card(parent, **kw):
    c = tk.Frame(
        parent,
        bg=C_CARD,
        padx=10,
        pady=4,
        highlightbackground=C_BORDER,
        highlightthickness=1,
        **kw
    )

    c.pack_propagate(True)
    return c
class ResponsiveForm(tk.Frame):
    """Responsive form whose height always follows its actual contents."""

    _PAD_X = 4
    _PAD_Y = 2

    def __init__(self, parent, bg=None, **kw):
        bg = bg or C_CARD
        super().__init__(parent, bg=bg, **kw)

        self._bg = bg
        self._pairs = []
        self._button_frame = None
        self._last_w = 0

        self.bind("<Configure>", self._on_resize)

    def add(self, label_text, field_widget):
        lbl = tk.Label(
            self,
            text=label_text,
            bg=self._bg,
            font=FNB,
            anchor="w"
        )

        self._pairs.append(("field", lbl, field_widget))
        return field_widget

    def add_span(self, widget):
        self._pairs.append(("span", None, widget))

    def add_btn_row(self, *widgets):
        self._pairs.append(("btn", None, list(widgets)))

    def _on_resize(self, event=None):
        w = self.winfo_width()

        if w < 10:
            return

        if w == self._last_w:
            return

        self._last_w = w
        self.after_idle(lambda: self._reflow(w))

    def _reflow(self, avail):

        # Remove previous button container completely.
        if self._button_frame is not None:
            try:
                self._button_frame.destroy()
            except Exception:
                pass

            self._button_frame = None

        # Remove existing grid placements.
        for item in self._pairs:

            if item[0] == "field":
                item[1].grid_forget()
                item[2].grid_forget()

            elif item[0] == "span":
                item[2].grid_forget()

        # Reset grid columns.
        for c in range(40):
            self.grid_columnconfigure(c, weight=0)

        row = 0
        col = 0
        x_used = 0

        for item in self._pairs:

            kind = item[0]

            # -------------------------------------------------
            # FULL WIDTH / BUTTON ROW
            # -------------------------------------------------
            if kind in ("span", "btn"):

                if col > 0:
                    row += 1
                    col = 0
                    x_used = 0

                if kind == "span":

                    item[2].grid(
                        row=row,
                        column=0,
                        columnspan=40,
                        sticky="ew",
                        padx=self._PAD_X,
                        pady=self._PAD_Y
                    )

                else:

                    self._button_frame = tk.Frame(
                        self,
                        bg=self._bg
                    )

                    self._button_frame.grid(
                        row=row,
                        column=0,
                        columnspan=40,
                        sticky="w",
                        padx=self._PAD_X,
                        pady=(self._PAD_Y, self._PAD_Y)
                    )

                    for b in item[2]:
                        b.pack(
                            in_=self._button_frame,
                            side="left",
                            padx=(0, 6)
                        )
                        b.lift() 

                row += 1
                col = 0
                x_used = 0
                continue

            # -------------------------------------------------
            # NORMAL LABEL + FIELD
            # -------------------------------------------------

            lbl = item[1]
            fld = item[2]

            lbl_w = len(lbl["text"]) * 7 + self._PAD_X * 2

            try:
                fld_w = max(
                    100,
                    int(fld["width"]) * 8
                )
            except Exception:
                fld_w = 140

            pair_w = (
                lbl_w +
                fld_w +
                self._PAD_X * 2
            )

            if x_used + pair_w > avail and col > 0:
                row += 1
                col = 0
                x_used = 0

            gc = col * 2

            lbl.grid(
                row=row,
                column=gc,
                sticky="w",
                padx=(self._PAD_X, 4),
                pady=self._PAD_Y
            )

            fld.grid(
                row=row,
                column=gc + 1,
                sticky="ew",
                padx=(0, self._PAD_X),
                pady=self._PAD_Y
            )

            self.grid_columnconfigure(
                gc + 1,
                weight=1
            )

            x_used += pair_w
            col += 1

        self.update_idletasks()

        # Tell the parent card to recalculate its own height.
        parent = self.master
        if parent is not None:
            parent.update_idletasks()

class FormFrame(tk.Frame):
    """Responsive form layout — label+field pairs auto-flow into columns
    based on available width. Label and its field always stay on same row."""

    # approx char-width in pixels at current font
    _CHAR_W = 8
    _MIN_FIELD_W = 120   # px — minimum field width before wrapping to next row
    _PAD_X = 10
    _PAD_Y = 5

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=kw.pop("bg", C_CARD), **kw)
        self._pairs = []   # list of (label_widget, field_widget, min_field_px)
        self.bind("<Configure>", self._reflow)

    def add(self, label_text, field_widget, min_field_px=None):
        """Add a label+field pair. min_field_px overrides auto sizing."""
        lbl = tk.Label(self, text=label_text, bg=self["bg"],
                       font=FNB, anchor="w")
        if min_field_px is None:
            # estimate from field width attribute if available
            try:
                min_field_px = max(self._MIN_FIELD_W,
                                   int(field_widget["width"]) * self._CHAR_W)
            except Exception:
                min_field_px = self._MIN_FIELD_W
        self._pairs.append((lbl, field_widget, min_field_px))
        return field_widget

    def add_button(self, btn_widget):
        """Add a full-width button row at the end."""
        self._pairs.append(("btn", btn_widget, 0))

    def _reflow(self, event=None):
        available = self.winfo_width()
        if available < 10:
            self.after(50, self._reflow)
            return

        # remove all children from grid
        for lbl, fld, _ in self._pairs:
            if lbl == "btn":
                fld.grid_forget()
            else:
                lbl.grid_forget()
                fld.grid_forget()

        row = 0; col = 0; x_used = 0

        for lbl, fld, min_fp in self._pairs:
            if lbl == "btn":
                # buttons always go on their own row spanning all columns
                fld.grid(row=row + 1, column=0, columnspan=99,
                         sticky="w", padx=self._PAD_X,
                         pady=(self._PAD_Y * 2, self._PAD_Y))
                row += 2; col = 0; x_used = 0
                continue

            # estimate label width
            lbl_w = len(lbl["text"]) * self._CHAR_W + self._PAD_X * 2
            pair_w = lbl_w + min_fp + self._PAD_X * 2

            if x_used + pair_w > available and col > 0:
                # wrap to next row
                row += 1; col = 0; x_used = 0

            lbl.grid(row=row, column=col * 2,
                     sticky="w", padx=(self._PAD_X, 4),
                     pady=self._PAD_Y)
            fld.grid(row=row, column=col * 2 + 1,
                     sticky="ew", padx=(0, self._PAD_X),
                     pady=self._PAD_Y)
            self.grid_columnconfigure(col * 2 + 1, weight=1)

            x_used += pair_w
            col += 1

        self._last_row = row

def confirm(msg):
    return messagebox.askyesno("Confirm", msg)

def edit_dialog(parent, title, fields, values, on_save):
    """Popup with pre-filled fields. `fields` = [(key, label, kind, extra), ...]
    kind is "entry", "date", "combo", or "listbox" (multi-select;
    extra = list of option strings). `values` = {key: current value}
    (a list of selected strings for "listbox" fields).
    `on_save(values_dict)` runs on Save; return False to keep the box open."""
    win = tk.Toplevel(parent)
    win.title(title)
    win.configure(bg=C_BG)
    win.transient(parent.winfo_toplevel())
    win.grab_set()

    body = tk.Frame(win, bg=C_BG, padx=16, pady=16)
    body.pack(fill="both", expand=True)
    body.grid_columnconfigure(1, weight=1)

    widgets = {}
    for i, (key, label, kind, extra) in enumerate(fields):
        tk.Label(body, text=label, bg=C_BG, font=FNB, anchor="w").grid(
            row=i, column=0, sticky="nw" if kind=="listbox" else "w", padx=(0,10), pady=6)
        default = values.get(key, "")
        if kind == "date":
            w = date_picker(body, default=default or today())
        elif kind == "combo":
            w = combo(body, extra or [], default)
        elif kind == "listbox":
            w = CheckList(body, height=min(8, max(4, len(extra or []))))
            for opt in (extra or []): w.insert("end", opt)
            selected = set(default or [])
            for idx, opt in enumerate(extra or []):
                if opt in selected: w.selection_set(idx)
        else:
            w = entry(body, default)
        w.grid(row=i, column=1, sticky="ew", pady=6)
        widgets[key] = w

    def _collect():
        out = {}
        for key, _, kind, extra in fields:
            w = widgets[key]
            if kind == "date":
                out[key] = w.get_date().strftime("%Y-%m-%d")
            elif kind == "listbox":
                out[key] = [extra[i] for i in w.curselection()]
            else:
                out[key] = w.get()
        return out

    def _do_save():
        if on_save(_collect()) is not False:
            win.destroy()

    btn_row = tk.Frame(win, bg=C_BG); btn_row.pack(fill="x", padx=16, pady=(0,16))
    btn(btn_row, "Save", _do_save, bg=C_SUCCESS, hov=C_SUC_HOV).pack(side="left")
    btn(btn_row, "Cancel", win.destroy, bg=C_GRAY, hov=C_GRAY).pack(side="left", padx=(8,0))
    win.bind("<Escape>", lambda e: win.destroy())
    widgets[fields[0][0]].focus_set()

def warn(msg):
    messagebox.showwarning("Warning", msg)

def err(msg):
    messagebox.showerror("Error", msg)

def hline(parent, color=C_BORDER):
    tk.Frame(parent, bg=color, height=1).pack(fill="x", pady=6)

# ── Scrollable frame ──────────────────────────────────────────────────────────
_THEME_READY = False

def _style_scrollbars():
    global _THEME_READY
    style = ttk.Style()
    if not _THEME_READY:
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", takefocus=0) 
        _THEME_READY = True

class ScrollFrame(tk.Frame):
    """A frame with a vertical scrollbar. Put widgets in .inner.

    Scrolling works anywhere the mouse is over this panel — including while
    hovering a label, entry, combo, or any nested widget — not just the bare
    canvas background, and it also works with Linux mouse wheels/trackpads.
    """
    def __init__(self, parent, bg=C_BG, **kw):
        super().__init__(parent, bg=bg, **kw)
        _style_scrollbars()
        canvas = tk.Canvas(self, bg=bg, highlightthickness=0)
        vsb    = ttk.Scrollbar(self, orient="vertical", command=canvas.yview,
                               style="Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(canvas, bg=bg)
        self._win  = canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(self._win, width=e.width))

        # mouse-wheel — bound/unbound globally while the cursor is anywhere
        # over this panel (via <Enter>/<Leave> on the outer container), so
        # scrolling works no matter which nested widget — label, entry,
        # combo, tree — the mouse happens to be sitting over. This also
        # means it keeps working for widgets added to .inner *after* this
        # ScrollFrame is built (the normal case for every panel).
        def _wheel(e):
            try:
                if getattr(e, "num", None) == 4:        # Linux scroll up
                    canvas.yview_scroll(-3, "units")
                elif getattr(e, "num", None) == 5:      # Linux scroll down
                    canvas.yview_scroll(3, "units")
                else:                                    # Windows / macOS
                    canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
            except Exception:
                pass

        self._wheel = _wheel
        def _bind_wheel(_e=None):
            canvas.bind_all("<MouseWheel>", _wheel)
            canvas.bind_all("<Button-4>", _wheel)
            canvas.bind_all("<Button-5>", _wheel)

        def _unbind_wheel(_e=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        self.bind("<Enter>", _bind_wheel)
        self.bind("<Leave>", _unbind_wheel)
        self.bind("<Destroy>", _unbind_wheel, add="+")

# ── Checkbox list (multi-select with visible tick boxes + scrollbar) ─────────
class CheckList(tk.Frame):
    """Drop-in replacement for tk.Listbox(selectmode="multiple").
    Every row has a visible checkbox, and a scrollbar is always shown."""
    def __init__(self, parent, height=5, **kw):
        super().__init__(parent, bg=C_CARD,
                         highlightbackground=C_BORDER, highlightthickness=1)
        import tkinter.font as tkfont
        row_h = tkfont.Font(font=FN).metrics("linespace") + 12
        self._rows = []                     # [(BooleanVar, Checkbutton, text)]
        self._canvas = tk.Canvas(self, bg=C_CARD, highlightthickness=0,
                                 height=max(1, int(height)) * row_h)
        self._vsb = ttk.Scrollbar(self, orient="vertical",
                                  command=self._canvas.yview,
                                  style="Vertical.TScrollbar")
        self._canvas.configure(yscrollcommand=self._vsb.set)
        self._vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self._inner = tk.Frame(self._canvas, bg=C_CARD)
        self._win = self._canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>", lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(
            self._win, width=e.width))
        self._bind_wheel(self._canvas)
        self._bind_wheel(self._inner)

    # wheel: scroll this list first; at its top/bottom (or if nothing to
    # scroll) let the event pass on to the page behind it.
    def _bind_wheel(self, w):
        w.bind("<MouseWheel>", self._on_wheel)
        w.bind("<Button-4>", self._on_wheel)
        w.bind("<Button-5>", self._on_wheel)

    def _on_wheel(self, e):
        first, last = self._canvas.yview()
        if first <= 0.0 and last >= 1.0:
            return
        up = getattr(e, "num", None) == 4 or getattr(e, "delta", 0) > 0
        if (up and first <= 0.0) or (not up and last >= 1.0):
            return
        step = -2 if getattr(e, "num", None) == 4 else (
            2 if getattr(e, "num", None) == 5 else int(-1 * (e.delta / 120)) * 2)
        self._canvas.yview_scroll(step, "units")
        return "break"

    # ── Listbox-compatible API ──────────────────────────────────────────────
    def insert(self, index, text):
        var = tk.BooleanVar(value=False)
        cb = tk.Checkbutton(self._inner, text=text, variable=var, anchor="w",
                            bg=C_CARD, activebackground=C_SEL, font=FN,
                            selectcolor="white", cursor="hand2")
        cb.pack(fill="x", anchor="w", padx=4, pady=1)
        self._bind_wheel(cb)
        self._rows.append((var, cb, text))

    def delete(self, first=0, last=None):
        for _, cb, _ in self._rows:
            cb.destroy()
        self._rows = []
        self._canvas.yview_moveto(0)

    def curselection(self):
        return tuple(i for i, (v, _, _) in enumerate(self._rows) if v.get())

    def selection_set(self, i):
        self._rows[i][0].set(True)

    def selection_clear(self, first=0, last=None):
        for v, _, _ in self._rows:
            v.set(False)

    def size(self):
        return len(self._rows)

    def get(self, i):
        return self._rows[i][2]

# ── Treeview table ────────────────────────────────────────────────────────────
def make_tree(parent, cols, heights=12, col_widths=None, stretch_col=None):
    """Returns (frame, tree). Hidden columns have width 0."""
    frame = tk.Frame(parent, bg=C_CARD)
    style = ttk.Style()
    style.theme_use("clam")          # vista theme ignores our selected-row colors, clam respects them
    style.configure("Treeview.Heading", background=C_TH_BG,
                    foreground=C_TH_FG, font=FNB, relief="flat",
                    padding=(8, 8))
    style.map("Treeview.Heading", background=[("active", C_PRI_HOV)])
    style.configure("Treeview", background="white",
                    fieldbackground="white", foreground=C_TEXT,
                rowheight=28, font=FN, borderwidth=0, relief="flat")
    style.map("Treeview",
            background=[("selected", C_SEL)],
            foreground=[("selected", C_TEXT)])   # keep text dark instead of white on select
    _style_scrollbars()

    tree = ttk.Treeview(frame, columns=cols, show="headings", height=heights)
    tree.tag_configure("alt", background=C_ROW_ALT)
    tree.tag_configure("selbold", font=FNB)
    def _bold_selected(e=None, tv=tree):
        sel = set(tv.selection())
        for iid in tv.get_children():
            tags = [t for t in tv.item(iid, "tags") if t != "selbold"]
            if iid in sel: tags.append("selbold")
            tv.item(iid, tags=tags)
    tree.bind("<<TreeviewSelect>>", _bold_selected)

    for col in cols:
        w = 0 if col.startswith("_") else (
            col_widths.get(col, 120) if col_widths else 120)
        if col == "Date":
            w = max(w, 115)   # room for "16 Sep 2026" in the display format

        tree.heading(col, text="" if col.startswith("_") else col, anchor="w")
        tree.column(col, width=w, stretch=(col == stretch_col), anchor="w",
                    minwidth=0 if col.startswith("_") else 30)

    vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview,
                        style="Vertical.TScrollbar")
    hsb = ttk.Scrollbar(frame, orient="horizontal",  command=tree.xview,
                        style="Horizontal.TScrollbar")
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)

    # mouse-wheel over a tree that lives inside a ScrollFrame: let the tree
    # scroll itself first, only bubble up to the outer panel once the tree
    # has no further room to scroll (so nested tables and the outer page
    # both scroll naturally instead of fighting each other).
    def _tree_wheel(e):
        first, last = tree.yview()
        going_up = getattr(e, "num", None) == 4 or getattr(e, "delta", 0) > 0
        if (going_up and first <= 0.0) or (not going_up and last >= 1.0):
            return  # let the event propagate to the ScrollFrame beneath
        step = -3 if getattr(e, "num", None) == 4 else (3 if getattr(e, "num", None) == 5 else int(-1 * (e.delta / 120)))
        tree.yview_scroll(step, "units")
        return "break"
    tree.bind("<MouseWheel>", _tree_wheel)
    tree.bind("<Button-4>", _tree_wheel)
    tree.bind("<Button-5>", _tree_wheel)
    return frame, tree

def tree_fill(tree, rows):
    """Clear and refill a treeview with alternating row colours.

    Purely cosmetic: if the tree has a column literally named "Date", that
    cell is reformatted for display (see fmt_date_display). This never
    touches the database — it only changes what's printed in the cell.
    """
    tree.delete(*tree.get_children())
    cols = tree.cget("columns")
    date_idx = cols.index("Date") if "Date" in cols else None
    for i, row in enumerate(rows):
        if date_idx is not None:
            row = list(row)
            row[date_idx] = fmt_date_display(row[date_idx])
        tag = "alt" if i % 2 else ""
        tree.insert("", "end", values=row, tags=(tag,))

def action_click(tv, event, col_id, on_edit, on_delete):
    """Call from a <ButtonRelease-1> handler of a table whose Action column
    shows 'Edit | Delete'.  Left half of that cell = Edit, right half = Delete.
    on_edit / on_delete both receive the row's values (last value = hidden id)."""
    if tv.identify_region(event.x, event.y) != "cell":
        return
    if tv.identify_column(event.x) != col_id:
        return
    row = tv.identify_row(event.y)
    if not row:
        return
    bbox = tv.bbox(row, col_id)
    if not bbox:
        return
    vals = tv.item(row, "values")
    if event.x - bbox[0] < bbox[2] // 2:
        on_edit(vals)
    else:
        on_delete(vals)

class EditDialog(tk.Toplevel):
    """fields = list of dicts:
         {"key": "date", "label": "Date:", "kind": "entry"|"combo"|"list"|"label",
          "value": <initial value>, "values": [options for combo/list]}
       on_save(data_dict) must return None on success, or an error text
       (the window then stays open and shows that text)."""
 
    def __init__(self, parent, title, fields, on_save):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=C_CARD)
        self.resizable(False, False)
        self.transient(parent.winfo_toplevel())
        self._on_save = on_save
        self._fields = {}
 
        body = tk.Frame(self, bg=C_CARD, padx=20, pady=16)
        body.pack(fill="both", expand=True)
        body.grid_columnconfigure(1, weight=1)
        tk.Label(body, text=title, font=FSH, bg=C_CARD, fg=C_PRIMARY).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
 
        for r, f in enumerate(fields, start=1):
            kind = f.get("kind", "entry")
            val = f.get("value")
            tk.Label(body, text=f["label"], bg=C_CARD, font=FNB,
                     anchor="w").grid(row=r, column=0, sticky="nw", pady=4)
            if kind == "entry":
                w = entry(body, "" if val is None else str(val), width=32)
            elif kind == "combo":
                values = list(f["values"])
                w = ttk.Combobox(body, values=values, font=FN,
                                 state="readonly", width=30)
                if val not in (None, ""):
                    w.set(val)
                elif values:
                    w.set(values[0])
            elif kind == "list":
                w = CheckList(body, height=7)
                for i, name in enumerate(f["values"]):
                    w.insert("end", name)
                    if name in (val or []):
                        w.selection_set(i)
            else:   # "label" = read-only text
                w = tk.Label(body, text="" if val is None else str(val),
                             bg=C_CARD, fg=C_GRAY, font=FN, anchor="w",
                             justify="left", wraplength=300)
            w.grid(row=r, column=1, sticky="ew", padx=(12, 0), pady=4)
            self._fields[f["key"]] = (kind, w, f)
 
        bar = tk.Frame(body, bg=C_CARD)
        bar.grid(row=len(fields) + 1, column=0, columnspan=2,
                 sticky="e", pady=(14, 0))
        btn(bar, "Save Changes", self._save,
            bg=C_SUCCESS, hov=C_SUC_HOV).pack(side="left")
        btn(bar, "Cancel", self.destroy,
            bg=C_GRAY, hov=C_GRA_HOV).pack(side="left", padx=(8, 0))
        self.bind("<Escape>", lambda e: self.destroy())
 
        self.update_idletasks()
        root = parent.winfo_toplevel()
        x = root.winfo_rootx() + max(0, (root.winfo_width() - self.winfo_width()) // 2)
        y = root.winfo_rooty() + max(0, (root.winfo_height() - self.winfo_height()) // 3)
        self.geometry(f"+{x}+{y}")
        self.after(80, self._take_focus)
 
    def _take_focus(self):
        try:
            self.grab_set()
            self.focus_force()
        except tk.TclError:
            pass
 
    def _save(self):
        data = {}
        for key, (kind, w, f) in self._fields.items():
            if kind in ("entry", "combo"):
                data[key] = w.get().strip()
            elif kind == "list":
                data[key] = [f["values"][i] for i in w.curselection()]
        try:
            msg = self._on_save(data)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save changes:\n\n{e}", parent=self)
            return
        if msg:
            messagebox.showwarning("Warning", msg, parent=self)
        else:
            self.destroy()
 

def make_action_list(parent, headers, col_widths, rows, height_rows=10):
    """Header row + one row per record, each with its own Edit/Delete buttons,
    directly inside `parent` (which already lives in the panel's scroll area).
    Every column has ONE pixel width shared by the header and all rows, so the
    header text always sits exactly above its values.
    Display-only: a "Date" header gets fmt_date_display(); the record each
    button acts on (row["buttons"]) is a separate, untouched copy."""
    from tkinter import font as tkfont
    f_norm, f_bold = tkfont.Font(font=FN), tkfont.Font(font=FNB)
    zero = f_norm.measure("0")

    def _px(h):
        w = col_widths.get(h, 15)
        w = max(w, 12) if h == "Date" else w
        return max(w * zero, f_bold.measure(h) + 4)      # never narrower than its header

    px = [_px(h) for h in headers]

    def _setup(frame):
        for i, w in enumerate(px):
            frame.grid_columnconfigure(i, minsize=w + 12)   # +12 = padx 6 each side
        frame.grid_columnconfigure(len(px), weight=1)

    hdr = tk.Frame(parent, bg=C_TH_BG); hdr.pack(fill="x"); _setup(hdr)
    for i, h in enumerate(headers):
        tk.Label(hdr, text=h, bg=C_TH_BG, fg=C_TH_FG, font=FNB, anchor="w").grid(
            row=0, column=i, sticky="w", padx=6, pady=8)
    tk.Label(hdr, text="Action", bg=C_TH_BG, fg=C_TH_FG, font=FNB, anchor="w").grid(
        row=0, column=len(px), sticky="w", padx=6, pady=8)

    date_idx = headers.index("Date") if "Date" in headers else None

    for i, row in enumerate(rows):
        rowbg = C_ROW_ALT if i % 2 else "white"
        rf = tk.Frame(parent, bg=rowbg); rf.pack(fill="x"); _setup(rf)
        cell_widgets = []
        vals = list(row["values"])
        if date_idx is not None and date_idx < len(vals):
            vals[date_idx] = fmt_date_display(vals[date_idx])
        for c, (h, val) in enumerate(zip(headers, vals)):
            w = tk.Label(rf, text=val, bg=rowbg, font=FN, anchor="w", justify="left",
                         wraplength=0 if h == "Date" else px[c])
            w.grid(row=0, column=c, sticky="w", padx=6, pady=6)
            cell_widgets.append(w)
        bf = tk.Frame(rf, bg=rowbg); bf.grid(row=0, column=len(px), sticky="w", padx=6, pady=4)
        for label, cmd, color in row["buttons"]:
            btn(bf, label, cmd, bg=color, hov=color, padx=10, pady=4).pack(side="left", padx=2)

        def _on_enter(_e, rf=rf, cells=cell_widgets, bf=bf):
            rf.config(bg=C_SEL)
            for w in cells: w.config(bg=C_SEL)
            bf.config(bg=C_SEL)
        def _on_leave(_e, rf=rf, cells=cell_widgets, bf=bf, bg=rowbg):
            rf.config(bg=bg)
            for w in cells: w.config(bg=bg)
            bf.config(bg=bg)
        rf.bind("<Enter>", _on_enter); rf.bind("<Leave>", _on_leave)
        for w in cell_widgets:
            w.bind("<Enter>", _on_enter); w.bind("<Leave>", _on_leave)
    parent.update_idletasks()

# ═══════════════════════════════════════════════════════════════════════════════
#  ALERT OVERLAY  (bottom-right, closeable)
# ═══════════════════════════════════════════════════════════════════════════════
class AlertOverlay(tk.Toplevel):
    """Small stacked alert box, similar to a compact error notification."""

    WIDTH = 320
    HEIGHT = 105
    GAP = 3

    def __init__(self, master, alert, stack=None):
        super().__init__(master)

        self.master_window = master
        self.alert = alert
        self.stack = stack

        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=C_DANGER)


        master.bind("<Configure>", lambda e: self._reposition())

        # -------------------------
        # Header
        # -------------------------
        hdr = tk.Frame(self, bg=C_DANGER)
        hdr.pack(fill="x", padx=10, pady=(6, 2))

        tk.Label(
            hdr,
            text="⚠  Low Stock Alert",
            bg=C_DANGER,
            fg="white",
            font=("Arial", 10, "bold")
        ).pack(side="left")

        tk.Button(
            hdr,
            text="×",
            bg=C_DANGER,
            fg="white",
            relief="flat",
            font=("Arial", 14, "bold"),
            cursor="hand2",
            bd=0,
            padx=2,
            pady=0,
            command=self.close
        ).pack(side="right")

        # -------------------------
        # Alert body
        # -------------------------
        name = alert.get("name", "")
        qty = alert.get("qty", 0)
        mq = alert.get("min_qty", 0)
        stage = alert.get("stage", "")

        msg = (
            f"{name} [{stage}]\n"
            f"Stock: {qty} pcs  |  Min: {mq} pcs"
        )

        tk.Label(
            self,
            text=msg,
            bg=C_DANGER,
            fg="white",
            font=FSM,
            justify="left",
            anchor="w"
        ).pack(
            fill="x",
            padx=12,
            pady=(2, 8)
        )

    def _reposition(self):
        """Position this alert according to its position in the stack."""
        try:
            if self.stack is None:
                index = 0
            else:
                index = self.stack.index(self)

            mw = self.master_window.winfo_width()
            mh = self.master_window.winfo_height()
            mx = self.master_window.winfo_rootx()
            my = self.master_window.winfo_rooty()

            x = mx + mw - self.WIDTH - 20

            # Bottom alert sits at the bottom.
            # Higher alerts move upward.
            y = (
                my
                + mh
                - self.HEIGHT
                - 20
                - index * (self.HEIGHT + self.GAP)
            )

            self.geometry(
                f"{self.WIDTH}x{self.HEIGHT}+{x}+{y}"
            )

        except Exception:
            pass

    def close(self):
        """Remove this alert and reposition the remaining alerts."""
        if self.stack is not None:
            if self in self.stack:
                self.stack.remove(self)

            self.destroy()

            # Move remaining alerts down
            for alert in self.stack:
                alert._reposition()
        else:
            self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  APPLICATION ROOT
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Registers System")
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = min(int(sw * 0.85), 1600)
        h = min(int(sh * 0.85), 1000)
        self.geometry(f"{w}x{h}")
        self.configure(bg=C_BG)
        self.minsize(int(sw * 0.6), int(sh * 0.6))
        self.current_user = None
        self.current_role = None
        self.alert_overlays = []
        style = ttk.Style(self)
        style.theme_use("clam")
        self.bind_all("<Return>", self._enter_click)
        self.bind_all("<KP_Enter>", self._enter_click)

        self.container = tk.Frame(self, bg=C_BG)
        self.container.pack(fill="both", expand=True)
        self._show_login()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        _close_db()
        self.destroy()

    # ── navigation ────────────────────────────────────────────────────────────
    def _clear(self):
        for w in self.container.winfo_children():
            w.destroy()

    def _enter_click(self, event):
        """Treat Enter as clicking the currently focused button."""
        widget = self.focus_get()
        if isinstance(widget, (tk.Button, ttk.Button)):
            widget.invoke()
            return "break"

    def _show_login(self):
        self._clear()
        LoginFrame(self.container, self).pack(fill="both", expand=True)

    def show_shell(self):
        self._clear()
        Shell(self.container, self).pack(fill="both", expand=True)
        self.after(400, self._check_alerts)

    def logout(self):
        if self.current_user:
            log_action(self.current_user, "LOGOUT", "User logged out")
        self.current_user = self.current_role = None
        for overlay in self.alert_overlays[:]:
            try:
                overlay.destroy()
            except Exception:
                pass

        self.alert_overlays.clear()
        self._show_login()

    # ── alerts ────────────────────────────────────────────────────────────────
    def _check_alerts(self):
        conn = get_conn()
        alerts = _get_low_stock_alerts(conn.cursor())
        conn.close()
        if alerts:
            self.show_alerts(alerts)

    
    def show_alerts(self, alerts=None):
        if alerts is None:
            conn = get_conn()
            alerts = _get_low_stock_alerts(conn.cursor())
            conn.close()

        # Close any existing alert boxes
        for overlay in self.alert_overlays[:]:
            try:
                overlay.destroy()
            except Exception:
                pass

        self.alert_overlays.clear()

        if not alerts:
            messagebox.showinfo(
                "Low Stock Alerts",
                "No low-stock items found."
            )
            return

        # Create one box for each alert
        for alert in alerts:
            overlay = AlertOverlay(
                self,
                alert,
                self.alert_overlays
            )

            self.alert_overlays.append(overlay)
            overlay._reposition()



# ═══════════════════════════════════════════════════════════════════════════════
#  LOGIN
# ═══════════════════════════════════════════════════════════════════════════════
class LoginFrame(tk.Frame):
    def __init__(self, parent, app: App):
        super().__init__(parent, bg=C_BG)
        self.app = app
        # a thin "shadow" frame behind the card gives a subtle sense of depth
        shadow = tk.Frame(self, bg="#d7dcec")
        shadow.place(relx=0.5, rely=0.5, anchor="center", x=4, y=4)
        c = tk.Frame(self, bg=C_CARD, padx=48, pady=44,
                     highlightbackground=C_BORDER, highlightthickness=1)
        c.place(relx=0.5, rely=0.5, anchor="center")
        shadow.lower(c)
        tk.Label(c, text="⚙", font=(_FONT_FAMILY, 28), bg=C_CARD,
                 fg=C_PRIMARY).grid(row=0, column=0, columnspan=2, pady=(0, 4))
        tk.Label(c, text="Registers System", font=(_FONT_FAMILY, 20, "bold"),
                 bg=C_CARD, fg=C_TEXT).grid(row=1, column=0, columnspan=2,
                                             pady=(0, 24))
        tk.Label(c, text="Username", bg=C_CARD, fg=C_MUTED, font=FSM).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(0, 2))
        self.uname = entry(c, width=26)
        self.uname.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 12))
        tk.Label(c, text="Password", bg=C_CARD, fg=C_MUTED, font=FSM).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(0, 2))
        self.pwd = entry(c, width=26)
        self.pwd.config(show="*")
        self.pwd.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        self.err = tk.Label(c, text="", fg=C_DANGER, bg=C_CARD, font=FSM)
        self.err.grid(row=6, column=0, columnspan=2, pady=4, sticky="w")
        btn(c, "Log In", self.do_login).grid(row=7, column=0,
                                             columnspan=2, sticky="ew",
                                             pady=(14, 0))
        self.pwd.bind("<Return>", lambda e: self.do_login())
        self.uname.focus_set()

    def do_login(self):
        row = qry("SELECT * FROM users WHERE username=? AND password=?",
                  (self.uname.get().strip(), self.pwd.get().strip()),
                  fetch="one")
        if row:
            self.app.current_user = row[1]
            self.app.current_role = row[3]
            log_action(row[1], "LOGIN", "User logged into the system")
            self.app.show_shell()
        else:
            self.err.config(text="Invalid credentials.")


# ═══════════════════════════════════════════════════════════════════════════════
#  SHELL  (topbar + sidebar + content)
# ═══════════════════════════════════════════════════════════════════════════════
class Shell(tk.Frame):
    ICONS = {"Add Items": "➕", "Registers": "📒", "Audit Logs": "🕘",
             "Manage Users": "👤", "Logout": "⏻"}

    def __init__(self, parent, app: App):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._nav_buttons = {}
        self._active = None

        # ── top bar ──────────────────────────────────────────────────────────
        top = tk.Frame(self, bg=C_TOPBAR, height=50,
                       highlightbackground=C_BORDER, highlightthickness=1)
        top.pack(side="top", fill="x")
        tk.Label(top, text="Registers System", bg=C_TOPBAR, fg=C_PRIMARY,
                 font=FSH).pack(side="left", padx=(18, 24), pady=12)
        tk.Label(top,
                 text=f"{app.current_user}  ·  {app.current_role.capitalize()}",
                 bg=C_TOPBAR, fg=C_MUTED, font=FN).pack(side="left", pady=12)
        # alert bell always visible in topbar
        btn(top, "⚠ Alerts", app.show_alerts,
            bg=C_DANGER, hov=C_DAN_HOV).pack(side="right", padx=14, pady=9)

        # ── body ─────────────────────────────────────────────────────────────
        body = tk.Frame(self, bg=C_BG)
        body.pack(fill="both", expand=True)

        # sidebar (dark, modern) — the active section is highlighted so it's
        # always clear at a glance which screen is open
        sb = tk.Frame(body, bg=C_SIDEBAR, width=200)
        sb.pack(side="left", fill="y")
        sb.pack_propagate(False)
        tk.Frame(sb, bg=C_SIDEBAR, height=8).pack(fill="x")

        self.content = tk.Frame(body, bg=C_BG)
        self.content.pack(side="left", fill="both", expand=True,
                          padx=16, pady=16)

        nav = [
            ("Add Items",    self._add_items),
            ("Registers",    self._registers),
            ("Audit Logs",   self._audit),
        ]
        if app.current_role == "admin":
            nav.append(("Manage Users", self._users))

        for lbl_text, cmd in nav:
            self._make_nav_button(sb, lbl_text, cmd)

        tk.Frame(sb, bg=C_SIDEBAR_HOV, height=1).pack(fill="x", pady=8, padx=14)
        self._make_nav_button(sb, "Logout", app.logout, active_key=False)

        self._add_items()

    def _make_nav_button(self, sb, lbl_text, cmd, active_key=True):
        b = tk.Button(sb, text=f"  {self.ICONS.get(lbl_text,'•')}   {lbl_text}",
                      command=cmd, bg=C_SIDEBAR, fg=C_SIDEBAR_TEXT,
                      activebackground=C_SIDEBAR_ACT, activeforeground="white",
                      relief="flat", anchor="w", font=FNB, padx=8, pady=12,
                      cursor="hand2", bd=0, highlightthickness=0)
        b.pack(fill="x", padx=8, pady=2)
        if active_key:
            def _enter(e, x=b):
                if self._active is not x: x.config(bg=C_SIDEBAR_HOV)
            def _leave(e, x=b):
                if self._active is not x: x.config(bg=C_SIDEBAR)
            b.bind("<Enter>", _enter); b.bind("<Leave>", _leave)
            self._nav_buttons[lbl_text] = b
        else:
            b.bind("<Enter>", lambda e, x=b: x.config(bg=C_DANGER))
            b.bind("<Leave>", lambda e, x=b: x.config(bg=C_SIDEBAR))
        return b

    def _set_active(self, lbl_text):
        if self._active is not None:
            self._active.config(bg=C_SIDEBAR, fg=C_SIDEBAR_TEXT)
        b = self._nav_buttons.get(lbl_text)
        if b:
            b.config(bg=C_SIDEBAR_ACT, fg="white")
        self._active = b

    def _swap(self, cls, *a):
        for w in self.content.winfo_children():
            w.destroy()
        cls(self.content, self.app, *a).pack(fill="both", expand=True)

    def _add_items(self): self._set_active("Add Items"); self._swap(AddItemsScreen)
    def _registers(self): self._set_active("Registers"); self._swap(RegistersScreen)
    def _audit(self):     self._set_active("Audit Logs"); self._swap(AuditLogsScreen)
    def _users(self):     self._set_active("Manage Users"); self._swap(ManageUsersScreen)


# ═══════════════════════════════════════════════════════════════════════════════
#  ADD ITEMS SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
class AddItemsScreen(tk.Frame):
    SECS = ["Items", "Vendors", "Unit of Measurement",
            "Box Pack Items", "Rate List"]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        tk.Label(self, text="Add Items & Management",
                 font=FH, bg=C_BG, fg="#222").pack(anchor="w",
                                                     pady=(0, 8))
        top = card(self)
        top.pack(fill="x", pady=(0, 10))
        tk.Label(top, text="Select Category:", bg=C_CARD,
                 font=FNB).pack(side="left")
        self.sec = combo(top, self.SECS, self.SECS[0], width=32)
        self.sec.pack(side="left", padx=10)
        self.sec.bind("<<ComboboxSelected>>", lambda e: self._render())
        self.body = tk.Frame(self, bg=C_BG)
        self.body.pack(fill="both", expand=True)
        self._render()

    def _render(self):
        for w in self.body.winfo_children():
            w.destroy()
        {
            "Items":               ItemsPanel,
            "Vendors":             VendorsPanel,
            "Unit of Measurement": UnitsPanel,
            "Box Pack Items":      BoxPackPanel,
            "Rate List":           RateListPanel,
        }[self.sec.get()](self.body, self.app).pack(fill="both", expand=True)


# ─────────────────────────────────────────────────────────────────────────────
#  ITEMS PANEL
# ─────────────────────────────────────────────────────────────────────────────
class ItemsPanel(tk.Frame):
    CATS = ["Brass Item", "Miscellaneous Item"]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._editing_id = None

        sf = ScrollFrame(self); sf.pack(fill="both", expand=True)
        p = sf.inner

        c = card(p); c.pack(fill="x", pady=(0,10))
        f = ResponsiveForm(c); f.pack(fill="both", expand=True)
        self.ne = f.add("Item Name:", entry(f))
        self.cc = f.add("Category:", combo(f, self.CATS, self.CATS[0]))
        self.cc.bind("<<ComboboxSelected>>", lambda e: self._toggle_fields())
        self.mq = f.add("Min Qty (Total Pcs):", entry(f, "0"))

        self._brass_frame = ResponsiveForm(c, bg=C_CARD)
        self._brass_frame.pack(fill="x")
        self._brass_entries = {}
        for key, lbl_text in [("mcq","Min Casting Qty (Pcs):"),
                               ("mmq","Min Machined Qty (Pcs):"),
                               ("maq","Min Assembly Qty (Pcs):")]:
            self._brass_entries[key] = self._brass_frame.add(lbl_text, entry(self._brass_frame, "0"))

        self._misc_frame = ResponsiveForm(c, bg=C_CARD)
        self._misc_frame.pack(fill="x")
        self.mpq = self._misc_frame.add("Min Purchase Qty (Pcs):", entry(self._misc_frame, "0"))

        self._btn_bar = tk.Frame(c, bg=C_CARD)
        self._btn_bar.pack(anchor="w", pady=(10,0))
        self.sb = btn(self._btn_bar, "Save Item", self.save, bg=C_SUCCESS, hov=C_SUC_HOV)
        self.cb_ = btn(self._btn_bar, "Cancel", self.cancel, bg=C_GRAY, hov=C_GRA_HOV)
        self.sb.pack(side="left")
        self.cb_.pack(side="left", padx=(8,0))
        self._toggle_fields()

        tk.Label(p, text="Existing Items", font=FSH, bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG)
        self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def _toggle_fields(self):
        cat = self.cc.get()
        if cat == "Brass Item":
            self._misc_frame.pack_forget()
            self._brass_frame.pack(fill="x", before=self._btn_bar)
        else:
            self._brass_frame.pack_forget()
            self._misc_frame.pack(fill="x", before=self._btn_bar)

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        rows = qry("""SELECT id, item_name, item_category,
                      COALESCE(min_quantity,0),
                      COALESCE(min_casting_quantity,0),
                      COALESCE(min_machined_quantity,0),
                      COALESCE(min_assembly_quantity,0),
                      COALESCE(min_purchase_quantity,0)
                      FROM master_items ORDER BY item_name""", fetch="all") or []
        data = []
        for r in rows:
            vals = (r[1], r[2] or "Brass Item", int(r[3]), int(r[4]),
                    int(r[5]), int(r[6]), int(r[7]), "", r[0])
            data.append({
                "values": [vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], vals[6]],
                "buttons": [
                    ("Edit",   lambda v=vals: self._load_edit(v), C_GRAY),
                    ("Delete", lambda v=vals: self._delete(v),   "#dc3545"),
                ]
            })
        make_action_list(self.list_holder,
                          ["Item Name","Category","Min Qty","Min Cast","Min Mach","Min Asm","Min Pur"],
                          {"Item Name":18,"Category":14,"Min Qty":8,"Min Cast":8,
                           "Min Mach":8,"Min Asm":8,"Min Pur":8},
                          data)

    # def _on_click(self, event):
    #     col = self.tv.identify_column(event.x)
    #     row = self.tv.identify_row(event.y)
    #     if not row: return
    #     vals = self.tv.item(row, "values")
    #     if col == "#8":  # Action column
    #         # determine sub-click: rough split
    #         x = event.x - self.tv.bbox(row, "#8")[0] if self.tv.bbox(row, "#8") else 0
    #         bbox = self.tv.bbox(row, "#8")
    #         if bbox:
    #             mid = bbox[2] // 2
    #             rel_x = event.x - bbox[0]
    #             if rel_x < mid:
    #                 self._load_edit(vals)
    #             else:
    #                 self._delete(vals)

    def _load_edit(self, vals):
        if self.app.current_role != "admin":
            warn("Only admins can edit items."); return
        iid = vals[8]
        def _save(v):
            name = v["name"].strip()
            cat = v["category"] or "Brass Item"
            if not name: warn("Item Name required."); return False
            try:
                mq  = float(v["min_qty"] or 0); mcq = float(v["min_cast"] or 0)
                mmq = float(v["min_mach"] or 0); maq = float(v["min_asm"] or 0)
                mpq = float(v["min_pur"] or 0)
            except ValueError:
                err("Qty fields must be numbers."); return False
            if cat == "Miscellaneous Item": mcq = mmq = maq = 0
            else: mpq = 0

            old_row = qry("SELECT item_name FROM master_items WHERE id=?", (iid,), fetch="one")
            old_name = old_row[0] if old_row else None

            try:
                qry("""UPDATE master_items SET item_name=?,item_category=?,
                    min_quantity=?,min_casting_quantity=?,
                    min_machined_quantity=?,min_assembly_quantity=?,
                    min_purchase_quantity=? WHERE id=?""",
                    (name, cat, mq, mcq, mmq, maq, mpq, iid), commit=False)
            except sqlite3.IntegrityError:
                err("Item name already exists."); return False

            if old_name and old_name != name:
                for table, col in [
                    ("casting_in","item"),("machining_register","item"),
                    ("rejection_register","item"),("plating_out","item"),
                    ("plating_in","item"),("grinder_register","item"),
                    ("opening_stock","item"),("fg_deductions","item"),
                ]:
                    try: qry(f"UPDATE {table} SET {col}=? WHERE {col}=?", (name, old_name), commit=False)
                    except: pass

            qry("SELECT 1", commit=True)
            log_action(self.app.current_user, "UPDATE ITEM", f"Renamed '{old_name}' → '{name}'")
            self.refresh()
        edit_dialog(self, "Edit Item",
                    [("name","Item Name:","entry",None),
                    ("category","Category:","combo",self.CATS),
                    ("min_qty","Min Qty (Total Pcs):","entry",None),
                    ("min_cast","Min Casting Qty (Pcs):","entry",None),
                    ("min_mach","Min Machined Qty (Pcs):","entry",None),
                    ("min_asm","Min Assembly Qty (Pcs):","entry",None),
                    ("min_pur","Min Purchase Qty (Pcs):","entry",None)],
                    {"name": vals[0], "category": vals[1], "min_qty": vals[2],
                    "min_cast": vals[3], "min_mach": vals[4], "min_asm": vals[5],
                    "min_pur": vals[6]},
                    _save)

    def save(self):
        name = self.ne.get().strip()
        cat  = self.cc.get() or "Brass Item"
        if not name: warn("Item Name required."); return
        try:
            mq   = float(self.mq.get() or 0)
            mcq  = float(self._brass_entries["mcq"].get() or 0)
            mmq  = float(self._brass_entries["mmq"].get() or 0)
            maq  = float(self._brass_entries["maq"].get() or 0)
            mpq  = float(self.mpq.get() or 0)
        except ValueError:
            err("Qty fields must be numbers."); return

        if cat == "Miscellaneous Item":
            mcq = mmq = maq = 0
        else:
            mpq = 0

        if self._editing_id is None:
            try:
                qry("""INSERT INTO master_items
                    (item_name,item_category,min_quantity,
                        min_casting_quantity,min_machined_quantity,
                        min_assembly_quantity,min_purchase_quantity)
                    VALUES (?,?,?,?,?,?,?)""",
                    (name, cat, mq, mcq, mmq, maq, mpq), commit=True)
                log_action(self.app.current_user, "CREATE ITEM", f"Added '{name}' ({cat})")
            except sqlite3.IntegrityError:
                err("Item name already exists."); return
        else:
            old_row = qry("SELECT item_name FROM master_items WHERE id=?", (self._editing_id,), fetch="one")
            old_name = old_row[0] if old_row else None

            try:
                qry("""UPDATE master_items SET item_name=?,item_category=?,
                    min_quantity=?,min_casting_quantity=?,
                    min_machined_quantity=?,min_assembly_quantity=?,
                    min_purchase_quantity=? WHERE id=?""",
                    (name, cat, mq, mcq, mmq, maq, mpq, self._editing_id), commit=False)
            except sqlite3.IntegrityError:
                err("Item name already exists."); return

            if old_name and old_name != name:
                for table, col in [
                    ("casting_in","item"),("machining_register","item"),
                    ("rejection_register","item"),("plating_out","item"),
                    ("plating_in","item"),("grinder_register","item"),
                    ("opening_stock","item"),("fg_deductions","item"),
                ]:
                    try: qry(f"UPDATE {table} SET {col}=? WHERE {col}=?", (name, old_name), commit=False)
                    except: pass

            qry("SELECT 1", commit=True)
            log_action(self.app.current_user, "UPDATE ITEM", f"Renamed '{old_name}' → '{name}'")

        self.cancel(); self.refresh()

    def cancel(self):
        self._editing_id = None
        self.ne.delete(0, "end")
        self.cc.set(self.CATS[0])
        self._toggle_fields()
        for e in self._brass_entries.values():
            e.delete(0, "end"); e.insert(0, "0")
        self.mq.delete(0, "end"); self.mq.insert(0, "0")
        self.mpq.delete(0, "end"); self.mpq.insert(0, "0")
        self.sb.config(text="Save Item")
        self.cb_.pack_forget()

    def _delete(self, vals):
        if self.app.current_role != "admin":
            warn("Only admins can delete items."); return
        if not confirm(f"Delete item '{vals[0]}'?"): return
        qry("DELETE FROM master_items WHERE id=?", (vals[8],), commit=True)
        log_action(self.app.current_user, "DELETE ITEM", f"Removed '{vals[0]}'")
        self.refresh()


# ─────────────────────────────────────────────────────────────────────────────
#  VENDORS PANEL
# ─────────────────────────────────────────────────────────────────────────────
class VendorsPanel(tk.Frame):
    REGS = ["Brass In","Brass Out","Casting In","Machining",
            "Grinding","Plating","Miscellaneous","Rejection"]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        self._eid = None

        sf = ScrollFrame(self); sf.pack(fill="both", expand=True)
        p = sf.inner

        f = card(p); f.pack(fill="x", pady=(0, 10))
        f.grid_columnconfigure(1, weight=1, minsize=250)
        self.ne = self._row(f, "Vendor Name:", 0)
        self.ce = self._row(f, "Contact Info:", 1)

        tk.Label(f, text="Registers:", bg=C_CARD, font=FNB, anchor="w").grid(
            row=2, column=0, sticky="nw", pady=4)
        self.rlb = CheckList(f, height=5)
        for r in self.REGS: self.rlb.insert("end", r)
        self.rlb.grid(row=2, column=1, sticky="ew", padx=(8,0), pady=4)

        tk.Label(f, text="Items:", bg=C_CARD, font=FNB, anchor="w").grid(
            row=3, column=0, sticky="nw", pady=4)
        self.ilb = CheckList(f, height=5)
        self.ilb.grid(row=3, column=1, sticky="ew", padx=(8,0), pady=4)

        br = tk.Frame(f, bg=C_CARD)
        br.grid(row=4, column=0, columnspan=2, pady=(10,0), sticky="w")
        self.sb  = btn(br, "Save Vendor", self.save, bg=C_SUCCESS, hov=C_SUC_HOV)
        self.sb.pack(side="left")
        self.cb_ = btn(br, "Cancel", self.cancel, bg=C_GRAY, hov=C_GRA_HOV)

        tk.Label(p, text="Existing Vendors", font=FSH, bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG)
        self.list_holder.pack(fill="both", expand=True)
        self._load_items(); self.refresh()

    def _row(self, parent, label, r):
        tk.Label(parent, text=label, bg=C_CARD, font=FNB, anchor="w").grid(
            row=r, column=0, sticky="w", pady=4)
        e = entry(parent)
        e.grid(row=r, column=1, sticky="ew", padx=(8,0), pady=4)
        return e

    def _load_items(self):
        self.mitems = qry("SELECT id,item_name FROM master_items ORDER BY item_name",
                          fetch="all") or []
        self.ilb.delete(0,"end")
        for _, n in self.mitems: self.ilb.insert("end", n)

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data = []
        for vid, vn, vc in qry("SELECT id,vendor_name,contact_info FROM vendors ORDER BY vendor_name", fetch="all") or []:
            deals = qry("SELECT m.item_name FROM master_items m JOIN vendor_items vi ON m.id=vi.item_id WHERE vi.vendor_id=?", (vid,), fetch="all") or []
            regs  = qry("SELECT register_type FROM vendor_registers WHERE vendor_id=?", (vid,), fetch="all") or []
            vals = (vn, vc or "", ", ".join(r[0] for r in deals) or "None", ", ".join(r[0] for r in regs) or "None", "", vid)
            data.append({
                "values": [vals[0], vals[1], vals[2], vals[3]],
                "buttons": [("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                            ("Delete", lambda v=vals: self._delete(v), "#dc3545")]
            })
        make_action_list(self.list_holder, ["Vendor","Contact","Items","Registers"],
                        {"Vendor":16,"Contact":14,"Items":26,"Registers":22}, data)

    def _load_edit(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        vid = vals[5]
        ai = {r[0] for r in qry("SELECT item_id FROM vendor_items WHERE vendor_id=?", (vid,), fetch="all") or []}
        ar = {r[0] for r in qry("SELECT register_type FROM vendor_registers WHERE vendor_id=?", (vid,), fetch="all") or []}
        item_names = [n for _, n in self.mitems]
        selected_items = [n for i, n in self.mitems if i in ai]
        selected_regs = [r for r in self.REGS if r in ar]
        def _save(v):
            name = v["name"].strip()
            if not name: warn("Vendor Name required."); return False
            sel_names = set(v["items"])
            sitems = [i for i, n in self.mitems if n in sel_names]
            try:
                qry("UPDATE vendors SET vendor_name=?,contact_info=? WHERE id=?", (name, v["contact"], vid), commit=True)
                qry("DELETE FROM vendor_items WHERE vendor_id=?", (vid,), commit=True)
                qry("DELETE FROM vendor_registers WHERE vendor_id=?", (vid,), commit=True)
            except sqlite3.IntegrityError:
                err("Vendor name already exists."); return False
            for iid in sitems:
                qry("INSERT INTO vendor_items (vendor_id,item_id) VALUES (?,?)", (vid, iid), commit=True)
            for rt in v["registers"]:
                qry("INSERT INTO vendor_registers (vendor_id,register_type) VALUES (?,?)", (vid, rt), commit=True)
            log_action(self.app.current_user, "UPDATE VENDOR", f"Edited '{name}'")
            self.refresh()
        edit_dialog(self, "Edit Vendor",
                    [("name","Vendor Name:","entry",None),
                     ("contact","Contact Info:","entry",None),
                     ("registers","Registers:","listbox",self.REGS),
                     ("items","Items:","listbox",item_names)],
                    {"name": vals[0], "contact": vals[1],
                     "registers": selected_regs, "items": selected_items},
                    _save)

    def save(self):
        name = self.ne.get().strip()
        if not name: warn("Vendor Name required."); return
        sregs  = [self.REGS[i] for i in self.rlb.curselection()]
        sitems = [self.mitems[i][0] for i in self.ilb.curselection()]
        contact = self.ce.get().strip()
        try:
            if self._eid is None:
                vid = qry("INSERT INTO vendors (vendor_name,contact_info) VALUES (?,?)",
                          (name, contact), commit=True)
                log_action(self.app.current_user,"CREATE VENDOR",f"Added '{name}'")
            else:
                vid = self._eid
                qry("UPDATE vendors SET vendor_name=?,contact_info=? WHERE id=?",
                    (name,contact,vid),commit=True)
                qry("DELETE FROM vendor_items WHERE vendor_id=?",(vid,),commit=True)
                qry("DELETE FROM vendor_registers WHERE vendor_id=?",(vid,),commit=True)
                log_action(self.app.current_user,"UPDATE VENDOR",f"Edited '{name}'")
            for iid in sitems:
                qry("INSERT INTO vendor_items (vendor_id,item_id) VALUES (?,?)",(vid,iid),commit=True)
            for rt in sregs:
                qry("INSERT INTO vendor_registers (vendor_id,register_type) VALUES (?,?)",(vid,rt),commit=True)
        except sqlite3.IntegrityError:
            err("Vendor name already exists."); return
        self.cancel(); self.refresh()

    def cancel(self):
        self._eid = None
        self.ne.delete(0,"end"); self.ce.delete(0,"end")
        self.rlb.selection_clear(0,"end"); self.ilb.selection_clear(0,"end")
        self.sb.config(text="Save Vendor"); self.cb_.pack_forget()

    def _delete(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        if not confirm(f"Delete vendor '{vals[0]}'?"): return
        qry("DELETE FROM vendors WHERE id=?", (vals[5],), commit=True)
        log_action(self.app.current_user,"DELETE VENDOR",f"Removed '{vals[0]}'")
        self.refresh()


# ─────────────────────────────────────────────────────────────────────────────
#  UNITS PANEL
# ─────────────────────────────────────────────────────────────────────────────
class UnitsPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app; self._eid = None

        sf = ScrollFrame(self); sf.pack(fill="both", expand=True)
        p = sf.inner

        f = card(p); f.pack(fill="x", pady=(0,10))
        f.grid_columnconfigure(1, weight=1, minsize=200)
        tk.Label(f, text="Unit Name:", bg=C_CARD, font=FNB, anchor="w").grid(row=0, column=0, sticky="w", pady=4)
        self.ne = entry(f); self.ne.grid(row=0, column=1, sticky="ew", padx=(8,0), pady=4)
        br = tk.Frame(f, bg=C_CARD); br.grid(row=1, column=0, columnspan=2, pady=(10,0), sticky="w")
        self.sb  = btn(br,"Save Unit",self.save,bg=C_SUCCESS,hov=C_SUC_HOV); self.sb.pack(side="left")
        self.cb_ = btn(br,"Cancel",self.cancel,bg=C_GRAY,hov=C_GRA_HOV)

        info_box(p, "Each unit shows the items currently using it. Click 'Assign Items' to check/uncheck items for that unit in bulk.")
        tk.Label(p, text="Units of Measurement", font=FSH, bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG)
        self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        units = qry("""SELECT u.id, u.unit_name, GROUP_CONCAT(m.item_name, ', ')
                    FROM units u
                    LEFT JOIN item_units iu ON iu.unit_id = u.id
                    LEFT JOIN master_items m ON m.id = iu.item_id
                    GROUP BY u.id, u.unit_name
                    ORDER BY u.unit_name""", fetch="all") or []
        rows = []
        for uid, uname, items_str in units:
            rows.append({
                "values": [uname, items_str or "No items assigned"],
                "buttons": [
                    ("Delete", lambda u=uid, n=uname: self._delete((n, "", "", u)), C_DANGER if 'C_DANGER' in globals() else "#dc3545"),
                    ("Edit",   lambda u=uid, n=uname: self._load_edit((n, "", "", u)), C_GRAY),
                    ("Assign Items", lambda u=uid, n=uname: self._assign_items((n, "", "", u)), C_PRIMARY),
                ]
            })
        make_action_list(self.list_holder, ["Unit Name","Items Using This Unit"],
                        {"Unit Name":18,"Items Using This Unit":45}, rows)

    def _load_edit(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        uid = vals[3]
        def _save(v):
            name = v["name"].strip()
            if not name: warn("Unit Name required."); return False
            try:
                qry("UPDATE units SET unit_name=? WHERE id=?", (name, uid), commit=True)
            except sqlite3.IntegrityError:
                err("Unit name already exists."); return False
            log_action(self.app.current_user, "UPDATE UNIT", f"Renamed id={uid} to '{name}'")
            self.refresh()
        edit_dialog(self, "Edit Unit", [("name","Unit Name:","entry",None)],
                    {"name": vals[0]}, _save)

    def save(self):
        name = self.ne.get().strip()
        if not name: warn("Unit Name required."); return
        try:
            if self._eid is None:
                qry("INSERT INTO units (unit_name) VALUES (?)",(name,),commit=True)
                log_action(self.app.current_user,"CREATE UNIT",f"Added '{name}'")
            else:
                qry("UPDATE units SET unit_name=? WHERE id=?",(name,self._eid),commit=True)
                log_action(self.app.current_user,"UPDATE UNIT",f"Renamed id={self._eid} to '{name}'")
        except sqlite3.IntegrityError:
            err("Unit name already exists."); return
        self.cancel(); self.refresh()

    def cancel(self):
        self._eid = None; self.ne.delete(0,"end")
        self.sb.config(text="Save Unit"); self.cb_.pack_forget()

    def _delete(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        if not confirm(f"Delete unit '{vals[0]}'?"): return
        qry("DELETE FROM units WHERE id=?",(vals[3],),commit=True)
        log_action(self.app.current_user,"DELETE UNIT",f"Removed '{vals[0]}'"); self.refresh()

    def _assign_items(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        uid = vals[3]; uname = vals[0]
        items = qry("""SELECT m.id, m.item_name, iu.unit_id
                       FROM master_items m
                       LEFT JOIN item_units iu ON iu.item_id = m.id
                       ORDER BY m.item_name""", fetch="all") or []
        dlg = tk.Toplevel(self); dlg.title(f"Assign Items to '{uname}'")
        dlg.configure(bg=C_CARD); dlg.geometry("400x480")
        tk.Label(dlg, text=f"Unit: {uname}", font=FSH, bg=C_CARD).pack(anchor="w", padx=15, pady=(12,4))
        tk.Label(dlg, text="Check every item that should use this unit:", font=FN, bg=C_CARD).pack(anchor="w", padx=15)
        sf = ScrollFrame(dlg); sf.pack(fill="both", expand=True, padx=15, pady=10)
        vars_ = {}
        for item_id, item_name, current_unit_id in items:
            v = tk.BooleanVar(value=(current_unit_id == uid))
            label = item_name
            if current_unit_id and current_unit_id != uid:
                other = qry("SELECT unit_name FROM units WHERE id=?", (current_unit_id,), fetch="one")
                if other: label += f"   (currently: {other[0]})"
            tk.Checkbutton(sf.inner, text=label, variable=v, bg=C_BG, font=FN, anchor="w").pack(fill="x", pady=1)
            vars_[item_id] = v

        def sv():
            for item_id, v in vars_.items():
                if v.get():
                    qry("INSERT INTO item_units (item_id,unit_id) VALUES (?,?) ON CONFLICT(item_id) DO UPDATE SET unit_id=excluded.unit_id",
                        (item_id, uid), commit=True)
                else:
                    qry("DELETE FROM item_units WHERE item_id=? AND unit_id=?", (item_id, uid), commit=True)
            log_action(self.app.current_user, "ASSIGN ITEMS TO UNIT", f"unit_id={uid} '{uname}'")
            dlg.destroy(); self.refresh()

        btn(dlg, "Save", sv, bg=C_SUCCESS, hov=C_SUC_HOV).pack(pady=10)

# ─────────────────────────────────────────────────────────────────────────────
#  BOX PACK PANEL
# ─────────────────────────────────────────────────────────────────────────────
class BoxPackPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app

        sf = ScrollFrame(self); sf.pack(fill="both", expand=True)
        p = sf.inner

        f = card(p); f.pack(fill="x", pady=(0,10))
        f.grid_columnconfigure(1, weight=1, minsize=250)
        tk.Label(f, text="Box Pack Name:", bg=C_CARD, font=FNB, anchor="w").grid(row=0,column=0,sticky="w",pady=4)
        self.ne = entry(f); self.ne.grid(row=0,column=1,sticky="ew",padx=(8,0),pady=4)
        tk.Label(f, text="Component Items:", bg=C_CARD, font=FNB, anchor="w").grid(row=1,column=0,sticky="nw",pady=4)
        self.ilb = CheckList(f, height=7)
        self.ilb.grid(row=1,column=1,sticky="ew",padx=(8,0),pady=4)
        btn(f,"Save Box Pack Item",self.save,bg=C_SUCCESS,hov=C_SUC_HOV).grid(row=2,column=0,columnspan=2,pady=(10,0),sticky="w")

        tk.Label(p, text="Existing Box Pack Items", font=FSH, bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG)
        self.list_holder.pack(fill="both", expand=True)
        self._load_items(); self.refresh()

    def _load_items(self):
        self.mitems = qry("SELECT id,item_name FROM master_items ORDER BY item_name",fetch="all") or []
        self.ilb.delete(0,"end")
        for _,n in self.mitems: self.ilb.insert("end",n)

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data = []
        for bid,bn in qry("SELECT id,box_pack_name FROM box_pack_items ORDER BY box_pack_name",fetch="all") or []:
            comps = qry("SELECT m.item_name FROM master_items m JOIN box_pack_components bc ON m.id=bc.item_id WHERE bc.box_pack_id=?",(bid,),fetch="all") or []
            vals = (bn, ", ".join(c[0] for c in comps) or "No items", "", bid)
            data.append({"values":[vals[0],vals[1]], "buttons":[("Edit", lambda v=vals: self._on_edit(v), "#0d6efd"), ("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Box Pack Name","Items Used"], {"Box Pack Name":20,"Items Used":40}, data)

    def _on_edit(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        bid = vals[3]
        comp_ids = {r[0] for r in qry("SELECT item_id FROM box_pack_components WHERE box_pack_id=?", (bid,), fetch="all") or []}
        item_names = [n for _, n in self.mitems]
        selected = [n for i, n in self.mitems if i in comp_ids]
        def _save(v):
            name = v["name"].strip()
            sel_names = set(v["items"])
            sids = [i for i, n in self.mitems if n in sel_names]
            if not name or not sids: warn("Name and at least one item required."); return False
            qry("UPDATE box_pack_items SET box_pack_name=? WHERE id=?", (name, bid), commit=True)
            qry("DELETE FROM box_pack_components WHERE box_pack_id=?", (bid,), commit=True)
            for iid in sids:
                qry("INSERT INTO box_pack_components (box_pack_id,item_id) VALUES (?,?)", (bid, iid), commit=True)
            log_action(self.app.current_user, "EDIT BOX PACK ITEM", f"Updated '{name}'")
            self.refresh()
        edit_dialog(self, "Edit Box Pack",
                    [("name","Box Pack Name:","entry",None),
                     ("items","Component Items:","listbox",item_names)],
                    {"name": vals[0], "items": selected},
                    _save)

    def _on_delete(self, vals):
        if not confirm(f"Delete box pack '{vals[0]}'?"): return
        if self.app.current_role != "admin": warn("Admins only."); return
        qry("DELETE FROM box_pack_components WHERE box_pack_id=?",(vals[3],),commit=True)
        qry("DELETE FROM box_pack_items WHERE id=?",(vals[3],),commit=True)
        log_action(self.app.current_user,"DELETE BOX PACK ITEM",f"Removed '{vals[0]}'")
        self.refresh()
    

    def save(self):
        name = self.ne.get().strip()
        sids = [self.mitems[i][0] for i in self.ilb.curselection()]
        if not name or not sids: warn("Name and at least one item required."); return
        if getattr(self, "_save_mode", None) == "edit":
            bid = self._edit_id
            qry("UPDATE box_pack_items SET box_pack_name=? WHERE id=?", (name, bid), commit=True)
            qry("DELETE FROM box_pack_components WHERE box_pack_id=?", (bid,), commit=True)
            for iid in sids:
                qry("INSERT INTO box_pack_components (box_pack_id,item_id) VALUES (?,?)", (bid, iid), commit=True)
            log_action(self.app.current_user, "EDIT BOX PACK ITEM", f"Updated '{name}'")
            self._save_mode = None; self._edit_id = None
        else:
            try: bid = qry("INSERT INTO box_pack_items (box_pack_name) VALUES (?)",(name,),commit=True)
            except sqlite3.IntegrityError: err("Box pack name already exists."); return
            for iid in sids:
                qry("INSERT INTO box_pack_components (box_pack_id,item_id) VALUES (?,?)",(bid,iid),commit=True)
            log_action(self.app.current_user,"CREATE BOX PACK ITEM",f"Added '{name}'")
        self.ne.delete(0,"end"); self.ilb.selection_clear(0,"end"); self.refresh()

# ─────────────────────────────────────────────────────────────────────────────
#  RATE LIST PANEL
# ─────────────────────────────────────────────────────────────────────────────
class RateListPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app; is_adm = app.current_role == "admin"

        sf = ScrollFrame(self); sf.pack(fill="both", expand=True)
        p = sf.inner

        if not is_adm:
            info_box(p,"Only admins can edit the rate list. Viewing in read-only mode.")
        if is_adm:
            f = card(p); f.pack(fill="x",pady=(0,10)); f.grid_columnconfigure(1,weight=1,minsize=220)
            self.items  = qry("SELECT id,item_name FROM master_items ORDER BY item_name",fetch="all") or []
            self.vendors= qry("SELECT id,vendor_name FROM vendors ORDER BY vendor_name",fetch="all") or []
            self.ic = combo(f,[i[1] for i in self.items]); tk.Label(f,text="Item:",bg=C_CARD,font=FNB).grid(row=0,column=0,sticky="w",pady=4); self.ic.grid(row=0,column=1,sticky="ew",padx=(8,0),pady=4)
            self.vc = combo(f,[v[1] for v in self.vendors]); tk.Label(f,text="Vendor:",bg=C_CARD,font=FNB).grid(row=1,column=0,sticky="w",pady=4); self.vc.grid(row=1,column=1,sticky="ew",padx=(8,0),pady=4)
            self.re = entry(f); tk.Label(f,text="Rate (Rs.):",bg=C_CARD,font=FNB).grid(row=2,column=0,sticky="w",pady=4); self.re.grid(row=2,column=1,sticky="ew",padx=(8,0),pady=4)
            btn(f,"Add / Update Rate",self.save,bg=C_SUCCESS,hov=C_SUC_HOV).grid(row=3,column=0,columnspan=2,pady=(10,0),sticky="w")

        tk.Label(p, text="Rate List", font=FSH, bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG)
        self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        is_adm = self.app.current_role=="admin"
        data = []
        for rid,iname,vname,rate in qry("SELECT rl.id,m.item_name,v.vendor_name,rl.rate FROM rate_list rl JOIN master_items m ON rl.item_id=m.id JOIN vendors v ON rl.vendor_id=v.id ORDER BY m.item_name,v.vendor_name",fetch="all") or []:
            vals=(iname,vname,f"Rs. {rate:.2f}","",rid)
            row = {"values":[vals[0],vals[1],vals[2]], "buttons":[]}
            if is_adm:
                row["buttons"] = [("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]
            data.append(row)
        make_action_list(self.list_holder, ["Item","Vendor","Rate"], {"Item":22,"Vendor":18,"Rate":10}, data)

    def _on_delete(self,vals):
        if not confirm(f"Delete rate for '{vals[0]}' / '{vals[1]}'?"): return
        qry("DELETE FROM rate_list WHERE id=?",(vals[4],),commit=True)
        log_action(self.app.current_user,"RATE LIST DELETE",f"id={vals[4]}"); self.refresh()

    def save(self):
        iname=self.ic.get(); vname=self.vc.get()
        if not iname or not vname: warn("Select item and vendor."); return
        try: rate=float(self.re.get())
        except: err("Rate must be a number."); return
        iid=next(i[0] for i in self.items if i[1]==iname)
        vid=next(v[0] for v in self.vendors if v[1]==vname)
        ex=qry("SELECT id FROM rate_list WHERE item_id=? AND vendor_id=?",(iid,vid),fetch="one")
        if ex:
            qry("UPDATE rate_list SET rate=? WHERE id=?",(rate,ex[0]),commit=True)
            log_action(self.app.current_user,"RATE LIST UPDATE",f"rate={rate}")
        else:
            qry("INSERT INTO rate_list (item_id,vendor_id,rate) VALUES (?,?,?)",(iid,vid,rate),commit=True)
            log_action(self.app.current_user,"RATE LIST INSERT",f"item={iname} vendor={vname} rate={rate}")
        self.re.delete(0,"end"); self.refresh()


# ═══════════════════════════════════════════════════════════════════════════════
#  REGISTERS SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
class RegistersScreen(tk.Frame):
    REGS = ["Opening Stock","Inventory","Brass In","Brass Out",
            "Casting In","Machining","Plating","Miscellaneous",
            "Rejection","Sale Register","Finished Goods",
            "Vendor Ledgers","Item Records"]

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG)
        self.app = app
        tk.Label(self, text="Registers", font=FH, bg=C_BG, fg="#222").pack(anchor="w", pady=(0,8))
        top = card(self); top.pack(fill="x", pady=(0,10))
        tk.Label(top, text="Select Register:", bg=C_CARD, font=FNB).pack(side="left")
        self.rc = combo(top, self.REGS, self.REGS[0], width=32)
        self.rc.pack(side="left", padx=10)
        self.rc.bind("<<ComboboxSelected>>", lambda e: self._render())
        self.body = tk.Frame(self, bg=C_BG)
        self.body.pack(fill="both", expand=True)
        self._render()

    def _render(self):
        for w in self.body.winfo_children(): w.destroy()
        m = {
            "Opening Stock":   OpeningStockPanel,
            "Inventory":       InventoryPanel,
            "Brass In":        BrassInPanel,
            "Brass Out":       BrassOutPanel,
            "Casting In":      CastingInPanel,
            "Machining":       MachiningPanel,
            "Plating":         PlatingPanel,
            "Miscellaneous":   MiscPanel,
            "Rejection":       RejectionPanel,
            "Sale Register":   SaleRegisterPanel,
            "Finished Goods":  FinishedGoodsPanel,
            "Vendor Ledgers":  VendorLedgersPanel,
            "Item Records":    ItemRecordsPanel,
        }
        m[self.rc.get()](self.body, self.app).pack(fill="both", expand=True)


# ─────────────────────────────────────────────────────────────────────────────
#  OPENING STOCK
# ─────────────────────────────────────────────────────────────────────────────
class OpeningStockPanel(tk.Frame):
    STAGES = {"casted":"Casted","machined":"Machined","assembly":"Assembly",
              "plating_short":"Plating Short","misc":"Miscellaneous Stock"}

    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app = app
        sf = ScrollFrame(self); sf.pack(fill="both", expand=True)
        p  = sf.inner
        tk.Label(p, text="Opening Stock", font=FSH, bg=C_BG).pack(anchor="w", pady=(0,4))
        info_box(p,"Set starting balances once per item+stage when migrating from a previous period.")

        items = [i[0] for i in qry("SELECT item_name FROM master_items ORDER BY item_name",fetch="all") or []]
        c=card(p)
        c.pack(fill="x", pady=(0,8))
        c.pack_propagate(True)
        f=ResponsiveForm(c); f.pack(fill="x", expand=False)
        self.de = f.add("Date:", date_picker(f, default=today()))
        self.ic = f.add("Item:", combo(f, items, items[0] if items else ""))
        self.sc = f.add("Stage:", combo(f, list(self.STAGES.values())))
        self.pe = f.add("Pieces:", entry(f, "0"))
        self.we = f.add("Weight (KG):", entry(f, "0"))
        f.add_btn_row(btn(f,"Save Opening Balance",self.save_item,bg=C_SUCCESS,hov=C_SUC_HOV))

        c2=card(p); c2.pack(fill="x",pady=(0,8))
        f2=ResponsiveForm(c2); f2.pack(fill="x", expand=False)
        self.sde = f2.add("Date:", date_picker(f2, default=today()))
        self.sbn = f2.add("Bag #:", entry(f2, "1"))
        self.swt = f2.add("Weight (KG):", entry(f2))
        f2.add_btn_row(btn(f2,"Save Opening Scrap",self.save_scrap,bg=C_SUCCESS,hov=C_SUC_HOV))

        tk.Label(p,text="Opening Balances — Item Stock",font=FSH,bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)    

        tk.Label(p,text="Opening Balances — Scrap Brass",font=FSH,bg=C_BG).pack(anchor="w",pady=(8,0))
        self.list_holder2 = tk.Frame(p, bg=C_BG); self.list_holder2.pack(fill="both", expand=True)
        self.refresh()

    def _row(self,f,lbl_text,r,default=""):
        tk.Label(f,text=lbl_text,bg=C_CARD,font=FNB,anchor="w").grid(row=r,column=0,sticky="w",pady=3)
        e=entry(f,default); e.grid(row=r,column=1,sticky="ew",padx=(8,0),pady=3); return e

    def _combo(self,f,lbl_text,r,values):
        tk.Label(f,text=lbl_text,bg=C_CARD,font=FNB,anchor="w").grid(row=r,column=0,sticky="w",pady=3)
        cb=combo(f,values,values[0] if values else ""); cb.grid(row=r,column=1,sticky="ew",padx=(8,0),pady=3); return cb

    def _stage_key(self,label):
        for k,v in self.STAGES.items():
            if v==label: return k
        return label

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,item,stage,pieces,weight FROM opening_stock ORDER BY item,stage",fetch="all") or []:
            vals=(r[1],r[2],self.STAGES.get(r[3],r[3]),r[4],f"{r[5]:.3f} KG","",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4]],
                        "buttons":[("Edit", lambda v=vals: self._edit_item(v), "#0d6efd"),
                                   ("Delete", lambda v=vals: self._del_item(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Item","Stage","Pieces","Weight"],
                        {"Date":10,"Item":18,"Stage":14,"Pieces":8,"Weight":10}, data)

        for w in self.list_holder2.winfo_children(): w.destroy()
        data2=[]
        for r in qry("SELECT id,date,bag_number,weight FROM brass_in WHERE vendor='Opening Balance' AND status='available' ORDER BY date",fetch="all") or []:
            vals=(r[1],r[2],f"{r[3]:.3f} KG","",r[0])
            data2.append({"values":[vals[0],vals[1],vals[2]],
                        "buttons":[("Delete", lambda v=vals: self._del_scrap(v), "#dc3545")]})
        make_action_list(self.list_holder2, ["Date","Bag #","Weight"], {"Date":10,"Bag #":8,"Weight":10}, data2)

    def _edit_item(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        oid = vals[6]
        items = [i[0] for i in qry("SELECT item_name FROM master_items ORDER BY item_name",fetch="all") or []]
        def _save(v):
            try: pcs = int(v["pieces"] or 0); wt = float(v["weight"] or 0)
            except ValueError: err("Pieces and Weight must be numbers."); return False
            qry("UPDATE opening_stock SET date=?,item=?,stage=?,pieces=?,weight=? WHERE id=?",
                (v["date"], v["item"], self._stage_key(v["stage"]), pcs, wt, oid), commit=True)
            log_action(self.app.current_user, "UPDATE OPENING STOCK", f"id={oid}")
            self.refresh()
        edit_dialog(self, "Edit Opening Stock",
                    [("date","Date:","date",None),
                     ("item","Item:","combo",items),
                     ("stage","Stage:","combo",list(self.STAGES.values())),
                     ("pieces","Pieces:","entry",None),
                     ("weight","Weight (KG):","entry",None)],
                    {"date": vals[0], "item": vals[1], "stage": vals[2],
                     "pieces": vals[3], "weight": vals[4].replace(" KG","")},
                    _save)

    def _del_item(self,vals):
        if self.app.current_role!="admin": return
        if not confirm("Delete this opening stock entry?"): return
        qry("DELETE FROM opening_stock WHERE id=?",(vals[6],),commit=True)
        log_action(self.app.current_user,"DELETE OPENING STOCK",f"'{vals[1]}' ({vals[2]})"); self.refresh()

    def _del_scrap(self,vals):
        if self.app.current_role!="admin": return
        if not confirm("Delete this opening scrap entry?"): return
        qry("DELETE FROM brass_in WHERE id=? AND vendor='Opening Balance'",(vals[4],),commit=True)
        log_action(self.app.current_user,"DELETE OPENING SCRAP",f"Bag #{vals[1]}"); self.refresh()

    def save_item(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); item=self.ic.get(); stage=self._stage_key(self.sc.get())
        if not item: warn("Select an item."); return
        try: pieces=int(self.pe.get() or 0); wt=float(self.we.get() or 0)
        except: err("Pieces/Weight must be numbers."); return
        if getattr(self, "_edit_mode", False):
            qry("UPDATE opening_stock SET date=?,item=?,stage=?,pieces=?,weight=? WHERE id=?",
                (date,item,stage,pieces,wt,self._edit_id),commit=True)
            log_action(self.app.current_user,"EDIT OPENING STOCK",f"'{item}' ({stage}) {pieces} pcs")
            self._edit_mode = False; self._edit_id = None
        else:
            qry("INSERT INTO opening_stock (item,stage,pieces,weight,date) VALUES (?,?,?,?,?) ON CONFLICT(item,stage) DO UPDATE SET pieces=excluded.pieces,weight=excluded.weight,date=excluded.date",
                (item,stage,pieces,wt,date),commit=True)
            log_action(self.app.current_user,"SET OPENING STOCK",f"'{item}' ({stage}) {pieces} pcs")
        self.refresh()

    def save_scrap(self):
        date=self.sde.get_date().strftime("%Y-%m-%d") or today()
        try: bag=int(self.sbn.get() or 1); wt=float(self.swt.get() or 0)
        except: err("Bag# and Weight must be numbers."); return
        qry("INSERT INTO brass_in (bag_number,date,vendor,weight,remaining_weight,status) VALUES (?,?,'Opening Balance',?,?,'available')",
            (bag,date,wt,wt),commit=True)
        log_action(self.app.current_user,"SET OPENING SCRAP STOCK",f"Bag #{bag} {wt:.3f} KG")
        self.swt.delete(0,"end"); self.refresh()

    # def _del_item(self,event):
    #     if self.app.current_role!="admin": return
    #     col=self.tv.identify_column(event.x); row=self.tv.identify_row(event.y)
    #     if not row or col!="#6": return
    #     vals=self.tv.item(row,"values")
    #     if not confirm("Delete this opening stock entry?"): return
    #     qry("DELETE FROM opening_stock WHERE id=?",(vals[6],),commit=True)
    #     log_action(self.app.current_user,"DELETE OPENING STOCK",f"'{vals[1]}' ({vals[2]})"); self.refresh()

    # def _del_scrap(self,event):
    #     if self.app.current_role!="admin": return
    #     col=self.tv2.identify_column(event.x); row=self.tv2.identify_row(event.y)
    #     if not row or col!="#4": return
    #     vals=self.tv2.item(row,"values")
    #     if not confirm("Delete this opening scrap entry?"): return
    #     qry("DELETE FROM brass_in WHERE id=? AND vendor='Opening Balance'",(vals[4],),commit=True)
    #     log_action(self.app.current_user,"DELETE OPENING SCRAP",f"Bag #{vals[1]}"); self.refresh()


# ─────────────────────────────────────────────────────────────────────────────
#  INVENTORY
# ─────────────────────────────────────────────────────────────────────────────
class InventoryPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app = app
        tk.Label(self,text="Inventory Register (Current Availability)",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        info_box(self,"Casted = Casting In − Machining − Rejections.  Machined = Machined − Plating Out − Grinder Rej.  Assembly = Plating Inward − Plated Rej − FG Deductions.")
        bar = tk.Frame(self, bg=C_BG); bar.pack(fill="x", pady=(0,6))
        btn(bar,"Refresh",self.refresh,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left")
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left",padx=(8,0))
        tf,self.tv = make_tree(self,["S.No","Item & Stage","Weight","Quantity"],18,
                               {"S.No":50,"Item & Stage":400,"Weight":120,"Quantity":120})
        tf.pack(fill="both",expand=True)
        self.refresh()

    def _calc(self):
        rows=[]
        r=qry("SELECT COALESCE(SUM(remaining_weight),0) FROM brass_in WHERE status='available' AND remaining_weight>0",fetch="one")
        scrap=r[0] if r else 0.0
        rows.append(("Scrap (Brass)",f"{scrap:.3f} KG" if scrap>0 else "",""))

        items=qry("SELECT id,item_name,item_category FROM master_items ORDER BY item_name",fetch="all") or []

        op_map={}
        for oi,os_,op,ow in qry("SELECT item,stage,pieces,weight FROM opening_stock",fetch="all") or []:
            op_map[(oi,os_)]=(op or 0,ow or 0.0)

        iu_map={r[0]:r[1] for r in qry("SELECT m.item_name,u.unit_name FROM item_units iu JOIN master_items m ON m.id=iu.item_id JOIN units u ON u.id=iu.unit_id",fetch="all") or []}

        # ── ONE query per table instead of one per item ──
        rej_map={}
        for it,rt,rp,rw in qry("SELECT item,rejection_type,COALESCE(SUM(pieces),0),COALESCE(SUM(weight),0) FROM rejection_register WHERE rejection_type!='Vendor Rejection' GROUP BY item,rejection_type",fetch="all") or []:
            rej_map[(it,rt)]=(rp,rw)

        cast_map={r[0]:(r[1],r[2]) for r in qry("SELECT item,COALESCE(SUM(pieces),0),COALESCE(SUM(weight),0) FROM casting_in GROUP BY item",fetch="all") or []}
        mach_map={r[0]:(r[1],r[2]) for r in qry("SELECT item,COALESCE(SUM(pieces),0),COALESCE(SUM(weight),0) FROM machining_register GROUP BY item",fetch="all") or []}

        plate_out_map={r[0]:(r[1],r[2],r[3]) for r in qry(
            "SELECT pi.item,COALESCE(SUM(pi.pieces),0),COALESCE(SUM(pi.weight),0),COALESCE(SUM(pi.pieces-pi.remaining_pieces),0) "
            "FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id WHERE pc.direction='Outward' GROUP BY pi.item",
            fetch="all") or []}

        plate_in_map={r[0]:(r[1],r[2]) for r in qry(
            "SELECT pi.item,COALESCE(SUM(pi.pieces),0),COALESCE(SUM(pi.weight),0) "
            "FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id WHERE pc.direction='Inward' GROUP BY pi.item",
            fetch="all") or []}

        misc_map={r[0]:(r[1],r[2],r[3],r[4]) for r in qry(
            "SELECT item,"
            "COALESCE(SUM(CASE WHEN option_type='purchase' THEN pieces ELSE 0 END),0),"
            "COALESCE(SUM(CASE WHEN option_type='purchase' THEN weight ELSE 0 END),0),"
            "COALESCE(SUM(CASE WHEN option_type='in assembly' THEN pieces ELSE 0 END),0),"
            "COALESCE(SUM(CASE WHEN option_type='in assembly' THEN weight ELSE 0 END),0) "
            "FROM miscellaneous_entries GROUP BY item", fetch="all") or []}

        fgd_map={r[0]:r[1] for r in qry(
            "SELECT fgd.item_name,COALESCE(SUM(fgd.pcs),0) FROM finished_goods_assembly_deductions fgd "
            "JOIN finished_goods_log fgl ON fgd.finished_goods_log_id=fgl.id GROUP BY fgd.item_name", fetch="all") or []}

        for iid,iname,icat in items:
            op_c_p,op_c_w   =op_map.get((iname,"casted"),(0,0.0))
            op_m_p,op_m_w   =op_map.get((iname,"machined"),(0,0.0))
            op_a_p,op_a_w   =op_map.get((iname,"assembly"),(0,0.0))
            op_ps_p,op_ps_w =op_map.get((iname,"plating_short"),(0,0.0))
            op_misc_p,op_misc_w=op_map.get((iname,"misc"),(0,0.0))

            cr_p=rej_map.get((iname,"Casted Rejection"),(0,0))[0]+rej_map.get((iname,"Machined Rejection"),(0,0))[0]
            cr_w=rej_map.get((iname,"Casted Rejection"),(0,0.0))[1]+rej_map.get((iname,"Machined Rejection"),(0,0.0))[1]
            mr_p,mr_w=rej_map.get((iname,"Grinder Rejection"),(0,0.0))
            ar_p,ar_w=rej_map.get((iname,"Plated Rejection"),(0,0.0))
            misc_rp,misc_rw=rej_map.get((iname,"Misc Vendor Rejection"),(0,0.0))

            cp,cw=cast_map.get(iname,(0,0.0))
            mp,mw=mach_map.get(iname,(0,0.0))
            po_p,po_w,pi_old_p=plate_out_map.get(iname,(0,0,0))
            pi_new_p,pi_new_w=plate_in_map.get(iname,(0,0.0))
            pi_p=pi_old_p+pi_new_p; ps_p=po_p-pi_p+op_ps_p; ps_w=max(0.0,po_w-pi_new_w+op_ps_w)

            misc_pp,misc_pw,misc_ap,misc_aw=misc_map.get(iname,(0,0.0,0,0.0))
            fg_ded=fgd_map.get(iname,0)

            nc_p=cp-mp-cr_p+op_c_p; nc_w=max(0.0,cw-mw-cr_w+op_c_w)
            nm_p=mp-po_p-mr_p+op_m_p; nm_w=max(0.0,mw-po_w-mr_w+op_m_w)
            na_p=pi_p-ar_p+op_a_p-fg_ded; na_w=max(0.0,pi_new_w-ar_w+op_a_w)
            nmisc_p=misc_pp-misc_ap-misc_rp+op_misc_p; nmisc_w=max(0.0,misc_pw-misc_aw-misc_rw+op_misc_w)

            def _w(x): return f"{x:.3f} KG"
            if icat=="Brass Item":
                # always show all 3 stages, 0 when empty
                rows.append((f"{iname} - Casted",   _w(max(nc_w,0)), str(max(nc_p,0))))
                rows.append((f"{iname} - Machined", _w(max(nm_w,0)), str(max(nm_p,0))))
                rows.append((f"{iname} - Assembly", _w(max(na_w,0)), str(max(na_p,0))))
                # Plating Short only when there is something short
                if ps_p>0 or ps_w>0:
                    rows.append((f"{iname} - Plating Short", _w(ps_w), str(ps_p)))
            else:
                unit=iu_map.get(iname,"pcs")
                rows.append((f"{iname} - Misc Stock",       _w(max(nmisc_w,0)), f"{max(nmisc_p,0)} {unit}"))
                rows.append((f"{iname} - Misc In Assembly", _w(misc_aw),        f"{misc_ap} {unit}"))
        return rows

    def refresh(self):
        data=self._calc()
        tree_fill(self.tv, [(i+1,)+r for i,r in enumerate(data)])


    def export_excel(self):
        if not EXCEL_OK:
            err("openpyxl not installed.")
            return

        from openpyxl.utils import get_column_letter

        vendor_sel    = self.vc.get()
        vendor_filter = "" if vendor_sel == "All" else vendor_sel
        date_from     = self.df.get_date().strftime("%Y-%m-%d")
        date_to       = self.dt.get_date().strftime("%Y-%m-%d")

        # ---------------------------------------------------------
        # Ask where to save
        # ---------------------------------------------------------
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if vendor_filter:
            safe_v = "".join(c for c in vendor_filter if c not in '\\/:*?"<>|')
            default_name = f"{safe_v} - Vendor Ledgers.xlsx"
        else:
            default_name = f"Vendor_Ledgers_brass_out_casting_in_{stamp}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Export Casting Vendor Ledger to Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
        )
        if not path:
            return

        # ---------------------------------------------------------
        # Pull the data with app.py's own ledger function.
        # That function reads the date filters from Flask's `request`,
        # so we temporarily swap in a tiny stand-in object.
        # ---------------------------------------------------------
        _app_mod = sys.modules[_get_vendor_ledger_data.__module__]

        class _FakeRequest:
            def __init__(self, values):
                self.values = values

        _orig_request = _app_mod.request
        _app_mod.request = _FakeRequest({
            "cast_date_from": date_from,
            "cast_date_to":   date_to,
        })
        conn = get_conn()
        try:
            vendor_data = _get_vendor_ledger_data(
                conn.cursor(), "brass_out_casting_in", vendor_filter
            )
        finally:
            conn.close()
            _app_mod.request = _orig_request

        # ---------------------------------------------------------
        # Styles / helpers
        # ---------------------------------------------------------
        thin   = Side(style="thin", color="B7B7B7")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        bold   = Font(bold=True)
        title_font = Font(bold=True, size=11)
        center = Alignment(horizontal="center", vertical="center")

        used_names = set()

        def safe_sheet_name(name):
            cleaned = name
            for ch in ["\\", "/", "*", "[", "]", ":", "?"]:
                cleaned = cleaned.replace(ch, "")
            cleaned = (cleaned.strip() or "Vendor")[:31]
            candidate = cleaned
            n = 2
            while candidate in used_names:
                suffix = f"_{n}"
                candidate = cleaned[:31 - len(suffix)] + suffix
                n += 1
            used_names.add(candidate)
            return candidate

        def to_float(text):
            try:
                return float(str(text).replace(" KG", "").replace(",", "").strip())
            except ValueError:
                return None

        def put(ws, r, c, value, is_bold=False, num_fmt=None, align=None):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = border
            if is_bold:
                cell.font = bold
            if num_fmt:
                cell.number_format = num_fmt
            if align:
                cell.alignment = align
            return cell

        # ---------------------------------------------------------
        # Layout constants (same as the web export)
        # ---------------------------------------------------------
        BO_COLS  = ["Date", "Bag Number", "Weight (KG)"]
        CI_COLS  = ["Date", "Item", "Pieces", "Weight (KG)"]
        BO_WIDTH = [18, 16, 14]
        CI_WIDTH = [18, 20, 12, 14]
        GAP      = 2
        CI_START = len(BO_COLS) + GAP + 1
        TOTAL_W  = len(BO_COLS) + GAP + len(CI_COLS)

        BO_BOLD = {"Total Brass Out (Period)"}
        CI_BOLD = {"Total", "Less: Vendor Rejection", "Adjusted Subtotal",
                   "Burning Loss (5%)", "Total Accounted Return"}

        wb = Workbook()
        wb.remove(wb.active)

        if not vendor_data:
            ws = wb.create_sheet("Export")
            ws["A1"] = "No data to export."

        # ---------------------------------------------------------
        # One sheet per vendor
        # ---------------------------------------------------------
        for v in vendor_data:
            ws = wb.create_sheet(title=safe_sheet_name(v["vendor"]))
            bo_all = v["tables"][0]["rows"]   # [date/label, bag, weight, paid]
            ci_all = v["tables"][1]["rows"]   # [date/label, item, pcs, weight, paid]

            # Row 1: vendor name across the full width
            t = ws.cell(row=1, column=1, value=v["vendor"])
            t.font = title_font
            t.alignment = center
            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=TOTAL_W)

            # Optional row 2: date filter info
            parts = []
            if v.get("cast_date_from"):
                parts.append(f"From: {v['cast_date_from']}")
            if v.get("cast_date_to"):
                parts.append(f"To: {v['cast_date_to']}")

            section_row = 2
            if parts:
                info = ws.cell(row=2, column=1,
                               value="Date Filter: " + "  |  ".join(parts))
                info.font = Font(italic=True, color="555555")
                ws.merge_cells(start_row=2, start_column=1,
                               end_row=2, end_column=TOTAL_W)
                section_row = 3

            # Section titles: "Brass Out" (left) and "Casting In" (right)
            h1 = ws.cell(row=section_row, column=1, value="Brass Out (Scrap Given)")
            h1.font = title_font
            h1.alignment = center
            h1.border = border
            ws.merge_cells(start_row=section_row, start_column=1,
                           end_row=section_row, end_column=len(BO_COLS))

            h2 = ws.cell(row=section_row, column=CI_START, value="Casting In")
            h2.font = title_font
            h2.alignment = center
            h2.border = border
            ws.merge_cells(start_row=section_row, start_column=CI_START,
                           end_row=section_row,
                           end_column=CI_START + len(CI_COLS) - 1)

            # Column headers
            hdr_row = section_row + 1
            for i, name in enumerate(BO_COLS):
                put(ws, hdr_row, 1 + i, name, True, align=center)
                ws.column_dimensions[get_column_letter(1 + i)].width = BO_WIDTH[i]
            for i, name in enumerate(CI_COLS):
                put(ws, hdr_row, CI_START + i, name, True, align=center)
                ws.column_dimensions[get_column_letter(CI_START + i)].width = CI_WIDTH[i]
            for g in range(len(BO_COLS) + 1, CI_START):
                ws.column_dimensions[get_column_letter(g)].width = 3

            # Data rows, both tables side by side
            data_start = hdr_row + 1
            max_rows   = max(len(bo_all), len(ci_all))

            for idx in range(max_rows):
                r = data_start + idx

                # ---- Brass Out side ----
                if idx < len(bo_all):
                    row   = bo_all[idx]
                    label = str(row[0]) if row[0] is not None else ""

                    if label.startswith("Brass Balance Before"):
                        put(ws, r, 1, label, True)
                        put(ws, r, 2, "", True)
                        ws.merge_cells(start_row=r, start_column=1,
                                       end_row=r, end_column=2)
                        wt = to_float(row[2])
                        put(ws, r, 3, wt if wt is not None else row[2],
                            True, "0.000")
                    else:
                        is_sum = label in BO_BOLD
                        put(ws, r, 1, label, is_sum)
                        put(ws, r, 2, str(row[1]) if row[1] is not None else "", is_sum)
                        wt = to_float(row[2]) if row[2] != "" else None
                        put(ws, r, 3, wt if wt is not None else "",
                            is_sum, "0.000" if wt is not None else None)
                else:
                    for c in range(1, len(BO_COLS) + 1):
                        ws.cell(row=r, column=c).border = border

                # ---- Casting In side ----
                if idx < len(ci_all):
                    row    = ci_all[idx]
                    label  = str(row[0]) if row[0] is not None else ""
                    is_sum = label in CI_BOLD

                    put(ws, r, CI_START, label, is_sum)
                    put(ws, r, CI_START + 1,
                        str(row[1]) if row[1] is not None else "", is_sum)

                    pcs_raw = str(row[2]) if row[2] is not None else ""
                    try:
                        pcs_val = int(pcs_raw)
                    except ValueError:
                        pcs_val = pcs_raw
                    put(ws, r, CI_START + 2, pcs_val, is_sum)

                    wt = to_float(row[3]) if row[3] != "" else None
                    put(ws, r, CI_START + 3, wt if wt is not None else "",
                        is_sum, "0.000" if wt is not None else None)
                else:
                    for c in range(CI_START, CI_START + len(CI_COLS)):
                        ws.cell(row=r, column=c).border = border

            # Net balance summary under both tables
            if v.get("summary"):
                sr = data_start + max_rows + 1
                put(ws, sr, 1, v["summary"], True)
                ws.merge_cells(start_row=sr, start_column=1,
                               end_row=sr, end_column=TOTAL_W)

        # ---------------------------------------------------------
        # Save
        # ---------------------------------------------------------
        try:
            wb.save(path)
        except PermissionError:
            messagebox.showerror(
                "Export Failed",
                "The Excel file is currently open.\n\n"
                "Please close the file and try again."
            )
            return
        except Exception as e:
            messagebox.showerror("Export Failed",
                                 f"Could not create Excel file:\n\n{e}")
            return

        try:
            log_action(
                self.app.current_user,
                "EXPORT EXCEL",
                "Exported Vendor Ledgers (brass_out_casting_in) for "
                + (f"vendor {vendor_filter}" if vendor_filter else "all vendors")
            )
        except Exception:
            pass

        messagebox.showinfo("Exported", f"Excel file saved successfully:\n\n{path}")

# ─────────────────────────────────────────────────────────────────────────────
#  BRASS IN
# ─────────────────────────────────────────────────────────────────────────────
class BrassInPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app; self._eid=None
        self._batch=[]
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        tk.Label(p,text="Brass In (Scrap) Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        vendors=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Brass In' ORDER BY v.vendor_name",fetch="all") or []]
        c=card(p)
        c.pack(fill="x", pady=(0,8))
        c.pack_propagate(True)
        f=ResponsiveForm(c); f.pack(fill="x", expand=False)
        self.de = f.add("Date:", date_picker(f, default=today()))
        self.vc = f.add("Vendor:", combo(f, vendors, vendors[0] if vendors else ""))
        self.be = f.add("Bag Number:", entry(f, "1"))
        self.we = f.add("Scrap Weight (KG):", entry(f))
        f.add_btn_row(btn(f,"Add Bag to Batch",self.add_to_batch,bg=C_WARNING,hov=C_WAR_HOV,fg="black"))

        tk.Label(p,text="Current Unsaved Scrap Batch",font=FSH,bg=C_BG).pack(anchor="w")
        tf,self.btv=make_tree(p,["Date","Vendor","Bag Number","Scrap Weight"],4,
                              {"Date":115,"Vendor":180,"Bag Number":110,"Scrap Weight":130})
        tf.pack(fill="x")
        bb=tk.Frame(p,bg=C_BG); bb.pack(anchor="w",pady=(6,8))
        btn(bb,"Save Batch",self.save_batch,bg=C_SUCCESS,hov=C_SUC_HOV).pack(side="left")
        btn(bb,"Remove Selected",self.remove_selected,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left",padx=(8,0))
        btn(bb,"Clear Batch",self.clear_batch,bg=C_DANGER,hov=C_DAN_HOV).pack(side="left",padx=(8,0))

        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="All Brass In Records",font=FSH,bg=C_BG).pack(anchor="w")
        cols=["Date","Vendor","Bag #","Original","Remaining","Status","Action","_id"]
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self._edit_row_id=None
        self.refresh()

    # ── batch helpers ────────────────────────────────────────────────────────
    def _refresh_batch(self):
        tree_fill(self.btv, [(b["date"], b["vendor"], b["bag"], f'{b["wt"]:.3f} KG')
                             for b in self._batch])

    def _reset_bag_no(self):
        nxt = max([b["bag"] for b in self._batch], default=0) + 1
        self.be.delete(0,"end"); self.be.insert(0,str(nxt))

    def add_to_batch(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); vendor=self.vc.get()
        if not vendor: warn("Select a vendor."); return
        try: bag=int(self.be.get()); wt=float(self.we.get())
        except: err("Bag# and Weight must be numbers."); return
        self._batch.append({"date":date,"vendor":vendor,"bag":bag,"wt":wt})
        self._refresh_batch()
        self.we.delete(0,"end"); self._reset_bag_no()

    def remove_selected(self):
        sel=self.btv.selection()
        if not sel: warn("Select a bag in the batch first."); return
        del self._batch[self.btv.index(sel[0])]
        self._refresh_batch(); self._reset_bag_no()

    def clear_batch(self):
        if not self._batch: return
        if not confirm("Clear all unsaved bags in this batch?"): return
        self._batch=[]; self._refresh_batch(); self._reset_bag_no()

    def save_batch(self):
        if not self._batch: warn("Batch is empty. Add at least one bag."); return
        for b in self._batch:
            qry("INSERT INTO brass_in (bag_number,date,vendor,weight,remaining_weight,status) VALUES (?,?,?,?,?,'available')",
                (b["bag"],b["date"],b["vendor"],b["wt"],b["wt"]),commit=True)
        log_action(self.app.current_user,"CREATE BRASS IN SCRAP",f"Saved scrap batch of {len(self._batch)} bags.")
        self._batch=[]; self._refresh_batch(); self._reset_bag_no(); self.refresh()

    def _form_row(form, label, widget):
        """Add a label+field pair to a ResponsiveForm. Returns widget."""
        return form.add(label, widget)

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,vendor,bag_number,weight,remaining_weight,status FROM brass_in ORDER BY date DESC,id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],f"{r[4]:.3f}",f"{r[5]:.3f}",r[6].capitalize(),"",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4],vals[5]],
                        "buttons":[("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                                    ("Delete", lambda v=vals: self._delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Vendor","Bag #","Original","Remaining","Status"],
                        {"Date":10,"Vendor":15,"Bag #":8,"Original":10,"Remaining":10,"Status":10}, data)

    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); vendor=self.vc.get()
        if not vendor: warn("Select a vendor."); return
        try: bag=int(self.be.get()); wt=float(self.we.get())
        except: err("Bag# and Weight must be numbers."); return
        if self._eid:
            qry("UPDATE brass_in SET date=?,vendor=?,bag_number=?,weight=?,remaining_weight=? WHERE id=?",
                (date,vendor,bag,wt,wt,self._eid),commit=True)
            log_action(self.app.current_user,"UPDATE BRASS IN",f"Bag #{bag} vendor '{vendor}'")
            self._eid=None; self.be.config(state="normal")
            btn_text="Save Brass In Entry"
        else:
            qry("INSERT INTO brass_in (bag_number,date,vendor,weight,remaining_weight,status) VALUES (?,?,?,?,?,'available')",
                (bag,date,vendor,wt,wt),commit=True)
            log_action(self.app.current_user,"CREATE BRASS IN SCRAP",f"Bag #{bag} vendor '{vendor}' {wt:.3f} KG")
        self.be.delete(0,"end"); self.we.delete(0,"end"); self.refresh()

    def _load_edit(self, vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        bid = vals[7]
        vendors=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Brass In' ORDER BY v.vendor_name",fetch="all") or []]
        def _save(v):
            try: bag = int(v["bag"]); wt = float(v["weight"])
            except ValueError: err("Bag# and Weight must be numbers."); return False
            qry("UPDATE brass_in SET date=?,vendor=?,bag_number=?,weight=?,remaining_weight=? WHERE id=?",
                (v["date"], v["vendor"], bag, wt, wt, bid), commit=True)
            log_action(self.app.current_user,"UPDATE BRASS IN",f"Bag #{bag} vendor '{v['vendor']}'")
            self.refresh()
        edit_dialog(self, "Edit Brass In Entry",
                    [("date","Date:","date",None),
                     ("vendor","Vendor:","combo",vendors),
                     ("bag","Bag Number:","entry",None),
                     ("weight","Scrap Weight (KG):","entry",None)],
                    {"date": vals[0], "vendor": vals[1], "bag": vals[2], "weight": vals[3]},
                    _save)

    def _delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        bid=vals[7]
        brass_out_exists=qry("SELECT COUNT(*) FROM brass_out WHERE brass_in_id=?",(bid,),fetch="one")[0]
        if brass_out_exists:
            warn(f"Cannot delete Bag #{vals[2]} ('{vals[1]}').\n"
                f"This bag has already been sent out in Brass Out.\n"
                f"Delete the Brass Out entry first."); return
        if not confirm(f"Delete Bag #{vals[2]} for '{vals[1]}'?"): return
        qry("DELETE FROM brass_in WHERE id=?",(bid,),commit=True)
        log_action(self.app.current_user,"DELETE BRASS IN",f"Bag #{vals[2]} vendor '{vals[1]}'"); self.refresh()
        
    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT date,vendor,bag_number,weight,remaining_weight,status FROM brass_in ORDER BY date DESC,id DESC",fetch="all") or []
        _simple_excel(path,["Date","Vendor","Bag #","Original (KG)","Remaining (KG)","Status"],rows,"Brass In")


# ─────────────────────────────────────────────────────────────────────────────
#  BRASS OUT
# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
#  BRASS OUT
# ─────────────────────────────────────────────────────────────────────────────
class BrassOutPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app; self._eid=None
        self._table_sf = None
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        self._title = tk.Label(p,text="Brass Out Register",font=FSH,bg=C_BG)
        self._title.pack(anchor="w",pady=(0,4))
        c=card(p)
        c.pack(fill="x", pady=(0,8))
        c.pack_propagate(True)
        self._card = c
        f=ResponsiveForm(c); f.pack(fill="x", expand=False)
        self.de = f.add("Date:", date_picker(f, default=today()))
        self.bc = f.add("Available Scrap Bag:", combo(f, []))
        self.bc.bind("<<ComboboxSelected>>",self._fill_wt)
        self.we = f.add("Weight (KG):", entry(f))
        casting_vs = [v[0] for v in qry(
            "SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id "
            "WHERE vr.register_type IN ('Brass Out','Casting In') ORDER BY v.vendor_name", fetch="all") or []]
        self.gvc = f.add("Given To:", combo(f, casting_vs, casting_vs[0] if casting_vs else ""))
        self.sb = btn(f,"Confirm & Save",self.save,bg=C_SUCCESS,hov=C_SUC_HOV)
        f.add_btn_row(self.sb)
        self.load_bags()

        self._bar=tk.Frame(p,bg=C_BG); self._bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(self._bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        self._title2 = tk.Label(p,text="All Brass Out Records",font=FSH,bg=C_BG)
        self._title2.pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()

        self.bind("<Configure>", lambda e: self._resize_table())
        self.after(50, self._resize_table)

    def _resize_table(self):
        if self._table_sf is None:
            return
        self.update_idletasks()
        used = (self._title.winfo_reqheight() + self._card.winfo_reqheight()
                + self._bar.winfo_reqheight() + self._title2.winfo_reqheight())
        avail = self.winfo_height() - used - 60
        self._table_sf.configure(height=max(220, avail))

    def _form_row(form, label, widget):
        return form.add(label, widget)

    def load_bags(self):
        self.bags=qry("SELECT id,bag_number,vendor,remaining_weight FROM brass_in WHERE status='available' AND remaining_weight>0",fetch="all") or []
        self.bc.config(values=[f"Bag #{b[1]} | {b[2]} | {b[3]:.3f} KG" for b in self.bags])

    def _fill_wt(self,_=None):
        idx=self.bc.current()
        if idx>=0: self.we.delete(0,"end"); self.we.insert(0,f"{self.bags[idx][3]:.3f}")

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,vendor,bag_number,weight,given_to FROM brass_out ORDER BY id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],f"{r[4]:.3f} KG",r[5] or "",r[0],r[1])  # vals[5]=given_to, vals[6]=id, vals[7]=raw date
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4]],
                        "buttons":[("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                                    ("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]})
        self._table_sf = make_action_list(self.list_holder, ["Date","Vendor","Bag #","Weight Out","Given To"],
                        {"Date":10,"Vendor":15,"Bag #":8,"Weight Out":12,"Given To":15}, data)
        self._resize_table()

    def _on_delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm(f"Delete Brass Out Bag #{vals[2]}? Scrap bag will be restored."): return
        r=qry("SELECT brass_in_id FROM brass_out WHERE id=?",(vals[5],),fetch="one")
        qry("DELETE FROM brass_out WHERE id=?",(vals[5],),commit=True)
        if r and r[0]:
            br2=qry("SELECT weight FROM brass_in WHERE id=?",(r[0],),fetch="one")
            if br2: qry("UPDATE brass_in SET status='available',remaining_weight=? WHERE id=?",(br2[0],r[0]),commit=True)
        log_action(self.app.current_user,"DELETE BRASS OUT",f"Bag #{vals[2]}")
        self.load_bags(); self.refresh()

    def _load_edit(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        eid = vals[5]
        casting_vs = [v[0] for v in qry(
            "SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id "
            "WHERE vr.register_type IN ('Brass Out','Casting In') ORDER BY v.vendor_name", fetch="all") or []]
        def _save(v):
            try: wt = float(v["weight"])
            except ValueError: err("Weight must be a number."); return False
            qry("UPDATE brass_out SET date=?,vendor=?,bag_number=?,weight=?,given_to=? WHERE id=?",
                (v["date"], v["vendor"], v["bag"], wt, v["given_to"], eid), commit=True)
            log_action(self.app.current_user, "UPDATE BRASS OUT", f"id={eid}")
            self.refresh()
        edit_dialog(self, "Edit Brass Out Entry",
                    [("date",     "Date:",         "date",  None),
                     ("vendor",   "Vendor:",        "entry", None),
                     ("bag",      "Bag #:",         "entry", None),
                     ("weight",   "Weight (KG):",   "entry", None),
                     ("given_to", "Given To:",      "combo", casting_vs)],
                    {"date":     vals[6],
                     "vendor":   vals[1],
                     "bag":      vals[2],
                     "weight":   vals[3].replace(" KG",""),
                     "given_to": vals[4]},
                    _save)

    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today()
        try: wt=float(self.we.get())
        except: err("Weight must be a number."); return

        if self._eid:
            qry("UPDATE brass_out SET date=?,weight=? WHERE id=?",(date,wt,self._eid),commit=True)
            log_action(self.app.current_user,"UPDATE BRASS OUT",f"id={self._eid}")
            self._eid=None; self.sb.config(text="Confirm & Save")
            self.we.delete(0,"end"); self.refresh()
            return

        idx=self.bc.current()
        if idx<0: warn("Select a scrap bag."); return
        bid,bnum,vendor,_=self.bags[idx]
        given_to=self.gvc.get()
        qry("INSERT INTO brass_out (date,vendor,bag_number,weight,brass_in_id,given_to) VALUES (?,?,?,?,?,?)",
            (date,vendor,bnum,wt,bid,given_to),commit=True)
        qry("UPDATE brass_in SET remaining_weight=0,status='out' WHERE id=?",(bid,),commit=True)
        log_action(self.app.current_user,"CREATE BRASS OUT",f"Bag #{bnum} {wt:.3f} KG")
        self.load_bags(); self.we.delete(0,"end"); self.refresh()

    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT date,vendor,bag_number,weight,given_to FROM brass_out ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Date","Vendor","Bag #","Weight (KG)","Given To"],rows,"Brass Out")
# ─────────────────────────────────────────────────────────────────────────────
#  CASTING IN
# ─────────────────────────────────────────────────────────────────────────────
class CastingInPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app; self._eid=None
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        tk.Label(p,text="Casting In Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        vendors=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type IN ('Brass Out','Casting In') ORDER BY v.vendor_name",fetch="all") or []]
        first_vendor = vendors[0] if vendors else ""
        items=[i[0] for i in qry(
            "SELECT mi.item_name FROM master_items mi "
            "JOIN vendor_items vi ON vi.item_id=mi.id "
            "JOIN vendors v ON v.id=vi.vendor_id "
            "WHERE mi.item_category='Brass Item' AND v.vendor_name=? "
            "ORDER BY mi.item_name", (first_vendor,), fetch="all") or []]
        c=card(p)
        c.pack(fill="x", pady=(0,8))
        c.pack_propagate(True)
        f=ResponsiveForm(c); f.pack(fill="x", expand=False)
        self.de = f.add("Date:", date_picker(f, default=today()))
        self.vc = f.add("Vendor:", combo(f, vendors, vendors[0] if vendors else ""))
        self.vc.bind("<<ComboboxSelected>>", self._on_vendor_change)
        self.ic = f.add("Item:", combo(f, items, items[0] if items else ""))
        self.pe = f.add("Pieces:", entry(f))
        self.we = f.add("Weight (KG):", entry(f))
        self.sb = btn(f,"Save Casting In",self.save,bg=C_SUCCESS,hov=C_SUC_HOV)
        f.add_btn_row(self.sb)
        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="All Casting In Records",font=FSH,bg=C_BG).pack(anchor="w")
        cols=["Date","Vendor","Item","Pieces","Weight","Action","_id"]
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()


    def _form_row(form, label, widget):
        """Add a label+field pair to a ResponsiveForm. Returns widget."""
        return form.add(label, widget)

    def _on_vendor_change(self, event=None):
        vname = self.vc.get()
        rows = qry(
            "SELECT mi.item_name FROM master_items mi "
            "JOIN vendor_items vi ON vi.item_id=mi.id "
            "JOIN vendors v ON v.id=vi.vendor_id "
            "WHERE mi.item_category='Brass Item' AND v.vendor_name=? "
            "ORDER BY mi.item_name", (vname,), fetch="all") or []
        items = [r[0] for r in rows]
        self.ic["values"] = items
        self.ic.set(items[0] if items else "")
            

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,vendor,item,pieces,weight FROM casting_in ORDER BY id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],r[4],f"{r[5]:.3f} KG","",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4]],
                        "buttons":[("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                                    ("Delete", lambda v=vals: self._delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Vendor","Item","Pieces","Weight"],
                        {"Date":10,"Vendor":15,"Item":18,"Pieces":8,"Weight":10}, data)
    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); vendor=self.vc.get(); item=self.ic.get()
        if not vendor or not item: warn("Select vendor and item."); return
        try: pieces=int(self.pe.get()); wt=float(self.we.get())
        except: err("Pieces and Weight must be numbers."); return
        if self._eid:
            qry("UPDATE casting_in SET date=?,vendor=?,item=?,pieces=?,weight=? WHERE id=?",
                (date,vendor,item,pieces,wt,self._eid),commit=True)
            log_action(self.app.current_user,"UPDATE CASTING IN",f"id={self._eid}")
            self._eid=None; self.sb.config(text="Save Casting In")
        else:
            qry("INSERT INTO casting_in (date,vendor,item,pieces,weight) VALUES (?,?,?,?,?)",
                (date,vendor,item,pieces,wt),commit=True)
            log_action(self.app.current_user,"CREATE CASTING IN",f"Item '{item}'")
        self.pe.delete(0,"end"); self.we.delete(0,"end"); self.refresh()

    def _load_edit(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        cid = vals[6]
        vendors=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type IN ('Brass Out','Casting In') ORDER BY v.vendor_name",fetch="all") or []]
        items=[i[0] for i in qry("SELECT item_name FROM master_items WHERE item_category='Brass Item' ORDER BY item_name",fetch="all") or []]
        def _save(v):
            try: pcs = int(v["pieces"]); wt = float(v["weight"])
            except ValueError: err("Pieces and Weight must be numbers."); return False
            qry("UPDATE casting_in SET date=?,vendor=?,item=?,pieces=?,weight=? WHERE id=?",
                (v["date"], v["vendor"], v["item"], pcs, wt, cid), commit=True)
            log_action(self.app.current_user,"UPDATE CASTING IN",f"id={cid}")
            self.refresh()
        edit_dialog(self, "Edit Casting In Entry",
                    [("date","Date:","date",None),
                     ("vendor","Vendor:","combo",vendors),
                     ("item","Item:","combo",items),
                     ("pieces","Pieces:","entry",None),
                     ("weight","Weight (KG):","entry",None)],
                    {"date": vals[0], "vendor": vals[1], "item": vals[2],
                     "pieces": vals[3], "weight": vals[4].replace(" KG","")},
                    _save)

    def _delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm(f"Delete casting entry for '{vals[2]}'?"): return
        qry("DELETE FROM casting_in WHERE id=?",(vals[6],),commit=True)
        log_action(self.app.current_user,"DELETE CASTING IN",f"Item '{vals[2]}'"); self.refresh()

    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT date,vendor,item,pieces,weight FROM casting_in ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Date","Vendor","Item","Pieces","Weight (KG)"],rows,"Casting In")


# ─────────────────────────────────────────────────────────────────────────────
#  MACHINING
# ─────────────────────────────────────────────────────────────────────────────
class MachiningPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app; self._eid=None
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        tk.Label(p,text="Machining Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        ci_items = set(i[0] for i in qry("SELECT DISTINCT item FROM casting_in ORDER BY item", fetch="all") or [])
        op_items = set(i[0] for i in qry("SELECT DISTINCT item FROM opening_stock WHERE stage='casted'", fetch="all") or [])
        items = sorted(ci_items | op_items) 
        c=card(p)
        c.pack(fill="x", pady=(0,8))
        c.pack_propagate(True)
        f=ResponsiveForm(c); f.pack(fill="x", expand=False)
        self.de = f.add("Date:", date_picker(f, default=today()))
        self.ic = f.add("Item:", combo(f, items, items[0] if items else ""))
        self.ic.bind("<<ComboboxSelected>>",self._update_avail)
        self.avail_lbl = tk.Label(f, text="", bg=C_CARD, fg=C_GRAY, font=FSM, anchor="w")
        f.add_span(self.avail_lbl)
        self.pe = f.add("Pieces:", entry(f))
        self.we = f.add("Weight (KG):", entry(f))
        self.sb = btn(f,"Save Machining Entry",self.save,bg=C_SUCCESS,hov=C_SUC_HOV)
        f.add_btn_row(self.sb)
        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="All Machining Records",font=FSH,bg=C_BG).pack(anchor="w")
        cols=["Date","Item","Direction","Pieces","Weight","Action","_id"]
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def _r(self,f,lbl_text,r,default=""):
        tk.Label(f,text=lbl_text,bg=C_CARD,font=FNB,anchor="w").grid(row=r,column=0,sticky="w",pady=3)
        e=entry(f,default); e.grid(row=r,column=1,sticky="ew",padx=(8,0),pady=3); return e

    def _update_avail(self,_=None):
        item=self.ic.get()
        if not item: return
        op=qry("SELECT COALESCE(pieces,0) FROM opening_stock WHERE item=? AND stage='casted'",(item,),fetch="one")
        op_p=op[0] if op else 0
        cp=qry("SELECT COALESCE(SUM(pieces),0) FROM casting_in WHERE item=?",(item,),fetch="one")[0]
        mp=qry("SELECT COALESCE(SUM(pieces),0) FROM machining_register WHERE item=?",(item,),fetch="one")[0]
        cr=qry("SELECT COALESCE(SUM(pieces),0) FROM rejection_register WHERE item=? AND rejection_type IN ('Casted Rejection','Machined Rejection')",(item,),fetch="one")[0]
        avail=cp-mp-cr+op_p
        self.avail_lbl.config(text=f"Available for machining: {avail} pcs")

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,item,direction,pieces,weight FROM machining_register ORDER BY id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],r[4] or "-",f"{r[5]:.3f} KG" if r[5] else "-","",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4]],
                        "buttons":[("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                                ("Delete", lambda v=vals: self._delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Item","Direction","Pieces","Weight"],
                          {"Date":10,"Item":18,"Direction":10,"Pieces":8,"Weight":10}, data)
        
    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); item=self.ic.get()
        if not item: warn("Select an item."); return
        try: pieces=int(self.pe.get() or 0); wt=float(self.we.get() or 0)
        except: err("Pieces and Weight must be numbers."); return
        if self._eid:
            qry("UPDATE machining_register SET date=?,item=?,pieces=?,weight=? WHERE id=?",
                (date,item,pieces,wt,self._eid),commit=True)
            log_action(self.app.current_user,"UPDATE MACHINING",f"id={self._eid}")
            self._eid=None; self.sb.config(text="Save Machining Entry")
        else:
            op=qry("SELECT COALESCE(pieces,0) FROM opening_stock WHERE item=? AND stage='casted'",(item,),fetch="one")
            op_p=op[0] if op else 0
            cp=qry("SELECT COALESCE(SUM(pieces),0) FROM casting_in WHERE item=?",(item,),fetch="one")[0]
            mp=qry("SELECT COALESCE(SUM(pieces),0) FROM machining_register WHERE item=?",(item,),fetch="one")[0]
            cr=qry("SELECT COALESCE(SUM(pieces),0) FROM rejection_register WHERE item=? AND rejection_type IN ('Casted Rejection','Machined Rejection')",(item,),fetch="one")[0]
            if pieces>(cp-mp-cr+op_p): err(f"Cannot machine {pieces} pcs. Only {cp-mp-cr+op_p} available."); return
            qry("INSERT INTO machining_register (date,item,direction,pieces,weight) VALUES (?,?,'Machined',?,?)",
                (date,item,pieces,wt),commit=True)
            log_action(self.app.current_user,"CREATE MACHINING ENTRY",f"Item '{item}' {pieces} pcs")
        self.pe.delete(0,"end"); self.we.delete(0,"end"); self._update_avail(); self.refresh()


    def _load_edit(self, vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        mid = vals[6]
        items=[i[0] for i in qry("SELECT DISTINCT item FROM casting_in ORDER BY item",fetch="all") or []]
        def _save(v):
            try: pcs = int(v["pieces"] or 0); wt = float(v["weight"] or 0)
            except ValueError: err("Pieces and Weight must be numbers."); return False
            qry("UPDATE machining_register SET date=?,item=?,pieces=?,weight=? WHERE id=?",
                (v["date"], v["item"], pcs, wt, mid), commit=True)
            log_action(self.app.current_user,"UPDATE MACHINING",f"id={mid}")
            self._update_avail(); self.refresh()
        edit_dialog(self, "Edit Machining Entry",
                    [("date","Date:","date",None),
                     ("item","Item:","combo",items),
                     ("pieces","Pieces:","entry",None),
                     ("weight","Weight (KG):","entry",None)],
                    {"date": vals[0], "item": vals[1],
                     "pieces": vals[3] if vals[3]!="-" else "0",
                     "weight": vals[4].replace(" KG","") if vals[4]!="-" else "0"},
                    _save)

    def _delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm(f"Delete machining entry for '{vals[1]}'?"): return
        qry("DELETE FROM machining_register WHERE id=?",(vals[6],),commit=True)
        log_action(self.app.current_user,"DELETE MACHINING ENTRY",f"Item '{vals[1]}'"); self.refresh()

    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT date,item,direction,pieces,weight FROM machining_register ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Date","Item","Direction","Pieces","Weight (KG)"],rows,"Machining")


# ─────────────────────────────────────────────────────────────────────────────
#  PLATING
# ─────────────────────────────────────────────────────────────────────────────
class PlatingPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app
        tk.Label(self,text="Plating Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        top=card(self); top.pack(fill="x",pady=(0,8))
        tk.Label(top,text="Direction:",bg=C_CARD,font=FNB).pack(side="left")
        self.dv=tk.StringVar(value="Outward")
        for d in ("Outward","Inward"):
            tk.Radiobutton(top,text=d,variable=self.dv,value=d,bg=C_CARD,command=self._render).pack(side="left",padx=6)
        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)
        self._render()

    def _render(self):
        for w in self.body.winfo_children(): w.destroy()
        if self.dv.get()=="Outward": PlatingOutPanel(self.body,self.app).pack(fill="both",expand=True)
        else: PlatingInPanel(self.body,self.app).pack(fill="both",expand=True)


class PlatingOutPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app; self.batch=[]; self._eid=None
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        plating_vs=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Plating' ORDER BY v.vendor_name",fetch="all") or []]
        grinding_vs=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Grinding' ORDER BY v.vendor_name",fetch="all") or []]
        _all_item_names=[i[0] for i in qry("SELECT item_name FROM master_items ORDER BY item_name",fetch="all") or []]
        _opening_mach_map={}
        for _op_item,_op_pcs,_op_wt in qry("SELECT item, pieces, weight FROM opening_stock WHERE stage='machined'",fetch="all") or []:
            _opening_mach_map[_op_item]=(_op_pcs,_op_wt)

        all_items=[]
        for item_name in _all_item_names:
            m_pcs,m_wt=qry("SELECT COALESCE(SUM(pieces),0), COALESCE(SUM(weight),0) FROM machining_register WHERE item=?",[item_name],fetch="one")
            plate_out_pcs,plate_out_wt=qry(
                "SELECT COALESCE(SUM(pi.pieces),0), COALESCE(SUM(pi.weight),0) "
                "FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id "
                "WHERE pi.item=? AND pc.direction='Outward'",[item_name],fetch="one")
            grind_rej_pcs,grind_rej_wt=qry(
                "SELECT COALESCE(SUM(pieces),0), COALESCE(SUM(weight),0) FROM rejection_register "
                "WHERE item=? AND rejection_type='Grinder Rejection'",[item_name],fetch="one")
            op_pcs,op_wt=_opening_mach_map.get(item_name,(0,0.0))

            net_mach_pcs=m_pcs-plate_out_pcs-grind_rej_pcs+op_pcs
            if net_mach_pcs>0:
                all_items.append(item_name)
        f=card(p); f.pack(fill="x",pady=(0,6)); f.grid_columnconfigure(1,weight=1,minsize=200)
        tk.Label(f,text="Challan Details",bg=C_CARD,font=FNB).grid(row=0,column=0,columnspan=2,sticky="w")
        self.cno=self._r(f,"Challan No:",1)

        tk.Label(f,text="Date:",bg=C_CARD,font=FNB,anchor="w").grid(
    row=2,column=0,sticky="w",pady=3)
        self.de = date_picker(f, default=today())
        self.de.grid(row=2,column=1,sticky="ew",padx=(8,0),pady=3)
        self.vc=self._cb(f,"Plating Vendor:",3,plating_vs)
        f2=card(p); f2.pack(fill="x",pady=(0,6)); f2.grid_columnconfigure(1,weight=1,minsize=200)
        tk.Label(f2,text="Add Item to Batch",bg=C_CARD,font=FNB).grid(row=0,column=0,columnspan=2,sticky="w")
        self.ic=self._cb(f2,"Item:",1,all_items); self.gvc=self._cb(f2,"Grinding Vendor:",2,["-- None --"]+grinding_vs)
        self.pe=self._r(f2,"Pieces:",3); self.we=self._r(f2,"Weight (KG):",4); self.bag=self._r(f2,"Bag #:",5,"1")
        self.abb=btn(f2,"Add to Batch",self._add_batch,bg=C_WARNING,hov=C_WAR_HOV,fg="black")
        self.abb.grid(row=6,column=0,columnspan=2,pady=(8,0),sticky="w")
        tk.Label(p,text="Current Batch",font=FSH,bg=C_BG).pack(anchor="w")
        tf,self.btv=make_tree(p,["Item","Grind Vendor","Pieces","Weight","Bag"],4)
        tf.pack(fill="x")
        btn(p,"Save Challan",self._save_challan,bg=C_SUCCESS,hov=C_SUC_HOV).pack(anchor="w",pady=(6,8))

        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="All Plating Outward Records",font=FSH,bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def _r(self, f, lbl_text, r, default=""):
        tk.Label(f, text=lbl_text, bg=C_CARD, font=FNB, anchor="w").grid(
            row=r, column=0, sticky="w", pady=3)
        e = entry(f, default)
        e.grid(row=r, column=1, sticky="ew", padx=(8,0), pady=3)
        return e

    def _cb(self, f, lbl_text, r, values, default=""):
        tk.Label(f, text=lbl_text, bg=C_CARD, font=FNB, anchor="w").grid(
            row=r, column=0, sticky="w", pady=3)
        c = combo(f, values, default)
        c.grid(row=r, column=1, sticky="ew", padx=(8,0), pady=3)
        return c

    def _form_row(form, label, widget):
        """Add a label+field pair to a ResponsiveForm. Returns widget."""
        return form.add(label, widget)

    def _add_batch(self):
        item=self.ic.get()
        if not item: warn("Select an item."); return
        try: pcs=int(self.pe.get()); wt=float(self.we.get()); bag=int(self.bag.get() or 1)
        except: err("Pieces, Weight, Bag must be numbers."); return
        gv=self.gvc.get(); gv="" if gv=="-- None --" else gv

        if self._eid:
            qry("UPDATE plating_items SET item=?,grinding_vendor=?,pieces=?,weight=?,bag_number=? WHERE id=?",
                (item,gv,pcs,wt,bag,self._eid),commit=True)
            log_action(self.app.current_user,"UPDATE PLATING OUTWARD ITEM",f"id={self._eid}")
            self._eid=None; self.abb.config(text="Add to Batch")
            self.gvc.set("-- None --"); self.pe.delete(0,"end"); self.we.delete(0,"end")
            self.bag.delete(0,"end"); self.bag.insert(0,"1")
            self.refresh()
            return

        # Calculate available machined pieces for this item
        m_pcs=qry("SELECT COALESCE(SUM(pieces),0) FROM machining_register WHERE item=?",[item],fetch="one")[0]
        plate_out_pcs=qry(
            "SELECT COALESCE(SUM(pi.pieces),0) FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id "
            "WHERE pi.item=? AND pc.direction='Outward'",[item],fetch="one")[0]
        grind_rej_pcs=qry(
            "SELECT COALESCE(SUM(pieces),0) FROM rejection_register WHERE item=? AND rejection_type='Grinder Rejection'",[item],fetch="one")[0]
        op_pcs=qry("SELECT COALESCE(pieces,0) FROM opening_stock WHERE item=? AND stage='machined'",[item],fetch="one")
        op_pcs=op_pcs[0] if op_pcs else 0
        # also subtract pieces already in current unsaved batch for this item
        batch_pcs_for_item=sum(b["pcs"] for b in self.batch if b["item"]==item)
        avl=m_pcs - plate_out_pcs - grind_rej_pcs + op_pcs - batch_pcs_for_item
        if pcs > avl:
            err(f"Cannot send {pcs} pcs for '{item}'.\nOnly {avl} pcs available in Machined stock."); return

        self.batch.append({"item":item,"gv":gv,"pcs":pcs,"wt":wt,"bag":bag})
        self.btv.insert("","end",values=(item,gv or "N/A",pcs,f"{wt:.3f} KG",bag))
        self.pe.delete(0,"end"); self.we.delete(0,"end")
        
    def _save_challan(self):
        cno=self.cno.get().strip()
        if not cno or not self.batch: warn("Challan number and at least one item required."); return
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); vendor=self.vc.get()
        if not vendor: warn("Select a plating vendor."); return
        cid=qry("INSERT INTO plating_challans (challan_number,date,vendor,direction,status) VALUES (?,?,?,'Outward','open')",
                (cno,date,vendor),commit=True)
        for b in self.batch:
            qry("INSERT INTO plating_items (challan_id,item,pieces,weight,bag_number,remaining_pieces,grinding_vendor) VALUES (?,?,?,?,?,?,?)",
                (cid,b["item"],b["pcs"],b["wt"],b["bag"],b["pcs"],b["gv"]),commit=True)
        log_action(self.app.current_user,"CREATE PLATING OUTWARD",f"Challan #{cno} vendor '{vendor}'")
        self.batch=[]; self.btv.delete(*self.btv.get_children()); self.cno.delete(0,"end"); self.refresh()

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT pc.id,pc.challan_number,pc.date,pc.vendor,pi.item,pi.grinding_vendor,pi.pieces,pi.weight,pi.id FROM plating_challans pc JOIN plating_items pi ON pc.id=pi.challan_id WHERE pc.direction='Outward' ORDER BY pc.id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],r[4],r[5] or "N/A",r[6],f"{r[7]:.3f} KG","",r[0],r[8])   # vals[9] = plating_items.id, for Edit
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4],vals[5],vals[6]],
                        "buttons":[("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                                    ("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Ch No","Date","Vendor","Item","Grind Vendor","Pcs","Weight"],
                        {"Ch No":8,"Date":10,"Vendor":14,"Item":16,"Grind Vendor":14,"Pcs":6,"Weight":10}, data)

    def _load_edit(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        pi_id = vals[9]

        bag_row = qry("SELECT bag_number FROM plating_items WHERE id=?", (pi_id,), fetch="one")
        current_bag = str(bag_row[0]) if bag_row and bag_row[0] is not None else "1"

        grinding_vs = [v[0] for v in qry(
            "SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id "
            "WHERE vr.register_type='Grinding' ORDER BY v.vendor_name", fetch="all") or []]
        all_items = [i[0] for i in qry("SELECT item_name FROM master_items ORDER BY item_name", fetch="all") or []]

        def _save(v):
            gv = v["grind_vendor"] if v["grind_vendor"] and v["grind_vendor"] != "-- None --" else ""
            try:
                pcs = int(v["pieces"] or 0)
                wt  = float(v["weight"] or 0)
                bag = int(v["bag"] or 1)
            except ValueError:
                err("Pieces / Weight / Bag must be numbers."); return False
            qry("UPDATE plating_items SET item=?, grinding_vendor=?, pieces=?, weight=?, bag_number=? WHERE id=?",
                (v["item"], gv, pcs, wt, bag, pi_id), commit=True)
            log_action(self.app.current_user, "UPDATE PLATING OUTWARD ITEM", f"id={pi_id}")
            self.refresh()

        edit_dialog(self, "Edit Plating Outward Item",
            [("item",         "Item:",            "combo", all_items),
            ("grind_vendor", "Grinding Vendor:", "combo", ["-- None --"] + grinding_vs),
            ("pieces",       "Pieces:",          "entry", None),
            ("weight",       "Weight (KG):",     "entry", None),
            ("bag",          "Bag #:",           "entry", None)],
            {"item":         vals[3],
            "grind_vendor": vals[4] if vals[4] and vals[4] != "N/A" else "-- None --",
            "pieces":       str(vals[5]),
            "weight":       vals[6].replace(" KG", ""),
            "bag":          current_bag},
            _save)
    
    def _on_delete(self, vals):
        if self.app.current_role != "admin": warn("Admins only."); return
        challan_id = vals[8]
        item = vals[3]
        vendor = vals[2]
        pcs = vals[5]  # pieces in this outward entry

        # Check plating short in opening_stock
        op_short_row = qry(
            "SELECT id, pieces FROM opening_stock WHERE item=? AND stage='plating_short'",
            (item,), fetch="one")
        op_short_pcs = op_short_row[1] if op_short_row else 0

        if op_short_pcs >= pcs:
            # Enough in plating short — deduct and allow delete
            if not confirm(f"Delete outward challan #{vals[0]}?\n"
                        f"{pcs} pcs will be deducted from Plating Short opening stock."): return
            new_short = op_short_pcs - pcs
            if new_short == 0:
                qry("DELETE FROM opening_stock WHERE id=?", (op_short_row[0],), commit=True)
            else:
                qry("UPDATE opening_stock SET pieces=? WHERE id=?",
                    (new_short, op_short_row[0]), commit=True)
            qry("DELETE FROM plating_items WHERE challan_id=?", (challan_id,), commit=True)
            qry("DELETE FROM plating_challans WHERE id=?", (challan_id,), commit=True)
            log_action(self.app.current_user, "DELETE PLATING ENTRY",
                    f"Challan #{vals[0]}, deducted {pcs} pcs from plating_short"); self.refresh()
        else:
            inward_pcs = qry(
                "SELECT COALESCE(SUM(pi.pieces),0) FROM plating_challans pc JOIN plating_items pi ON pc.id=pi.challan_id "
                "WHERE pc.direction='Inward' AND pi.item=?",
                (item,), fetch="one")[0]
            total_out = qry(
                "SELECT COALESCE(SUM(pi.pieces),0) FROM plating_challans pc JOIN plating_items pi ON pc.id=pi.challan_id "
                "WHERE pc.direction='Outward' AND pi.item=?",
                (item,), fetch="one")[0]
            # after deleting this entry, remaining outward = total_out - pcs
            remaining_out = total_out - pcs
            if inward_pcs > remaining_out:
                warn(f"Cannot delete outward challan #{vals[0]}.\n"
                    f"Inward has {inward_pcs} pcs but only {remaining_out} pcs would remain in outward.\n"
                    f"Delete the inward entry first."); return
            if not confirm(f"Delete outward challan #{vals[0]}?"): return
            qry("DELETE FROM plating_items WHERE challan_id=?", (challan_id,), commit=True)
            qry("DELETE FROM plating_challans WHERE id=?", (challan_id,), commit=True)
            log_action(self.app.current_user, "DELETE PLATING ENTRY",
                    f"Challan #{vals[0]}"); self.refresh()
            
    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT pc.challan_number,pc.date,pc.vendor,pi.item,pi.grinding_vendor,pi.pieces,pi.weight FROM plating_challans pc JOIN plating_items pi ON pc.id=pi.challan_id WHERE pc.direction='Outward' ORDER BY pc.id DESC",fetch="all") or []
        _simple_excel(path,["Challan No","Date","Vendor","Item","Grind Vendor","Pieces","Weight (KG)"],rows,"Plating Outward")


class PlatingInPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app; self._eid=None
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        plating_vs=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Plating' ORDER BY v.vendor_name",fetch="all") or []]

        def _pending_items_for_vendor(vendor):
            if not vendor: return []
            out_rows=qry(
                "SELECT pi.item, COALESCE(SUM(pi.pieces),0), COALESCE(SUM(pi.weight),0), "
                "COALESCE(SUM(pi.pieces-pi.remaining_pieces),0), "
                "COALESCE(SUM(CASE WHEN pi.pieces>0 THEN pi.weight*(pi.pieces-pi.remaining_pieces)*1.0/pi.pieces ELSE 0 END),0) "
                "FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id "
                "WHERE pc.direction='Outward' AND pc.vendor=? GROUP BY pi.item",
                [vendor],fetch="all") or []
            in_rows=qry(
                "SELECT pi.item, COALESCE(SUM(pi.pieces),0), COALESCE(SUM(pi.weight),0) "
                "FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id "
                "WHERE pc.direction='Inward' AND pc.vendor=? GROUP BY pi.item",
                [vendor],fetch="all") or []
            # opening stock plating_short — vendor-agnostic, treat as already-out stock
            op_short_rows=qry(
                "SELECT item, COALESCE(pieces,0) FROM opening_stock WHERE stage='plating_short'",
                fetch="all") or []
            op_short_map={r[0]:r[1] for r in op_short_rows}

            new_in_map={r[0]:(r[1],r[2]) for r in in_rows}
            out_map={r[0]:(r[1],r[2],r[3],r[4]) for r in out_rows}

            # merge: items from outward challans + items from opening plating_short
            all_items=set(out_map.keys()) | set(op_short_map.keys())
            pending=[]
            for item in all_items:
                out_pcs,out_wt,old_in_pcs,old_in_wt=out_map.get(item,(0,0.0,0,0.0))
                new_in_pcs,new_in_wt=new_in_map.get(item,(0,0.0))
                op_s=op_short_map.get(item,0)
                short_pcs=(out_pcs+op_s)-(old_in_pcs+new_in_pcs)
                if short_pcs>0: pending.append(item)
            return pending

        f=card(p); f.pack(fill="x",pady=(0,8)); f.grid_columnconfigure(1,weight=1,minsize=200)
        tk.Label(f,text="Date:",bg=C_CARD,font=FNB,anchor="w").grid(
            row=0,column=0,sticky="w",pady=3
        )
        self.de = date_picker(f, default=today())
        self.de.grid(row=0,column=1,sticky="ew",padx=(8,0),pady=3)
        self.vc=self._cb(f,"Plating Vendor:",1,plating_vs)
        self.ic=self._cb(f,"Item:",2,_pending_items_for_vendor(self.vc.get()))
        self.sb=btn(f,"Save Inward Receipt",self.save,bg=C_SUCCESS,hov=C_SUC_HOV)
        self.sb.grid(row=5,column=0,columnspan=2,pady=(10,0),sticky="w")   # <-- ADD
        self.vc.bind("<<ComboboxSelected>>", self._refresh_items)
        self.pe=self._r(f,"Pieces Received:",3); self.we=self._r(f,"Weight (KG):",4)
        self._pending_items_for_vendor=_pending_items_for_vendor
        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="All Plating Inward Records",font=FSH,bg=C_BG).pack(anchor="w")
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()


    def _r(self, f, lbl_text, r, default=""):
        tk.Label(f, text=lbl_text, bg=C_CARD, font=FNB, anchor="w").grid(
            row=r, column=0, sticky="w", pady=3)
        e = entry(f, default)
        e.grid(row=r, column=1, sticky="ew", padx=(8,0), pady=3)
        return e

    def _cb(self, f, lbl_text, r, values, default=""):
        tk.Label(f, text=lbl_text, bg=C_CARD, font=FNB, anchor="w").grid(
            row=r, column=0, sticky="w", pady=3)
        c = combo(f, values, default)
        c.grid(row=r, column=1, sticky="ew", padx=(8,0), pady=3)
        return c

    def _form_row(form, label, widget):
        """Add a label+field pair to a ResponsiveForm. Returns widget."""
        return form.add(label, widget)

    def _refresh_items(self,_=None):
        items=self._pending_items_for_vendor(self.vc.get())
        self.ic.config(values=items)
        self.ic.set(items[0] if items else "")
    
    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT pc.id,pc.date,pc.vendor,pi.item,pi.pieces,pi.weight FROM plating_challans pc JOIN plating_items pi ON pc.id=pi.challan_id WHERE pc.direction='Inward' ORDER BY pc.id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],r[4],f"{r[5]:.3f} KG","",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4]],
                         "buttons":[("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                                    ("Delete", lambda v=vals: self._delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Vendor","Item","Pcs Received","Weight"],
                         {"Date":10,"Vendor":15,"Item":18,"Pcs Received":12,"Weight":10}, data)

    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); vendor=self.vc.get(); item=self.ic.get()
        if not vendor or not item: warn("Select vendor and item."); return
        try: pcs=int(self.pe.get()); wt=float(self.we.get())
        except: err("Pieces and Weight must be numbers."); return
        # total already sent out for this item (outward challans + opening plating_short)
        out_pcs = qry(
            "SELECT COALESCE(SUM(pi.pieces),0) FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id "
            "WHERE pc.direction='Outward' AND pc.vendor=? AND pi.item=?",
            (vendor, item), fetch="one")[0]
        op_short = qry(
            "SELECT COALESCE(pieces,0) FROM opening_stock WHERE item=? AND stage='plating_short'",
            (item,), fetch="one")
        op_short = op_short[0] if op_short else 0
        total_out = out_pcs + op_short

        # total already received back inward for this item from this vendor
        already_in = qry(
            "SELECT COALESCE(SUM(pi.pieces),0) FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id "
            "WHERE pc.direction='Inward' AND pc.vendor=? AND pi.item=?",
            (vendor, item), fetch="one")[0]

        if already_in + pcs > total_out:
            err(f"Cannot receive {pcs} pcs for '{item}'.\n"
                f"Total sent out: {total_out} pcs, already received: {already_in} pcs, "
                f"max you can receive now: {total_out - already_in} pcs."); return
        cid=qry("INSERT INTO plating_challans (challan_number,date,vendor,direction,status) VALUES (NULL,?,?,'Inward','received')",
                (date,vendor),commit=True)
        qry("INSERT INTO plating_items (challan_id,item,pieces,weight,bag_number,remaining_pieces,grinding_vendor) VALUES (?,?,?,?,NULL,0,NULL)",
            (cid,item,pcs,wt),commit=True)
        log_action(self.app.current_user,"CREATE PLATING INWARD",f"'{item}' {pcs} pcs from '{vendor}'")
        self.pe.delete(0,"end"); self.we.delete(0,"end"); self.refresh()


    def _load_edit(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        cid=vals[6]
        def _save(v):
            try: pcs=int(v["pieces"]); wt=float(v["weight"])
            except ValueError: err("Pieces and Weight must be numbers."); return False
            qry("UPDATE plating_challans SET date=? WHERE id=?",(v["date"],cid),commit=True)
            qry("UPDATE plating_items SET pieces=?,weight=? WHERE challan_id=?",(pcs,wt,cid),commit=True)
            log_action(self.app.current_user,"UPDATE PLATING INWARD",f"challan_id={cid}")
            self.refresh()
        edit_dialog(self, f"Edit Plating Inward — {vals[1]} / {vals[2]}",
                    [("date","Date:","date",None),
                     ("pieces","Pieces Received:","entry",None),
                     ("weight","Weight (KG):","entry",None)],
                    {"date": vals[0], "pieces": vals[3], "weight": vals[4].replace(" KG","")},
                    _save)

    def _delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm(f"Delete the plating inward receipt for '{vals[2]}' ({vals[1]})?"): return

        qry("DELETE FROM plating_items WHERE challan_id=?",(vals[6],),commit=True)
        qry("DELETE FROM plating_challans WHERE id=?",(vals[6],),commit=True)
        log_action(self.app.current_user,"DELETE PLATING INWARD",f"id={vals[6]}"); self.refresh()


    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT pc.date,pc.vendor,pi.item,pi.pieces,pi.weight FROM plating_challans pc JOIN plating_items pi ON pc.id=pi.challan_id WHERE pc.direction='Inward' ORDER BY pc.id DESC",fetch="all") or []
        _simple_excel(path,["Date","Vendor","Item","Pieces","Weight (KG)"],rows,"Plating Inward")


# ─────────────────────────────────────────────────────────────────────────────
#  MISCELLANEOUS
# ─────────────────────────────────────────────────────────────────────────────
class MiscPanel(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=C_BG); self.app=app; self._eid=None
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        tk.Label(p,text="Miscellaneous Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        misc_vs=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Miscellaneous' ORDER BY v.vendor_name",fetch="all") or []]
        all_items=[i[0] for i in qry("SELECT item_name FROM master_items ORDER BY item_name",fetch="all") or []]
        all_units=[u[0] for u in qry("SELECT unit_name FROM units ORDER BY unit_name",fetch="all") or []]
        self.vitem_map={}
        for vid,vn in qry("SELECT v.id,v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Miscellaneous'",fetch="all") or []:
            self.vitem_map[vn]=[r[0] for r in qry("SELECT m.item_name FROM master_items m JOIN vendor_items vi ON m.id=vi.item_id WHERE vi.vendor_id=?",(vid,),fetch="all") or []]
        self.all_items=all_items

        c=card(p)
        c.pack(fill="x", pady=(0,8))
        c.pack_propagate(True)
        f=ResponsiveForm(c); f.pack(fill="x", expand=False)
        self.de = f.add("Date:", date_picker(f, default=today()))
        self.ot = f.add("Option Type:", combo(f, ["purchase","in assembly"], "purchase"))
        self.ot.bind("<<ComboboxSelected>>",self._toggle)
        self.vc = f.add("Vendor:", combo(f, misc_vs, misc_vs[0] if misc_vs else ""))
        self.vc.bind("<<ComboboxSelected>>",self._upd_items)
        self.ic = f.add("Item:", combo(f, [], ""))
        self.pe = f.add("Pieces / Qty:", entry(f))
        self.uc = f.add("Unit:", combo(f, all_units, all_units[0] if all_units else ""))
        self.we = f.add("Weight (KG):", entry(f))
        self.sb = btn(f,"Save Entry",self.save,bg=C_SUCCESS,hov=C_SUC_HOV)
        f.add_btn_row(self.sb)
        self._toggle()

        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="All Miscellaneous Entries",font=FSH,bg=C_BG).pack(anchor="w")
        cols=["Date","Type","Vendor","Item","Pieces","Weight","Action","_id"]
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def _r(self,f,lbl_text,r,default=""):
        tk.Label(f,text=lbl_text,bg=C_CARD,font=FNB,anchor="w").grid(row=r,column=0,sticky="w",pady=3)
        e=entry(f,default); e.grid(row=r,column=1,sticky="ew",padx=(8,0),pady=3); return e

    def _toggle(self,_=None):
        if self.ot.get()=="purchase":
            self.vc.config(state="readonly")
            self._upd_items()
        else:
            self.vc.config(state="disabled")
            self.ic.config(values=self.all_items)
            if self.all_items: self.ic.set(self.all_items[0])

    def _upd_items(self,_=None):
        vn=self.vc.get(); items=self.vitem_map.get(vn,[])
        self.ic.config(values=items)
        if items: self.ic.set(items[0])

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,option_type,vendor,item,pieces,weight FROM miscellaneous_entries ORDER BY id DESC",fetch="all") or []:
            vals=(r[1],r[2].capitalize(),r[3] or "N/A",r[4],r[5] or "-",f"{r[6]:.3f} KG" if r[6] else "-","",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4],vals[5]],
                        "buttons":[("Edit", lambda v=vals: self._load_edit(v), C_GRAY),
                                    ("Delete", lambda v=vals: self._delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Type","Vendor","Item","Pieces","Weight"],
                        {"Date":10,"Type":10,"Vendor":14,"Item":16,"Pieces":8,"Weight":10}, data)

    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); ot=self.ot.get()
        vendor=self.vc.get() if ot=="purchase" else ""
        item=self.ic.get()
        if not item: warn("Select an item."); return
        try: pcs=int(self.pe.get() or 0)
        except: err("Pieces must be a number."); return
        try: wt=float(self.we.get() or 0)
        except: err("Weight must be a number."); return
        unit=self.uc.get()
        if self._eid:
            qry("UPDATE miscellaneous_entries SET date=?,option_type=?,vendor=?,item=?,pieces=?,weight=?,quantity_unit=? WHERE id=?",
                (date,ot,vendor,item,pcs,wt,unit,self._eid),commit=True)
            log_action(self.app.current_user,"UPDATE MISC ENTRY",f"id={self._eid}")
            self._eid=None; self.sb.config(text="Save Entry")
        else:
            qry("INSERT INTO miscellaneous_entries (date,option_type,vendor,item,pieces,weight,quantity_unit) VALUES (?,?,?,?,?,?,?)",
                (date,ot,vendor,item,pcs,wt,unit),commit=True)
            log_action(self.app.current_user,"CREATE MISCELLANEOUS ENTRY",f"'{item}' ({ot})")
        self.pe.delete(0,"end"); self.we.delete(0,"end"); self.refresh()

    def _on_click(self,event):
        col=self.tv.identify_column(event.x); row=self.tv.identify_row(event.y)
        if not row or col!="#7": return
        vals=self.tv.item(row,"values")
        bbox=self.tv.bbox(row,"#7")
        if bbox:
            rel_x=event.x-bbox[0]
            if rel_x<bbox[2]//2: self._load_edit(vals)
            else: self._delete(vals)

    def _load_edit(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        eid = vals[7]
        def _save(v):
            ot = v["type"].lower()
            vendor = v["vendor"] if ot=="purchase" else ""
            try: pcs = int(v["pieces"] or 0); wt = float(v["weight"] or 0)
            except ValueError: err("Pieces/Weight must be numbers."); return False
            qry("UPDATE miscellaneous_entries SET date=?,option_type=?,vendor=?,item=?,pieces=?,weight=? WHERE id=?",
                (v["date"], ot, vendor, v["item"], pcs, wt, eid), commit=True)
            log_action(self.app.current_user,"UPDATE MISC ENTRY",f"id={eid}")
            self.refresh()
        edit_dialog(self, "Edit Miscellaneous Entry",
                    [("date","Date:","date",None),
                     ("type","Option Type:","combo",["purchase","in assembly"]),
                     ("vendor","Vendor:","combo",list(self.vitem_map.keys())),
                     ("item","Item:","combo",self.all_items),
                     ("pieces","Pieces / Qty:","entry",None),
                     ("weight","Weight (KG):","entry",None)],
                    {"date": vals[0], "type": vals[1].lower(),
                     "vendor": vals[2] if vals[2]!="N/A" else "",
                     "item": vals[3],
                     "pieces": vals[4] if vals[4]!="-" else "0",
                     "weight": vals[5].replace(" KG","") if vals[5]!="-" else "0"},
                    _save)

    def _delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm(f"Delete misc entry for '{vals[3]}'?"): return
        qry("DELETE FROM miscellaneous_entries WHERE id=?",(vals[7],),commit=True)
        log_action(self.app.current_user,"DELETE MISCELLANEOUS ENTRY",f"Item '{vals[3]}'"); self.refresh()

    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT date,option_type,vendor,item,pieces,weight FROM miscellaneous_entries ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Date","Type","Vendor","Item","Pieces","Weight (KG)"],rows,"Miscellaneous")


# ─────────────────────────────────────────────────────────────────────────────
#  REJECTION
# ─────────────────────────────────────────────────────────────────────────────
class RejectionPanel(tk.Frame):
    SUBS=["inventory","brass_in_vendor","casting_vendor","misc_vendor"]
    SLABELS=["1. Process Rejection (Inventory)","2. Brass In Vendor",
             "3. Casting/Brass Out Vendor","4. Miscellaneous Vendor"]
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        tk.Label(self,text="Rejection Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        top=card(self); top.pack(fill="x",pady=(0,8))
        tk.Label(top,text="Category:",bg=C_CARD,font=FNB).pack(side="left")
        self.sc=combo(top,self.SLABELS,self.SLABELS[0],width=42)
        self.sc.pack(side="left",padx=10)
        self.sc.bind("<<ComboboxSelected>>",lambda e:self._render())
        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)
        self._render()

    def _render(self):
        for w in self.body.winfo_children(): w.destroy()
        idx=self.SLABELS.index(self.sc.get())
        RejectionSubPanel(self.body,self.app,self.SUBS[idx]).pack(fill="both",expand=True)


class RejectionSubPanel(tk.Frame):
    INV_TYPES=["Casted Rejection","Machined Rejection","Grinder Rejection","Plated Rejection"]
    def __init__(self,parent,app,sub):
        super().__init__(parent,bg=C_BG); self.app=app; self.sub=sub; self._eid=None
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        all_items=[i[0] for i in qry("SELECT item_name FROM master_items ORDER BY item_name",fetch="all") or []]
        all_vendors=[v[0] for v in qry("SELECT vendor_name FROM vendors ORDER BY vendor_name",fetch="all") or []]
        c=card(p)
        c.pack(fill="x", pady=(0,8))
        c.pack_propagate(True)
        f=ResponsiveForm(c); f.pack(fill="x", expand=False)
        self.de = f.add("Date:", date_picker(f, default=today()))
        if sub=="inventory":
            self.ic  = f.add("Item:", combo(f, all_items, all_items[0] if all_items else ""))
            self.rtc = f.add("Rejection Type:", combo(f, self.INV_TYPES, self.INV_TYPES[0]))
            self.vc  = f.add("Vendor:", combo(f, all_vendors, all_vendors[0] if all_vendors else ""))
            self.pe  = f.add("Pieces:", entry(f))
            self.we  = f.add("Weight (KG):", entry(f))
        elif sub=="brass_in_vendor":
            biv=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Brass In' ORDER BY v.vendor_name",fetch="all") or []]
            self.vc = f.add("Brass In Vendor:", combo(f, biv, biv[0] if biv else ""))
            self.we = f.add("Rejection Weight (KG):", entry(f))
        elif sub=="casting_vendor":
            cvs=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type IN ('Brass Out','Casting In') ORDER BY v.vendor_name",fetch="all") or []]
            self.vc = f.add("Vendor:", combo(f, cvs, cvs[0] if cvs else ""))
            self.we = f.add("Rejection Weight (KG):", entry(f))
        else:
            mvs=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Miscellaneous' ORDER BY v.vendor_name",fetch="all") or []]
            self.vc = f.add("Misc Vendor:", combo(f, mvs, mvs[0] if mvs else ""))
            self.ic = f.add("Item:", combo(f, all_items, all_items[0] if all_items else ""))
            self.pe = f.add("Pieces:", entry(f))
            self.we = f.add("Weight (KG):", entry(f))
        self.sb = btn(f,"Save Rejection Entry",self.save,bg=C_DANGER,hov=C_DAN_HOV)
        f.add_btn_row(self.sb)
        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="Rejection Records",font=FSH,bg=C_BG).pack(anchor="w")
        rf={"inventory":"rejection_type NOT IN ('Vendor Rejection','Brass In Vendor Rejection','Misc Vendor Rejection')",
            "brass_in_vendor":"rejection_type='Brass In Vendor Rejection'",
            "casting_vendor":"rejection_type='Vendor Rejection'",
            "misc_vendor":"rejection_type='Misc Vendor Rejection'"}
        self._filter=rf[sub]
        cols=["Date","Item","Type","Vendor","Pcs","Weight","Action","_id"]
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()


    def _form_row(form, label, widget):
        """Add a label+field pair to a ResponsiveForm. Returns widget."""
        return form.add(label, widget)
    
    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry(f"SELECT id,date,item,rejection_type,vendor,pieces,weight FROM rejection_register WHERE {self._filter} ORDER BY id DESC",fetch="all") or []:
            vals=(r[1],r[2] or "",r[3],r[4] or "",r[5] or "",f"{r[6]:.3f} KG" if r[6] else "","",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4],vals[5]],
                        "buttons":[("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Item","Type","Vendor","Pcs","Weight"],
                        {"Date":10,"Item":16,"Type":16,"Vendor":14,"Pcs":6,"Weight":10}, data)

    def _on_delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm("Delete this rejection entry?"): return
        qry("DELETE FROM rejection_register WHERE id=?",(vals[7],),commit=True)
        log_action(self.app.current_user,"DELETE REJECTION ENTRY",f"id={vals[7]}"); self.refresh()

    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today()
        vendor=self.vc.get() if hasattr(self,"vc") else ""
        item=self.ic.get() if hasattr(self,"ic") else ""
        try: wt=float(self.we.get() or 0)
        except: err("Weight must be a number."); return
        pcs=0
        if hasattr(self,"pe"):
            try: pcs=int(self.pe.get() or 0)
            except: err("Pieces must be a number."); return
        rt={"inventory":self.rtc.get() if hasattr(self,"rtc") else "",
            "brass_in_vendor":"Brass In Vendor Rejection",
            "casting_vendor":"Vendor Rejection",
            "misc_vendor":"Misc Vendor Rejection"}[self.sub]
        if not vendor: warn("Select a vendor."); return
        qry("INSERT INTO rejection_register (date,item,rejection_type,vendor,pieces,weight) VALUES (?,?,?,?,?,?)",
            (date,item,rt,vendor,pcs,wt),commit=True)
        if self.sub=="brass_in_vendor":
            bags=qry("SELECT id,remaining_weight FROM brass_in WHERE vendor=? AND status='available' AND remaining_weight>0 ORDER BY date ASC,id ASC",
                     (vendor,),fetch="all") or []
            rem=wt
            for bid,brem in bags:
                if rem<=0: break
                ded=min(brem,rem); new_rem=round(brem-ded,3)
                qry("UPDATE brass_in SET remaining_weight=?,status=? WHERE id=?",
                    (max(0.0,new_rem),"out" if new_rem<=0.0001 else "available",bid),commit=True)
                rem-=ded
        log_action(self.app.current_user,"CREATE REJECTION ENTRY",f"sub={self.sub} vendor='{vendor}' {wt:.3f} KG")
        if hasattr(self,"pe"): self.pe.delete(0,"end")
        self.we.delete(0,"end"); self.refresh()


    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry(f"SELECT date,item,rejection_type,vendor,pieces,weight FROM rejection_register WHERE {self._filter} ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Date","Item","Type","Vendor","Pieces","Weight (KG)"],rows,"Rejection")


# ─────────────────────────────────────────────────────────────────────────────
#  SALE REGISTER
# ─────────────────────────────────────────────────────────────────────────────
class SaleRegisterPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        tk.Label(p,text="Sale Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        info_box(p,"Selling a Box Pack deducts pcs from Assembly stock of all component items.")
        bps=qry("SELECT id,box_pack_name FROM box_pack_items ORDER BY box_pack_name",fetch="all") or []
        self.bp_map={b[1]:b[0] for b in bps}
        f=card(p); f.pack(fill="x",pady=(0,8)); f.grid_columnconfigure(1,weight=1,minsize=200)
        tk.Label(f,text="Date:",bg=C_CARD,font=FNB,anchor="w").grid(row=0,column=0,sticky="w",pady=3)
        self.de = date_picker(f, default=today())
        self.de.grid(row=0,column=1,sticky="ew",padx=(8,0),pady=3)
        self.pty=self._r(f,"Party / Customer:",1)
        self.bn=self._r(f,"Bill No.:",2)
        tk.Label(f,text="Box Pack Item:",bg=C_CARD,font=FNB,anchor="w").grid(row=3,column=0,sticky="w",pady=3)
        self.bpc=combo(f,[b[1] for b in bps],bps[0][1] if bps else "")
        self.bpc.grid(row=3,column=1,sticky="ew",padx=(8,0),pady=3)
        self.pce=self._r(f,"Pieces:",4)
        self.nce=self._r(f,"No. of Cartons:",5,"0")
        self.carton_f=tk.Frame(f,bg=C_CARD); self.carton_f.grid(row=6,column=0,columnspan=2,sticky="ew")
        self.carton_f.grid_columnconfigure(1,weight=1,minsize=150)
        self.carton_entries=[]
        self.nce.bind("<FocusOut>",self._rebuild_cartons)
        btn(f,"Save Sale Entry",self.save,bg=C_SUCCESS,hov=C_SUC_HOV).grid(row=7,column=0,columnspan=2,pady=(10,0),sticky="w")

        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        tk.Label(p,text="All Sale Entries",font=FSH,bg=C_BG).pack(anchor="w")
        # cols=["Date","Party","Bill No","Box Pack","Pcs","Cartons","Action","_id"]
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def _r(self,f,lbl_text,r,default=""):
        tk.Label(f,text=lbl_text,bg=C_CARD,font=FNB,anchor="w").grid(row=r,column=0,sticky="w",pady=3)
        e=entry(f,default); e.grid(row=r,column=1,sticky="ew",padx=(8,0),pady=3); return e

    def _rebuild_cartons(self,_=None):
        for w in self.carton_f.winfo_children(): w.destroy()
        self.carton_entries=[]
        try: n=int(self.nce.get() or 0)
        except: return
        for i in range(1,n+1):
            tk.Label(self.carton_f,text=f"Carton {i} Weight (KG):",bg=C_CARD,font=FN).grid(row=i-1,column=0,sticky="w",pady=2)
            e=entry(self.carton_f); e.grid(row=i-1,column=1,sticky="ew",pady=2,padx=(8,0)); self.carton_entries.append(e)

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,party,bill_no,box_pack_name,pcs,no_of_cartons FROM sale_register ORDER BY id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],r[4],r[5],r[6],"",r[0])
            data.append({"values":[vals[0],vals[1],vals[2],vals[3],vals[4],vals[5]],
                        "buttons":[("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Party","Bill No","Box Pack","Pcs","Cartons"],
                        {"Date":10,"Party":14,"Bill No":10,"Box Pack":16,"Pcs":6,"Cartons":8}, data)

    def _on_delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm(f"Delete sale entry? FG stock will be restored."): return
        sid=vals[7]
        sr=qry("SELECT box_pack_id,pcs FROM sale_register WHERE id=?",(sid,),fetch="one")
        if sr: qry("UPDATE finished_goods SET pcs=pcs+? WHERE box_pack_id=?",(sr[1],sr[0]),commit=True)
        qry("DELETE FROM sale_register_cartons WHERE sale_id=?",(sid,),commit=True)
        qry("DELETE FROM sale_register_deductions WHERE sale_id=?",(sid,),commit=True)
        qry("DELETE FROM sale_register WHERE id=?",(sid,),commit=True)
        log_action(self.app.current_user,"DELETE SALE REGISTER ENTRY",f"id={sid}"); self.refresh()
    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); party=self.pty.get().strip(); bill=self.bn.get().strip()
        bp_name=self.bpc.get()
        if not party or not bill or not bp_name: warn("Party, Bill No, Box Pack required."); return
        bp_id=self.bp_map.get(bp_name)
        try: pcs=int(self.pce.get()); nc=int(self.nce.get() or 0)
        except: err("Pieces/Cartons must be numbers."); return
        sid=qry("INSERT INTO sale_register (date,party,bill_no,box_pack_id,box_pack_name,pcs,no_of_cartons) VALUES (?,?,?,?,?,?,?)",
                (date,party,bill,bp_id,bp_name,pcs,nc),commit=True)
        for i,ce in enumerate(self.carton_entries,1):
            try: cw=float(ce.get() or 0)
            except: cw=0.0
            qry("INSERT INTO sale_register_cartons (sale_id,carton_number,weight) VALUES (?,?,?)",(sid,i,cw),commit=True)
        qry("UPDATE finished_goods SET pcs=pcs-? WHERE box_pack_id=?",(pcs,bp_id),commit=True)
        comps=qry("SELECT m.id,m.item_name FROM master_items m JOIN box_pack_components bc ON m.id=bc.item_id WHERE bc.box_pack_id=?",(bp_id,),fetch="all") or []
        for comp_id,comp_name in comps:
            qry("INSERT INTO sale_register_deductions (sale_id,item_id,item_name,pcs) VALUES (?,?,?,?)",(sid,comp_id,comp_name,pcs),commit=True)
        log_action(self.app.current_user,"CREATE SALE REGISTER ENTRY",f"'{bp_name}' {pcs} pcs to '{party}'")
        self.pty.delete(0,"end"); self.bn.delete(0,"end"); self.pce.delete(0,"end"); self.nce.delete(0,"end"); self.nce.insert(0,"0")
        for w in self.carton_f.winfo_children(): w.destroy(); self.carton_entries=[]; self.refresh()

    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT date,party,bill_no,box_pack_name,pcs,no_of_cartons FROM sale_register ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Date","Party","Bill No","Box Pack","Pcs","Cartons"],rows,"Sale Register")


# ─────────────────────────────────────────────────────────────────────────────
#  FINISHED GOODS
# ─────────────────────────────────────────────────────────────────────────────

def assembly_stock_map():
    """{item_name: pieces currently sitting in Assembly}. Same maths as the
    Inventory register: Plating Inward - Plated Rejection + Opening Assembly
    - Finished-Goods deductions (misc items: 'In Assembly' entries instead of
    plating)."""
    q = lambda sql: qry(sql, fetch="all") or []
    cat = {r[0]: r[1] for r in q("SELECT item_name,item_category FROM master_items")}
    op_a = {r[0]: (r[1] or 0) for r in q("SELECT item,pieces FROM opening_stock WHERE stage='assembly'")}
    out_done = {r[0]: r[1] for r in q(
        "SELECT pi.item,COALESCE(SUM(pi.pieces-pi.remaining_pieces),0) FROM plating_items pi "
        "JOIN plating_challans pc ON pi.challan_id=pc.id WHERE pc.direction='Outward' GROUP BY pi.item")}
    p_in = {r[0]: r[1] for r in q(
        "SELECT pi.item,COALESCE(SUM(pi.pieces),0) FROM plating_items pi "
        "JOIN plating_challans pc ON pi.challan_id=pc.id WHERE pc.direction='Inward' GROUP BY pi.item")}
    pl_rej = {r[0]: r[1] for r in q(
        "SELECT item,COALESCE(SUM(pieces),0) FROM rejection_register "
        "WHERE rejection_type='Plated Rejection' GROUP BY item")}
    misc_asm = {r[0]: r[1] for r in q(
        "SELECT item,COALESCE(SUM(pieces),0) FROM miscellaneous_entries "
        "WHERE option_type='in assembly' GROUP BY item")}
    fgd = {r[0]: r[1] for r in q(
        "SELECT fgd.item_name,COALESCE(SUM(fgd.pcs),0) FROM finished_goods_assembly_deductions fgd "
        "JOIN finished_goods_log fgl ON fgd.finished_goods_log_id=fgl.id GROUP BY fgd.item_name")}
    out = {}
    for name, c in cat.items():
        base = (out_done.get(name, 0) + p_in.get(name, 0) - pl_rej.get(name, 0)) \
               if c == "Brass Item" else misc_asm.get(name, 0)
        out[name] = base + op_a.get(name, 0) - fgd.get(name, 0)
    return out

class FinishedGoodsPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        tk.Label(p,text="Finished Goods Register",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        info_box(p,"Adding items here deducts component items from Assembly stock.")
        bps=qry("SELECT id,box_pack_name FROM box_pack_items ORDER BY box_pack_name",fetch="all") or []
        self.bp_map={b[1]:b[0] for b in bps}
        f=card(p); f.pack(fill="x",pady=(0,8)); f.grid_columnconfigure(1,weight=1,minsize=200)
        tk.Label(f,text="Date:",bg=C_CARD,font=FNB,anchor="w").grid(
            row=0,column=0,sticky="w",pady=3
        )
        self.de = date_picker(f, default=today())
        self.de.grid(row=0,column=1,sticky="ew",padx=(8,0),pady=3)
        tk.Label(f,text="Box Pack Item:",bg=C_CARD,font=FNB,anchor="w").grid(row=1,column=0,sticky="w",pady=3)
        self.bc=combo(f,[b[1] for b in bps],bps[0][1] if bps else ""); self.bc.grid(row=1,column=1,sticky="ew",padx=(8,0),pady=3)
        self.bc.bind("<<ComboboxSelected>>", lambda e: self._update_max())
        self.pe=self._r(f,"Pieces:",2)
        self.maxl=tk.Label(f,text="",bg=C_CARD,font=FN,fg=C_GRAY,anchor="w")
        self.maxl.grid(row=3,column=1,sticky="w",padx=(8,0))
        btn(f,"Add to Finished Goods",self.save,bg=C_SUCCESS,hov=C_SUC_HOV).grid(row=4,column=0,columnspan=2,pady=(10,0),sticky="w")

        tk.Label(p,text="Current Finished Goods Stock",font=FSH,bg=C_BG).pack(anchor="w",pady=(8,0))
        tf,self.stv=make_tree(p,["Box Pack Item","Available Pcs"],4,{"Box Pack Item":250,"Available Pcs":100})
        tf.pack(fill="x")

        tk.Label(p,text="Finished Goods Addition Log",font=FSH,bg=C_BG).pack(anchor="w",pady=(8,0))
        cols=["Date","Box Pack","Pcs Added","Action","_id"]
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        bar=tk.Frame(p,bg=C_BG); bar.pack(fill="x",pady=(0,4))
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left")
        self.refresh()

    def _r(self,f,lbl_text,r,default=""):
        tk.Label(f,text=lbl_text,bg=C_CARD,font=FNB,anchor="w").grid(row=r,column=0,sticky="w",pady=3)
        e=entry(f,default); e.grid(row=r,column=1,sticky="ew",padx=(8,0),pady=3); return e

    def _max_pcs(self, bp_name):
        """(max pieces, limiting item): the box pack can only be made as many
        times as its scarcest component has pieces in Assembly."""
        bp_id=self.bp_map.get(bp_name)
        comps=[r[0] for r in qry("SELECT m.item_name FROM master_items m JOIN box_pack_components bc ON m.id=bc.item_id WHERE bc.box_pack_id=?",(bp_id,),fetch="all") or []]
        if not comps: return None, None
        stock=assembly_stock_map()
        low=min(comps,key=lambda n: stock.get(n,0))
        return max(0,stock.get(low,0)), low

    def _update_max(self):
        n,item=self._max_pcs(self.bc.get())
        self.maxl.config(text="No component items set for this box pack." if n is None
                         else f"Max you can add: {n} pcs  (limited by {item} in Assembly)")

    def refresh(self):
        self._update_max()
        tree_fill(self.stv,qry("SELECT box_pack_name,pcs FROM finished_goods ORDER BY box_pack_name",fetch="all") or [])
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,date,box_pack_name,pcs FROM finished_goods_log ORDER BY id DESC",fetch="all") or []:
            vals=(r[1],r[2],r[3],"",r[0])
            data.append({"values":[vals[0],vals[1],vals[2]],
                        "buttons":[("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Date","Box Pack","Pcs Added"], {"Date":10,"Box Pack":20,"Pcs Added":10}, data)

    def _on_delete(self,vals):
        if self.app.current_role!="admin": warn("Admins only."); return
        if not confirm("Delete this entry? Assembly deductions will be reversed."): return
        lid=vals[4]
        r=qry("SELECT box_pack_id,pcs FROM finished_goods_log WHERE id=?",(lid,),fetch="one")
        if r: qry("UPDATE finished_goods SET pcs=pcs-? WHERE box_pack_id=?",(r[1],r[0]),commit=True)
        qry("DELETE FROM finished_goods_assembly_deductions WHERE finished_goods_log_id=?",(lid,),commit=True)
        qry("DELETE FROM finished_goods_log WHERE id=?",(lid,),commit=True)
        log_action(self.app.current_user,"DELETE FINISHED GOODS ENTRY",f"log id={lid}"); self.refresh()

    def save(self):
        date=self.de.get_date().strftime("%Y-%m-%d") or today(); bp_name=self.bc.get()
        if not bp_name: warn("Select a Box Pack."); return
        try: pcs=int(self.pe.get())
        except: err("Pieces must be a number."); return
        if pcs<=0: warn("Pieces must be more than 0."); return                      # NEW
        mx,low=self._max_pcs(bp_name)                                                # NEW
        if mx is not None and pcs>mx:                                                # NEW
            warn(f"Only {mx} pcs of '{bp_name}' can be made.\n'{low}' has just {mx} pcs in Assembly."); return   # NEW
        bp_id=self.bp_map[bp_name]
        ex=qry("SELECT id FROM finished_goods WHERE box_pack_id=?",(bp_id,),fetch="one")
        if ex: qry("UPDATE finished_goods SET pcs=pcs+? WHERE box_pack_id=?",(pcs,bp_id),commit=True)
        else: qry("INSERT INTO finished_goods (box_pack_id,box_pack_name,pcs) VALUES (?,?,?)",(bp_id,bp_name,pcs),commit=True)
        log_id=qry("INSERT INTO finished_goods_log (date,box_pack_id,box_pack_name,pcs) VALUES (?,?,?,?)",(date,bp_id,bp_name,pcs),commit=True)
        comps=qry("SELECT m.id,m.item_name FROM master_items m JOIN box_pack_components bc ON m.id=bc.item_id WHERE bc.box_pack_id=?",(bp_id,),fetch="all") or []
        for cid,cname in comps:
            qry("INSERT INTO finished_goods_assembly_deductions (finished_goods_log_id,item_id,item_name,pcs) VALUES (?,?,?,?)",(log_id,cid,cname,pcs),commit=True)
        log_action(self.app.current_user,"CREATE FINISHED GOODS ENTRY",f"'{bp_name}' {pcs} pcs")
        self.pe.delete(0,"end"); self.refresh()

    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT date,box_pack_name,pcs FROM finished_goods_log ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Date","Box Pack","Pcs Added"],rows,"Finished Goods Log")


# ─────────────────────────────────────────────────────────────────────────────
#  VENDOR LEDGERS
# ─────────────────────────────────────────────────────────────────────────────
class VendorLedgersPanel(tk.Frame):
    TYPES=["Casting Vendor (Brass Out & Casting In)",
           "Scrap Vendor (Brass In)",
           "Miscellaneous Vendor",
           "Grinding Vendor",
           "Plating Vendor"]

    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        tk.Label(self,text="Vendor Ledgers",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        top=card(self); top.pack(fill="x",pady=(0,8))
        tk.Label(top,text="Ledger Type:",bg=C_CARD,font=FNB).pack(side="left")
        self.tc=combo(top,self.TYPES,self.TYPES[0],width=45)
        self.tc.pack(side="left",padx=10)
        self.tc.bind("<<ComboboxSelected>>",lambda e:self._render())
        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)
        self._render()

    def _render(self):
        for w in self.body.winfo_children(): w.destroy()
        m={
            self.TYPES[0]: CastingLedgerPanel,
            self.TYPES[1]: BrassInLedgerPanel,
            self.TYPES[2]: MiscLedgerPanel,
            self.TYPES[3]: GrindLedgerPanel,
            self.TYPES[4]: PlateLedgerPanel,
        }
        m[self.tc.get()](self.body,self.app).pack(fill="both",expand=True)


class CastingLedgerPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        top=card(self); top.pack(fill="x",pady=(0,6))
        vendors=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type IN ('Brass Out','Casting In') ORDER BY v.vendor_name",fetch="all") or []]
        tk.Label(top,text="Vendor:",bg=C_CARD,font=FNB).pack(side="left")
        self.vc=combo(top,["All"]+vendors,"All",width=22); self.vc.pack(side="left",padx=6)
        tk.Label(top,text="From:",bg=C_CARD,font=FNB).pack(side="left",padx=(10,0))
        self.df=date_picker(top, default=(date.today() - timedelta(days=30)).isoformat()); self.df.pack(side="left",padx=4)
        tk.Label(top,text="To:",bg=C_CARD,font=FNB).pack(side="left")
        self.dt=date_picker(top); self.dt.pack(side="left",padx=4)
        btn(top,"Show",self.refresh,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left",padx=(8,0))
        if EXCEL_OK: btn(top,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left",padx=(6,0))

        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)
        self.refresh()

    def refresh(self):
        for w in self.body.winfo_children(): w.destroy()
        vf=self.vc.get(); df=self.df.get_date().strftime("%Y-%m-%d"); dt=self.dt.get_date().strftime("%Y-%m-%d")
        vendors=[vf] if vf!="All" else [v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type IN ('Brass Out','Casting In') ORDER BY v.vendor_name",fetch="all") or []]
        sf=ScrollFrame(self.body); sf.pack(fill="both",expand=True); p=sf.inner

        for vname in vendors:
            # ── Brass Out rows for the selected From→To window ──
            bq="SELECT date,bag_number,weight,given_to FROM brass_out WHERE given_to=?"
            bp=[vname]
            if df: bq+=" AND date>=?"; bp.append(df)
            if dt: bq+=" AND date<=?"; bp.append(dt)
            bq+=" ORDER BY date ASC"
            bo_rows=qry(bq,bp,fetch="all") or []
            total_bo_period=sum(r[2] for r in bo_rows)

            # ── Casting In rows for the selected From→To window ──
            cq="SELECT date,item,pieces,weight FROM casting_in WHERE vendor=?"
            cp_=[vname]
            if df: cq+=" AND date>=?"; cp_.append(df)
            if dt: cq+=" AND date<=?"; cp_.append(dt)
            cq+=" ORDER BY date ASC"
            ci_rows=qry(cq,cp_,fetch="all") or []
            total_ci_period=sum(r[3] for r in ci_rows)

            # ── Vendor rejection for the selected window ──
            rq="SELECT COALESCE(SUM(weight),0) FROM rejection_register WHERE vendor=? AND rejection_type='Vendor Rejection'"
            rp=[vname]
            if df: rq+=" AND date>=?"; rp.append(df)
            if dt: rq+=" AND date<=?"; rp.append(dt)
            rej_wt_period=qry(rq,rp,fetch="one")[0] or 0.0

            adj_period=max(0.0,total_ci_period-rej_wt_period)
            burn_period=adj_period*0.05
            accounted_period=adj_period+burn_period

            # ── Opening balance = running Brass balance strictly BEFORE "From" date ──
            # (same formula app.py / _get_vendor_ledger_data uses, so the on-screen
            #  number always matches the webpage and the Excel export.)
            opening_balance=None
            if df:
                pre_bo=qry("SELECT COALESCE(SUM(weight),0) FROM brass_out WHERE given_to=? AND date<?",[ vname,df],fetch="one")[0] or 0.0
                pre_ci=qry("SELECT COALESCE(SUM(weight),0) FROM casting_in WHERE vendor=? AND date<?",[vname,df],fetch="one")[0] or 0.0
                pre_rej=qry("SELECT COALESCE(SUM(weight),0) FROM rejection_register WHERE vendor=? AND rejection_type='Vendor Rejection' AND date<?",[vname,df],fetch="one")[0] or 0.0
                pre_adj=max(0.0,pre_ci-pre_rej); pre_burn=pre_adj*0.05; pre_accounted=pre_adj+pre_burn
                opening_balance=pre_bo-pre_accounted

            # Net balance for the selected From→To window = opening balance (everything
            # before "From") + this period's Brass Out − this period's Accounted Return.
            period_net=(opening_balance or 0.0)+total_bo_period-accounted_period
            # ── Balance from the very start of records through "To" date only —
            #    independent of whichever "From" date happens to be selected. ──

            # ── Overall (all-time, unfiltered) balance — this NEVER changes with the
            #    date filter, unlike period_net above, so it's the number to trust as
            #    "the" balance for the vendor. ──
            all_bo=qry("SELECT COALESCE(SUM(weight),0) FROM brass_out WHERE given_to=?",[vname],fetch="one")[0] or 0.0
            all_ci=qry("SELECT COALESCE(SUM(weight),0) FROM casting_in WHERE vendor=?",[vname],fetch="one")[0] or 0.0
            all_rej=qry("SELECT COALESCE(SUM(weight),0) FROM rejection_register WHERE vendor=? AND rejection_type='Vendor Rejection'",[vname],fetch="one")[0] or 0.0
            all_adj=max(0.0,all_ci-all_rej); all_burn=all_adj*0.05; all_accounted=all_adj+all_burn
            overall_net=all_bo-all_accounted
            # card
            c=card(p); c.pack(fill="x",pady=(0,10))
            tk.Label(c,text=f"Vendor: {vname}",font=FSH,bg=C_CARD,fg=C_PRIMARY).grid(row=0,column=0,columnspan=4,sticky="w",pady=(0,6))

            # Brass Out table
            tk.Label(c,text="Brass Out (Scrap Given)",bg=C_DANGER,fg="white",font=FNB,padx=6,pady=3).grid(row=1,column=0,columnspan=2,sticky="ew")
            n_bo_rows=len(bo_rows)+(2 if opening_balance is not None else 1)
            tf,btv=make_tree(c,["Date","Bag #","Weight (KG)"],min(n_bo_rows,7),{"Date":190,"Bag #":80,"Weight (KG)":110})
            tf.grid(row=2,column=0,columnspan=2,sticky="ew",pady=(4,0))
            for col in ("Date","Bag #","Weight (KG)"):
                btv.column(col, anchor="center")
                btv.heading(col, anchor="center")
            if opening_balance is not None:
                btv.insert("","end",values=(f"Brass Balance Before {df}","",f"{opening_balance:.3f}"),tags=("ob",))
            for d,bag,wt,gt in bo_rows: btv.insert("","end",values=(fmt_date_display(d),f"#{bag}",f"{wt:.3f}"))
            btv.insert("","end",values=("Total Brass Out (Period)","",f"{total_bo_period:.3f}"),tags=("tot",))
            btv.tag_configure("tot",background=C_GREEN_ROW,font=FNB)
            btv.tag_configure("ob",background=C_YEL_ROW,font=FNB)

            # Casting In table
            tk.Label(c,text="Casting In",bg=C_SUCCESS,fg="white",font=FNB,padx=6,pady=3).grid(row=1,column=2,columnspan=2,sticky="ew",padx=(20,0))
            tf2,ctv=make_tree(c,["Date","Item","Pcs","Weight (KG)"],min(len(ci_rows)+4,7),{"Date":118,"Item":190,"Pcs":60,"Weight (KG)":110})
            tf2.grid(row=2,column=2,columnspan=2,sticky="ew",pady=(4,0),padx=(20,0))
            for col in ("Date","Item","Pcs","Weight (KG)"):
                ctv.column(col, anchor="center")
                ctv.heading(col, anchor="center")
            for d,it,pcs_,wt in ci_rows: ctv.insert("","end",values=(fmt_date_display(d),it,pcs_,f"{wt:.3f}"))
            for lbl_t,val in [("Total Casting In Weight",f"{total_ci_period:.3f}"),
                               ("Less: Vendor Rejection",f"-{rej_wt_period:.3f}"),
                               ("Adjusted Subtotal",f"{adj_period:.3f}"),
                               ("Burning Loss (5%)",f"{burn_period:.3f}"),
                               ("Total Accounted Return",f"{accounted_period:.3f}")]:
                ctv.insert("","end",values=("",lbl_t,"",val),tags=("sumrow",))
            ctv.tag_configure("sumrow",background=C_YEL_ROW,font=FNB)

            # ── Overall net balance (always the same, regardless of the date filter) ──
            ov_c = C_DANGER if overall_net>0 else (C_SUCCESS if overall_net<0 else C_PRIMARY)
            ov_t = (f"Net Balance (Overall): {overall_net:+.3f} KG  "
                    f"({'Due From Vendor' if overall_net>0 else 'Advance With Vendor' if overall_net<0 else 'Fully Settled'})")
            tk.Label(c,text=ov_t,bg=C_YEL_ROW,fg=ov_c,font=FNB,padx=8,pady=6).grid(
                row=3,column=0,columnspan=4,sticky="ew",pady=(10,0))

            # ── Net balance for the selected From→To window ──
            pd_c = C_DANGER if period_net>0 else (C_SUCCESS if period_net<0 else C_PRIMARY)
            pd_t = (f"Net Balance from {df} to {dt}: {period_net:+.3f} KG  "
                    f"({'Due From Vendor' if period_net>0 else 'Advance With Vendor' if period_net<0 else 'Fully Settled'})")
            tk.Label(c,text=pd_t,bg=C_YEL_ROW,fg=pd_c,font=FNB,padx=8,pady=6).grid(
                row=4,column=0,columnspan=4,sticky="ew",pady=(4,0))

            c.grid_columnconfigure(0,weight=1); c.grid_columnconfigure(2,weight=1)
    def export_excel(self):
        if not EXCEL_OK:
            err("openpyxl not installed.")
            return

        from openpyxl.utils import get_column_letter

        vendor_sel    = self.vc.get()
        vendor_filter = "" if vendor_sel == "All" else vendor_sel
        date_from     = self.df.get_date().strftime("%Y-%m-%d")
        date_to       = self.dt.get_date().strftime("%Y-%m-%d")

        # ---- ask where to save ----
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if vendor_filter:
            safe_v = "".join(c for c in vendor_filter if c not in '\\/:*?"<>|')
            default_name = f"{safe_v} - Vendor Ledgers.xlsx"
        else:
            default_name = f"Vendor_Ledgers_brass_out_casting_in_{stamp}.xlsx"

        path = filedialog.asksaveasfilename(
            title="Export Casting Vendor Ledger to Excel",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
        )
        if not path:
            return

        # ---- pull data with app.py's ledger function ----
        _app_mod = sys.modules[_get_vendor_ledger_data.__module__]

        class _FakeRequest:
            def __init__(self, values):
                self.values = values

        _orig_request = _app_mod.request
        _app_mod.request = _FakeRequest({
            "cast_date_from": date_from,
            "cast_date_to":   date_to,
        })
        conn = get_conn()
        try:
            vendor_data = _get_vendor_ledger_data(
                conn.cursor(), "brass_out_casting_in", vendor_filter
            )
        finally:
            conn.close()
            _app_mod.request = _orig_request

        # ---- styles / helpers ----
        thin   = Side(style="thin", color="B7B7B7")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        bold   = Font(bold=True)
        title_font = Font(bold=True, size=11)
        center = Alignment(horizontal="center", vertical="center")

        used_names = set()

        def safe_sheet_name(name):
            cleaned = name
            for ch in ["\\", "/", "*", "[", "]", ":", "?"]:
                cleaned = cleaned.replace(ch, "")
            cleaned = (cleaned.strip() or "Vendor")[:31]
            candidate = cleaned
            n = 2
            while candidate in used_names:
                suffix = f"_{n}"
                candidate = cleaned[:31 - len(suffix)] + suffix
                n += 1
            used_names.add(candidate)
            return candidate

        def to_float(text):
            try:
                return float(str(text).replace(" KG", "").replace(",", "").strip())
            except ValueError:
                return None

        def put(ws, r, c, value, is_bold=False, num_fmt=None, align=None):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = border
            if is_bold:
                cell.font = bold
            if num_fmt:
                cell.number_format = num_fmt
            if align:
                cell.alignment = align
            return cell

        # ---- layout: Brass Out (left) | gap | Casting In (right) ----
        BO_COLS  = ["Date", "Bag Number", "Weight (KG)"]
        CI_COLS  = ["Date", "Item", "Pieces", "Weight (KG)"]
        BO_WIDTH = [18, 16, 14]
        CI_WIDTH = [18, 20, 12, 14]
        GAP      = 2
        CI_START = len(BO_COLS) + GAP + 1
        TOTAL_W  = len(BO_COLS) + GAP + len(CI_COLS)

        BO_BOLD = {"Total Brass Out (Period)"}
        CI_BOLD = {"Total", "Less: Vendor Rejection", "Adjusted Subtotal",
                   "Burning Loss (5%)", "Total Accounted Return"}

        wb = Workbook()
        wb.remove(wb.active)

        if not vendor_data:
            ws = wb.create_sheet("Export")
            ws["A1"] = "No data to export."

        # ---- one sheet per vendor ----
        for v in vendor_data:
            ws = wb.create_sheet(title=safe_sheet_name(v["vendor"]))
            bo_all = v["tables"][0]["rows"]
            ci_all = v["tables"][1]["rows"]

            t = ws.cell(row=1, column=1, value=v["vendor"])
            t.font = title_font
            t.alignment = center
            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=TOTAL_W)

            parts = []
            if v.get("cast_date_from"):
                parts.append(f"From: {v['cast_date_from']}")
            if v.get("cast_date_to"):
                parts.append(f"To: {v['cast_date_to']}")

            section_row = 2
            if parts:
                info = ws.cell(row=2, column=1,
                               value="Date Filter: " + "  |  ".join(parts))
                info.font = Font(italic=True, color="555555")
                ws.merge_cells(start_row=2, start_column=1,
                               end_row=2, end_column=TOTAL_W)
                section_row = 3

            h1 = ws.cell(row=section_row, column=1, value="Brass Out (Scrap Given)")
            h1.font = title_font
            h1.alignment = center
            h1.border = border
            ws.merge_cells(start_row=section_row, start_column=1,
                           end_row=section_row, end_column=len(BO_COLS))

            h2 = ws.cell(row=section_row, column=CI_START, value="Casting In")
            h2.font = title_font
            h2.alignment = center
            h2.border = border
            ws.merge_cells(start_row=section_row, start_column=CI_START,
                           end_row=section_row,
                           end_column=CI_START + len(CI_COLS) - 1)

            hdr_row = section_row + 1
            for i, name in enumerate(BO_COLS):
                put(ws, hdr_row, 1 + i, name, True, align=center)
                ws.column_dimensions[get_column_letter(1 + i)].width = BO_WIDTH[i]
            for i, name in enumerate(CI_COLS):
                put(ws, hdr_row, CI_START + i, name, True, align=center)
                ws.column_dimensions[get_column_letter(CI_START + i)].width = CI_WIDTH[i]
            for g in range(len(BO_COLS) + 1, CI_START):
                ws.column_dimensions[get_column_letter(g)].width = 3

            data_start = hdr_row + 1
            max_rows   = max(len(bo_all), len(ci_all))

            for idx in range(max_rows):
                r = data_start + idx

                # Brass Out side
                if idx < len(bo_all):
                    row   = bo_all[idx]
                    label = str(row[0]) if row[0] is not None else ""

                    if label.startswith("Brass Balance Before"):
                        put(ws, r, 1, label, True)
                        put(ws, r, 2, "", True)
                        ws.merge_cells(start_row=r, start_column=1,
                                       end_row=r, end_column=2)
                        wt = to_float(row[2])
                        put(ws, r, 3, wt if wt is not None else row[2],
                            True, "0.000")
                    else:
                        is_sum = label in BO_BOLD
                        put(ws, r, 1, label, is_sum)
                        put(ws, r, 2, str(row[1]) if row[1] is not None else "", is_sum)
                        wt = to_float(row[2]) if row[2] != "" else None
                        put(ws, r, 3, wt if wt is not None else "",
                            is_sum, "0.000" if wt is not None else None)
                else:
                    for c in range(1, len(BO_COLS) + 1):
                        ws.cell(row=r, column=c).border = border

                # Casting In side
                if idx < len(ci_all):
                    row    = ci_all[idx]
                    label  = str(row[0]) if row[0] is not None else ""
                    is_sum = label in CI_BOLD

                    put(ws, r, CI_START, label, is_sum)
                    put(ws, r, CI_START + 1,
                        str(row[1]) if row[1] is not None else "", is_sum)

                    pcs_raw = str(row[2]) if row[2] is not None else ""
                    try:
                        pcs_val = int(pcs_raw)
                    except ValueError:
                        pcs_val = pcs_raw
                    put(ws, r, CI_START + 2, pcs_val, is_sum)

                    wt = to_float(row[3]) if row[3] != "" else None
                    put(ws, r, CI_START + 3, wt if wt is not None else "",
                        is_sum, "0.000" if wt is not None else None)
                else:
                    for c in range(CI_START, CI_START + len(CI_COLS)):
                        ws.cell(row=r, column=c).border = border

            if v.get("summary"):
                sr = data_start + max_rows + 1
                put(ws, sr, 1, v["summary"], True)
                ws.merge_cells(start_row=sr, start_column=1,
                               end_row=sr, end_column=TOTAL_W)

        # ---- save ----
        try:
            wb.save(path)
        except PermissionError:
            messagebox.showerror(
                "Export Failed",
                "The Excel file is currently open.\n\n"
                "Please close the file and try again."
            )
            return
        except Exception as e:
            messagebox.showerror("Export Failed",
                                 f"Could not create Excel file:\n\n{e}")
            return

        try:
            log_action(
                self.app.current_user,
                "EXPORT EXCEL",
                "Exported Vendor Ledgers (brass_out_casting_in) for "
                + (f"vendor {vendor_filter}" if vendor_filter else "all vendors")
            )
        except Exception:
            pass

        messagebox.showinfo("Exported", f"Excel file saved successfully:\n\n{path}")

class BrassInLedgerPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        top=card(self); top.pack(fill="x",pady=(0,6))
        vendors=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Brass In' ORDER BY v.vendor_name",fetch="all") or []]
        tk.Label(top,text="Vendor:",bg=C_CARD,font=FNB).pack(side="left")
        self.vc=combo(top,["All"]+vendors,"All",width=28); self.vc.pack(side="left",padx=10)
        btn(top,"Show",self.refresh,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left",padx=(8,0))
        if EXCEL_OK: btn(top,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left",padx=(6,0))
        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)
        self.refresh()

    def refresh(self):
        for w in self.body.winfo_children(): w.destroy()
        vf=self.vc.get()
        vendors=[vf] if vf!="All" else [v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Brass In' ORDER BY v.vendor_name",fetch="all") or []]
        sf=ScrollFrame(self.body); sf.pack(fill="both",expand=True); p=sf.inner
        for vname in vendors:
            rows=qry("SELECT date,bag_number,weight,remaining_weight,status FROM brass_in WHERE vendor=? ORDER BY date ASC",(vname,),fetch="all") or []
            c=card(p)
            c.pack(fill="x", pady=(0,8))
            c.pack_propagate(True)
            tk.Label(c,text=f"Vendor: {vname}",font=FSH,bg=C_CARD,fg=C_SUCCESS).pack(anchor="w")
            tf,tv=make_tree(c,["Date","Bag #","Original","Remaining","Status"],min(len(rows)+2,10),{"Date":118,"Bag #":70,"Original":90,"Remaining":90,"Status":80})
            tf.pack(fill="x",pady=(4,0))
            total_in=0.0; total_rem=0.0
            for d,bag,wt,rem,st in rows:
                tv.insert("","end",values=(fmt_date_display(d),f"#{bag}",f"{wt:.3f}",f"{rem:.3f}",st.capitalize()))
                total_in+=wt
                if st=="available": total_rem+=rem
            rej=qry("SELECT COALESCE(SUM(weight),0) FROM rejection_register WHERE vendor=? AND rejection_type='Brass In Vendor Rejection'",(vname,),fetch="one")[0] or 0.0
            for lbl_t,val in [("Total Supplied",f"{total_in:.3f} KG"),("Total Rejected",f"-{rej:.3f} KG"),("Available Stock",f"{total_rem:.3f} KG")]:
                tv.insert("","end",values=("",lbl_t,"",val,""),tags=("sum",))
            tv.tag_configure("sum",background=C_GREEN_ROW,font=FNB)

    def export_excel(self):
        if not EXCEL_OK: return
        vf = self.vc.get()
        _export_vendor_ledger(self.app.current_user, "brass_in",
                              "" if vf == "All" else vf)

class MiscLedgerPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        top=card(self); top.pack(fill="x",pady=(0,6))
        vendors=[v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Miscellaneous' ORDER BY v.vendor_name",fetch="all") or []]
        tk.Label(top,text="Vendor:",bg=C_CARD,font=FNB).pack(side="left")
        self.vc=combo(top,["All"]+vendors,"All",width=28); self.vc.pack(side="left",padx=10)
        btn(top,"Show",self.refresh,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left",padx=(8,0))
        if EXCEL_OK: btn(top,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left",padx=(6,0))
        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)
        self.refresh()

    def refresh(self):
        for w in self.body.winfo_children(): w.destroy()
        vf=self.vc.get()
        vendors=[vf] if vf!="All" else [v[0] for v in qry("SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='Miscellaneous' ORDER BY v.vendor_name",fetch="all") or []]
        sf=ScrollFrame(self.body); sf.pack(fill="both",expand=True); p=sf.inner
        for vname in vendors:
            rows=qry("SELECT date,item,pieces,weight FROM miscellaneous_entries WHERE vendor=? AND option_type='purchase' ORDER BY date ASC",(vname,),fetch="all") or []
            c=card(p)
            c.pack(fill="x", pady=(0,8))
            c.pack_propagate(True)
            tk.Label(c,text=f"Vendor: {vname}",font=FSH,bg=C_CARD,fg="#6f42c1").pack(anchor="w")
            tf,tv=make_tree(c,["Date","Item","Pieces","Weight"],min(len(rows)+1,10),{"Date":118,"Item":160,"Pieces":80,"Weight":100})
            tf.pack(fill="x",pady=(4,0))
            tp=0; tw=0.0
            for d,it,pcs_,wt in rows:
                tv.insert("","end",values=(fmt_date_display(d),it,pcs_,f"{wt:.3f} KG")); tp+=pcs_; tw+=wt
            if rows:
                tv.insert("","end",values=("Total","",tp,f"{tw:.3f} KG"),tags=("sum",))
                tv.tag_configure("sum",background=C_GREEN_ROW,font=FNB)

    def export_excel(self):
        if not EXCEL_OK: return
        vf = self.vc.get()
        _export_vendor_ledger(self.app.current_user, "miscellaneous",
                              "" if vf == "All" else vf)

class _GrindPlateLedgerBase(tk.Frame):
    def __init__(self,parent,app,reg_type,vendor_col):
        super().__init__(parent,bg=C_BG); self.app=app; self.reg_type=reg_type; self.vendor_col=vendor_col
        vendors=[v[0] for v in qry(f"SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='{reg_type}' ORDER BY v.vendor_name",fetch="all") or []]
        top=card(self); top.pack(fill="x",pady=(0,6))
        tk.Label(top,text="Vendor:",bg=C_CARD,font=FNB).pack(side="left")
        self.vc=combo(top,["All"]+vendors,"All",width=22); self.vc.pack(side="left",padx=6)
        tk.Label(top,text="From:",bg=C_CARD,font=FNB).pack(side="left",padx=(8,0))
        self.df=date_picker(top, default=(date.today() - timedelta(days=30)).isoformat()); self.df.pack(side="left",padx=4)
        tk.Label(top,text="To:",bg=C_CARD,font=FNB).pack(side="left")
        self.dt=date_picker(top); self.dt.pack(side="left",padx=4)
        btn(top,"Show",self.refresh,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left",padx=(8,0))
        if EXCEL_OK: btn(top,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left",padx=(6,0))
        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)
        self.refresh()

    def refresh(self):
        for w in self.body.winfo_children(): w.destroy()
        vf=self.vc.get(); df=self.df.get_date().strftime("%Y-%m-%d"); dt=self.dt.get_date().strftime("%Y-%m-%d")
        vendors=[vf] if vf!="All" else [v[0] for v in qry(f"SELECT DISTINCT v.vendor_name FROM vendors v JOIN vendor_registers vr ON v.id=vr.vendor_id WHERE vr.register_type='{self.reg_type}' ORDER BY v.vendor_name",fetch="all") or []]
        sf=ScrollFrame(self.body); sf.pack(fill="both",expand=True); p=sf.inner
        for vname in vendors:
            q=f"SELECT pc.date,pc.challan_number,pi.item,pi.pieces,COALESCE((SELECT rl.rate FROM rate_list rl JOIN master_items mi ON mi.id=rl.item_id JOIN vendors vnd ON vnd.id=rl.vendor_id WHERE mi.item_name=pi.item AND vnd.vendor_name=? LIMIT 1),0) FROM plating_items pi JOIN plating_challans pc ON pi.challan_id=pc.id WHERE {self.vendor_col}=? AND pc.direction='Outward'"
            qp=[vname,vname]
            if df: q+=" AND pc.date>=?"; qp.append(df)
            if dt: q+=" AND pc.date<=?"; qp.append(dt)
            q+=" ORDER BY pc.date ASC"
            rows=qry(q,qp,fetch="all") or []
            c=card(p)
            c.pack(fill="x", pady=(0,8))
            c.pack_propagate(True)
            tk.Label(c,text=f"Vendor: {vname}",font=FSH,bg=C_CARD,fg="#fd7e14").pack(anchor="w")
            tf,tv=make_tree(c,["Date","Ch.No","Item","Pcs","Rate","Total Cost"],min(len(rows)+1,10),
                            {"Date":118,"Ch.No":80,"Item":150,"Pcs":60,"Rate":80,"Total Cost":100})
            tf.pack(fill="x",pady=(4,0))
            tp=0; tc=0.0
            for d,ch,it,pcs_,rate in rows:
                rate=float(rate); cost=pcs_*rate
                tv.insert("","end",values=(fmt_date_display(d),ch or "",it,pcs_,f"Rs.{rate:.2f}" if rate>0 else "-",f"Rs.{cost:.2f}" if rate>0 else "-"))
                tp+=pcs_; tc+=cost
            if rows:
                tv.insert("","end",values=("","","Total",tp,"",f"Rs.{tc:.2f}"),tags=("sum",))
                tv.tag_configure("sum",background=C_GREEN_ROW,font=FNB)

    def export_excel(self):
        if not EXCEL_OK: return
        vf = self.vc.get()
        _export_vendor_ledger(
            self.app.current_user,
            "grinding" if self.reg_type == "Grinding" else "plating",
            "" if vf == "All" else vf,
            self.df.get_date().strftime("%Y-%m-%d"),
            self.dt.get_date().strftime("%Y-%m-%d")
        )

class GrindLedgerPanel(_GrindPlateLedgerBase):
    def __init__(self,p,a): super().__init__(p,a,"Grinding","pi.grinding_vendor")

class PlateLedgerPanel(_GrindPlateLedgerBase):
    def __init__(self,p,a): super().__init__(p,a,"Plating","pc.vendor")


# ─────────────────────────────────────────────────────────────────────────────
#  ITEM RECORDS
# ─────────────────────────────────────────────────────────────────────────────
class ItemRecordsPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        tk.Label(self,text="Item Records",font=FSH,bg=C_BG).pack(anchor="w",pady=(0,4))
        top=card(self); top.pack(fill="x",pady=(0,8))
        items=[i[0] for i in qry("SELECT item_name FROM master_items ORDER BY item_name",fetch="all") or []]
        tk.Label(top,text="Select Item:",bg=C_CARD,font=FNB).pack(side="left")
        self.ic=combo(top,items,items[0] if items else "",width=35); self.ic.pack(side="left",padx=10)
        btn(top,"Show Records",self.show,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left")
        if EXCEL_OK: btn(top,"Export All (Excel)",self.export_all,bg="#217346",hov="#1a5e38").pack(side="left",padx=(8,0))
        self.body=tk.Frame(self,bg=C_BG); self.body.pack(fill="both",expand=True)

    def show(self):
        for w in self.body.winfo_children(): w.destroy()
        item=self.ic.get()
        if not item: warn("Select an item."); return
        cat_row=qry("SELECT item_category FROM master_items WHERE item_name=?",(item,),fetch="one")
        cat=cat_row[0] if cat_row else "Brass Item"

        sf=ScrollFrame(self.body); sf.pack(fill="both",expand=True)
        p=sf.inner; p.configure(bg=C_BG)

        def section(title,cols,rows_data,r,c,colspan=1):
            frame=card(p); frame.grid(row=r,column=c,columnspan=colspan,sticky="nsew",padx=4,pady=4)
            tk.Label(frame,text=title,font=FNB,bg=C_CARD,fg=C_PRIMARY).pack(anchor="w")
            tf,tv=make_tree(frame,cols,min(len(rows_data)+1,8)); tf.pack(fill="x")
            for row_ in rows_data: tv.insert("","end",values=(fmt_date_display(row_[0]),)+tuple(row_[1:]) if row_ else row_)

        p.grid_columnconfigure(0,weight=1); p.grid_columnconfigure(1,weight=1)

        if cat=="Brass Item":
            cr=qry("SELECT date,vendor,pieces,weight FROM casting_in WHERE item=? ORDER BY id DESC",(item,),fetch="all") or []
            section("1. Casting",["Date","Vendor","Pcs","Weight"],[(r[0],r[1],r[2],f"{r[3]:.3f}") for r in cr],0,0)
            mr=qry("SELECT date,pieces,weight FROM machining_register WHERE item=? ORDER BY id DESC",(item,),fetch="all") or []
            section("2. Machined",["Date","Pcs","Weight"],[(r[0],r[1] or "-",f"{r[2]:.3f}" if r[2] else "-") for r in mr],0,1)
            gr=qry("SELECT c.date,c.challan_number,i.grinding_vendor,i.pieces,i.weight FROM plating_challans c JOIN plating_items i ON c.id=i.challan_id WHERE i.item=? AND c.direction='Outward' AND i.grinding_vendor IS NOT NULL AND i.grinding_vendor!='' ORDER BY c.id DESC",(item,),fetch="all") or []
            section("3. Grinding",["Date","Ch#","Vendor","Pcs","Wt"],[(r[0],r[1],r[2],r[3],f"{r[4]:.3f}") for r in gr],1,0)
            pr=qry("SELECT c.date,COALESCE(c.challan_number,'(Inward)'),c.vendor,CASE WHEN c.direction='Inward' THEN i.pieces ELSE (i.pieces-i.remaining_pieces) END FROM plating_challans c JOIN plating_items i ON c.id=i.challan_id WHERE i.item=? AND ((c.direction='Outward' AND (i.pieces-i.remaining_pieces)>0) OR (c.direction='Inward' AND i.pieces>0)) ORDER BY c.date DESC,c.id DESC",(item,),fetch="all") or []
            section("4. Plating (Received)",["Date","Ch#","Vendor","Pcs"],pr,1,1)
            sr=qry("SELECT sr.date,sr.bill_no,sd.pcs FROM sale_register_deductions sd JOIN sale_register sr ON sd.sale_id=sr.id WHERE sd.item_name=? ORDER BY sr.id DESC",(item,),fetch="all") or []
            section("5. Sales",["Date","Bill No","Pcs"],sr,2,0,2)
        else:
            pur=qry("SELECT date,vendor,pieces,weight FROM miscellaneous_entries WHERE item=? AND option_type='purchase' ORDER BY id DESC",(item,),fetch="all") or []
            section("Purchase",["Date","Vendor","Pcs","Wt"],[(r[0],r[1],r[2],f"{r[3]:.3f}") for r in pur],0,0)
            asm=qry("SELECT date,pieces,weight FROM miscellaneous_entries WHERE item=? AND option_type='in assembly' ORDER BY id DESC",(item,),fetch="all") or []
            section("In Assembly",["Date","Pcs","Wt"],[(r[0],r[1] or "-",f"{r[2]:.3f}" if r[2] else "-") for r in asm],0,1)
            rj=qry("SELECT date,vendor,pieces,weight FROM rejection_register WHERE item=? AND rejection_type='Misc Vendor Rejection' ORDER BY id DESC",(item,),fetch="all") or []
            section("Misc Rejection",["Date","Vendor","Pcs","Wt"],[(r[0],r[1],r[2],f"{r[3]:.3f}") for r in rj],1,0)
            sr=qry("SELECT sr.date,sr.bill_no,sd.pcs FROM sale_register_deductions sd JOIN sale_register sr ON sd.sale_id=sr.id WHERE sd.item_name=? ORDER BY sr.id DESC",(item,),fetch="all") or []
            section("Sales",["Date","Bill No","Pcs"],sr,1,1)

    def export_all(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        items=qry("SELECT item_name,item_category FROM master_items ORDER BY item_name",fetch="all") or []
        wb=Workbook()
        used=set()
        for iname,icat in items:
            sname=iname[:31]
            if sname in used:
                sname=iname[:28]+"_"+str(len(used))
            used.add(sname)
            ws=wb.create_sheet(title=sname)
            col=1
            def block(title,headers,rows):
                nonlocal col
                ws.cell(1,col,title).font=Font(bold=True)
                for ci,h in enumerate(headers):
                    ws.cell(2,col+ci,h).font=Font(bold=True)
                for ri,r in enumerate(rows):
                    for ci,v in enumerate(r):
                        ws.cell(3+ri,col+ci,v)
                col+=len(headers)+1
            if icat=="Brass Item":
                block("Casting",["Date","Vendor","Pcs","Weight"],qry("SELECT date,vendor,pieces,weight FROM casting_in WHERE item=? ORDER BY id DESC",(iname,),fetch="all") or [])
                block("Machining",["Date","Pcs","Weight"],qry("SELECT date,pieces,weight FROM machining_register WHERE item=? ORDER BY id DESC",(iname,),fetch="all") or [])
            else:
                block("Purchase",["Date","Vendor","Pcs","Weight"],qry("SELECT date,vendor,pieces,weight FROM miscellaneous_entries WHERE item=? AND option_type='purchase' ORDER BY id DESC",(iname,),fetch="all") or [])
                block("Assembly",["Date","Pcs","Weight"],qry("SELECT date,pieces,weight FROM miscellaneous_entries WHERE item=? AND option_type='in assembly' ORDER BY id DESC",(iname,),fetch="all") or [])
        if "Sheet" in wb.sheetnames: wb.remove(wb["Sheet"])
        wb.save(path); messagebox.showinfo("Exported",f"Saved to {path}")


# ═══════════════════════════════════════════════════════════════════════════════
#  AUDIT LOGS
# ═══════════════════════════════════════════════════════════════════════════════
class AuditLogsScreen(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        tk.Label(self,text="System Audit Logs",font=FH,bg=C_BG,fg="#222").pack(anchor="w",pady=(0,8))
        bar=tk.Frame(self,bg=C_BG); bar.pack(fill="x",pady=(0,6))
        btn(bar,"Refresh",self.refresh,bg=C_GRAY,hov=C_GRA_HOV).pack(side="left")
        if EXCEL_OK: btn(bar,"Export Excel",self.export_excel,bg="#217346",hov="#1a5e38").pack(side="left",padx=(8,0))
        tf,self.tv=make_tree(self,["Timestamp","User","Action","Details"],25,
                             {"Timestamp":160,"User":100,"Action":160,"Details":450})
        tf.pack(fill="both",expand=True)
        self.refresh()

    def refresh(self):
        rows=qry("SELECT timestamp,username,action,details FROM audit_logs ORDER BY id DESC",fetch="all") or []
        tree_fill(self.tv,rows)

    def export_excel(self):
        if not EXCEL_OK: return
        path=filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not path: return
        rows=qry("SELECT timestamp,username,action,details FROM audit_logs ORDER BY id DESC",fetch="all") or []
        _simple_excel(path,["Timestamp","User","Action","Details"],rows,"Audit Logs")


# ═══════════════════════════════════════════════════════════════════════════════
#  MANAGE USERS
# ═══════════════════════════════════════════════════════════════════════════════
class ManageUsersScreen(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg=C_BG); self.app=app
        sf=ScrollFrame(self); sf.pack(fill="both",expand=True); p=sf.inner
        tk.Label(p,text="Manage Users",font=FH,bg=C_BG,fg="#222").pack(anchor="w",pady=(0,8))
        f=card(p); f.pack(fill="x",pady=(0,10)); f.grid_columnconfigure(1,weight=1,minsize=200)
        tk.Label(f,text="Username:",bg=C_CARD,font=FNB,anchor="w").grid(row=0,column=0,sticky="w",pady=4)
        self.ue=entry(f); self.ue.grid(row=0,column=1,sticky="ew",padx=(8,0),pady=4)
        tk.Label(f,text="Password:",bg=C_CARD,font=FNB,anchor="w").grid(row=1,column=0,sticky="w",pady=4)
        self.pe_=entry(f); self.pe_.grid(row=1,column=1,sticky="ew",padx=(8,0),pady=4)
        tk.Label(f,text="Role:",bg=C_CARD,font=FNB,anchor="w").grid(row=2,column=0,sticky="w",pady=4)
        self.re=combo(f,["employee","admin"],"employee"); self.re.grid(row=2,column=1,sticky="ew",padx=(8,0),pady=4)
        btn(f,"Create User",self.create,bg=C_SUCCESS,hov=C_SUC_HOV).grid(row=3,column=0,columnspan=2,pady=(10,0),sticky="w")
        self.list_holder = tk.Frame(p, bg=C_BG); self.list_holder.pack(fill="both", expand=True)
        self.refresh()

    def refresh(self):
        for w in self.list_holder.winfo_children(): w.destroy()
        data=[]
        for r in qry("SELECT id,username,password,role FROM users",fetch="all") or []:
            vals=(r[1],r[2],r[3],"",r[0])
            data.append({"values":[vals[0],vals[1],vals[2]],
                        "buttons":[("Delete", lambda v=vals: self._on_delete(v), "#dc3545")]})
        make_action_list(self.list_holder, ["Username","Password","Role"], {"Username":14,"Password":12,"Role":8}, data)

    def _on_delete(self,vals):
        if vals[0]==self.app.current_user: warn("Cannot delete your own account."); return
        if not confirm(f"Delete user '{vals[0]}'?"): return
        qry("DELETE FROM users WHERE id=?",(vals[4],),commit=True)
        log_action(self.app.current_user,"DELETE USER",f"Removed '{vals[0]}'"); self.refresh()

    def create(self):
        u=self.ue.get().strip(); pw=self.pe_.get().strip(); role=self.re.get() or "employee"
        if not u or not pw: warn("Username and Password required."); return
        try:
            qry("INSERT INTO users (username,password,role) VALUES (?,?,?)",(u,pw,role),commit=True)
            log_action(self.app.current_user,"CREATE USER",f"Created '{u}'")
        except sqlite3.IntegrityError:
            err("Username already exists."); return
        self.ue.delete(0,"end"); self.pe_.delete(0,"end"); self.refresh()


# ═══════════════════════════════════════════════════════════════════════════════
#  EXPORT HELPER
# ═══════════════════════════════════════════════════════════════════════════════
def _simple_excel(path, headers, rows, sheet_name="Sheet"):
    if not EXCEL_OK: return
    wb = Workbook(); ws = wb.active; ws.title = sheet_name[:31]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="007bff")
    for row in rows:
        ws.append(list(row))
    wb.save(path)
    messagebox.showinfo("Exported", f"Saved to {path}")

def _safe_sheet_name(name, used):
    cleaned = name
    for ch in ["\\", "/", "*", "[", "]", ":", "?"]:
        cleaned = cleaned.replace(ch, "")
    cleaned = (cleaned.strip() or "Vendor")[:31]
    candidate = cleaned
    n = 2
    while candidate in used:
        suffix = f"_{n}"
        candidate = cleaned[:31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def _export_vendor_ledger(user, ledger_sub_type, vendor_filter="",
                          date_from="", date_to="", include_paid=True):
    """
    Same layout as the web export: one sheet per vendor, named after the vendor.
    Works for: 'brass_in', 'miscellaneous', 'grinding', 'plating'.
    Set include_paid=False if you don't want the Paid (Yes/No) column.
    """
    if not EXCEL_OK:
        err("openpyxl not installed.")
        return

    from openpyxl.utils import get_column_letter

    # ---- ask where to save ----
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if vendor_filter:
        safe_v = "".join(c for c in vendor_filter if c not in '\\/:*?"<>|')
        default_name = f"{safe_v} - Vendor Ledgers.xlsx"
    else:
        default_name = f"Vendor_Ledgers_{ledger_sub_type}_{stamp}.xlsx"

    path = filedialog.asksaveasfilename(
        title="Export Vendor Ledger to Excel",
        defaultextension=".xlsx",
        initialfile=default_name,
        filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
    )
    if not path:
        return

    # ---- pull data using app.py's ledger function ----
    app_mod = sys.modules[_get_vendor_ledger_data.__module__]

    class _FakeRequest:
        def __init__(self, values):
            self.values = values

    orig_request = app_mod.request
    app_mod.request = _FakeRequest({
        "cast_date_from":  date_from, "cast_date_to":  date_to,
        "grind_date_from": date_from, "grind_date_to": date_to,
        "plate_date_from": date_from, "plate_date_to": date_to,
    })
    conn = get_conn()
    try:
        vendor_data = _get_vendor_ledger_data(
            conn.cursor(), ledger_sub_type, vendor_filter
        )
    finally:
        conn.close()
        app_mod.request = orig_request

    # ---- styles ----
    thin   = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold   = Font(bold=True)
    title_font = Font(bold=True, size=11)
    center = Alignment(horizontal="center", vertical="center")

    def cell_value(header, val):
        if val is None or val == "":
            return None
        h = header.lower()
        if any(k in h for k in ("weight", "pieces", "rate", "total cost")):
            try:
                f = float(str(val).replace(",", ""))
                return int(f) if ("pieces" in h and f.is_integer()) else f
            except ValueError:
                return val
        return val

    def number_format(header):
        h = header.lower()
        if "weight" in h:
            return "0.000"
        if "rate" in h or "total cost" in h:
            return "0.00"
        if "pieces" in h:
            return "0"
        return None

    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()

    if not vendor_data:
        ws = wb.create_sheet("Export")
        ws["A1"] = "No data to export."

    for v in vendor_data:
        ws = wb.create_sheet(title=_safe_sheet_name(v["vendor"], used_names))

        # drop the Paid column if not wanted
        prepared = []
        max_cols = 1
        for t in v["tables"]:
            cols = list(t["columns"])
            keep = [i for i, c in enumerate(cols) if include_paid or c != "Paid"]
            hdrs = [cols[i] for i in keep]
            rows = [[(r[i] if i < len(r) else "") for i in keep] for r in t["rows"]]
            prepared.append((t["title"], hdrs, rows))
            max_cols = max(max_cols, len(hdrs))

        # Row 1: vendor name
        tc = ws.cell(row=1, column=1, value=v["vendor"])
        tc.font = title_font
        tc.alignment = center
        if max_cols > 1:
            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=max_cols)
        r = 2

        # Optional date filter line
        parts = []
        if date_from:
            parts.append(f"From: {date_from}")
        if date_to:
            parts.append(f"To: {date_to}")
        if parts:
            ic = ws.cell(row=r, column=1, value="Date Filter: " + "  |  ".join(parts))
            ic.font = Font(italic=True, color="555555")
            if max_cols > 1:
                ws.merge_cells(start_row=r, start_column=1,
                               end_row=r, end_column=max_cols)
            r += 1
        r += 1  # blank line

        widths = {}

        for title, hdrs, rows in prepared:
            if len(prepared) > 1:
                ws.cell(row=r, column=1, value=title).font = title_font
                r += 1

            # header row
            for ci, h in enumerate(hdrs, start=1):
                c = ws.cell(row=r, column=ci, value=h)
                c.font = bold
                c.border = border
                c.alignment = center
                w = 14 if h == "Date" else 22 if h == "Item" else 16
                widths[ci] = max(widths.get(ci, 0), w)
            r += 1

            if not rows:
                ws.cell(row=r, column=1, value="No records.").border = border
                r += 1

            for row in rows:
                label = str(row[0]) if row and row[0] is not None else ""
                is_sum = label.startswith(("Total", "Currently Available"))
                for ci, h in enumerate(hdrs, start=1):
                    val = cell_value(h, row[ci - 1])
                    c = ws.cell(row=r, column=ci, value=val)
                    c.border = border
                    if is_sum:
                        c.font = bold
                    fmt = number_format(h)
                    if fmt and isinstance(val, (int, float)):
                        c.number_format = fmt
                r += 1

            r += 1  # gap between tables

        if v.get("summary"):
            ws.cell(row=r, column=1, value=v["summary"]).font = bold

        for ci, w in widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w

    # ---- save ----
    try:
        wb.save(path)
    except PermissionError:
        messagebox.showerror("Export Failed",
                             "The Excel file is currently open.\n\n"
                             "Please close the file and try again.")
        return
    except Exception as e:
        messagebox.showerror("Export Failed", f"Could not create Excel file:\n\n{e}")
        return

    try:
        log_action(user, "EXPORT EXCEL",
                   f"Exported Vendor Ledgers ({ledger_sub_type}) for "
                   + (f"vendor {vendor_filter}" if vendor_filter else "all vendors"))
    except Exception:
        pass

    messagebox.showinfo("Exported", f"Excel file saved successfully:\n\n{path}")
# ═══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    init_db()

    def _run_flask():
        flask_app.run(host="0.0.0.0", port=5000, debug=False, use_reloader=False)

    threading.Thread(target=_run_flask, daemon=True).start()

    App().mainloop()
