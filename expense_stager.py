
import os, re, csv, json, sqlite3, shutil, subprocess, webbrowser, sys, calendar, base64, zipfile, difflib
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font as tkfont

APP_TITLE = "Expense Stager"
APP_VERSION = "2026.08.11.2"  # date-based (YYYY.MM.DD; append .N for an Nth release same day). Shown in the title bar
# + footer, and mirrored by the repo-root VERSION_<APP_VERSION>.txt marker so GitHub shows it at a glance.
# Bump this AND rename the marker together on every release — dev/run_tests.py fails if they diverge.
DB_NAME = "expense_stager.sqlite3"
CONCUR_URL = "https://us2.concursolutions.com/"
# A record can hold BOTH a receipt and an invoice; status tracks where the item is in the staging lifecycle.
STATUSES = ["Draft", "Awaiting invoice", "Awaiting receipt", "Receipt received", "Ready to file", "Filed"]
# Old status spellings -> new (applied as a one-time data migration on existing DBs).
STATUS_MIGRATE = {"Expected": "Awaiting receipt", "Receipt Received": "Receipt received", "Ready": "Ready to file"}
PAYMENT_TYPES = ["Corporate Card", "Personal / reimbursable", "Other"]
CURRENCIES = ["USD"]
TRAVEL_TYPES = ["No Travel", "Domestic Travel", "International Travel"]
GRANT_TYPES = ["(GL) Non-Grant", "Grant"]
# Seed list of oracle/account codes (editable later in the Oracle Codes glossary).
ORACLE_SEED = [
    ("", "Unknown", ""),
]
# Expense codes bookmarked (favorited) on first run. Fill in the GL codes YOUR reports use most.
# Names/definitions come from your own expense_codes.csv (see README: export it from your org's Concur
# expense-type picker; it is your org's internal data and deliberately does not ship with this app).
# NB: a GL code can name more than one official expense type — favorites star every row sharing the code.
FAVORITE_CODES = []
CC_FEE_RATE = 0.03  # industry-standard credit-card surcharge used by the "Amount after CC fee" auto-calc
NO_REPORT = '(No report)'  # sentinel shown in the expense dialog's report dropdown = leave the expense unassigned
# Expense fields that make sense to bake into a reusable template. Per-instance fields (transaction date,
# attached file paths, invoice number, OCR text, missing-receipt ack) are intentionally excluded.
TEMPLATE_FIELDS = [
    ('vendor', 'Vendor'), ('amount', 'Amount'), ('amount_after_fee', 'Amount after CC fee'),
    ('currency', 'Currency'), ('status', 'Status'), ('expense_type_code', 'Expense type code'),
    ('expense_type_label', 'Expense type label'), ('business_purpose', 'Business purpose'),
    ('business_name', 'Business name'), ('city', 'City'), ('state', 'State'), ('country', 'Country'),
    ('payment_type', 'Payment type'), ('is_vendor_invoice', 'Is vendor invoice'),
    ('personal_no_reimburse', 'Personal / no reimburse'), ('comment', 'Comment'),
    ('loose_notes', 'Loose notes'), ('attendees', 'Attendees'),
]
TEMPLATE_BOOL_FIELDS = {'is_vendor_invoice', 'personal_no_reimburse'}
# Default templates seeded on first run. Vendor is always required in a template. Ships with common
# vendors but NO expense-type codes/labels — those are your org's internal data. Fill in each template's
# code + label (from your own expense_codes.csv) via the template editor or Save-as-Template.
TEMPLATE_SEED = [
    ('Uber ride', {'vendor': 'Uber', 'payment_type': PAYMENT_TYPES[0]}),
    ('Amazon order', {'vendor': 'Amazon', 'payment_type': PAYMENT_TYPES[0]}),
    ('ClickUp subscription', {'vendor': 'ClickUp', 'business_purpose': 'Productivity software', 'payment_type': PAYMENT_TYPES[0]}),
    ('Calendly subscription', {'vendor': 'Calendly', 'business_purpose': 'Scheduling software', 'payment_type': PAYMENT_TYPES[0]}),
    ('Doodle subscription', {'vendor': 'Doodle', 'business_purpose': 'Productivity software', 'payment_type': PAYMENT_TYPES[0]}),
    ('Staples office supplies', {'vendor': 'Staples', 'business_purpose': 'Office supplies', 'payment_type': PAYMENT_TYPES[0]}),
]
APP_DIR = Path(os.getenv("APPDATA", str(Path.home()))) / "ExpenseStager"
APP_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_DB_PATH = APP_DIR / DB_NAME
# The DB itself can be relocated (e.g. to a Google-Drive-synced folder). Its location is recorded in a tiny
# pointer file that ALWAYS lives in the fixed %APPDATA% dir -- the DB can't store its own location (chicken/egg).
DB_POINTER = APP_DIR / "db_location.txt"

def read_db_pointer():
    try:
        p = DB_POINTER.read_text(encoding="utf-8").strip()
        return Path(p) if p else None
    except Exception:
        return None
def write_db_pointer(path):
    try: DB_POINTER.write_text(str(path), encoding="utf-8")
    except Exception: pass
def resolve_db_path():
    """Where the DB lives now: the pointer target if its parent folder is reachable, else the default."""
    p = read_db_pointer()
    if p and p.parent.exists(): return p
    return DEFAULT_DB_PATH
def first_run_needed():
    """True when there is no recorded DB and no default DB yet -- a genuine first launch."""
    return read_db_pointer() is None and not DEFAULT_DB_PATH.exists()

# inbox_path / receipt_root are MACHINE-SPECIFIC folders (a Drive mount is J:\ on one PC, C:\Users\x\My Drive
# on another). They must NOT live in the shared/synced DB, or a second machine inherits paths it can't use
# (attach then tries to mkdir a foreign C:\Users\<other-user>\... and dies with Access denied). So they live in
# a tiny per-machine JSON next to the DB pointer, in the fixed local %APPDATA% dir.
LOCAL_SETTINGS = APP_DIR / "local_settings.json"
LOCAL_DEFAULTS = {'inbox_path': str(Path.home() / 'Downloads'),
                  'receipt_root': str(Path.home() / 'Documents' / 'Expense Receipts')}
def read_local_settings():
    try: return json.loads(LOCAL_SETTINGS.read_text(encoding="utf-8"))
    except Exception: return {}
def write_local_settings(d):
    try: LOCAL_SETTINGS.write_text(json.dumps(d, indent=2), encoding="utf-8")
    except Exception: pass
def get_local_setting(key, default=None):
    v = read_local_settings().get(key)
    return v if v not in (None, '') else (default if default is not None else LOCAL_DEFAULTS.get(key, ''))
def set_local_setting(key, value):
    d = read_local_settings(); d[key] = value; write_local_settings(d)
def ensure_local_paths():
    """Seed this machine's local inbox/receipt paths once. If the (legacy) synced DB carried a path that actually
    exists here, adopt it; otherwise fall back to a safe, creatable local default (never a foreign machine's path)."""
    d = read_local_settings(); changed = False
    for key, default in LOCAL_DEFAULTS.items():
        if key not in d:
            legacy = S.get_setting(key) if S is not None else None
            d[key] = legacy if (legacy and os.path.exists(legacy)) else default
            changed = True
    if changed: write_local_settings(d)

def q(s): return '' if s is None else str(s)
def today(): return date.today().isoformat()
def safe_filename(s):
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', q(s)).strip()
    s = re.sub(r'\s+', ' ', s)
    return s[:120] or 'Unknown Vendor'
def money_part(amount):
    try: return f"{float(clean_number(amount)):.2f}"
    except Exception: return "0.00"
def clean_number(s):
    """Coerce free-typed money text into a clean float string. Strips $, commas, spaces, stray chars;
    keeps digits, one dot, and a leading minus. '' / junk -> '0'. Used so numeric fields never store '1,405.00'."""
    s = q(s).strip()
    if not s: return '0'
    neg = s.lstrip().startswith('-')
    s = re.sub(r'[^0-9.]', '', s)            # drop $ , spaces letters etc.
    if s.count('.') > 1:                      # keep only the first dot
        head, _, tail = s.partition('.'); s = head + '.' + tail.replace('.', '')
    s = s.lstrip('0') or '0'                   # tidy leading zeros but keep one
    if s.startswith('.'): s = '0' + s
    if not re.search(r'\d', s): return '0'
    return ('-' + s) if neg else s
def _num_validate(proposed):
    """Tk validatecommand: allow only chars that can build a money value as the user types (digits . , $ -)."""
    return re.fullmatch(r'[-$ ]?[0-9,]*\.?[0-9]*', proposed or '') is not None
def open_path(path):
    if not path: return
    try:
        if sys.platform.startswith('win'): os.startfile(path)
        elif sys.platform == 'darwin': subprocess.Popen(['open', path])
        else: subprocess.Popen(['xdg-open', path])
    except Exception as e: messagebox.showerror("Open failed", str(e))
def open_folder(path, label='folder'):
    """Open a configured folder (inbox / receipt root). If it's unset or missing, guide the user instead
    of throwing a raw WinError — a Drive path that never synced is a config problem, not a crash."""
    if not path:
        messagebox.showinfo("Not set", f"No {label} is configured yet.\nSet it in Settings."); return
    p=Path(path)
    if not p.exists():
        if messagebox.askyesno("Folder missing", f"The {label} doesn't exist:\n{path}\n\nCreate it now?\n\n(If this is a Google Drive path, make sure Drive for desktop is running and pointing at the right location first — fix the path in Settings if not.)"):
            try: p.mkdir(parents=True, exist_ok=True)
            except Exception as e: messagebox.showerror("Couldn't create folder", str(e)); return
        else: return
    open_path(str(p))
def _path_parts(s):
    """Split a stored path into components regardless of separator or drive prefix."""
    return [c for c in re.split(r'[\\/]+', str(s)) if c and not c.endswith(':')]
def resolve_attachment(stored):
    """Re-anchor a stored attachment path onto THIS machine's receipt root.

    Attachment paths are saved absolute, so the prefix BEFORE the receipt root
    (drive letter, Windows username, Google-Drive mount point) differs per machine.
    But the folder tree UNDER the receipt root is synced identically across machines,
    so when the stored path doesn't resolve as-is we re-root its tail onto the current
    receipt root(s). Requires only that the receipt root points at the synced folder."""
    if not stored: return stored
    try:
        if os.path.exists(stored): return stored
    except Exception: return stored
    parts=_path_parts(stored)
    if not parts: return stored
    anchors=[]
    try:
        rr=S.get_setting('receipt_root')
        if rr: anchors.append(rr)
        for r in S.rows("SELECT receipt_root FROM users WHERE receipt_root IS NOT NULL AND receipt_root!=''"):
            if r[0]: anchors.append(r[0])
    except Exception: pass
    for anchor in anchors:
        for k in range(min(4,len(parts)),0,-1):  # try longest tail first (User/Year/file) down to just the filename
            try:
                cand=Path(anchor).joinpath(*parts[-k:])
                if cand.exists(): return str(cand)
            except Exception: pass
    return stored  # no anchor matched -> hand back the original so the caller's error is truthful
def open_file_location(path):
    """Open the containing folder for an attached file, selecting the file where the OS supports it."""
    if not path: messagebox.showinfo("No file", "No file is attached."); return
    path=resolve_attachment(path)
    p=Path(path)
    if not p.exists(): messagebox.showinfo("Not found", f"File no longer exists:\n{path}"); return
    try:
        if sys.platform.startswith('win'): subprocess.Popen(f'explorer /select,"{p}"')
        elif sys.platform == 'darwin': subprocess.Popen(['open', '-R', str(p)])
        else: open_path(str(p.parent))
    except Exception as e: messagebox.showerror("Open failed", str(e))
def work_area():
    """The usable desktop rectangle EXCLUDING the taskbar (Windows SPI_GETWORKAREA), as (x, y, w, h).
    winfo_screenheight() ignores the taskbar, which is why a near-full-height dialog tucks under it."""
    try:
        import ctypes
        class RECT(ctypes.Structure):
            _fields_=[("left",ctypes.c_long),("top",ctypes.c_long),("right",ctypes.c_long),("bottom",ctypes.c_long)]
        r=RECT(); ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(r), 0)  # 0x0030 = SPI_GETWORKAREA
        if r.right>r.left and r.bottom>r.top: return (r.left, r.top, r.right-r.left, r.bottom-r.top)
    except Exception: pass
    return None
def fit_to_screen(win, min_w=700, min_h=480, margin=24):
    """Size a dialog to its content (capped to the taskbar-free work area) and CENTER it on its parent window,
    clamped to stay fully on-screen. Uses withdraw -> idle -> geometry -> deiconify (Tk best practice) so the
    window never flashes in the wrong spot first. Replaces the old fixed top-left +30+10 placement."""
    try: win.withdraw()
    except Exception: pass
    win.update_idletasks()
    wa = work_area()
    ax, ay, aw, ah = wa if wa else (0, 0, win.winfo_screenwidth(), win.winfo_screenheight()-48)
    w = max(min(win.winfo_reqwidth(),  aw-margin), min(min_w, aw-margin))
    h = max(min(win.winfo_reqheight(), ah-margin), min(min_h, ah-margin))
    parent = win.master
    try:
        if parent and parent.winfo_ismapped() and parent.winfo_width() > 1:
            cx = parent.winfo_rootx() + parent.winfo_width()//2; cy = parent.winfo_rooty() + parent.winfo_height()//2
        else: cx, cy = ax + aw//2, ay + ah//2
    except Exception: cx, cy = ax + aw//2, ay + ah//2
    x = max(ax, min(int(cx - w/2), ax + aw - w))  # clamp fully inside the work area (never under taskbar / off-top)
    y = max(ay, min(int(cy - h/2), ay + ah - h))
    win.geometry(f"{w}x{h}+{x}+{y}"); win.minsize(min(w, min_w), min(h, min_h))
    # Esc closes the dialog (standard Windows behavior). Dialogs with an unsaved-changes guard expose
    # a .close() method; Esc routes through it so edits are never silently discarded.
    win.bind('<Escape>', lambda e: getattr(win, 'close', win.destroy)())
    try: win.deiconify()
    except Exception: pass

class FlowBar(ttk.Frame):
    """A row of buttons that WRAPS to additional rows when the window is too narrow, so buttons
    can never be clipped off the right edge at any width. Add buttons via .button(text, cmd),
    thin vertical group dividers via .separator(), and any pre-built widget via .attach()."""
    def __init__(self, master, pad=4, **kw):
        super().__init__(master, **kw); self.pad=pad; self.items=[]
        self.bind('<Configure>', self._reflow)
    def button(self, text, command):
        b=ttk.Button(self, text=text, command=command); self.items.append(b); self.after_idle(self._reflow); return b
    def separator(self):
        s=ttk.Separator(self, orient='vertical'); s._flow_w=9  # small fixed slot in the flow
        self.items.append(s); self.after_idle(self._reflow); return s
    def attach(self, widget):
        """Add an externally-created child widget (e.g. a Menubutton) into the wrap flow."""
        self.items.append(widget); self.after_idle(self._reflow); return widget
    def _reflow(self, e=None):
        width=self.winfo_width()
        if width<=1: width=self.winfo_reqwidth() or self.winfo_screenwidth()
        btnh=max([w.winfo_reqheight() for w in self.items if not isinstance(w, ttk.Separator)] or [26])
        x=y=rowh=0
        for w in self.items:
            sep=isinstance(w, ttk.Separator)
            ww=getattr(w, '_flow_w', None) or w.winfo_reqwidth(); wh=btnh if sep else w.winfo_reqheight()
            if x>0 and x+ww>width: x=0; y+=rowh+self.pad; rowh=0  # wrap to next row
            if sep and x==0: w.place_forget(); continue  # never lead a wrapped row with a divider
            if sep: w.place(x=x+ww//2, y=y+2, height=btnh-4)
            else: w.place(x=x, y=y)
            x+=ww+self.pad; rowh=max(rowh, wh)
        self.configure(height=y+rowh+self.pad+self._pad_v())
    def _pad_v(self):
        """The frame's OWN vertical ttk padding. `place` coordinates start after it, but -height sets the
        frame's total requested height — so leaving it out makes the bar request less room than its buttons
        occupy, and a dialog sized to its content clips the last few pixels of the button row."""
        p=self.cget('padding')
        if isinstance(p,str): p=p.split()
        try: p=[int(float(str(v))) for v in (p or ())]
        except Exception: return 0
        if not p: return 0
        if len(p)==1: return p[0]*2          # one value pads every side
        if len(p)>=4: return p[1]+p[3]       # left top right bottom
        return p[1]*2                        # 2 or 3 values: bottom mirrors top
class Tooltip:
    """Lightweight hover tooltip for any widget — used to explain non-obvious toolbar buttons in-app."""
    def __init__(self, widget, text):
        self.widget=widget; self.text=text; self.tip=None
        widget.bind('<Enter>', self._show, add='+'); widget.bind('<Leave>', self._hide, add='+')
    def _show(self, e=None):
        if self.tip or not self.text: return
        x=self.widget.winfo_rootx()+12; y=self.widget.winfo_rooty()+self.widget.winfo_height()+4
        self.tip=tk.Toplevel(self.widget); self.tip.wm_overrideredirect(True); self.tip.wm_geometry(f'+{x}+{y}')
        tk.Label(self.tip, text=self.text, justify='left', background='#ffffe0', relief='solid', borderwidth=1,
                 wraplength=320, padx=6, pady=3).pack()
    def _hide(self, e=None):
        if self.tip: self.tip.destroy(); self.tip=None
CHECKBOX_PX = 16  # tk Treeview has no checkbox column, so every row carries a small check-box IMAGE in the
# tree column (left-aligned, like Concur's own "Available Expenses" list). The images are drawn in code —
# no icon files to ship, and nothing to go missing in a frozen build.
def make_checkbox_images(master):
    """{False: empty box, True: ticked box} as PhotoImages. A fresh PhotoImage starts fully transparent
    and only the pixels we `put` become opaque, so the box sits cleanly on a selected (highlighted) row."""
    border, body, ticked, tick = '#5a5a5a', '#ffffff', '#1668b8', '#ffffff'
    imgs = {}
    for checked in (False, True):
        img = tk.PhotoImage(master=master, width=CHECKBOX_PX, height=CHECKBOX_PX)
        img.put(ticked if checked else body, to=(2, 2, 14, 14))  # box body
        for edge in [(2, 2, 14, 3), (2, 13, 14, 14), (2, 2, 3, 14), (13, 2, 14, 14)]:
            img.put(ticked if checked else border, to=edge)      # 1px outline
        if checked:
            for x, y in [(3, 7), (4, 8), (5, 9), (6, 8), (7, 7), (8, 6), (9, 5), (10, 4)]:
                img.put(tick, to=(x, y, x + 2, y + 2))           # the tick, 2px thick
        imgs[checked] = img
    return imgs
def is_within(child, parent):
    try: return parent and Path(child).resolve().parent == Path(parent).resolve() or str(Path(child).resolve()).startswith(str(Path(parent).resolve()) + os.sep)
    except Exception: return False

class Store:
    def __init__(self, path=None):
        self.path = Path(path) if path else resolve_db_path()
        self._connect()
    def _connect(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.con = sqlite3.connect(self.path); self.con.row_factory = sqlite3.Row; self.init()
    def relocate(self, new_dir, delete_old=False):
        """Move the live DB to new_dir (copying its data), repoint, and reconnect. Optionally remove the old file."""
        new_path = Path(new_dir) / DB_NAME; old = self.path
        try: self.con.commit(); self.con.close()
        except Exception: pass
        new_path.parent.mkdir(parents=True, exist_ok=True)
        if old.resolve() != new_path.resolve(): shutil.copy2(old, new_path)
        self.path = new_path; self._connect(); write_db_pointer(str(new_path))
        if delete_old and old.resolve() != new_path.resolve():
            try: old.unlink()
            except Exception: pass
        return new_path
    def use_existing(self, db_file):
        """Point at an already-existing DB elsewhere (e.g. a Drive-synced file on another device) and reconnect."""
        try: self.con.close()
        except Exception: pass
        self.path = Path(db_file); self._connect(); write_db_pointer(str(self.path))
        import_codes_if_present(); self.apply_favorite_bookmarks()
    def init(self):
        c=self.con.cursor()
        c.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)")
        c.execute("CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL)")
        # Codes are NOT unique in Concur (one GL code can name several official expense types),
        # and the rule is that a code needs MULTIPLE identities: the official listing
        # AND the colloquial name accounting actually uses (searchable via `tags`, explained in `notes`).
        # So the key is (code, name), not code alone.
        c.execute("CREATE TABLE IF NOT EXISTS expense_codes(id INTEGER PRIMARY KEY, code TEXT, name TEXT, tags TEXT DEFAULT '', favorite INTEGER DEFAULT 0, notes TEXT DEFAULT '', UNIQUE(code,name))")
        c.execute("CREATE TABLE IF NOT EXISTS oracle_codes(id INTEGER PRIMARY KEY, code TEXT DEFAULT '', name TEXT, notes TEXT DEFAULT '', favorite INTEGER DEFAULT 0, UNIQUE(code,name))")
        c.execute("CREATE TABLE IF NOT EXISTS vendors(name TEXT PRIMARY KEY, favorite INTEGER DEFAULT 0, notes TEXT DEFAULT '', last_code TEXT DEFAULT '', last_label TEXT DEFAULT '')")
        c.execute("""CREATE TABLE IF NOT EXISTS reports(
            id INTEGER PRIMARY KEY, user_id INTEGER, name TEXT NOT NULL, report_date TEXT, start_date TEXT, end_date TEXT,
            travel_purpose TEXT, travel_type TEXT DEFAULT 'No Travel', expense_group_id TEXT DEFAULT '',
            report_id TEXT, currency TEXT DEFAULT 'US, Dollar', approval_status TEXT DEFAULT 'Not Submitted', payment_status TEXT DEFAULT 'Not Paid',
            grant_type TEXT DEFAULT '(GL) Non-Grant', oracle_alias TEXT, expense_report_for TEXT, business_unit TEXT, comment TEXT,
            status TEXT DEFAULT 'Draft', created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        c.execute("""CREATE TABLE IF NOT EXISTS expenses(
            id INTEGER PRIMARY KEY, user_id INTEGER NOT NULL, report_id INTEGER, transaction_date TEXT NOT NULL, vendor TEXT NOT NULL,
            amount REAL DEFAULT 0, amount_after_fee REAL DEFAULT 0, currency TEXT DEFAULT 'USD', status TEXT DEFAULT 'Draft', doc_type TEXT DEFAULT 'Receipt',
            expense_type_code TEXT, expense_type_label TEXT, business_purpose TEXT, business_name TEXT, city TEXT, state TEXT, country TEXT DEFAULT 'US',
            payment_type TEXT DEFAULT 'Corporate Card', is_vendor_invoice INTEGER DEFAULT 0,
            personal_no_reimburse INTEGER DEFAULT 0, missing_receipt_ack INTEGER DEFAULT 0, comment TEXT, loose_notes TEXT,
            attendees TEXT, receipt_path TEXT, invoice_path TEXT, invoice_number TEXT DEFAULT '', ocr_text TEXT, filed_at TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        # A template is a named, reusable set of expense field values (stored as a JSON dict in `fields`); vendor is required.
        c.execute("""CREATE TABLE IF NOT EXISTS templates(
            id INTEGER PRIMARY KEY, name TEXT UNIQUE NOT NULL, vendor TEXT NOT NULL DEFAULT '',
            fields TEXT DEFAULT '{}', is_default INTEGER DEFAULT 0, created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        self.con.commit()
        # --- migrations for DBs created before these columns/statuses existed ---
        # expense_codes: rebuild old code-PK tables into the (code,name)-unique shape, keeping
        # favorites/tags/notes. SQLite can't alter a PK, so rename -> recreate -> copy -> drop.
        ccols=[r['name'] for r in self.rows("PRAGMA table_info(expense_codes)")]
        if 'id' not in ccols:
            self.execute("ALTER TABLE expense_codes RENAME TO expense_codes_old")
            self.execute("CREATE TABLE expense_codes(id INTEGER PRIMARY KEY, code TEXT, name TEXT, tags TEXT DEFAULT '', favorite INTEGER DEFAULT 0, notes TEXT DEFAULT '', UNIQUE(code,name))")
            self.execute("INSERT OR IGNORE INTO expense_codes(code,name,tags,favorite,notes) SELECT code,name,COALESCE(tags,''),COALESCE(favorite,0),COALESCE(notes,'') FROM expense_codes_old")
            self.execute("DROP TABLE expense_codes_old")
        # Drop old code-as-name placeholder rows once a real named row exists for the same code.
        self.execute("DELETE FROM expense_codes WHERE name=code AND EXISTS(SELECT 1 FROM expense_codes e2 WHERE e2.code=expense_codes.code AND e2.name<>e2.code)")
        cols=[r['name'] for r in self.rows("PRAGMA table_info(expenses)")]
        if 'invoice_number' not in cols: self.execute("ALTER TABLE expenses ADD COLUMN invoice_number TEXT DEFAULT ''")
        if 'amount_after_fee' not in cols: self.execute("ALTER TABLE expenses ADD COLUMN amount_after_fee REAL DEFAULT 0")
        ucols=[r['name'] for r in self.rows("PRAGMA table_info(users)")]
        if 'receipt_root' not in ucols: self.execute("ALTER TABLE users ADD COLUMN receipt_root TEXT DEFAULT ''")
        # Clean any money values stored as text (e.g. '1,405.00', '$59.69') from older builds into real numbers.
        for col in ('amount','amount_after_fee'):
            for r in self.rows(f"SELECT id,{col} v FROM expenses WHERE typeof({col})='text'"):
                self.execute(f"UPDATE expenses SET {col}=? WHERE id=?", (float(clean_number(r['v'])), r['id']))
        for old,new in STATUS_MIGRATE.items(): self.execute("UPDATE expenses SET status=? WHERE status=?", (new, old))
        if self.scalar('SELECT COUNT(*) FROM users') == 0:
            self.execute('INSERT INTO users(name) VALUES (?)', ('Me',))
        # Seed the default templates only when the table is empty, so a user-deleted default stays deleted.
        if self.scalar('SELECT COUNT(*) FROM templates') == 0:
            for name, fields in TEMPLATE_SEED:
                self.execute('INSERT OR IGNORE INTO templates(name,vendor,fields,is_default) VALUES(?,?,?,1)', (name, fields.get('vendor',''), json.dumps(fields)))
        # One-time top-up for DBs seeded by an older build that only had the single "Uber ride" default:
        # add the newer common-vendor templates once (guarded by a flag so a user who deletes one keeps it gone).
        elif self.get_setting('seed_templates_v2') is None:
            for name, fields in TEMPLATE_SEED:
                if not self.row('SELECT id FROM templates WHERE name=?', (name,)):
                    self.execute('INSERT INTO templates(name,vendor,fields,is_default) VALUES(?,?,?,1)', (name, fields.get('vendor',''), json.dumps(fields)))
        if self.get_setting('seed_templates_v2') is None:
            self.set_setting('seed_templates_v2', '1')
        for k,v in [('inbox_path', str(Path.home()/ 'Downloads')), ('receipt_root', str(Path.home()/ 'Documents' / 'Expense Receipts'))]:
            if self.get_setting(k) is None: self.set_setting(k,v)
        for code,name,notes in ORACLE_SEED:
            self.execute('INSERT OR IGNORE INTO oracle_codes(code,name,notes) VALUES(?,?,?)', (code,name,notes))
        pass  # legacy private-build data migrations removed in the public build
    def apply_favorite_bookmarks(self):
        """Flag the report-derived favorite codes. Run AFTER the CSV import so names come from the master list.
        A code with multiple official names gets ALL its rows starred (they share the GL code the reports used)."""
        # Retire any old code-as-name placeholder once the CSV import has supplied real named rows.
        self.execute("DELETE FROM expense_codes WHERE name=code AND EXISTS(SELECT 1 FROM expense_codes e2 WHERE e2.code=expense_codes.code AND e2.name<>e2.code)")
        for code in FAVORITE_CODES:
            if not self.row('SELECT 1 FROM expense_codes WHERE code=?', (code,)):  # placeholder only when truly unknown
                self.execute('INSERT INTO expense_codes(code,name) VALUES(?,?)', (code, code))
            self.execute('UPDATE expense_codes SET favorite=1 WHERE code=?', (code,))
    def execute(self, sql, params=()): cur=self.con.execute(sql, params); self.con.commit(); return cur
    def rows(self, sql, params=()): return self.con.execute(sql, params).fetchall()
    def row(self, sql, params=()): return self.con.execute(sql, params).fetchone()
    def scalar(self, sql, params=()):
        r=self.con.execute(sql, params).fetchone(); return None if r is None else r[0]
    def get_setting(self,k): return self.scalar('SELECT value FROM settings WHERE key=?',(k,))
    def set_setting(self,k,v): self.execute('INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)',(k,v))
    def upsert_vendor(self, name, code='', label=''):
        name=q(name).strip()
        if not name: return
        if self.row('SELECT 1 FROM vendors WHERE name=?', (name,)):
            if code: self.execute('UPDATE vendors SET last_code=?, last_label=? WHERE name=?', (code,label,name))
        else:
            self.execute('INSERT INTO vendors(name,last_code,last_label) VALUES(?,?,?)', (name,code,label))
    def add_oracle(self, text):
        text=q(text).strip()
        if not text: return
        code = text if text.replace('-','').isdigit() else ''
        name = '' if code else text
        # If the entered value already matches a known code or name, do nothing.
        if self.row('SELECT 1 FROM oracle_codes WHERE code=? OR name=?', (text, text)): return
        try: self.execute('INSERT OR IGNORE INTO oracle_codes(code,name) VALUES(?,?)', (code, name or text))
        except Exception: pass
    def get_templates(self): return self.rows('SELECT * FROM templates ORDER BY is_default DESC, name')
    def template_fields(self, name):
        """The saved field dict for a template name ({} if missing/corrupt/non-dict — e.g. a hand-edited
        config with "fields": null, which json.loads decodes to None without raising)."""
        r=self.row('SELECT fields FROM templates WHERE name=?', (name,))
        try: d=json.loads(r['fields']) if r and r['fields'] else {}
        except Exception: d={}
        return d if isinstance(d, dict) else {}
    def upsert_template(self, name, vendor, fields, is_default=None):
        """Create or replace a template BY NAME (used by Save-as-Template and config import). On a name
        collision this UPDATEs in place, so the existing row's id and is_default survive an overwrite
        (an INSERT OR REPLACE would silently drop a seeded default's star). is_default is preserved unless
        passed explicitly. `fields` is normalized to a dict so a null/non-dict import value can't poison it."""
        name=q(name).strip(); vendor=q(vendor).strip()
        if not name or not vendor: return False
        if not isinstance(fields, (str, dict)): fields={}
        body=fields if isinstance(fields, str) else json.dumps(fields)
        existing=self.row('SELECT id,is_default FROM templates WHERE name=?', (name,))
        if existing:
            flag=existing['is_default'] if is_default is None else is_default
            self.execute('UPDATE templates SET vendor=?,fields=?,is_default=? WHERE id=?', (vendor, body, flag, existing['id']))
        else:
            self.execute('INSERT INTO templates(name,vendor,fields,is_default) VALUES(?,?,?,?)', (name, vendor, body, is_default or 0))
        return True
    def update_template(self, tpl_id, name, vendor, fields):
        """Update an existing template BY ID (lets the editor rename without losing the row)."""
        self.execute('UPDATE templates SET name=?,vendor=?,fields=? WHERE id=?', (q(name).strip(), q(vendor).strip(), json.dumps(fields), tpl_id))
_FIRST_RUN = first_run_needed()  # capture BEFORE the Store creates the default DB file
_SUPPRESS_PROMPTS = bool(os.getenv('CONCUR_BUDDY_SUPPRESS_PROMPTS'))  # set by selfcheck/headless tests to skip modal dialogs
S = None  # the live Store; created by open_store() (after the optional first-run location prompt)

def import_codes_if_present():
    for p in [Path('expense_codes.csv'), Path(__file__).with_name('expense_codes.csv')]:
        if p.exists():
            with p.open(newline='', encoding='utf-8-sig') as f:
                for r in csv.DictReader(f):
                    code,name,tags=r.get('code',''),r.get('name',''),r.get('tags','')
                    S.execute('INSERT OR IGNORE INTO expense_codes(code,name,tags) VALUES(?,?,?)',(code,name,tags))
                    # Backfill the CSV's category tag onto pre-existing rows that have no tags yet —
                    # user-entered tags (colloquial aliases) are never overwritten.
                    if tags: S.execute("UPDATE expense_codes SET tags=? WHERE code=? AND name=? AND (tags IS NULL OR tags='')",(tags,code,name))
            return
def open_store():
    global S
    S = Store()
    import_codes_if_present()
    S.apply_favorite_bookmarks()  # bookmark the report-derived codes now that the master list is loaded
    ensure_local_paths()  # migrate/seed this machine's own inbox + receipt-root paths (never from a foreign machine)
    return S
open_store()  # default startup; the GUI entrypoint re-runs the first-run prompt before this when launched directly

def duplicate_expense(exp_id):
    """Duplicate an expense as a fresh Draft: drop the identity/lifecycle + attachment/OCR fields, keep the
    reusable data (vendor, amount, code, purpose, report grouping). Returns the new expense id."""
    src=dict(S.row('SELECT * FROM expenses WHERE id=?',(exp_id,)))
    for k in ('id','created_at','filed_at','receipt_path','invoice_path','ocr_text','missing_receipt_ack'): src.pop(k, None)
    src['status']='Draft'
    cols=list(src)
    return S.execute('INSERT INTO expenses('+','.join(cols)+') VALUES('+','.join(['?']*len(cols))+')', tuple(src[c] for c in cols)).lastrowid

# ---------------------------------------------------------------- Concur export (.xlsx) import
# The reverse of the filing flow: read a report exported OUT of Concur and reconcile it against what is
# already staged here. An .xlsx is just a zip of XML, so this is parsed with the standard library alone —
# the app keeps its "no installs" promise (openpyxl/pandas would break it).
CONCUR_COLUMNS = {  # normalised export header -> expense field ('_' names are handled specially below)
    'date': 'transaction_date', 'transaction date': 'transaction_date', 'expense date': 'transaction_date',
    'vendor details': 'vendor', 'vendor': 'vendor', 'vendor name': 'vendor', 'merchant': 'vendor', 'merchant name': 'vendor',
    'requested': 'amount', 'amount': 'amount', 'transaction amount': 'amount', 'approved': 'amount', 'expense amount': 'amount',
    'payment type': 'payment_type', 'currency': 'currency', 'expense type': '_expense_type', 'expense type name': '_expense_type',
    'business purpose': 'business_purpose', 'purpose': 'business_purpose', 'business name': 'business_name',
    'city': 'city', 'city of purchase': 'city', 'location': 'city', 'state': 'state', 'country': 'country',
    'comment': 'comment', 'comments': 'comment', 'attendees': 'attendees',
    'is this a vendor invoice': 'is_vendor_invoice', 'vendor invoice': 'is_vendor_invoice',
    'personal expense (do not reimburse)': 'personal_no_reimburse', 'personal expense': 'personal_no_reimburse',
    'receipt': '_receipt_in_concur', 'receipt status': '_receipt_in_concur', 'report name': '_report_name',
}
# Concur greys these out on the expense form because they arrive from the card feed and nothing in Concur
# (or here) can edit them. The export is therefore the authoritative copy: a merge always takes them from
# the file. They are shown before you apply, but never asked about — there is nothing to decide.
CARD_FIELDS = ('transaction_date', 'vendor', 'amount', 'payment_type', 'currency')
# Everything else in an export IS editable in Concur, so a filled-in Concur Buddy value that disagrees is a
# real conflict and the user picks a side. Empty here = just take the export's value.
MERGE_FIELDS = ('business_purpose', 'business_name', 'city', 'state', 'country', 'comment', 'attendees',
                'is_vendor_invoice', 'personal_no_reimburse')
# The expense type's code and its name are two halves of ONE fact, so they are decided TOGETHER under the
# pseudo-field '_expense_type'. Deciding them apart would let "keep my code" + "take their name" mint a
# code/name pair that exists in neither system (one code carrying another type's official name).
TYPE_PAIR = ('expense_type_code', 'expense_type_label')
FIELD_LABELS = {'_expense_type': 'Expense type', 'transaction_date': 'Transaction date', 'vendor': 'Vendor name', 'amount': 'Amount',
                'payment_type': 'Payment type', 'currency': 'Currency', 'expense_type_code': 'Expense type code',
                'expense_type_label': 'Expense type', 'business_purpose': 'Business purpose', 'business_name': 'Business name',
                'city': 'City', 'state': 'State', 'country': 'Country', 'comment': 'Comment', 'attendees': 'Attendees',
                'is_vendor_invoice': 'Is a vendor invoice', 'personal_no_reimburse': 'Personal / no reimburse'}
MATCH_DAY_WINDOW = 31  # how far apart two dates can be before an amount match needs the name to back it up
MATCH_NAME_FLOOR = 0.5  # …or how alike the two vendor strings must look instead
def _col_index(ref):
    n = 0
    for ch in ref: n = n * 26 + (ord(ch) - 64)
    return n - 1
def _col_name(i):
    s = ''; i += 1
    while i: i, r = divmod(i - 1, 26); s = chr(65 + r) + s
    return s
def read_xlsx(path):
    """Minimal .xlsx reader: first worksheet -> list of rows, each a list of cell strings. Stdlib only."""
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')).findall(f'{ns}si'):
                shared.append(''.join(t.text or '' for t in si.iter(f'{ns}t')))
        sheets = sorted(n for n in z.namelist() if n.startswith('xl/worksheets/') and n.endswith('.xml'))
        if not sheets: raise ValueError('no worksheet found in this .xlsx')
        sheet = ET.fromstring(z.read(sheets[0]))
    rows = []
    for row in sheet.iter(f'{ns}row'):
        cells = {}
        for c in row.findall(f'{ns}c'):
            col = re.match(r'[A-Z]+', c.get('r') or 'A').group(0)
            t = c.get('t'); inline = c.find(f'{ns}is'); v = c.find(f'{ns}v')
            if inline is not None: val = ''.join(x.text or '' for x in inline.iter(f'{ns}t'))
            elif v is None or v.text is None: val = ''
            elif t == 's': val = shared[int(v.text)] if int(v.text) < len(shared) else ''
            else: val = v.text
            cells[col] = val.strip()
        width = max((_col_index(k) for k in cells), default=-1) + 1
        rows.append([cells.get(_col_name(i), '') for i in range(width)])
    return rows
def parse_export_money(s):
    """'$1,405.00' -> 1405.0 ; '($23.00)' -> -23.0 (Concur brackets credits/refunds rather than signing them)."""
    s = q(s).strip()
    v = float(clean_number(s) or 0)
    return -abs(v) if (s.startswith('(') and s.endswith(')')) else v
def parse_export_date(s):
    """Concur exports MM/DD/YYYY text; a sheet edited in Excel can hand back a serial number instead."""
    s = q(s).strip()
    if re.fullmatch(r'\d{1,6}(\.\d+)?', s):
        try: return (date(1899, 12, 30) + timedelta(days=int(float(s)))).isoformat()
        except Exception: return ''
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%d %b %Y', '%b %d, %Y'):
        try: return datetime.strptime(s, fmt).date().isoformat()
        except ValueError: pass
    return s
def _yes(s): return q(s).strip().lower()[:1] in ('y', 't', '1', 'x')
def parse_concur_export(path):
    """A Concur report export -> (rows, report_name, unknown_headers).

    Each row is a dict of expense fields plus `_receipt_in_concur` (Concur's own Yes/No receipt column)
    and `_extra` — any column this app has no field for, kept so an import never silently drops data."""
    raw = read_xlsx(path)
    head = head_raw = None; rows = []; unknown = []
    for cells in raw:
        cells = [q(c).strip() for c in cells]
        if not any(cells): continue
        if head is None:  # first non-empty row is the header
            head_raw = cells
            head = [re.sub(r'\s+', ' ', c.lower()).strip(' *') for c in cells]
            unknown = [r for r, h in zip(head_raw, head) if h and h not in CONCUR_COLUMNS]
            continue
        row = {'_extra': {}, '_receipt_in_concur': None}
        for i, val in enumerate(cells):
            h = head[i] if i < len(head) else ''
            if not h or not val: continue
            f = CONCUR_COLUMNS.get(h)
            if f is None: row['_extra'][head_raw[i]] = val
            elif f == 'transaction_date': row[f] = parse_export_date(val)
            elif f == 'amount': row[f] = parse_export_money(val)
            elif f == '_receipt_in_concur': row[f] = _yes(val)
            elif f in ('is_vendor_invoice', 'personal_no_reimburse'): row[f] = 1 if _yes(val) else 0
            elif f == '_expense_type':
                # "53103-US LOCAL TRANSPORTATION" -> ('53103', 'US LOCAL TRANSPORTATION'). Split on the FIRST
                # hyphen only, and never inside the code: type names carry hyphens of their own.
                m = re.match(r'\s*([0-9][0-9.]*)\s*[-–]\s*(.+)$', val)
                row['expense_type_code'], row['expense_type_label'] = (m.group(1), m.group(2).strip()) if m else ('', val)
            else: row[f] = val
        # Total/footer rows carry neither a vendor nor a date — they are not expenses.
        if not row.get('vendor') and not row.get('transaction_date'): continue
        rows.append(row)
    name = q(rows[0].get('_report_name')) if rows and rows[0].get('_report_name') else Path(path).stem
    return rows, re.sub(r'\s+', ' ', name.replace('_', ' ')).strip(), unknown
def vendor_similarity(a, b):
    """0..1 likeness of two vendor strings, ignoring case/punctuation. The card feed's descriptor rarely
    matches what you typed ('COLPARK LOC 958' vs 'Colonial parking'), so this only RANKS candidates."""
    norm = lambda s: re.sub(r'[^a-z0-9 ]', ' ', q(s).lower()).strip()
    return difflib.SequenceMatcher(None, norm(a), norm(b)).ratio()
def date_gap(a, b):
    try: return abs((date.fromisoformat(q(a)[:10]) - date.fromisoformat(q(b)[:10])).days)
    except Exception: return 999
def match_concur_row(row, expenses, taken=()):
    """Rank staged expenses as candidates for one imported row, best first.

    AMOUNT is the anchor: at this volume a to-the-cent match is close to unique, so only same-amount rows
    are candidates at all. Date distance and vendor similarity merely ORDER them — they never promote a
    row whose amount differs. Returns [(expense, days_apart, vendor_similarity), ...]."""
    cents = round(float(row.get('amount') or 0) * 100)
    out = [(e, date_gap(row.get('transaction_date'), e['transaction_date']), vendor_similarity(row.get('vendor'), e['vendor']))
           for e in expenses if e['id'] not in taken and round(float(e['amount'] or 0) * 100) == cents]
    out.sort(key=lambda t: (t[1], -t[2]))
    return out
def merge_plan(row, exp):
    """What merging `row` into staged expense `exp` would do, as (card, conflicts, fills) — each a list of
    (field, staged_value, export_value):
      card      — card-fed fields the export overrules outright (no decision to make);
      conflicts — editable fields where BOTH sides have a value and they differ (the user picks a side);
      fills     — editable fields empty here that the export can fill in for free."""
    card, conflicts, fills = [], [], []
    for f in CARD_FIELDS:
        if f not in row or row[f] in ('', None): continue
        old, new = exp[f] if f in exp.keys() else '', row[f]
        if f == 'amount':
            if round(float(old or 0) * 100) == round(float(new) * 100): continue
            old, new = money_part(old), money_part(new)
        elif q(old).strip() == q(new).strip(): continue
        card.append((f, q(old), q(new)))
    for f in MERGE_FIELDS:
        if f not in row or row[f] in ('', None): continue
        old, new = exp[f] if f in exp.keys() else '', row[f]
        if q(old).strip() == q(new).strip(): continue
        (fills if q(old).strip() == '' else conflicts).append((f, q(old), q(new)))
    pair = lambda d: ' '.join(x for x in (q(d[TYPE_PAIR[0]] if TYPE_PAIR[0] in (d.keys() if hasattr(d,'keys') else d) else '').strip(),
                                          q(d[TYPE_PAIR[1]] if TYPE_PAIR[1] in (d.keys() if hasattr(d,'keys') else d) else '').strip()) if x)
    old_t, new_t = pair(exp), pair(row)
    if new_t and old_t != new_t: (fills if not old_t else conflicts).append(('_expense_type', old_t, new_t))
    return card, conflicts, fills
def apply_field(vals, row, field):
    """Copy one decided field out of an export row into a column->value dict, expanding the expense-type
    pseudo-field back into its two real columns."""
    if field == '_expense_type':
        for f in TYPE_PAIR: vals[f] = row.get(f, '')
    else: vals[field] = row[field]
    return vals

class CodePicker(ttk.Frame):
    """Autocomplete entry for expense codes. Filters as you type a number or name; with the field empty
    (e.g. when you click into it) it lists your FAVORITE codes so they're reachable straight from the dialog.
    Favorites are marked with a leading star. A '★ Favorites' button re-opens the favorites list any time."""
    def __init__(self, master, code_var, label_var):
        super().__init__(master); self.code_var=code_var; self.label_var=label_var
        top=ttk.Frame(self); top.pack(fill='x')
        self.entry=ttk.Entry(top, textvariable=code_var); self.entry.pack(side='left', fill='x', expand=True)
        ttk.Button(top, text='★ Favorites', width=11, command=self.show_favorites).pack(side='left')
        self.lb=tk.Listbox(self, height=6); self.lb.pack(fill='x'); self.lb.pack_forget()
        self.entry.bind('<KeyRelease>', self.search); self.entry.bind('<FocusIn>', self.search); self.lb.bind('<<ListboxSelect>>', self.pick)
    def show_favorites(self):
        self.code_var.set(''); self.entry.focus_set(); self.search()
    def search(self, e=None):
        term=self.code_var.get().strip(); self.lb.delete(0,'end')
        if term:
            rows=S.rows("SELECT code,name,favorite FROM expense_codes WHERE code LIKE ? OR name LIKE ? OR tags LIKE ? ORDER BY favorite DESC, code LIMIT 12", (term+'%', '%'+term+'%', '%'+term+'%'))
        else:  # empty -> surface favorites (the whole point: pick a usual code without remembering its number)
            rows=S.rows("SELECT code,name,favorite FROM expense_codes WHERE favorite=1 ORDER BY code")
        for r in rows: self.lb.insert('end', f"{'★ ' if r['favorite'] else ''}{r['code']} — {r['name']}")
        self.lb.pack(fill='x') if rows else self.lb.pack_forget()
    def pick(self, e=None):
        sel=self.lb.curselection()
        if not sel: return
        # Parse code AND name from the picked line — a code can name several official types,
        # so a code-only lookup could land on the wrong row.
        line=self.lb.get(sel[0]).lstrip('★ '); code,_,name=line.partition(' — ')
        r=S.row('SELECT * FROM expense_codes WHERE code=? AND name=?',(code,name)) or S.row('SELECT * FROM expense_codes WHERE code=?',(code,))
        if r: self.code_var.set(r['code']); self.label_var.set(r['name'])
        self.lb.pack_forget()

class OraclePicker(ttk.Frame):
    """Autocomplete entry for oracle/account codes. New values are saved to grow the list on report save."""
    def __init__(self, master, value_var):
        super().__init__(master); self.value_var=value_var
        self.entry=ttk.Entry(self, textvariable=value_var); self.entry.pack(fill='x', expand=True)
        self.lb=tk.Listbox(self, height=5); self.lb.pack(fill='x'); self.lb.pack_forget()
        self.entry.bind('<KeyRelease>', self.search); self.lb.bind('<<ListboxSelect>>', self.pick)
    def search(self, e=None):
        term=self.value_var.get().strip(); self.lb.delete(0,'end')
        rows=S.rows("SELECT code,name FROM oracle_codes WHERE code LIKE ? OR name LIKE ? ORDER BY favorite DESC, name LIMIT 12", (term+'%' if term else '%','%'+term+'%'))
        for r in rows: self.lb.insert('end', f"{r['code']} — {r['name']}" if r['code'] else r['name'])
        self.lb.pack(fill='x') if rows else self.lb.pack_forget()
    def pick(self, e=None):
        sel=self.lb.curselection()
        if not sel: return
        val=self.lb.get(sel[0]); code=val.split(' — ')[0] if ' — ' in val else ''
        self.value_var.set(code or val); self.lb.pack_forget()

class CalendarPopup(tk.Toplevel):
    """Tiny dependency-free month calendar. Clicking a day writes YYYY-MM-DD into the bound StringVar."""
    def __init__(self, master, date_var):
        super().__init__(master); self.date_var=date_var; self.title('Pick a date'); self.transient(master); self.grab_set()
        try: d=datetime.strptime(date_var.get().strip(), '%Y-%m-%d').date()
        except Exception: d=date.today()
        self.year, self.month = d.year, d.month
        self.head=ttk.Frame(self,padding=6); self.head.pack(fill='x')
        ttk.Button(self.head,text='◀',width=3,command=lambda:self.shift(-1)).pack(side='left')
        self.title_var=tk.StringVar(); ttk.Label(self.head,textvariable=self.title_var,anchor='center',width=18).pack(side='left',expand=True)
        ttk.Button(self.head,text='▶',width=3,command=lambda:self.shift(1)).pack(side='left')
        self.grid_frame=ttk.Frame(self,padding=(6,0,6,6)); self.grid_frame.pack()
        bottom=ttk.Frame(self,padding=4); bottom.pack(fill='x')
        ttk.Button(bottom,text='Today',command=self.set_today).pack(side='left')
        ttk.Button(bottom,text='Close',command=self.destroy).pack(side='right')
        self.bind('<Escape>', lambda e: self.destroy()); self.draw()
    def shift(self, n):
        m=self.month+n; self.year+=(m-1)//12; self.month=(m-1)%12+1; self.draw()
    def set_today(self):
        self._choose(date.today())
    def draw(self):
        for w in self.grid_frame.winfo_children(): w.destroy()
        self.title_var.set(f"{calendar.month_name[self.month]} {self.year}")
        for c,wd in enumerate(['Mo','Tu','We','Th','Fr','Sa','Su']):
            ttk.Label(self.grid_frame,text=wd,width=4,anchor='center').grid(row=0,column=c,padx=1,pady=1)
        for r,week in enumerate(calendar.Calendar().monthdayscalendar(self.year,self.month), start=1):
            for c,day in enumerate(week):
                if day==0: continue
                ttk.Button(self.grid_frame,text=str(day),width=4,
                           command=lambda dd=day:self._choose(date(self.year,self.month,dd))).grid(row=r,column=c,padx=1,pady=1)
    def _choose(self, d):
        self.date_var.set(d.isoformat()); self.destroy()

class ExpenseDialog(tk.Toplevel):
    def __init__(self, master, exp_id=None, default_user_id=None):
        super().__init__(master); self.exp_id=exp_id; self.title('Expense'); self.transient(master); self.grab_set()
        self.vars={}; self.user_map={r['name']:r['id'] for r in S.rows('SELECT * FROM users ORDER BY name')}
        data=dict(S.row('SELECT * FROM expenses WHERE id=?',(exp_id,))) if exp_id else {}
        def v(name, default=''): self.vars[name]=tk.StringVar(value=q(data.get(name, default))); return self.vars[name]
        def b(name): self.vars[name]=tk.IntVar(value=int(data.get(name,0) or 0)); return self.vars[name]
        user_name=next((n for n,i in self.user_map.items() if i==data.get('user_id')), None) or next((n for n,i in self.user_map.items() if i==default_user_id), None) or list(self.user_map)[0]
        self.user_var=tk.StringVar(value=user_name)
        # Button bar pinned to the bottom (packed first) so it is ALWAYS visible; FlowBar wraps it when narrow.
        btn=FlowBar(self,padding=4); btn.pack(side='bottom',fill='x')
        for txt,cmd in [('Save',self.save),('Attach Receipt',lambda:self.attach('Receipt')),('Attach Invoice',lambda:self.attach('Invoice')),
                        ('Detach Receipt',lambda:self.detach('Receipt')),('Detach Invoice',lambda:self.detach('Invoice')),
                        ('OCR Receipt',lambda:self.ocr('receipt_path')),('OCR Invoice',lambda:self.ocr('invoice_path')),
                        ('Open Receipt Loc',lambda:open_file_location(self.vars['receipt_path'].get())),('Open Invoice Loc',lambda:open_file_location(self.vars['invoice_path'].get())),
                        ('Mark Ready',lambda:self.vars['status'].set('Ready to file')),('Copy to New',self.copy_to_new),('Save as Template',self.save_as_template),('Close',self.close)]:
            btn.button(txt,cmd)
        # Live "what Concur will require" hints — passive, never blocks a save (Concur enforces; we just warn).
        self.hint_var=tk.StringVar(); self.hint_lbl=ttk.Label(self, textvariable=self.hint_var, padding=(8,0), foreground='#996600')
        self.hint_lbl.pack(side='bottom', fill='x')
        row=0; frm=ttk.Frame(self,padding=10); frm.pack(side='top',fill='both',expand=True); frm.columnconfigure(1, weight=1)
        def add(label, widget):
            nonlocal row; ttk.Label(frm,text=label).grid(row=row,column=0,sticky='w',pady=3); widget.grid(row=row,column=1,sticky='ew',pady=3); row+=1
        # Apply-template selector: picking a template fills the matching fields (also re-appliable via the Apply button).
        tpl_row=ttk.Frame(frm); tpl_row.columnconfigure(0, weight=1)
        self.tpl_var=tk.StringVar(); self.tpl_cb=ttk.Combobox(tpl_row, textvariable=self.tpl_var, values=[r['name'] for r in S.get_templates()], state='readonly')
        self.tpl_cb.grid(row=0,column=0,sticky='ew'); self.tpl_cb.bind('<<ComboboxSelected>>', self.apply_template)
        ttk.Button(tpl_row, text='Apply', width=8, command=self.apply_template).grid(row=0,column=1,padx=(4,0))
        add('Apply template', tpl_row)
        add('User', ttk.Combobox(frm, textvariable=self.user_var, values=list(self.user_map), state='readonly'))
        # Assign the expense to one of this user's reports right here (no need for the separate Assign-to-Report step).
        self.report_map={r['name']:r['id'] for r in S.rows('SELECT id,name FROM reports WHERE user_id=? ORDER BY created_at DESC',(self.user_map.get(user_name),))}
        cur_rep=next((n for n,i in self.report_map.items() if i==data.get('report_id')), NO_REPORT)
        self.report_var=tk.StringVar(value=cur_rep)
        add('Expense report', ttk.Combobox(frm, textvariable=self.report_var, values=[NO_REPORT]+list(self.report_map), state='readonly'))
        date_frame=ttk.Frame(frm); date_frame.columnconfigure(0, weight=1)
        ttk.Entry(date_frame, textvariable=v('transaction_date', today())).grid(row=0,column=0,sticky='ew')
        ttk.Button(date_frame, text='📅', width=3, command=lambda:CalendarPopup(self, self.vars['transaction_date'])).grid(row=0,column=1,padx=(4,0))
        add('Transaction date', date_frame)
        vendor_entry=ttk.Entry(frm, textvariable=v('vendor','')); add('Vendor *', vendor_entry); vendor_entry.bind('<FocusOut>', self.apply_vendor_template)
        vcmd=(self.register(_num_validate), '%P')  # restrict numeric fields to money-shaped input as typed
        add('Amount', ttk.Entry(frm, textvariable=v('amount','0.00'), validate='key', validatecommand=vcmd))
        fee_frame=ttk.Frame(frm); fee_frame.columnconfigure(0, weight=1)
        ttk.Entry(fee_frame, textvariable=v('amount_after_fee','0.00'), validate='key', validatecommand=vcmd).grid(row=0,column=0,sticky='ew')
        ttk.Button(fee_frame, text=f'Auto ({int(CC_FEE_RATE*100)}%)', width=10, command=self.calc_after_fee).grid(row=0,column=1,padx=(4,0))
        add('Amount after CC fee', fee_frame)
        add('Currency', ttk.Combobox(frm, textvariable=v('currency','USD'), values=CURRENCIES)); add('Status', ttk.Combobox(frm, textvariable=v('status','Draft'), values=STATUSES, state='readonly'))
        # No binary "Document type" dropdown: a single record holds both a Receipt and an Invoice (see paths + attach buttons below).
        add('Expense type code', CodePicker(frm, v('expense_type_code',''), v('expense_type_label',''))); add('Expense type label', ttk.Entry(frm, textvariable=self.vars['expense_type_label']))
        for fld,label in [('business_purpose','Business purpose'),('business_name','Business name'),('city','City'),('state','State'),('country','Country')]: add(label, ttk.Entry(frm, textvariable=v(fld, 'US' if fld=='country' else '')))
        add('Payment type', ttk.Combobox(frm, textvariable=v('payment_type',PAYMENT_TYPES[0]), values=PAYMENT_TYPES))
        for fld,label in [('is_vendor_invoice','Is this a vendor invoice?'),('personal_no_reimburse','Personal expense / do not reimburse'),('missing_receipt_ack','Missing receipt acknowledgement attached')]: ttk.Checkbutton(frm,text=label,variable=b(fld)).grid(row=row,column=1,sticky='w',pady=1); row+=1
        for fld,label,h in [('comment','Comment',2),('loose_notes','Loose notes',2),('attendees','Attendees',2),('ocr_text','OCR / extracted text',3)]:
            ttk.Label(frm,text=label).grid(row=row,column=0,sticky='nw'); t=tk.Text(frm,height=h); t.insert('1.0', q(data.get(fld,''))); t.grid(row=row,column=1,sticky='ew',pady=3); self.vars[fld]=t; row+=1
        add('Receipt path', ttk.Entry(frm, textvariable=v('receipt_path',''))); add('Invoice path', ttk.Entry(frm, textvariable=v('invoice_path','')))
        add('Invoice number', ttk.Entry(frm, textvariable=v('invoice_number','')))
        fit_to_screen(self, min_w=820, min_h=520)
        # Concur required-field hints update live as the relevant fields change.
        for fld in ('amount','receipt_path','invoice_number','business_purpose','is_vendor_invoice','missing_receipt_ack'):
            self.vars[fld].trace_add('write', self._concur_hints)
        self._concur_hints()
        # Unsaved-changes guard: window [X] and Esc both route through close(), which prompts if edits are pending.
        self._baseline=self._snapshot()
        self.protocol('WM_DELETE_WINDOW', self.close)
    def _snapshot(self):
        out={k:(v.get('1.0','end') if isinstance(v, tk.Text) else v.get()) for k,v in self.vars.items()}
        out['_user']=self.user_var.get(); out['_report']=self.report_var.get()
        return out
    def close(self):
        """Close with a save prompt when there are unsaved edits (Yes=save, No=discard, Cancel=stay)."""
        if self._snapshot()!=self._baseline:
            ans=messagebox.askyesnocancel('Unsaved changes','Save your changes to this expense before closing?', parent=self)
            if ans is None: return
            if ans: self.save(); return  # save() closes on success; on a validation warning the dialog stays open
        self.destroy()
    def _concur_hints(self, *a):
        msgs=[]
        try: amt=float(clean_number(self.vars['amount'].get()))
        except Exception: amt=0
        if amt>75 and not self.vars['receipt_path'].get().strip() and not self.vars['missing_receipt_ack'].get():
            msgs.append('Concur requires a receipt over $75')
        if self.vars['is_vendor_invoice'].get() and not self.vars['invoice_number'].get().strip():
            msgs.append('a vendor invoice needs an invoice number')
        if not self.vars['business_purpose'].get().strip():
            msgs.append('Business purpose is blank (Concur requires it)')
        self.hint_var.set(('⚠ '+ ' · '.join(msgs)) if msgs else '✓ Meets the Concur required-field checks')
        try: self.hint_lbl.configure(foreground='#996600' if msgs else '#2a7a2a')
        except Exception: pass
    def copy_to_new(self):
        """Save the expense as shown, then open a fresh Draft copy of it (the in-dialog Copy button)."""
        if not self.save(close=False): return
        newid=duplicate_expense(self.exp_id); master=self.master
        self.destroy(); master.refresh(); ExpenseDialog(master, newid)
    def calc_after_fee(self):
        try: base=float(clean_number(self.vars['amount'].get()))
        except ValueError: messagebox.showwarning('Amount','Enter a numeric Amount first.'); return
        self.vars['amount_after_fee'].set(f"{base*(1+CC_FEE_RATE):.2f}")
    def apply_vendor_template(self, e=None):
        name=self.vars['vendor'].get().strip()
        if not name: return
        r=S.row('SELECT last_code,last_label FROM vendors WHERE name=?', (name,))
        if r and r['last_code'] and not self.vars['expense_type_code'].get().strip():
            self.vars['expense_type_code'].set(r['last_code']); self.vars['expense_type_label'].set(r['last_label'] or '')
    def set_field(self, name, value):
        """Write a value into one of this dialog's fields, handling StringVar / IntVar / Text widgets alike."""
        v=self.vars.get(name)
        if v is None: return
        if isinstance(v, tk.Text): v.delete('1.0','end'); v.insert('1.0', q(value))
        elif isinstance(v, tk.IntVar):
            try: v.set(int(value) if q(value) not in ('','None') else 0)
            except Exception: pass
        else: v.set(q(value))
    def apply_template(self, e=None):
        name=self.tpl_var.get().strip()
        if not name: return
        for k,val in S.template_fields(name).items(): self.set_field(k, val)
    def current_template_values(self):
        """Snapshot every template-eligible field's current value (string-ified) for Save-as-Template."""
        out={}
        for f,_ in TEMPLATE_FIELDS:
            v=self.vars.get(f)
            if v is None: continue
            out[f]=self.txt(f) if isinstance(v, tk.Text) else (str(v.get()) if isinstance(v, tk.IntVar) else v.get().strip())
        return out
    def save_as_template(self):
        if not self.txt('vendor').strip(): messagebox.showwarning('Vendor required','A template must have a vendor — enter the vendor first.', parent=self); return
        SaveTemplateDialog(self, self.current_template_values(), on_saved=self.refresh_templates)
    def refresh_templates(self):
        self.tpl_cb['values']=[r['name'] for r in S.get_templates()]
    def txt(self, name):
        v=self.vars[name]; return v.get('1.0','end').strip() if isinstance(v, tk.Text) else v.get()
    def save(self, close=True):
        if not self.txt('vendor').strip(): messagebox.showwarning('Required','Vendor name is required.', parent=self); return False
        keys=['transaction_date','vendor','amount','amount_after_fee','currency','status','expense_type_code','expense_type_label','business_purpose','business_name','city','state','country','payment_type','comment','loose_notes','attendees','receipt_path','invoice_path','invoice_number','ocr_text']
        vals={k:self.txt(k) for k in keys}; vals.update({k:self.vars[k].get() for k in ['is_vendor_invoice','personal_no_reimburse','missing_receipt_ack']}); vals['user_id']=self.user_map[self.user_var.get()]
        for nk in ('amount','amount_after_fee'): vals[nk]=float(clean_number(vals.get(nk)))  # store numbers, never '1,405.00'
        vals['report_id']=self.report_map.get(self.report_var.get())  # chosen report, or None when '(No report)'
        if self.exp_id: S.execute('UPDATE expenses SET '+', '.join([f'{k}=?' for k in vals])+' WHERE id=?', tuple(vals.values())+(self.exp_id,))
        else: self.exp_id=S.execute('INSERT INTO expenses('+', '.join(vals)+') VALUES('+', '.join(['?']*len(vals))+')', tuple(vals.values())).lastrowid
        S.upsert_vendor(vals['vendor'], vals['expense_type_code'], vals['expense_type_label'])  # remember last code for this vendor
        self._baseline=self._snapshot()  # what's on screen is now what's on disk
        self.master.refresh()
        if close: self.destroy()
        return True
    def attach(self, doc):
        _inbox=get_local_setting('inbox_path'); _initial=_inbox if (_inbox and Path(_inbox).exists()) else str(Path.home())
        p=filedialog.askopenfilename(initialdir=_initial, title=f'Select {doc.lower()}')
        if not p: return
        year=(self.vars['transaction_date'].get() or today())[:4]
        user_root=S.scalar('SELECT receipt_root FROM users WHERE id=?', (self.user_map[self.user_var.get()],))
        # A user with their own chosen save folder files straight under it\Year; otherwise use the shared root\User\Year.
        root=(Path(user_root) / year) if user_root else (Path(get_local_setting('receipt_root')) / safe_filename(self.user_var.get()) / year)
        try:
            root.mkdir(parents=True, exist_ok=True)
            ext=Path(p).suffix; dest=root / f"{self.vars['transaction_date'].get()} {safe_filename(self.vars['vendor'].get())} {money_part(self.vars['amount'].get())} {doc}{ext}"; i=2
            while dest.exists(): dest=root/f"{dest.stem} ({i}){ext}"; i+=1
            inbox=get_local_setting('inbox_path')
            if is_within(p, inbox):  # came from the inbox -> move it out entirely (copy+rename then drop original)
                shutil.move(p, dest)
            else:
                shutil.copy2(p, dest)
        except Exception as e:
            # Most often the Receipt root is unset/unreachable on THIS machine — tell the user where to fix it
            # instead of silently doing nothing (the old behavior swallowed the error in the button callback).
            messagebox.showerror('Could not save the receipt',
                f"Couldn't file the {doc.lower()} under:\n{root}\n\n{type(e).__name__}: {e}\n\n"
                f"Fix your Receipt root in Settings (it must be a folder that exists on THIS computer — "
                f"e.g. your Google Drive receipts folder on this machine).", parent=self)
            return
        (self.vars['receipt_path'] if doc=='Receipt' else self.vars['invoice_path']).set(str(dest))
        # A record can hold both docs; nudge the status forward without clobbering a later stage.
        cur=self.vars['status'].get()
        if doc=='Receipt':
            if cur in ('', 'Draft', 'Awaiting invoice', 'Awaiting receipt'): self.vars['status'].set('Receipt received')
        else:  # invoice attached; if no receipt yet, we're awaiting the paid receipt
            if not self.vars['receipt_path'].get() and cur in ('', 'Draft', 'Awaiting invoice'): self.vars['status'].set('Awaiting receipt')
    def detach(self, doc):
        """Un-link a misattributed receipt/invoice from this expense. NEVER deletes anything:
        clears the path (saved when you click Save) and optionally moves the file back to the
        inbox so it can be re-attached to the RIGHT expense. NB: the file keeps the name it got
        when it was attached here (date/vendor/amount of THIS expense) — rename if misleading."""
        fld='receipt_path' if doc=='Receipt' else 'invoice_path'
        p=self.vars[fld].get().strip()
        if not p: messagebox.showinfo(f'Detach {doc}', f'No {doc.lower()} is attached.', parent=self); return
        rp=resolve_attachment(p)
        self.vars[fld].set('')
        # Undo attach()'s forward status nudge, but only the receipt-driven stage and only if untouched since.
        if doc=='Receipt' and self.vars['status'].get()=='Receipt received':
            self.vars['status'].set('Awaiting receipt')
        if rp and os.path.exists(rp):
            inbox=get_local_setting('inbox_path')
            if inbox and Path(inbox).exists() and messagebox.askyesno(f'Detach {doc}',
                    f'{doc} detached (takes effect when you Save).\n\nAlso move the file back to your inbox so you can re-attach it to the right expense?\n\n{rp}\n→ {inbox}\n\n(No = leave the file where it is. Nothing is ever deleted.)', parent=self):
                try:
                    stem,ext=Path(rp).stem, Path(rp).suffix
                    dest=Path(inbox)/Path(rp).name; i=2
                    while dest.exists(): dest=Path(inbox)/f"{stem} ({i}){ext}"; i+=1
                    shutil.move(rp, str(dest))
                    messagebox.showinfo(f'Detach {doc}', f'File moved back to the inbox:\n{dest}', parent=self)
                except Exception as e:
                    messagebox.showerror('Move failed', f'{type(e).__name__}: {e}\n\nThe file is untouched at:\n{rp}', parent=self)
    def ocr(self, path_field):
        p=resolve_attachment(self.vars[path_field].get())
        if not p: messagebox.showinfo('No file', f'Attach a {("receipt" if path_field=="receipt_path" else "invoice")} first.'); return
        text=''
        try:
            if p.lower().endswith('.pdf'):
                try: from pypdf import PdfReader
                except Exception: from PyPDF2 import PdfReader
                text='\n'.join([(page.extract_text() or '') for page in PdfReader(p).pages])
            if not text:
                try:
                    import pytesseract
                    from PIL import Image
                    text=pytesseract.image_to_string(Image.open(p))
                except Exception as e: text=f"OCR unavailable. Install Tesseract for Windows and pytesseract. Details: {e}"
        except Exception as e: text=f"Text extraction failed: {e}"
        self.vars['ocr_text'].delete('1.0','end'); self.vars['ocr_text'].insert('1.0', text)

class ReportDialog(tk.Toplevel):
    def __init__(self, master, rep_id=None, default_user_id=None):
        super().__init__(master); self.rep_id=rep_id; self.title('Expense Report'); self.transient(master); self.grab_set()
        data=dict(S.row('SELECT * FROM reports WHERE id=?',(rep_id,))) if rep_id else {}; self.vars={}; self.user_map={r['name']:r['id'] for r in S.rows('SELECT * FROM users ORDER BY name')}
        user_name=next((n for n,i in self.user_map.items() if i==data.get('user_id')), None) or next((n for n,i in self.user_map.items() if i==default_user_id), None) or list(self.user_map)[0]
        self.user_var=tk.StringVar(value=user_name)
        btn=FlowBar(self,padding=4); btn.pack(side='bottom',fill='x'); btn.button('Save',self.save); btn.button('Oracle Codes',lambda:OracleGlossary(self.master)); btn.button('Close',self.close)
        frm=ttk.Frame(self,padding=10); frm.pack(side='top',fill='both',expand=True); frm.columnconfigure(1,weight=1); row=0
        def add(label, var, values=None):
            nonlocal row; ttk.Label(frm,text=label).grid(row=row,column=0,sticky='w',pady=4); w=ttk.Combobox(frm,textvariable=var,values=values) if values else ttk.Entry(frm,textvariable=var); w.grid(row=row,column=1,sticky='ew',pady=4); row+=1
        add('User', self.user_var, list(self.user_map))
        fields=[('name','Report Name *',''),('report_date','Report Date',today()),('start_date','Start Date',''),('end_date','End Date',''),('travel_purpose','Travel Destination/Business Purpose',''),('travel_type','Travel type',TRAVEL_TYPES),('expense_group_id','Expense Group ID',''),('report_id','Report Id',''),('currency','Report Currency','US, Dollar'),('approval_status','Approval Status','Not Submitted'),('payment_status','Payment Status','Not Paid'),('grant_type','Grant/Non Grant',GRANT_TYPES),('expense_report_for','Expense Report For',''),('business_unit','Business Unit',''),('comment','Comment','')]
        for f,label,default in fields:
            self.vars[f]=tk.StringVar(value=q(data.get(f, default if isinstance(default,str) else ''))); add(label,self.vars[f], default if isinstance(default,list) else None)
        # Oracle alias gets an autocomplete picker; new reports default to "Unknown"; new entries grow the list on save.
        self.vars['oracle_alias']=tk.StringVar(value=q(data.get('oracle_alias','')) or ('' if rep_id else 'Unknown'))
        ttk.Label(frm,text='Oracle Alias / Code').grid(row=row,column=0,sticky='w',pady=4); OraclePicker(frm, self.vars['oracle_alias']).grid(row=row,column=1,sticky='ew',pady=4); row+=1
        fit_to_screen(self, min_w=720, min_h=500)
        self._baseline=self._snapshot(); self.protocol('WM_DELETE_WINDOW', self.close)
    def _snapshot(self):
        out={k:v.get() for k,v in self.vars.items()}; out['_user']=self.user_var.get(); return out
    def close(self):
        if self._snapshot()!=self._baseline:
            ans=messagebox.askyesnocancel('Unsaved changes','Save your changes to this report before closing?', parent=self)
            if ans is None: return
            if ans: self.save(); return
        self.destroy()
    def save(self):
        if not self.vars['name'].get().strip(): messagebox.showwarning('Required','Report name is required.', parent=self); return
        vals={k:v.get() for k,v in self.vars.items()}; vals['user_id']=self.user_map[self.user_var.get()]
        S.add_oracle(vals.get('oracle_alias',''))  # grow the oracle list if this is a new code
        if self.rep_id: S.execute('UPDATE reports SET '+', '.join(f'{k}=?' for k in vals)+' WHERE id=?', tuple(vals.values())+(self.rep_id,))
        else: S.execute('INSERT INTO reports('+','.join(vals)+') VALUES('+','.join('?'*len(vals))+')', tuple(vals.values()))
        self.master.refresh(); self.destroy()

class ListGlossary(tk.Toplevel):
    """Shared editor for the generated lists (expense codes, oracle codes, vendors): edit notes, toggle favorite, delete."""
    def __init__(self, master, title, columns, load_fn, key_index, fav_col, notes_col, delete_fn):
        super().__init__(master); self.title(title); self.geometry('780x520'); self.minsize(640, 400)
        self.load_fn=load_fn; self.key_index=key_index; self.fav_col=fav_col; self.notes_col=notes_col; self.delete_fn=delete_fn; self.cols=columns
        top=ttk.Frame(self,padding=8); top.pack(fill='x')
        self.search=tk.StringVar(); ttk.Entry(top,textvariable=self.search).pack(side='left',fill='x',expand=True); ttk.Button(top,text='Search',command=self.refresh).pack(side='left')
        self.bar=FlowBar(self,padding=(8,2)); self.bar.pack(fill='x')
        self.bar.button('Toggle Favorite',self.toggle_fav); self.bar.button('Edit Notes',self.edit_notes); self.bar.button('Delete',self.delete)
        self.tree=ttk.Treeview(self,columns=columns,show='headings')
        for c in columns: self.tree.heading(c,text=c.title()); self.tree.column(c,width=130)
        self.tree.pack(fill='both',expand=True); self.tree.bind('<Double-1>', lambda e: self.edit_notes())
        self.bind('<Escape>', lambda e: self.destroy()); self.refresh()
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for vals in self.load_fn(self.search.get().strip()): self.tree.insert('', 'end', values=vals)
    def _sel(self):
        item=self.tree.focus(); vals=self.tree.item(item,'values'); return vals or None
    def toggle_fav(self):
        vals=self._sel()
        if not vals: return
        new=0 if (vals[self.fav_col] in ('1','★','Yes')) else 1; self._update(vals, favorite=new); self.refresh()
    def edit_notes(self):
        vals=self._sel()
        if not vals: return
        notes=simpledialog.askstring('Notes', f'Notes for {vals[self.key_index]}:', initialvalue=vals[self.notes_col], parent=self)
        if notes is not None: self._update(vals, notes=notes); self.refresh()
    def delete(self):
        vals=self._sel()
        if vals and messagebox.askyesno('Confirm delete', f'Delete {vals[self.key_index]}?', parent=self): self.delete_fn(vals); self.refresh()
    def _update(self, vals, favorite=None, notes=None): pass  # overridden per subclass; receives the full row

class CodeGlossary(ListGlossary):
    """Keyed by (code, name) — the same GL code can be listed as several official expense types.
    `tags` holds the colloquial names/aliases your accounting team actually uses (searchable from
    the picker), `notes` explains WHY a code gets used for things that don't literally match it."""
    def __init__(self, master):
        super().__init__(master, 'Expense Code Glossary', ('code','name','tags','favorite','notes'),
                         load_fn=self._load, key_index=0, fav_col=3, notes_col=4,
                         delete_fn=lambda vals: S.execute('DELETE FROM expense_codes WHERE code=? AND name=?',(vals[0],vals[1])))
        self.bar.button('Edit Tags/Aliases', self.edit_tags)
    def _load(self, term):
        t='%'+term+'%'; return [(r['code'],r['name'],r['tags'],'★' if r['favorite'] else '',r['notes']) for r in S.rows('SELECT * FROM expense_codes WHERE code LIKE ? OR name LIKE ? OR tags LIKE ? ORDER BY favorite DESC, code',(t,t,t))]
    def _update(self, vals, favorite=None, notes=None):
        if favorite is not None: S.execute('UPDATE expense_codes SET favorite=? WHERE code=? AND name=?',(favorite,vals[0],vals[1]))
        if notes is not None: S.execute('UPDATE expense_codes SET notes=? WHERE code=? AND name=?',(notes,vals[0],vals[1]))
    def edit_tags(self):
        vals=self._sel()
        if not vals: return
        tags=simpledialog.askstring('Tags / aliases', f'Colloquial names & search tags for {vals[0]} — {vals[1]}\n(comma-separated; the code picker matches these too):', initialvalue=vals[2], parent=self)
        if tags is not None: S.execute('UPDATE expense_codes SET tags=? WHERE code=? AND name=?',(tags,vals[0],vals[1])); self.refresh()

class OracleGlossary(ListGlossary):
    def __init__(self, master):
        super().__init__(master, 'Oracle / Account Codes', ('code','name','favorite','notes'),
                         load_fn=self._load, key_index=1, fav_col=2, notes_col=3, delete_fn=lambda vals: S.execute('DELETE FROM oracle_codes WHERE name=?',(vals[1],)))
        topadd=ttk.Frame(self,padding=4); topadd.pack(fill='x'); ttk.Button(topadd,text='Add Oracle Code',command=self.add).pack(side='left')
    def _load(self, term):
        t='%'+term+'%'; return [(r['code'],r['name'],'★' if r['favorite'] else '',r['notes']) for r in S.rows('SELECT * FROM oracle_codes WHERE code LIKE ? OR name LIKE ? ORDER BY favorite DESC, name',(t,t))]
    def _update(self, vals, favorite=None, notes=None):
        if favorite is not None: S.execute('UPDATE oracle_codes SET favorite=? WHERE name=?',(favorite,vals[1]))
        if notes is not None: S.execute('UPDATE oracle_codes SET notes=? WHERE name=?',(notes,vals[1]))
    def add(self):
        name=simpledialog.askstring('Add Oracle Code','Name (e.g. Facilities Admin):', parent=self)
        if not name: return
        code=simpledialog.askstring('Add Oracle Code', f'Code for {name} (e.g. 12345):', parent=self) or ''
        try: S.execute('INSERT OR IGNORE INTO oracle_codes(code,name) VALUES(?,?)',(code.strip(),name.strip()))
        except Exception as e: messagebox.showerror('Add failed', str(e), parent=self)
        self.refresh()

class VendorGlossary(ListGlossary):
    def __init__(self, master):
        super().__init__(master, 'Vendors & Templates', ('name','favorite','last_code','last_label','notes'),
                         load_fn=self._load, key_index=0, fav_col=1, notes_col=4, delete_fn=lambda vals: S.execute('DELETE FROM vendors WHERE name=?',(vals[0],)))
    def _load(self, term):
        t='%'+term+'%'; return [(r['name'],'★' if r['favorite'] else '',r['last_code'],r['last_label'],r['notes']) for r in S.rows('SELECT * FROM vendors WHERE name LIKE ? OR last_code LIKE ? ORDER BY favorite DESC, name',(t,t))]
    def _update(self, vals, favorite=None, notes=None):
        if favorite is not None: S.execute('UPDATE vendors SET favorite=? WHERE name=?',(favorite,vals[0]))
        if notes is not None: S.execute('UPDATE vendors SET notes=? WHERE name=?',(notes,vals[0]))

class SaveTemplateDialog(tk.Toplevel):
    """Field-picker for 'Save this expense as a template': name it, then tick which of the current expense's
    filled fields to bake in. Vendor is always required and always included (it is the template's anchor)."""
    def __init__(self, master, values, on_saved=None):
        super().__init__(master); self.title('Save as Template'); self.transient(master); self.grab_set()
        self.values=values or {}; self.on_saved=on_saved; self.checks={}
        frm=ttk.Frame(self,padding=10); frm.pack(side='top',fill='both',expand=True); frm.columnconfigure(1,weight=1); row=0
        ttk.Label(frm,text='Template name *').grid(row=row,column=0,sticky='w',pady=3)
        self.name_var=tk.StringVar(value=(q(self.values.get('vendor','')).strip() or 'New template'))
        ttk.Entry(frm,textvariable=self.name_var).grid(row=row,column=1,sticky='ew',pady=3); row+=1
        ttk.Label(frm,text='Include these fields:').grid(row=row,column=0,columnspan=2,sticky='w',pady=(8,2)); row+=1
        for f,label in TEMPLATE_FIELDS:
            val=q(self.values.get(f,'')).strip()
            if not val and f!='vendor': continue  # only offer fields that actually have a value (vendor always shown)
            var=tk.IntVar(value=1); self.checks[f]=var
            cb=ttk.Checkbutton(frm, variable=var, text=label); cb.grid(row=row,column=0,sticky='w')
            if f=='vendor': cb.state(['disabled','selected'])  # vendor is mandatory in every template
            ttk.Label(frm,text=(val[:60] or '(empty)')).grid(row=row,column=1,sticky='w'); row+=1
        btn=FlowBar(self,padding=4); btn.pack(side='bottom',fill='x'); btn.button('Save Template',self.save); btn.button('Cancel',self.destroy)
        fit_to_screen(self, min_w=480, min_h=320)
    def save(self):
        name=self.name_var.get().strip()
        if not name: messagebox.showwarning('Name required','Enter a template name.', parent=self); return
        vendor=q(self.values.get('vendor','')).strip()
        if not vendor: messagebox.showwarning('Vendor required','A template must include a vendor.', parent=self); return
        fields={f:self.values.get(f,'') for f,var in self.checks.items() if var.get()}; fields['vendor']=vendor
        if S.row('SELECT 1 FROM templates WHERE name=?',(name,)) and not messagebox.askyesno('Overwrite?', f'A template named "{name}" already exists. Replace it?', parent=self): return
        S.upsert_template(name, vendor, fields)
        if self.on_saved: self.on_saved()
        self.destroy()

class TemplateEditor(tk.Toplevel):
    """Create or edit a template directly (home-screen management). Name + every template-eligible field."""
    def __init__(self, master, tpl_id=None, on_saved=None):
        super().__init__(master); self.title('Template'); self.transient(master); self.grab_set()
        self.tpl_id=tpl_id; self.on_saved=on_saved; self.vars={}
        data=dict(S.row('SELECT * FROM templates WHERE id=?',(tpl_id,))) if tpl_id else {}
        try: fields=json.loads(data['fields']) if data.get('fields') else {}
        except Exception: fields={}
        btn=FlowBar(self,padding=4); btn.pack(side='bottom',fill='x'); btn.button('Save',self.save); btn.button('Close',self.destroy)
        frm=ttk.Frame(self,padding=10); frm.pack(side='top',fill='both',expand=True); frm.columnconfigure(1,weight=1); row=0
        ttk.Label(frm,text='Template name *').grid(row=row,column=0,sticky='w',pady=3)
        self.name_var=tk.StringVar(value=q(data.get('name',''))); ttk.Entry(frm,textvariable=self.name_var).grid(row=row,column=1,sticky='ew',pady=3); row+=1
        # Pre-create all vars first so the CodePicker can bind the label var that appears later in the field list.
        for f,_ in TEMPLATE_FIELDS:
            self.vars[f]=tk.IntVar(value=int(fields.get(f,0) or 0)) if f in TEMPLATE_BOOL_FIELDS else tk.StringVar(value=q(fields.get(f,'')))
        for f,label in TEMPLATE_FIELDS:
            if f in TEMPLATE_BOOL_FIELDS:
                ttk.Checkbutton(frm,text=label,variable=self.vars[f]).grid(row=row,column=1,sticky='w',pady=1); row+=1; continue
            ttk.Label(frm,text=label+(' *' if f=='vendor' else '')).grid(row=row,column=0,sticky='w',pady=2)
            if f=='expense_type_code': CodePicker(frm, self.vars['expense_type_code'], self.vars['expense_type_label']).grid(row=row,column=1,sticky='ew',pady=2)
            else: ttk.Entry(frm,textvariable=self.vars[f]).grid(row=row,column=1,sticky='ew',pady=2)
            row+=1
        fit_to_screen(self, min_w=520, min_h=480)
    def save(self):
        name=self.name_var.get().strip(); vendor=self.vars['vendor'].get().strip()
        if not name: messagebox.showwarning('Name required','Enter a template name.', parent=self); return
        if not vendor: messagebox.showwarning('Vendor required','A template must have a vendor.', parent=self); return
        fields={}
        for f,_ in TEMPLATE_FIELDS:
            v=self.vars[f]
            if isinstance(v, tk.IntVar):
                if v.get(): fields[f]=int(v.get())
            elif v.get().strip(): fields[f]=v.get().strip()
        fields['vendor']=vendor
        if self.tpl_id: S.update_template(self.tpl_id, name, vendor, fields)
        else:
            if S.row('SELECT 1 FROM templates WHERE name=?',(name,)) and not messagebox.askyesno('Overwrite?', f'A template named "{name}" already exists. Replace it?', parent=self): return
            S.upsert_template(name, vendor, fields)
        if self.on_saved: self.on_saved()
        self.destroy()

class TemplateGlossary(tk.Toplevel):
    """Home-screen template management: list, create, edit, and delete expense templates."""
    def __init__(self, master):
        super().__init__(master); self.title('Templates'); self.transient(master)
        bar=FlowBar(self,padding=8); bar.pack(side='top',fill='x')
        for txt,cmd in [('New',self.new),('Edit',self.edit),('Delete',self.delete),('Close',self.destroy)]: bar.button(txt,cmd)
        self.tree=ttk.Treeview(self,columns=('name','vendor','code','summary'),show='headings')
        for c,w in [('name',160),('vendor',140),('code',90),('summary',340)]: self.tree.heading(c,text=c.title()); self.tree.column(c,width=w)
        self.tree.pack(fill='both',expand=True,padx=6,pady=6); self.tree.bind('<Double-1>', lambda e: self.edit())
        self.ids={}; self.refresh(); fit_to_screen(self, min_w=640, min_h=380)
    def refresh(self):
        self.tree.delete(*self.tree.get_children()); self.ids={}
        for r in S.get_templates():
            try: f=json.loads(r['fields']) if r['fields'] else {}
            except Exception: f={}
            summary=', '.join(f"{k}={v}" for k,v in f.items() if k!='vendor' and v not in ('',0,'0'))[:120]
            iid=self.tree.insert('', 'end', values=(('★ ' if r['is_default'] else '')+r['name'], r['vendor'], f.get('expense_type_code',''), summary))
            self.ids[iid]=r['id']
    def _sel_id(self): return self.ids.get(self.tree.focus())
    def new(self): TemplateEditor(self, on_saved=self.refresh)
    def edit(self):
        tid=self._sel_id()
        if tid: TemplateEditor(self, tpl_id=tid, on_saved=self.refresh)
        else: messagebox.showinfo('Edit template','Select a template first.', parent=self)
    def delete(self):
        tid=self._sel_id()
        if tid and messagebox.askyesno('Delete template','Delete the selected template?', parent=self):
            S.execute('DELETE FROM templates WHERE id=?',(tid,)); self.refresh()

class AssignReportDialog(tk.Toplevel):
    """Pick an existing report from a list (or unassign) instead of typing the name. Accepts one expense id
    or a list of them (bulk assign from a multi-selection)."""
    def __init__(self, master, exp_id, user_id, on_done):
        super().__init__(master); self.title('Assign to Report'); self.geometry('420x360'); self.minsize(360, 300)
        self.transient(master); self.grab_set(); self.on_done=on_done
        self.exp_ids=list(exp_id) if isinstance(exp_id,(list,tuple,set)) else [exp_id]
        self.reps=S.rows('SELECT id,name FROM reports WHERE user_id=? ORDER BY name',(user_id,))
        n=len(self.exp_ids)
        ttk.Label(self,text=(f'Assign these {n} expenses to:' if n!=1 else 'Assign this expense to:'),padding=8).pack(anchor='w')
        self.lb=tk.Listbox(self); self.lb.pack(fill='both',expand=True,padx=8)
        self.lb.insert('end','<Unassigned>')
        for r in self.reps: self.lb.insert('end', r['name'])
        btn=FlowBar(self,padding=8); btn.pack(side='bottom',fill='x'); btn.button('Assign',self.assign); btn.button('Cancel',self.destroy)
        self.lb.bind('<Double-1>', lambda e: self.assign()); self.bind('<Escape>', lambda e: self.destroy())
    def assign(self):
        sel=self.lb.curselection()
        if not sel: return
        idx=sel[0]; rid=None if idx==0 else self.reps[idx-1]['id']
        for eid in self.exp_ids: S.execute('UPDATE expenses SET report_id=? WHERE id=?',(rid,eid))
        self.on_done(); self.destroy()

class ImportRowDialog(tk.Toplevel):
    """Review ONE imported row: confirm (or change) which staged expense it is, then settle each field the
    two sides disagree on. Card-fed fields are listed read-only — Concur greys them out, so the export wins."""
    def __init__(self, master, plan, on_done=None):
        super().__init__(master); self.title('Review imported expense'); self.transient(master); self.grab_set()
        self.plan=plan; self.on_done=on_done; self.row=plan['row']; self.choice_vars={}
        r=self.row
        head=f"From the export:   {q(r.get('transaction_date'))}   {q(r.get('vendor'))}   {money_part(r.get('amount'))}"
        if r.get('expense_type_label'): head+=f"   {q(r.get('expense_type_code'))} {q(r.get('expense_type_label'))}"
        ttk.Label(self,text=head,padding=(8,6)).pack(anchor='w')
        ttk.Label(self,text='Match it to:',padding=(8,0)).pack(anchor='w')
        self.lb=tk.Listbox(self,height=5,exportselection=False); self.lb.pack(fill='x',padx=8)
        self.lb.insert('end','Import as a NEW expense')
        for e,gap,sim in plan['candidates']:
            self.lb.insert('end', f"#{e['id']}  {q(e['transaction_date'])}  {q(e['vendor'])}  {money_part(e['amount'])}"
                                  f"   ({gap} day{'' if gap==1 else 's'} apart, name match {int(sim*100)}%)")
        self.lb.selection_set(0 if plan['action']!='merge' else 1+plan['cand_index'])
        self.lb.bind('<<ListboxSelect>>', lambda e: self.build_fields())
        self.fields=ttk.Frame(self,padding=(8,4)); self.fields.pack(fill='both',expand=True)
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        bar.button('OK',self.ok); bar.button('Skip this row',self.skip); bar.button('Cancel',self.destroy)
        self.bind('<Escape>', lambda e: self.destroy())
        self.build_fields(); fit_to_screen(self, min_w=560, min_h=420)
    def _selected_candidate(self):
        sel=self.lb.curselection(); i=(sel[0] if sel else 0)-1
        return i if 0<=i<len(self.plan['candidates']) else None
    def build_fields(self):
        for w in self.fields.winfo_children(): w.destroy()
        self.choice_vars={}
        i=self._selected_candidate()
        if i is None:
            ttk.Label(self.fields,text='This row will be added as a new expense — nothing to reconcile.').grid(row=0,column=0,sticky='w')
            return
        exp=self.plan['candidates'][i][0]
        card,conflicts,fills=merge_plan(self.row, exp)
        self.fields.columnconfigure(0,weight=1)
        row=0
        # Two lines per field — name (+ the picker, when there IS a choice) then both values underneath.
        # Everything wraps inside a fixed width so the dialog can be squeezed narrow without clipping.
        for title,items,editable in [('From the card — always taken from the export:',card,False),
                                     ('Both sides filled in — pick one:',conflicts,True),
                                     ('Empty here — the export fills it in:',fills,False)]:
            if not items: continue
            ttk.Label(self.fields,text=title).grid(row=row,column=0,columnspan=2,sticky='w',pady=(8,2)); row+=1
            for f,old,new in items:
                ttk.Label(self.fields,text=FIELD_LABELS.get(f,f)).grid(row=row,column=0,sticky='w',padx=(14,6))
                if editable:
                    v=tk.StringVar(value='Use Concur export' if self.plan['choices'].get(f)=='concur' else 'Keep Concur Buddy')
                    ttk.Combobox(self.fields,textvariable=v,state='readonly',width=18,
                                 values=['Keep Concur Buddy','Use Concur export']).grid(row=row,column=1,sticky='e',padx=6)
                    self.choice_vars[f]=v
                row+=1
                here=f'here “{old}”' if old else 'here (empty)'
                ttk.Label(self.fields,text=f'{here}      export “{new}”',wraplength=420,foreground='#444').grid(
                    row=row,column=0,columnspan=2,sticky='w',padx=(26,6)); row+=1
        if row==0: ttk.Label(self.fields,text='Nothing differs — merging just links the two records.').grid(row=0,column=0,sticky='w')
    def ok(self):
        i=self._selected_candidate()
        if i is None: self.plan['action']='new'; self.plan['cand_index']=None
        else:
            self.plan['action']='merge'; self.plan['cand_index']=i
            self.plan['choices']={f:('concur' if v.get().startswith('Use') else 'buddy') for f,v in self.choice_vars.items()}
        self._close()
    def skip(self): self.plan['action']='skip'; self._close()
    def _close(self):
        if self.on_done: self.on_done()
        self.destroy()

class ConcurImportDialog(tk.Toplevel):
    """Import a report exported OUT of Concur and reconcile it with what is already staged here.

    Nothing is written until Apply. Every row's proposed action is on screen first: merge into a staged
    expense (amount-matched), add as new, or skip. Attachments are never touched by a merge — the export
    has no files — so a staged receipt survives being reconciled with Concur."""
    def __init__(self, master, path, user_id, on_done=None):
        super().__init__(master); self.title('Import Concur export'); self.transient(master); self.grab_set()
        self.user_id=user_id; self.on_done=on_done; self.path=path
        rows, report_name, unknown = parse_concur_export(path)
        if not rows: raise ValueError('no expense rows found in that file')
        self.unknown=unknown
        staged=S.rows('SELECT * FROM expenses WHERE user_id=? ORDER BY transaction_date DESC',(user_id,))
        self.plans=[]; taken=set()
        for r in rows:
            cands=match_concur_row(r, staged, taken)
            # A same-amount row is only PROPOSED as a merge when something corroborates it: a nearby date
            # (card post dates lag the transaction, and a statement cycle is about a month) or a vendor name
            # that actually looks alike. Otherwise it defaults to "add as new" — the candidate is still
            # listed, one click away, so a coincidental amount match can't quietly rewrite the wrong row.
            propose=bool(cands) and (cands[0][1]<=MATCH_DAY_WINDOW or cands[0][2]>=MATCH_NAME_FLOOR)
            plan={'row':r,'candidates':cands,'cand_index':0 if cands else None,
                  'action':'merge' if propose else 'new','choices':{}}
            if propose: taken.add(cands[0][0]['id'])  # one staged expense can only absorb one export row
            self.plans.append(plan)
        top=ttk.Frame(self,padding=(8,6)); top.pack(fill='x')
        ttk.Label(top,text=f"{len(rows)} expenses in {Path(path).name}   ·   matching against "
                           f"{len(staged)} staged for this user").pack(anchor='w')
        rep=ttk.Frame(self,padding=(8,0)); rep.pack(fill='x')
        ttk.Label(rep,text='Group them under report:').pack(side='left')
        self.report_name=tk.StringVar(value=report_name)
        ttk.Entry(rep,textvariable=self.report_name,width=34).pack(side='left',padx=4)
        ttk.Label(rep,text='(blank = leave them loose)').pack(side='left')
        bar=FlowBar(self,padding=(8,4)); bar.pack(fill='x')
        bar.button('Review…',self.review); bar.button('Merge',lambda:self.set_action('merge'))
        bar.button('Add as new',lambda:self.set_action('new')); bar.button('Skip',lambda:self.set_action('skip'))
        bar.separator(); bar.button('Apply',self.apply); bar.button('Cancel',self.destroy)
        self.tree=ttk.Treeview(self,columns=('date','vendor','amount','type','action','detail'),show='headings',selectmode='extended')
        for c,w,t in [('date',90,'Date'),('vendor',185,'Vendor (export)'),('amount',80,'Amount'),
                      ('type',215,'Expense type'),('action',205,'What will happen'),('detail',270,'Notes')]:
            self.tree.heading(c,text=t); self.tree.column(c,width=w)
        self.tree.pack(fill='both',expand=True,padx=8,pady=6); self.tree.bind('<Double-1>', lambda e: self.review())
        hint='Double-click a row to change its match or settle a field. Conflicts keep the Concur Buddy value unless you say otherwise.'
        if unknown: hint+=f"\nColumns with no field here are kept as notes on new expenses: {', '.join(unknown[:6])}"
        ttk.Label(self,text=hint,padding=(8,0),foreground='#555',wraplength=900).pack(anchor='w',side='bottom')
        self.bind('<Escape>', lambda e: self.destroy())
        self.refresh(); fit_to_screen(self, min_w=900, min_h=480)
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i,p in enumerate(self.plans):
            r=p['row']
            if p['action']=='merge':
                e=p['candidates'][p['cand_index']][0]
                card,conflicts,fills=merge_plan(r,e)
                act=f"Merge into #{e['id']} {q(e['vendor'])}"
                bits=[]
                if conflicts: bits.append(f"{len(conflicts)} to decide")
                if card: bits.append(f"{len(card)} from card")
                if fills: bits.append(f"{len(fills)} filled in")
                if len(p['candidates'])>1: bits.append(f"{len(p['candidates'])} candidates")
                if r.get('_receipt_in_concur') is False and (e['receipt_path'] or e['invoice_path']):
                    bits.append('file here, not in Concur')
                detail=' · '.join(bits) or 'nothing differs'
            elif p['action']=='new': act='Add as a new expense'; detail='no amount match' if not p['candidates'] else f"{len(p['candidates'])} possible match(es)"
            else: act='Skip'; detail=''
            etype=' '.join(x for x in (q(r.get('expense_type_code')), q(r.get('expense_type_label'))) if x)
            self.tree.insert('','end',iid=str(i),values=(q(r.get('transaction_date')),q(r.get('vendor')),
                                                         money_part(r.get('amount')),etype,act,detail))
    def _sel(self): return [self.plans[int(i)] for i in self.tree.selection()]
    def set_action(self, action):
        for p in self._sel():
            if action=='merge' and not p['candidates']: continue  # nothing to merge into
            p['action']=action
            if action=='merge' and p['cand_index'] is None: p['cand_index']=0
        self.refresh()
    def review(self):
        sel=self._sel()
        if not sel: messagebox.showinfo('Review','Select a row first.',parent=self); return
        ImportRowDialog(self, sel[0], on_done=self.refresh)
    def apply(self):
        name=self.report_name.get().strip(); rid=None
        if name:
            got=S.row('SELECT id FROM reports WHERE user_id=? AND name=?',(self.user_id,name))
            rid=got['id'] if got else S.execute('INSERT INTO reports(user_id,name,report_date) VALUES(?,?,?)',(self.user_id,name,today())).lastrowid
        merged=new=skipped=codes=0; flags=[]
        for p in self.plans:
            r=p['row']
            if p['action']=='skip': skipped+=1; continue
            if r.get('expense_type_code') and r.get('expense_type_label'):
                if not S.row('SELECT 1 FROM expense_codes WHERE code=? AND name=?',(r['expense_type_code'],r['expense_type_label'])):
                    S.execute('INSERT OR IGNORE INTO expense_codes(code,name) VALUES(?,?)',(r['expense_type_code'],r['expense_type_label'])); codes+=1
            if p['action']=='merge':
                exp=p['candidates'][p['cand_index']][0]
                card,conflicts,fills=merge_plan(r,exp)
                vals={}
                for f,_,_ in card+fills: apply_field(vals, r, f)
                for f,_,_ in conflicts:
                    if p['choices'].get(f)=='concur': apply_field(vals, r, f)
                if rid and exp['report_id'] is None: vals['report_id']=rid  # never move one already in a report
                if vals:
                    S.execute('UPDATE expenses SET '+','.join(f'{f}=?' for f in vals)+' WHERE id=?', tuple(vals.values())+(exp['id'],))
                if r.get('_receipt_in_concur') is False and (exp['receipt_path'] or exp['invoice_path']):
                    flags.append(f"#{exp['id']} {q(exp['vendor'])}")
                merged+=1
            else:
                vals={f:r[f] for f in CARD_FIELDS+MERGE_FIELDS+TYPE_PAIR if r.get(f) not in ('',None)}
                vals.setdefault('transaction_date', today()); vals.setdefault('vendor','(unknown vendor)')
                vals['user_id']=self.user_id; vals['report_id']=rid
                # Concur already holds the receipt image when the export says so; otherwise it's still owed.
                vals['status']='Receipt received' if r.get('_receipt_in_concur') else 'Awaiting receipt'
                if r['_extra']: vals['loose_notes']='\n'.join(f'{k}: {v}' for k,v in r['_extra'].items())
                S.execute('INSERT INTO expenses('+','.join(vals)+') VALUES('+','.join(['?']*len(vals))+')', tuple(vals.values()))
                new+=1
            if r.get('vendor'): S.upsert_vendor(r['vendor'], q(r.get('expense_type_code')), q(r.get('expense_type_label')))
        msg=f"Merged {merged}, added {new} new, skipped {skipped}."
        if rid: msg+=f'\nGrouped under report "{name}".'
        if codes: msg+=f"\nLearned {codes} expense type(s) into the glossary."
        if flags: msg+=('\n\nThese have a receipt staged here that Concur says it does NOT have yet:\n  '
                        +'\n  '.join(flags[:12])+('\n  …' if len(flags)>12 else ''))
        messagebox.showinfo('Import complete', msg, parent=self)
        if self.on_done: self.on_done()
        self.destroy()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        if _FIRST_RUN and not _SUPPRESS_PROMPTS: self.withdraw(); self.prompt_first_run_location(); self.deiconify()
        self.title(f"{APP_TITLE} — v{APP_VERSION}"); self.geometry('1240x740'); self.minsize(1000, 600); self.user_map={}; self.build(); self.refresh()
    def prompt_first_run_location(self):
        """On a genuine first launch, let the user park the DB somewhere synced (e.g. Google Drive) instead of %APPDATA%."""
        use_default=messagebox.askyesno('Welcome to Concur Buddy',
            'Where should Concur Buddy keep its database?\n\n'
            f'• Yes — default location:\n   {DEFAULT_DB_PATH}\n\n'
            '• No — pick a folder yourself (e.g. a Google Drive folder, so it backs up and can be shared\n'
            '   across your devices). You can create a new folder in the next dialog.',
            parent=self)
        if use_default: return
        d=filedialog.askdirectory(title='Choose a folder for the Concur Buddy database (you can create a new folder here)', parent=self)
        if not d: return  # cancelled -> stay on the default
        try:
            S.relocate(d, delete_old=True)  # the just-created default DB is empty; move it to the chosen folder
            messagebox.showinfo('Database location', f'Database will be kept at:\n{S.path}\n\n'
                                'Tip: don\'t open the app on two devices at the same time while it syncs.', parent=self)
        except Exception as e:
            messagebox.showerror('Could not set location', f'{e}\n\nStaying on the default location.', parent=self)
    def build(self):
        self.search=tk.StringVar(); self.status=tk.StringVar(value='All'); self.user=tk.StringVar()
        self.filter_col=tk.StringVar(value='Any'); self.show_filed=tk.IntVar(value=0)
        # Row 1: filters/search only (kept short). Action buttons live in wrapping FlowBars below — never on this row.
        top=ttk.Frame(self,padding=4); top.pack(fill='x')
        ttk.Label(top,text='User').pack(side='left'); self.user_cb=ttk.Combobox(top,textvariable=self.user,state='readonly',width=14); self.user_cb.pack(side='left',padx=3); self.user_cb.bind('<<ComboboxSelected>>',lambda e:self.refresh())
        ttk.Label(top,text='Filter').pack(side='left'); ttk.Combobox(top,textvariable=self.filter_col,values=['Any','vendor','amount','status','doc','code','purpose','date'],state='readonly',width=8).pack(side='left',padx=2); self.filter_col.trace_add('write',lambda *a:self.refresh())
        ttk.Entry(top,textvariable=self.search,width=26).pack(side='left'); self.search.trace_add('write',lambda *a:self.refresh())
        ttk.Label(top,text='Status').pack(side='left',padx=(8,0)); ttk.Combobox(top,textvariable=self.status,values=['All']+STATUSES,state='readonly',width=14).pack(side='left'); self.status.trace_add('write',lambda *a:self.refresh())
        ttk.Checkbutton(top,text='Show Filed',variable=self.show_filed,command=self.refresh).pack(side='left',padx=4)
        # ONE slim, grouped toolbar of the frequent actions; the long tail lives in the "More ▾" dropdown
        # (single window + More dropdown, no menu bar).
        actions=FlowBar(self,padding=(4,2)); actions.pack(fill='x'); self.actions_bar=actions
        hints={'Quick Add':'Log a new expense (Ctrl+N).',
               'New Report':'Create a new expense report (Ctrl+Shift+N).',
               'Edit':'Open the selected expense or report (F2, or double-click an expense).',
               'Attach File':'Attach a receipt or invoice file to the selected expense.',
               'Mark Filed':'Mark the selection Filed in Concur (hides it from the default view).',
               'Add to Report':'Put the selected expense(s) into a report. Tip: you can also drag rows onto a report.',
               'Open Receipt/Invoice':'Open the file attached to the selected expense.',
               'Delete':'Delete the selection (asks for confirmation first).',
               'Open Concur':'Open SAP Concur in your browser to do the real filing.',
               'More ▾':'Everything else: glossaries, copy helpers, folders, import/export, settings.'}
        toolbar=[('Quick Add',self.quick_add),('New Report',self.new_report),None,
                 ('Edit',self.edit),('Attach File',self.attach_selected),('Mark Filed',self.mark_filed),None,
                 ('Add to Report',self.assign_report),('Open Receipt/Invoice',self.open_doc),None,
                 ('Delete',self.delete),None,
                 ('Open Concur',lambda:webbrowser.open(CONCUR_URL))]
        for item in toolbar:
            if item is None: actions.separator(); continue
            txt,cmd=item; b=actions.button(txt,cmd)
            if txt in hints: Tooltip(b, hints[txt])
        self.more_btn=ttk.Menubutton(actions, text='More ▾'); actions.attach(self.more_btn)
        Tooltip(self.more_btn, hints['More ▾'])
        self.more_menu=self._build_more_menu(self.more_btn); self.more_btn.configure(menu=self.more_menu)
        # At-a-glance counter of what's still on the pile: everything not yet marked Filed, plus the two
        # sub-counts that actually drive the next action (no file attached yet / already Ready to file).
        # Recomputed on every refresh; scoped to the user picked in the filter row.
        countbar=ttk.Frame(self,padding=(6,2)); countbar.pack(fill='x')
        self.open_count=tk.StringVar()
        self.count_font=tkfont.nametofont('TkDefaultFont').copy(); self.count_font.configure(weight='bold')
        self.count_lbl=ttk.Label(countbar,textvariable=self.open_count,font=self.count_font,foreground='#1668b8')
        self.count_lbl.pack(side='left')
        Tooltip(self.count_lbl,'Open = every expense for this user that is not marked Filed.\n'
                               '"need a file" = no receipt and no invoice attached yet.\n'
                               'Counts ignore the search box, so this is always the full pile.')
        # Version footer at the very bottom-left — always visible so it's obvious which build is running.
        # The right side doubles as a transient status line (e.g. "Copied to clipboard").
        ver=ttk.Frame(self,padding=(6,1)); ver.pack(side='bottom',fill='x')
        ttk.Label(ver,text=f"Concur Buddy  v{APP_VERSION}",foreground='#777').pack(side='left')
        self.status_msg=tk.StringVar(); ttk.Label(ver,textvariable=self.status_msg,foreground='#2a7a2a').pack(side='right')
        # Taller rows give the report expand/collapse (+/−) control a much bigger hit target.
        ttk.Style().configure('Treeview', rowheight=30)
        # 'id' stays in `columns` (values tuples are positional) but is hidden from view via displaycolumns —
        # the internal row id isn't useful to the user and just adds clutter.
        self.tree=ttk.Treeview(self,columns=('id','date','vendor','amount','status','doc','code','purpose','ready'),displaycolumns=('date','vendor','amount','status','doc','code','purpose','ready'),show='tree headings',selectmode='extended'); self.tree.column('#0',width=265)
        # Left-aligned row check boxes, like Concur's expense list. The box IS the selection: ticking one
        # does exactly what Ctrl-click does (adds/removes that row without clearing the rest), and
        # Ctrl/Shift-click ticks the boxes back. The heading box selects/clears every expense at once.
        self.chk_img=make_checkbox_images(self); self._chk_state={}
        # ttk's stock heading layout parks the heading image on the RIGHT; copy the current theme's layout
        # with the image moved left so the select-all box sits directly above the row boxes. Purely
        # cosmetic and theme-specific, so a theme that won't take it just keeps the stock placement.
        try:
            stl=ttk.Style()
            def image_left(nodes):
                out=[]
                for name,opts in nodes:
                    opts=dict(opts)
                    if name.endswith('image'): opts['side']='left'
                    if 'children' in opts: opts['children']=image_left(opts['children'])
                    out.append((name,opts))
                return out
            stl.layout('Checkbox.Treeview.Heading', image_left(stl.layout('Treeview.Heading')))
            self.tree.configure(style='Checkbox.Treeview')
        except Exception: pass
        self.tree.heading('#0',text='Report / Expense',image=self.chk_img[False],anchor='w',command=self.toggle_select_all)
        for c,w in [('id',55),('date',95),('vendor',190),('amount',85),('status',120),('doc',70),('code',95),('purpose',210),('ready',60)]: self.tree.heading(c,text=c.title()); self.tree.column(c,width=w)
        self.tree.pack(fill='both',expand=True,padx=4,pady=4); self.tree.bind('<Double-1>',self.on_double)
        # Multi-select (Ctrl/Shift-click) + two ways to file a batch into a report: right-click menu, or drag onto a report row.
        self.menu=tk.Menu(self, tearoff=0)
        self.menu.add_command(label='Add to report…', command=self.assign_report)
        self.menu.add_command(label='Copy expense', command=self.copy_expense)
        self.menu.add_command(label='Edit', command=self.edit)
        self.menu.add_separator()
        self.menu.add_command(label='Open receipt/invoice', command=self.open_doc)
        self.menu.add_command(label='Open file location', command=self.open_doc_location)
        self.menu.add_command(label='Copy Concur summary', command=self.copy_summary)
        self.menu.add_separator()
        self.menu.add_command(label='Move to user…', command=self.switch_user)
        self.menu.add_command(label='Mark filed', command=self.mark_filed)
        self.menu.add_command(label='Delete', command=self.delete)
        self.tree.bind('<Button-3>', self.popup_menu)
        self.tree.tag_configure('droptarget', background='#cce6ff')  # report row lights up while a drag hovers it
        self._drag_ids=[]; self._drag_from=None; self._dragging=False; self._press_xy=(0,0); self._drop_row=None; self._ghost=None
        # NB: the check-box handler must be bound BEFORE the drag handler — it returns 'break' on a box
        # click, which stops the rest of the chain (including tk's own click-collapses-the-selection).
        self.tree.bind('<ButtonPress-1>', self._check_click, add='+')
        self.tree.bind('<<TreeviewSelect>>', self._sync_checks, add='+')
        self.tree.bind('<space>', self._space_toggle)
        self.tree.bind('<ButtonPress-1>', self._drag_start, add='+')
        self.tree.bind('<B1-Motion>', self._drag_motion, add='+')
        self.tree.bind('<ButtonRelease-1>', self._drag_drop, add='+')
        self._bind_accelerators()
    def _build_more_menu(self, parent):
        """The 'More ▾' dropdown: the complete long tail of infrequent actions, grouped."""
        m=tk.Menu(parent, tearoff=0)
        gloss=tk.Menu(m, tearoff=0)
        gloss.add_command(label='Expense Codes…', command=lambda:CodeGlossary(self))
        gloss.add_command(label='Oracle Codes…', command=lambda:OracleGlossary(self))
        gloss.add_command(label='Vendors…', command=lambda:VendorGlossary(self))
        gloss.add_command(label='Templates…', command=lambda:TemplateGlossary(self))
        m.add_cascade(label='Glossaries', menu=gloss)
        m.add_separator()
        m.add_command(label='Copy Expense', command=self.copy_expense, accelerator='Ctrl+D')
        m.add_command(label='Move to User…', command=self.switch_user)
        m.add_command(label='Add User…', command=self.add_user)
        m.add_separator()
        m.add_command(label='Copy Concur Summary', command=self.copy_summary)
        m.add_command(label='Copy Attendees', command=self.copy_attendees)
        m.add_separator()
        m.add_command(label='Open File Location', command=self.open_doc_location)
        m.add_command(label='Open Receipt Root', command=lambda:open_folder(get_local_setting('receipt_root'),'receipt root'))
        m.add_command(label='Open Inbox', command=lambda:open_folder(get_local_setting('inbox_path'),'inbox'))
        m.add_separator()
        m.add_command(label='Import Concur Export (.xlsx)…', command=self.import_concur_export)
        m.add_command(label='Export Expenses to CSV…', command=self.export)
        m.add_command(label='Export for Autofill…', command=self.export_autofill)
        m.add_command(label='Export Setup…', command=self.export_config)
        m.add_command(label='Import Setup…', command=self.import_config)
        m.add_separator()
        m.add_command(label='Refresh', command=self.refresh, accelerator='F5')
        m.add_command(label='Settings…', command=self.settings)
        m.add_command(label='About Concur Buddy', command=self.about)
        return m
    def _bind_accelerators(self):
        """Keyboard shortcuts for the frequent actions. Skipped while a modal dialog owns input; the
        selection-scoped ones (Delete / Ctrl+D) only fire when the table itself has keyboard focus,
        so pressing Delete inside the search box can never delete an expense."""
        def h(fn, tree_only=False):
            def run(e):
                try:
                    if self.grab_current() is not None: return
                    if tree_only and self.focus_get() is not self.tree: return
                except Exception: pass
                fn()
            return run
        self.bind_all('<Control-n>', h(self.quick_add)); self.bind_all('<Control-N>', h(self.new_report))
        self.bind_all('<F2>', h(self.edit)); self.bind_all('<F5>', h(self.refresh))
        self.bind_all('<Control-d>', h(self.copy_expense, tree_only=True))
        self.bind_all('<Delete>', h(self.delete, tree_only=True))
    def about(self):
        messagebox.showinfo('About Concur Buddy',
            f'Concur Buddy v{APP_VERSION}\n\nReceipt & expense staging companion for SAP Concur.\n\nDatabase: {S.path}', parent=self)
    def flash_status(self, msg, ms=4000):
        """Transient feedback in the footer (e.g. after a clipboard copy) — visible but never modal."""
        self.status_msg.set(msg)
        self.after(ms, lambda: self.status_msg.set('') if self.status_msg.get()==msg else None)
    def current_user_id(self): return self.user_map.get(self.user.get())
    def refresh_users(self):
        self.user_map={r['name']:r['id'] for r in S.rows('SELECT * FROM users ORDER BY name')}; self.user_cb['values']=list(self.user_map)
        if not self.user.get() and self.user_map: self.user.set(list(self.user_map)[0])
    def refresh(self):
        # Remember which reports are expanded so a rebuild (e.g. after attaching to an expense) doesn't collapse them.
        open_ids={iid for iid in self.tree.get_children('') if self.tree.item(iid,'open')}
        self.refresh_users(); self.tree.delete(*self.tree.get_children()); self._chk_state={}; uid=self.current_user_id(); term=self.search.get(); st=self.status.get()
        for r in S.rows('SELECT * FROM reports WHERE user_id=? ORDER BY created_at DESC',(uid,)) if uid else []:
            # A Filed report disappears from the default view (with its filed expenses) just like a filed expense does.
            if not self.show_filed.get() and st!='Filed' and r['status']=='Filed': continue
            total=S.scalar('SELECT COALESCE(SUM(amount),0) FROM expenses WHERE report_id=?',(r['id'],)); parent=self.tree.insert('', 'end', iid=f"R{r['id']}", text='[Report] '+r['name'], image=self.chk_img[False], values=(r['id'],r['report_date'] or '', '', f"{total:.2f}", r['status'], 'Report', '', r['travel_purpose'] or '', ''), open=(f"R{r['id']}" in open_ids))
            self._chk_state[parent]=False
            for e in S.rows('SELECT * FROM expenses WHERE report_id=? ORDER BY transaction_date DESC',(r['id'],)):
                if self.match(e,term,st): self.insert_exp(parent,e)
        for e in S.rows('SELECT * FROM expenses WHERE user_id=? AND report_id IS NULL ORDER BY transaction_date DESC, id DESC',(uid,)) if uid else []:
            if self.match(e,term,st): self.insert_exp('',e)
        self._sync_checks(); self.refresh_counter()
    def match(self,e,term,st):
        if not self.show_filed.get() and st!='Filed' and e['status']=='Filed': return False  # default-hide the filed archive
        if not (st=='All' or e['status']==st): return False
        term=term.strip().lower()
        if not term: return True
        col=self.filter_col.get()
        field_map={'vendor':'vendor','amount':'amount','status':'status','code':'expense_type_code','purpose':'business_purpose','date':'transaction_date'}
        if col=='Any': hay=' '.join(q(e[k]) for k in e.keys()).lower()
        elif col=='doc': hay=self._docs(e).lower()
        else: hay=q(e[field_map.get(col,'vendor')]).lower()
        return term in hay
    @staticmethod
    def _docs(e):
        return '+'.join(x for x,p in [('Receipt', e['receipt_path']),('Invoice', e['invoice_path'])] if p) or '—'  # both docs can coexist on one record
    def insert_exp(self,parent,e):
        self.tree.insert(parent,'end',iid=f"E{e['id']}", text='[Expense] '+e['vendor'], image=self.chk_img[False], values=(e['id'],e['transaction_date'],e['vendor'],money_part(e['amount']),e['status'],self._docs(e),e['expense_type_code'] or '',e['business_purpose'] or '','✓' if e['status'] in ('Ready to file','Filed') else ''))
        self._chk_state[f"E{e['id']}"]=False
    def refresh_counter(self):
        """The home-screen counter. Deliberately NOT filtered by the search box / status dropdown — it
        answers "how much is still on my plate", which a filtered view would understate."""
        uid=self.current_user_id()
        if not uid: self.open_count.set(''); return
        n=lambda where: S.scalar(f"SELECT COUNT(*) FROM expenses WHERE user_id=? {where}",(uid,)) or 0
        openn=n("AND status<>'Filed'")
        if not openn: self.open_count.set('No open receipts — all caught up.'); return
        need=n("AND status<>'Filed' AND COALESCE(receipt_path,'')='' AND COALESCE(invoice_path,'')=''")
        ready=n("AND status='Ready to file'")
        parts=[f"{openn} open receipt{'' if openn==1 else 's'}"]
        if need: parts.append(f"{need} need{'s' if need==1 else ''} a file")
        if ready: parts.append(f"{ready} ready to file")
        self.open_count.set('   ·   '.join(parts))
    def _hits_checkbox(self, e, row):
        """True when the click landed on the row's check box. Element hit-test first (tk names it
        'image'/'Treeitem.image'); geometry fallback for any theme that names its elements differently."""
        el=str(self.tree.identify_element(e.x, e.y) or '')
        if 'image' in el: return True
        if 'text' in el or 'indicator' in el: return False
        bb=self.tree.bbox(row,'#0')
        if not bb: return False
        try: indent=int(ttk.Style().lookup('Treeview','indent') or 20)
        except Exception: indent=20
        x0=bb[0]+indent*(2 if self.tree.parent(row) else 1)  # the indicator (expander) owns the first slot
        return x0-6 <= e.x <= x0+CHECKBOX_PX+6
    def _check_click(self, e):
        """Clicking a row's box == Ctrl-clicking the row: it toggles just that row in/out of the
        selection and leaves the rest alone. Clicks anywhere else on the row fall through untouched."""
        if self.tree.identify_region(e.x, e.y)!='tree': return
        row=self.tree.identify_row(e.y)
        if not row or not self._hits_checkbox(e, row): return
        self._drag_from=None; self._dragging=False  # a box click never starts a drag
        if row in self.tree.selection():
            self.tree.selection_remove(row)
            if self.tree.focus()==row: self.tree.focus('')
        else: self.tree.selection_add(row); self.tree.focus(row)
        self._sync_checks(); return 'break'  # repaint now — <<TreeviewSelect>> only lands on the next idle pass
    def _space_toggle(self, e):
        """Space ticks/unticks the row you're on — keyboard parity with clicking its box."""
        row=self.tree.focus()
        if not row: return
        if row in self.tree.selection(): self.tree.selection_remove(row)
        else: self.tree.selection_add(row)
        self._sync_checks(); return 'break'
    def toggle_select_all(self):
        """The heading check box: tick every expense row, or clear them all if they're already ticked."""
        rows=[iid for iid in self._chk_state if iid.startswith('E')]
        if not rows: return
        sel=set(self.tree.selection())
        if all(r in sel for r in rows): self.tree.selection_remove(*rows)
        else: self.tree.selection_add(*rows)
        self._sync_checks()
    def _sync_checks(self, e=None):
        """Keep every box in step with the real selection, so Ctrl/Shift-click (and the right-click menu)
        tick the boxes too — the box is a view of the selection, never a second source of truth."""
        sel=set(self.tree.selection())
        for iid,was in list(self._chk_state.items()):
            if not self.tree.exists(iid): self._chk_state.pop(iid,None); continue
            now=iid in sel
            if now!=was: self.tree.item(iid,image=self.chk_img[now]); self._chk_state[iid]=now
        exp=[iid for iid in self._chk_state if iid.startswith('E')]
        self.tree.heading('#0',image=self.chk_img[bool(exp) and all(i in sel for i in exp)])
    def selected(self):
        iid=self.tree.focus(); return (iid[0], int(iid[1:])) if iid else (None,None)
    def quick_add(self): ExpenseDialog(self, default_user_id=self.current_user_id())
    def new_report(self): ReportDialog(self, default_user_id=self.current_user_id())
    def on_double(self, e):
        # Double-click an EXPENSE -> edit it. Double-click a REPORT row -> toggle expand/collapse (so adding
        # expenses to a report is easy); open the report editor via 'Edit Selected' instead.
        iid=self.tree.identify_row(e.y)
        if iid.startswith('R'): self.tree.item(iid, open=not self.tree.item(iid,'open'))
        elif iid.startswith('E'): self.edit()
    def edit(self):
        typ,id=self.selected()
        if typ=='E': ExpenseDialog(self,id)
        elif typ=='R': ReportDialog(self,id)
    def copy_expense(self):
        typ,id=self.selected()
        if typ!='E': messagebox.showinfo('Copy Expense','Select an expense to copy first.'); return
        newid=duplicate_expense(id)  # a copy is a fresh Draft; shared with the in-dialog Copy to New button
        self.refresh(); ExpenseDialog(self, newid)
    def add_user(self):
        name=simpledialog.askstring('Add User','New user name:', parent=self)
        if not name or not name.strip(): return
        name=name.strip()
        # Ask where THIS user's receipts should be saved (first use). Folder dialog lets them make a new folder;
        # Cancel = fall back to the shared receipt root (<root>\<User>\<Year>).
        d=filedialog.askdirectory(title=f"Choose {name}'s receipt save folder (you can create a new folder). Cancel = use the default receipt root.", parent=self)
        S.execute('INSERT OR IGNORE INTO users(name,receipt_root) VALUES(?,?)', (name, d or ''))
        if d:  # also set it on an existing user of the same name (INSERT OR IGNORE wouldn't update)
            S.execute('UPDATE users SET receipt_root=? WHERE name=? AND (receipt_root IS NULL OR receipt_root=\'\')', (d, name))
        self.refresh_users(); self.user.set(name); self.refresh()
    def selected_expense_ids(self):
        """Every EXPENSE row currently selected (multi-select via Ctrl/Shift-click). Report rows are ignored."""
        return [int(iid[1:]) for iid in self.tree.selection() if iid.startswith('E')]
    def assign_report(self):
        ids=self.selected_expense_ids()
        if not ids: messagebox.showinfo('Assign to Report','Select one or more expenses first.'); return
        AssignReportDialog(self, ids, self.current_user_id(), self.refresh)
    def _assign_expenses_to_report(self, ids, rid):
        for eid in ids: S.execute('UPDATE expenses SET report_id=? WHERE id=?',(rid,eid))
        self.refresh()
    def popup_menu(self, e):
        # Right-click: if the clicked row isn't part of the current selection, select just it, then show the menu.
        row=self.tree.identify_row(e.y)
        if row and row not in self.tree.selection(): self.tree.selection_set(row); self.tree.focus(row)
        if self.tree.selection(): self.menu.tk_popup(e.x_root, e.y_root)
    def _drag_start(self, e):
        # Capture the selection BEFORE the default press handler collapses it, so a multi-row drag keeps the group.
        self._drag_from=self.tree.identify_row(e.y); self._drag_ids=self.selected_expense_ids(); self._dragging=False
        self._press_xy=(e.x_root, e.y_root); self._drop_row=None
    def _drag_motion(self, e):
        if not (self._drag_from and self._drag_from.startswith('E')): return  # only expenses are draggable
        if not self._dragging:
            # Movement threshold: a plain click (or a sloppy one) never reads as a drag.
            if abs(e.x_root-self._press_xy[0])+abs(e.y_root-self._press_xy[1]) < 7: return
            self._dragging=True; self.tree.configure(cursor='hand2')
            n=len(self._drag_ids) or 1  # ghost badge follows the cursor so it's obvious what's in hand
            self._ghost=tk.Toplevel(self); self._ghost.wm_overrideredirect(True)
            try: self._ghost.attributes('-topmost', True)
            except Exception: pass
            tk.Label(self._ghost, text=f"{n} expense{'s' if n!=1 else ''}", background='#ffffe0',
                     relief='solid', borderwidth=1, padx=6, pady=2).pack()
        if self._ghost: self._ghost.wm_geometry(f'+{e.x_root+14}+{e.y_root+10}')
        row=self.tree.identify_row(e.y); tgt=row if row.startswith('R') else None
        if tgt!=self._drop_row:  # highlight the report row under the cursor; un-highlight the previous one
            if self._drop_row and self.tree.exists(self._drop_row): self.tree.item(self._drop_row, tags=())
            if tgt: self.tree.item(tgt, tags=('droptarget',))
            self._drop_row=tgt
    def _drag_drop(self, e):
        self.tree.configure(cursor='')
        if getattr(self,'_drop_row',None) and self.tree.exists(self._drop_row): self.tree.item(self._drop_row, tags=())
        self._drop_row=None
        if getattr(self,'_ghost',None): self._ghost.destroy(); self._ghost=None
        if not self._dragging: self._drag_from=None; return
        self._dragging=False; target=self.tree.identify_row(e.y); ids=list(self._drag_ids)
        if self._drag_from and self._drag_from.startswith('E'):
            di=int(self._drag_from[1:])
            if di not in ids: ids.append(di)  # include the row you grabbed even if the press collapsed selection
        self._drag_from=None
        if target.startswith('R') and ids: self._assign_expenses_to_report(ids, int(target[1:]))  # dropped on a report
    def switch_user(self):
        # NB: this MOVES the selected expense/report to another user (reassigns the owner) — it does not
        # change whose expenses you're viewing (that's the User dropdown in the filter row). Labeled
        # "Move to User…" everywhere the user sees it; the method name stays for compatibility.
        typ,id=self.selected()
        if typ not in ('E','R'): messagebox.showinfo('Move to User','Select an expense or report first.'); return
        names=list(self.user_map)
        target=simpledialog.askstring('Move to User', f'Move to which user?\nOptions: {", ".join(names)}', parent=self)
        if not target or target not in self.user_map: return
        uid=self.user_map[target]
        if typ=='E': S.execute('UPDATE expenses SET user_id=?, report_id=NULL WHERE id=?',(uid,id))  # detach from report to avoid cross-user grouping
        else: S.execute('UPDATE reports SET user_id=? WHERE id=?',(uid,id)); S.execute('UPDATE expenses SET user_id=? WHERE report_id=?',(uid,id))
        self.refresh()
    def attach_selected(self):
        typ,id=self.selected()
        if typ=='E': ExpenseDialog(self,id)
        else: open_folder(get_local_setting('receipt_root'), 'receipt root')  # attaching with nothing selected opens the root folder
    def delete(self):
        ids=self.selected_expense_ids(); typ,id=self.selected()
        if ids:  # one or more expenses selected
            msg=f'Delete {len(ids)} selected expenses?' if len(ids)>1 else 'Delete selected item?'
            if messagebox.askyesno('Confirm delete', msg):
                for eid in ids: S.execute('DELETE FROM expenses WHERE id=?',(eid,))
                self.refresh()
        elif typ=='R' and messagebox.askyesno('Confirm delete','Delete selected item?'):
            S.execute('UPDATE expenses SET report_id=NULL WHERE report_id=?',(id,)); S.execute('DELETE FROM reports WHERE id=?',(id,))
            self.refresh()
    def mark_filed(self):
        ids=self.selected_expense_ids(); typ,id=self.selected()
        now=datetime.now().isoformat(timespec='seconds')
        if ids:
            for eid in ids: S.execute("UPDATE expenses SET status='Filed', filed_at=? WHERE id=?",(now,eid))
            self.refresh()
        elif typ=='R':
            # Filing a REPORT files the report and every expense under it (that's what "filed in Concur" means).
            n=S.scalar('SELECT COUNT(*) FROM expenses WHERE report_id=?',(id,)) or 0
            name=S.scalar('SELECT name FROM reports WHERE id=?',(id,)) or 'this report'
            if not messagebox.askyesno('Mark report filed', f'Mark report "{name}" and its {n} expense(s) as Filed?', parent=self): return
            S.execute("UPDATE expenses SET status='Filed', filed_at=? WHERE report_id=?",(now,id))
            S.execute("UPDATE reports SET status='Filed' WHERE id=?",(id,))
            self.refresh()
        else:
            messagebox.showinfo('Mark Filed','Select an expense or a report first.', parent=self)
    def get_exp(self):
        typ,id=self.selected(); return S.row('SELECT * FROM expenses WHERE id=?',(id,)) if typ=='E' else None
    def copy_summary(self):
        e=self.get_exp()
        if not e: self.flash_status('Select an expense first'); return
        self.clipboard_clear(); self.clipboard_append(f"{e['transaction_date']} | {e['vendor']} | ${money_part(e['amount'])} | {e['expense_type_code']} {e['expense_type_label']} | {e['business_purpose']} | {e['comment']}")
        self.flash_status('Concur summary copied to clipboard')
    def copy_attendees(self):
        e=self.get_exp()
        if not e: self.flash_status('Select an expense first'); return
        self.clipboard_clear(); self.clipboard_append(e['attendees'] or '')
        self.flash_status('Attendees copied to clipboard')
    def open_doc(self):
        e=self.get_exp()
        if e: open_path(resolve_attachment(e['receipt_path'] or e['invoice_path']))
    def open_doc_location(self):
        e=self.get_exp()
        if e: open_file_location(e['receipt_path'] or e['invoice_path'])
    def settings(self):
        win=tk.Toplevel(self); win.title('Settings'); win.minsize(620, 300); vars={}; frm=ttk.Frame(win,padding=10); frm.pack(fill='both',expand=True); frm.columnconfigure(1,weight=1)
        # These two are per-machine (stored locally, not in the synced DB) — label them so it's clear.
        for i,k in enumerate(['inbox_path','receipt_root']):
            vars[k]=tk.StringVar(value=get_local_setting(k)); ttk.Label(frm,text=k+'  (this PC)').grid(row=i,column=0,sticky='w'); ttk.Entry(frm,textvariable=vars[k]).grid(row=i,column=1,sticky='ew'); ttk.Button(frm,text='Browse',command=lambda kk=k: vars[kk].set(filedialog.askdirectory() or vars[kk].get())).grid(row=i,column=2)
        # --- Database location (move to a Drive-synced folder, or point at an existing DB on another device) ---
        ttk.Separator(frm,orient='horizontal').grid(row=2,column=0,columnspan=3,sticky='ew',pady=8)
        ttk.Label(frm,text='Database').grid(row=3,column=0,sticky='w')
        db_var=tk.StringVar(value=str(S.path)); ttk.Entry(frm,textvariable=db_var,state='readonly').grid(row=3,column=1,sticky='ew')
        dbbar=FlowBar(frm,padding=(0,2)); dbbar.grid(row=4,column=0,columnspan=3,sticky='ew')
        def move_db():
            d=filedialog.askdirectory(title='Move the database to which folder? (you can create a new folder)', parent=win)
            if not d: return
            try:
                S.relocate(d, delete_old=messagebox.askyesno('Move database','Move complete after you confirm.\n\nAlso delete the old copy at the previous location?', parent=win))
                db_var.set(str(S.path)); self.refresh()
                messagebox.showinfo('Database moved', f'Now using:\n{S.path}\n\nTip: don\'t open the app on two devices at once while it syncs.', parent=win)
            except Exception as e: messagebox.showerror('Move failed', str(e), parent=win)
        def use_existing_db():
            p=filedialog.askopenfilename(title='Select an existing Concur Buddy database', filetypes=[('SQLite DB','*.sqlite3 *.db'),('All files','*.*')], parent=win)
            if not p: return
            try:
                S.use_existing(p); db_var.set(str(S.path)); self.refresh()
                messagebox.showinfo('Database switched', f'Now using:\n{S.path}', parent=win)
            except Exception as e: messagebox.showerror('Switch failed', str(e), parent=win)
        dbbar.button('Move database…', move_db); dbbar.button('Use existing database…', use_existing_db)
        def save():
            for k,v in vars.items(): set_local_setting(k, v.get())  # inbox_path/receipt_root are per-machine local
            win.destroy()
        actbar=FlowBar(frm,padding=(0,8)); actbar.grid(row=5,column=0,columnspan=3,sticky='ew')
        actbar.button('Add User', self.add_user); actbar.button('Save', save)
        fit_to_screen(win, min_w=620, min_h=300)
    def import_concur_export(self):
        """More ▾ → Import Concur Export: reconcile a report exported out of Concur with what's staged here."""
        uid=self.current_user_id()
        if not uid: messagebox.showinfo('Import Concur Export','Pick a user first.',parent=self); return
        p=filedialog.askopenfilename(title='Choose a Concur report export (.xlsx)',
                                     filetypes=[('Excel export','*.xlsx'),('All files','*.*')], parent=self)
        if not p: return
        try: ConcurImportDialog(self, p, uid, self.refresh)
        except Exception as e:
            messagebox.showerror('Import failed', f"Could not read that file:\n{e}\n\nExport the report from Concur "
                                                  "as Excel (.xlsx) — the entries view, one row per expense.", parent=self)
    def export(self):
        p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV','*.csv')])
        if not p: return
        rows=S.rows('SELECT * FROM expenses ORDER BY transaction_date DESC')
        with open(p,'w',newline='',encoding='utf-8') as f:
            w=csv.writer(f); w.writerow(rows[0].keys() if rows else ['no rows']); [w.writerow(tuple(r)) for r in rows]
    def export_autofill(self):
        """Write staged expenses as JSON for the Concur Buddy Autofill browser extension.
        Exports the current selection; with nothing selected, every 'Ready to file' expense of the
        current user. Keys under concur_fields/concur_checkboxes are Concur's own field names
        (the extension matches them to data-nuiexp selectors); manual_fields is the human checklist
        (comboboxes and per-type fields the MVP doesn't auto-fill)."""
        ids=self.selected_expense_ids()
        rows=[S.row('SELECT * FROM expenses WHERE id=?',(i,)) for i in ids] if ids \
            else S.rows("SELECT * FROM expenses WHERE user_id=? AND status='Ready to file' ORDER BY transaction_date",(self.current_user_id(),))
        rows=[r for r in rows if r]
        if not rows:
            messagebox.showinfo('Export for Autofill','Select one or more expenses first (or mark some "Ready to file").'); return
        p=filedialog.asksaveasfilename(defaultextension='.json', initialfile='staged_expenses.json',
                                       filetypes=[('JSON','*.json')], title='Export staged expenses for the autofill extension')
        if not p: return
        def us_date(iso):
            try: y,m,d=q(iso).split('-'); return f'{m}/{d}/{y}'  # Concur types dates as MM/DD/YYYY
            except Exception: return q(iso)
        def receipt_payload(path):
            """Embed the receipt file as base64 so the extension can drop it into Concur's upload
            input (verified live 2026-07-05: DataTransfer + change event uploads for real). Only
            Concur-accepted types; capped at 15 MB to keep the JSON manageable."""
            rp=resolve_attachment(q(path))
            if not rp: return None
            try:
                if not os.path.exists(rp) or os.path.getsize(rp) > 15*1024*1024: return None
                ext=Path(rp).suffix.lower()
                if ext not in ('.png','.jpg','.jpeg','.pdf','.tif','.tiff'): return None
                with open(rp,'rb') as f: return {'name': Path(rp).name, 'b64': base64.b64encode(f.read()).decode()}
            except Exception:
                return None
        out=[]
        for e in rows:
            d=dict(e)
            out.append({
                'id': d['id'], 'vendor': d['vendor'],
                'expense_type_code': q(d['expense_type_code']), 'expense_type_label': q(d['expense_type_label']),
                'concur_fields': {
                    'transactionDate': us_date(d['transaction_date']), 'businessPurpose': q(d['business_purpose']),
                    'vendorName': d['vendor'], 'custom34': q(d['business_name']),
                    'transactionAmount': money_part(d['amount']), 'comment': q(d['comment'])},
                'concur_checkboxes': {'isPersonalExpense': bool(d['personal_no_reimburse']),
                                      'custom3': bool(d['missing_receipt_ack'])},
                'manual_fields': {'city': q(d['city']), 'payment_type': q(d['payment_type']),
                                  'is_vendor_invoice': bool(d['is_vendor_invoice']),
                                  'invoice_number': q(d['invoice_number']), 'attendees': q(d['attendees'])},
                'receipt': receipt_payload(d['receipt_path']),
            })
        with open(p,'w',encoding='utf-8') as f:
            json.dump({'generated': datetime.now().isoformat(timespec='seconds'), 'app_version': APP_VERSION,
                       'expenses': out}, f, indent=2)
        self.flash_status(f'Exported {len(out)} expense(s) for autofill')
    def export_config(self):
        p=filedialog.asksaveasfilename(defaultextension='.json',filetypes=[('JSON','*.json')], title='Export config (users, codes, vendors — no expenses)')
        if not p: return
        config={
            # inbox_path/receipt_root are per-machine (local, not the DB) — intentionally NOT exported so a shared
            # config can't stamp one machine's folders onto another.
            'users':[r['name'] for r in S.rows('SELECT name FROM users ORDER BY name')],
            'oracle_codes':[dict(r) for r in S.rows('SELECT code,name,notes,favorite FROM oracle_codes')],
            'vendors':[dict(r) for r in S.rows('SELECT name,favorite,notes,last_code,last_label FROM vendors')],
            'expense_codes':[dict(r) for r in S.rows('SELECT code,name,tags,favorite,notes FROM expense_codes')],
            'templates':[dict(r) for r in S.rows('SELECT name,vendor,fields,is_default FROM templates')],
        }
        with open(p,'w',encoding='utf-8') as f: json.dump(config, f, indent=2)
        messagebox.showinfo('Export Config', f'Config written to:\n{p}')
    def import_config(self):
        p=filedialog.askopenfilename(filetypes=[('JSON','*.json')], title='Import config')
        if not p: return
        try:
            with open(p,encoding='utf-8') as f: config=json.load(f)
        except Exception as e: messagebox.showerror('Import failed', str(e)); return
        for k,v in (config.get('settings') or {}).items():  # legacy configs may carry these; keep them machine-local
            if k in ('inbox_path','receipt_root'): set_local_setting(k, v)
        for n in config.get('users',[]): S.execute('INSERT OR IGNORE INTO users(name) VALUES(?)',(n,))
        for r in config.get('oracle_codes',[]): S.execute('INSERT OR IGNORE INTO oracle_codes(code,name,notes,favorite) VALUES(?,?,?,?)',(r.get('code',''),r.get('name',''),r.get('notes',''),r.get('favorite',0)))
        for r in config.get('vendors',[]): S.execute('INSERT OR REPLACE INTO vendors(name,favorite,notes,last_code,last_label) VALUES(?,?,?,?,?)',(r.get('name',''),r.get('favorite',0),r.get('notes',''),r.get('last_code',''),r.get('last_label','')))
        for r in config.get('expense_codes',[]): S.execute('INSERT OR REPLACE INTO expense_codes(code,name,tags,favorite,notes) VALUES(?,?,?,?,?)',(r.get('code',''),r.get('name',''),r.get('tags',''),r.get('favorite',0),r.get('notes','')))
        for r in config.get('templates',[]): S.upsert_template(r.get('name',''), r.get('vendor',''), r.get('fields','{}'), r.get('is_default',0))
        self.refresh(); messagebox.showinfo('Import Config','Config imported.')

if __name__ == '__main__': App().mainloop()
