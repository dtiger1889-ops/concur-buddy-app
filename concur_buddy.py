
import os, re, csv, json, sqlite3, shutil, subprocess, webbrowser, sys, calendar, base64, zipfile, difflib
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import date, datetime, timedelta, timezone
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, font as tkfont

APP_TITLE = "Concur Buddy"
APP_VERSION = "2026.08.31.2"  # date-based (YYYY.MM.DD; append .N for an Nth release same day). Shown in the title bar
# + footer, and mirrored by the repo-root VERSION_<APP_VERSION>.txt marker so GitHub shows it at a glance.
# Bump this AND rename the marker together on every release — dev/run_tests.py fails if they diverge.
DB_NAME = "concur_buddy.sqlite3"
LEGACY_DB_NAME = "expense_stager.sqlite3"
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
RECEIPT_LIMIT_DEFAULT = 75.0  # spend at or above this needs a receipt (a common expense-policy figure).
# Adjustable in Settings -> "Receipt needed at or above"; stored in the DB, not the per-machine file, because
# it is a policy number rather than a path. Cached per run so painting a long list doesn't hit the DB per row.
_RECEIPT_LIMIT = [None]
def receipt_limit(reload=False):
    if reload or _RECEIPT_LIMIT[0] is None:
        try: _RECEIPT_LIMIT[0] = float(clean_number(S.get_setting('receipt_limit') or RECEIPT_LIMIT_DEFAULT))
        except Exception: _RECEIPT_LIMIT[0] = RECEIPT_LIMIT_DEFAULT
    return _RECEIPT_LIMIT[0]
def needs_receipt(e):
    """True when this expense actually OWES a receipt: still open, nothing attached, no missing-receipt
    acknowledgement on file, and big enough for the policy to care. Small spend is not chasing work."""
    if q(e['status']) == 'Filed': return False
    if q(e['receipt_path']).strip() or q(e['invoice_path']).strip(): return False
    if 'missing_receipt_ack' in e.keys() and e['missing_receipt_ack']: return False  # column absent on pre-migration rows
    try: return abs(float(e['amount'] or 0)) >= receipt_limit()
    except Exception: return False
# --- attendees -------------------------------------------------------------------------------------------
# Read straight out of Concur's own AttendeeImportTemplate.xls (the lookup sheet inside it). The CODE is what
# the import file wants; the label is what the person picking sees. Employee leads the list because it is the
# common case — the sheet's own order is alphabetical and would bury it.
ATTENDEE_TYPES = [('SYSEMP', 'Employee'), ('BUSGUEST', 'Business Guest'), ('SPOUSE', 'Spouse'),
                  ('STUDENT', 'Student'), ('PARTNER', 'Partner/Collaborator')]
# The template's Attendees sheet, header for header — read out of the file itself, not guessed. It carries
# TWO header rows: Concur's internal field names, then the labels a person reads. Title is the FOURTH column
# and Company the fifth. (Shipped the other way round until 2026-08-11; a real import had only names filled
# in, so the swap went unnoticed.) A successful import needed only the first three — the rest are optional.
ATTENDEE_KEYS = ['AtnTypeKey', 'LastName', 'FirstName', 'Title', 'Company']
ATTENDEE_COLUMNS = ['Attendee Type', 'Last Name', 'First Name', 'Attendee Title', 'Company']
EMAIL_RE = re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+')
def email_rewrites():
    """Optional `from=to` domain fixes, one per line in Settings. OFF by default: Concur turned out not to
    care which of an organisation's domains an address uses, so rewriting was solving a problem nobody had."""
    out = []
    for line in q(S.get_setting('email_rewrites') or '').splitlines():
        if '=' in line:
            a, b = line.split('=', 1)
            if a.strip(): out.append((a.strip().lower(), b.strip().lower()))
    return out
def normalize_email(addr):
    """Trim, unwrap <>, lowercase — enough to spot the same person twice. No domain surgery unless asked."""
    a = q(addr).strip().strip('<>').lower()
    if '@' not in a: return a
    user, _, dom = a.partition('@')
    for frm, to in email_rewrites():
        if dom == frm: dom = to; break
    return f'{user}@{dom}'
def split_name(raw):
    """'Okafor, Chidi' / 'Chidi Okafor' / 'c.okafor@example.edu' -> (last, first). Concur's import is keyed on the two
    names, so guessing them well matters more than anything else here."""
    s = re.sub(r'\s+', ' ', q(raw).strip().strip(',;'))
    if not s: return '', ''
    if ',' in s:                                   # "Last, First" — the unambiguous form
        last, _, first = s.partition(',')
        return last.strip(), first.strip()
    if '@' in s and ' ' not in s:                  # bare address: fall back to the local part
        local = s.split('@')[0]
        bits = re.split(r'[._-]+', local)
        if len(bits) >= 2: return bits[-1].title(), bits[0].title()
        return local.title(), ''
    bits = s.split(' ')
    if len(bits) == 1: return bits[0], ''
    return bits[-1], ' '.join(bits[:-1])           # "Chidi Okafor" -> Okafor / Chidi
def parse_contact_blob(text, default_type='SYSEMP'):
    """Turn pasted anything into contacts. Deliberately forgiving — Concur's own importer is not, and the
    point of doing it here is that a paste out of an email, a calendar invite or a spreadsheet should just
    work. One person per line; these all read correctly:

        Chidi Okafor <c.okafor@mail.example.edu>  Whitfield, Dana; dwhitfield@example.edu; Example Co
        c.okafor@mail.example.edu                 Reyes<TAB>Marisol<TAB>mreyes@example.edu

    Two rules earn their keep. A comma only separates FIELDS when there is no tab or semicolon and there are
    at least two of them — otherwise "Okafor, Chidi" would be torn in half. And a line whose first field is a
    single word with no space is read as spreadsheet columns in the template's own order (Last, First, …);
    anything else treats the first field as a whole name, so "Chidi Okafor; Example Co" keeps the company.
    De-duplicates on email, or on the name when there is none."""
    people = []
    for raw in q(text).replace('\r', '\n').split('\n'):
        line = raw.strip()
        if not line: continue
        emails = EMAIL_RE.findall(line)
        if '\t' in line: fields = line.split('\t')
        elif ';' in line: fields = line.split(';')
        elif line.count(',') >= 2: fields = line.split(',')
        else: fields = [line]
        fields = [EMAIL_RE.sub('', f).strip(' <>,;\t') for f in fields]
        fields = [f for f in fields if f]
        if len(fields) >= 2 and ' ' not in fields[0] and ',' not in fields[0]:
            last, first = fields[0], fields[1]          # spreadsheet columns, template order
            company, title = (fields + ['', ''])[2:4]
        elif fields:
            last, first = split_name(fields[0])         # a whole name, then company/title
            company, title = (fields + ['', ''])[1:3]
        elif emails:
            last, first = split_name(emails[0]); company = title = ''
        else: continue
        if not last and not first: continue
        people.append({'last_name': last.strip(), 'first_name': first.strip(),
                       'email': normalize_email(emails[0]) if emails else '',
                       'company': company.strip(), 'title': title.strip(), 'atn_type': default_type})
    seen = set(); out = []
    for p in people:
        key = p['email'] or f"{p['last_name'].lower()}|{p['first_name'].lower()}"
        if key in seen: continue
        seen.add(key); out.append(p)
    return out
def write_xlsx(path, rows, sheet='Attendees'):
    """Write a minimal .xlsx (stdlib zip + XML, same no-installs promise as the reader). Concur's attendee
    import accepts .xls/.xlsx/.xlsm, and .xlsx is the one that can be written honestly without Excel."""
    def esc(v): return q(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    body = []
    for r, row in enumerate(rows, start=1):
        cells = ''.join(f'<c r="{_col_name(c)}{r}" t="inlineStr"><is><t xml:space="preserve">{esc(v)}</t></is></c>'
                        for c, v in enumerate(row) if q(v) != '')
        body.append(f'<row r="{r}">{cells}</row>')
    sheet_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns='
                 '"http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + ''.join(body) +
                 '</sheetData></worksheet>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" '
                   'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-'
                   'officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns='
                   '"http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type='
                   '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                   'Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns='
                   '"http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.'
                   'openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + esc(sheet) +
                   '" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr('xl/worksheets/sheet1.xml', sheet_xml)
    return path
def attendee_rows(contacts):
    """Contacts -> the template's own sheet, both header rows and all.

    Both headers go in because the file Concur hands out has both, so its importer cannot be skipping
    exactly one row — a file with one header risks losing whoever is on the first data line. Type CODE in
    the first column: that is what a real import needed."""
    return [list(ATTENDEE_KEYS), list(ATTENDEE_COLUMNS)] + \
           [[q(c['atn_type']) or 'SYSEMP', q(c['last_name']), q(c['first_name']),
             q(c['title']), q(c['company'])] for c in contacts]
REIMBURSE_ALERT_DAYS = 7    # how much warning you get before either deadline below
REIMBURSE_DEADLINE_DAYS = 30  # personal spend has to be claimed back within 30 days of the transaction
def days_to_month_end(d=None):
    d = d or date.today()
    return (date(d.year + (d.month == 12), (d.month % 12) + 1, 1) - d).days - 1
def days_since(iso, d=None):
    try: return ((d or date.today()) - date.fromisoformat(q(iso)[:10])).days
    except Exception: return 0
def reimbursement_due(e, d=None):
    """True when this is money you are owed and one of the two clocks is running out.

    There are two, and either is enough. The expense's OWN 30-day claim deadline — that is the rule the
    policy actually states, and it is per-expense, so a mid-month expense can be nearly out of time while
    the month is young. And the month-end reimbursement RUN, which is when a claim has to be in. Warning
    window for both is REIMBURSE_ALERT_DAYS."""
    if q(e['status']) == 'Filed': return False
    if q(e['payment_type']).strip() != PAYMENT_TYPES[1]: return False  # 'Personal / reimbursable'
    if 'personal_no_reimburse' in e.keys() and e['personal_no_reimburse']: return False  # not claiming this one
    if days_since(e['transaction_date'], d) >= REIMBURSE_DEADLINE_DAYS - REIMBURSE_ALERT_DAYS: return True
    return days_to_month_end(d) <= REIMBURSE_ALERT_DAYS
def days_left_to_claim(e, d=None):
    """Days remaining on this expense's own 30-day clock (negative once it is over)."""
    return REIMBURSE_DEADLINE_DAYS - days_since(e['transaction_date'], d)
# --- reading a card off a receipt -------------------------------------------------------------------------
# Receipts, card slips and airline emails all name a card differently: "Visa ending in 1234", "MASTERCARD
# *1234", "Chip Card: Mastercard" with "XXXXXXXXXXXX1234" three lines away, "AMEX ...1234". Matching on the
# number alone hits invoice numbers and phone digits; matching on the network alone hits any payment footer.
# So a card is only recognised when BOTH its network and its last four appear, and how CLOSE they sit decides
# how sure we are.
CARD_NETWORKS = {
    'Visa': r'visas?\b',
    'Mastercard': r'master\s?card|\bmc\b|\bmcard\b',
    'Amex': r'american\s+express|\bamex\b|\bax\b',
    'Discover': r'discover(?:\s+card)?|\bdisc\b',
    'Diners': r'diners(?:\s+club)?',
    'JCB': r'\bjcb\b',
}
# Masked-number shapes seen in the wild, plus the spelled-out "ending in" phrasings.
LAST4_PATTERNS = [
    r'(?:ending(?:\s+(?:in|with))?|end(?:s)?\s+in|last\s*4(?:\s*digits)?(?:\s*[:#])?)\s*[:#-]?\s*(\d{4})\b',
    r'[*x#•·]{2,}[\s-]*(\d{4})\b',            # ****1234, xxxx-1234, ••••1234
    r'(?:\d{4}[\s-]){2,3}(\d{4})\b',           # 4111 1111 1111 1234
    r'\(\s*(\d{4})\s*\)',                      # (1234)
    r'[*x#]\s*(\d{4})\b',                      # *1234, x1234  (also catches "Mastercard\t*1234")
    r'(?:card|acct|account)\s*(?:no\.?|number|#)?\s*[:#]?\s*(?:[*x#•·\d]{4,}[\s-]*)?(\d{4})\b',
]
CARD_PROXIMITY = 120  # characters between the network word and the digits before the match is called "loose"
def cards_for(user_id):
    """That user's cards, plus any left as everyone's. Two people's cards are different cards — matching
    one person's receipt against another's personal Visa would set the wrong payment type on their
    expense."""
    return S.rows('SELECT * FROM cards WHERE user_id=? OR user_id IS NULL ORDER BY last4', (user_id,))
def find_card_mentions(text):
    """Every (network, position) and (last4, position) the text contains, lowercased and de-duplicated."""
    t = q(text); low = t.lower()
    nets = [(name, m.start()) for name, pat in CARD_NETWORKS.items() for m in re.finditer(pat, low)]
    fours = []
    for pat in LAST4_PATTERNS:
        for m in re.finditer(pat, low):
            fours.append((m.group(1), m.start(1)))
    seen = set(); uniq = []
    for d, pos in fours:
        if (d, pos) in seen: continue
        seen.add((d, pos)); uniq.append((d, pos))
    return nets, uniq
def detect_card(text, cards):
    """Match the OCR'd text against the cards set up in Settings.

    Returns (card_row, confidence, why) or (None, None, reason). Confidence is 'sure' when the network and
    the last four sit close together, 'likely' when both appear but far apart (a receipt can print the network
    in a header and the masked number in the footer). Two different cards matching = no guess at all."""
    nets, fours = find_card_mentions(text)
    if not nets and not fours: return None, None, 'no card details found in the text'
    if not nets: return None, None, 'found digits that could be a card, but nothing naming the network'
    if not fours: return None, None, 'found a card network, but no last four digits'
    hits = []
    for c in cards:
        last4 = q(c['last4']).strip(); net = q(c['network']).strip()
        d_pos = [p for d, p in fours if d == last4]
        n_pos = [p for n, p in nets if not net or n == net]
        if not d_pos or not n_pos: continue
        gap = min(abs(dp - np) for dp in d_pos for np in n_pos)
        hits.append((c, 'sure' if gap <= CARD_PROXIMITY else 'likely', gap))
    if not hits: return None, None, 'card details found, but none match a card you have set up'
    hits.sort(key=lambda h: h[2])
    if len(hits) > 1 and hits[0][2] == hits[1][2]:
        return None, None, 'more than one of your cards matches this text — set the payment type by hand'
    c, conf, gap = hits[0]
    why = f"{q(c['network']) or 'card'} ending {q(c['last4'])} found in the text"
    return c, conf, why + ('' if conf == 'sure' else ' (network and digits were far apart, so double-check)')
CC_FEE_RATE = 0.03  # industry-standard credit-card surcharge used by the "Amount after CC fee" auto-calc
NO_REPORT = '(No report)'  # sentinel shown in the expense dialog's report dropdown = leave the expense unassigned
NEW_REPORT = '(new report from the name)'  # sentinel in the importer's merge picker = don't join an existing report

def report_choice_label(r):
    """How a report reads in a one-line picker. A filed report keeps its name but says so, because
    'which of these can I still add to' is the only question a report list has to answer."""
    return f"{q(r['name'])}   (filed)" if q(r['status'])=='Filed' else q(r['name'])
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
# Recurring vendors pre-loaded on first run so they autocomplete and appear in the Vendor glossary without
# being typed by hand. Name only — the app learns each vendor's usual expense code the first time you file
# with it. (The public build ships this list empty; add your own recurring vendors here or in the glossary.)
VENDOR_SEED = []
APPDATA_ROOT = Path(os.getenv("APPDATA", str(Path.home())))
APP_DIR = APPDATA_ROOT / "ConcurBuddy"
LEGACY_APP_DIR = APPDATA_ROOT / "ExpenseStager"

def migrate_legacy_app_data(app_dir=APP_DIR, legacy_app_dir=LEGACY_APP_DIR):
    """Copy pre-rebrand local state forward once, preserving the old files as a rollback copy.

    A custom/synced database pointer stays pointed at the same file. Only the old default database is
    copied to the new Concur Buddy default path and repointed, so upgrades do not strand existing data.
    """
    app_dir = Path(app_dir); legacy_app_dir = Path(legacy_app_dir)
    app_dir.mkdir(parents=True, exist_ok=True)
    if not legacy_app_dir.exists() or legacy_app_dir.resolve() == app_dir.resolve():
        return False
    changed = False
    for name in ("db_location.txt", "local_settings.json"):
        old, new = legacy_app_dir / name, app_dir / name
        if old.exists() and not new.exists():
            shutil.copy2(old, new); changed = True
    old_default = legacy_app_dir / LEGACY_DB_NAME
    new_default = app_dir / DB_NAME
    if old_default.exists() and not new_default.exists():
        shutil.copy2(old_default, new_default); changed = True
    pointer = app_dir / "db_location.txt"
    if pointer.exists() and new_default.exists():
        try:
            if Path(pointer.read_text(encoding="utf-8").strip()).resolve() == old_default.resolve():
                pointer.write_text(str(new_default), encoding="utf-8"); changed = True
        except Exception:
            pass
    return changed

migrate_legacy_app_data()
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
def autofill_filename(n, total):
    """A unique-per-export name so successive exports never collide or overwrite each other.
    Everything in it is procedurally derived: a UTC stamp to the second, how many expenses the
    file holds, and their dollar total — e.g. staged_expenses_20260820T143012Z_7exp_1234.56.json."""
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
    return f"staged_expenses_{stamp}_{n}exp_{total:.2f}.json"
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
def open_attachment(path):
    """Open an attached file ITSELF (in the OS default viewer), re-anchored to this machine's receipt root.
    Same friendly empty/missing handling as open_file_location, which opens the containing folder instead."""
    if not path: messagebox.showinfo("No file", "No file is attached."); return
    path=resolve_attachment(path)
    if not Path(path).exists(): messagebox.showinfo("Not found", f"File no longer exists:\n{path}"); return
    open_path(path)
def scrolled_text(parent, height):
    """A word-wrapped multi-line text box with its own vertical scrollbar, in a frame that stretches to fill
    its grid cell. Returns (frame, text) — grid the frame with sticky='nsew' and give its row weight so the
    box grows with the window. Fixes the "only a sliver visible" detail panes (comment/notes/attendees/OCR)."""
    f=ttk.Frame(parent); f.columnconfigure(0, weight=1); f.rowconfigure(0, weight=1)
    t=tk.Text(f, height=height, width=1, wrap='word'); t.grid(row=0, column=0, sticky='nsew')
    sb=ttk.Scrollbar(f, orient='vertical', command=t.yview); sb.grid(row=0, column=1, sticky='ns')
    t.configure(yscrollcommand=sb.set)
    return f, t
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
    # Belt and braces: the numbers above are an ESTIMATE made before the window manager has placed the
    # window. Display scaling and title-bar/border height both make that estimate too tall, and the result
    # is a dialog whose bottom row (Save, Attach…) sits under the taskbar until you drag the window up.
    # So measure where it actually landed and correct: slide it up first, and only then shrink it.
    try:
        win.update_idletasks()
        h_now = win.winfo_height(); top = win.winfo_rooty(); chrome = max(0, top - win.winfo_y())
        overflow = (top + h_now) - (ay + ah)
        if overflow > 0:
            new_y = max(ay + chrome, top - overflow)
            room = (ay + ah) - new_y
            if room < h_now: win.geometry(f"{win.winfo_width()}x{room}")
            win.geometry(f"+{win.winfo_x()}+{max(ay, new_y - chrome)}")
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
        # Cards you might pay with, so OCR can tell a personal swipe from the corporate one. last4+network
        # together are the key: two cards can share either half.
        # People who turn up on meal/event expenses. Concur's own import wants type code + last + first;
        # email is kept HERE because that is how you find an employee in Concur's search, not in the file.
        c.execute("CREATE TABLE IF NOT EXISTS contacts(id INTEGER PRIMARY KEY, last_name TEXT DEFAULT '', first_name TEXT DEFAULT '', "
                  "email TEXT DEFAULT '', company TEXT DEFAULT '', title TEXT DEFAULT '', atn_type TEXT DEFAULT 'SYSEMP', "
                  "notes TEXT DEFAULT '', favorite INTEGER DEFAULT 0, use_count INTEGER DEFAULT 0)")
        c.execute("CREATE UNIQUE INDEX IF NOT EXISTS contacts_key ON contacts(last_name,first_name,email)")
        c.execute("CREATE TABLE IF NOT EXISTS cards(id INTEGER PRIMARY KEY, user_id INTEGER, last4 TEXT NOT NULL, "
                  "network TEXT DEFAULT '', payment_type TEXT DEFAULT '', label TEXT DEFAULT '', UNIQUE(last4,network,user_id))")
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
            oracle_alias TEXT DEFAULT '', attendees TEXT, receipt_path TEXT, invoice_path TEXT, invoice_number TEXT DEFAULT '', ocr_text TEXT, filed_at TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
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
        # cards gained an owner (two people's cards are different cards). SQLite cannot alter a
        # UNIQUE constraint, so rebuild: rename -> recreate -> copy -> drop. Existing rows keep user_id NULL,
        # which reads as "everyone's", so nobody loses a card they had set up.
        if 'user_id' not in [r['name'] for r in self.rows("PRAGMA table_info(cards)")]:
            self.execute("ALTER TABLE cards RENAME TO cards_old")
            self.execute("CREATE TABLE cards(id INTEGER PRIMARY KEY, user_id INTEGER, last4 TEXT NOT NULL, network TEXT DEFAULT '', "
                         "payment_type TEXT DEFAULT '', label TEXT DEFAULT '', UNIQUE(last4,network,user_id))")
            self.execute("INSERT INTO cards(id,last4,network,payment_type,label) SELECT id,last4,network,payment_type,label FROM cards_old")
            self.execute("DROP TABLE cards_old")
        cols=[r['name'] for r in self.rows("PRAGMA table_info(expenses)")]
        if 'invoice_number' not in cols: self.execute("ALTER TABLE expenses ADD COLUMN invoice_number TEXT DEFAULT ''")
        if 'amount_after_fee' not in cols: self.execute("ALTER TABLE expenses ADD COLUMN amount_after_fee REAL DEFAULT 0")
        # An expense can be coded to an org alias on its own — most spend never becomes a report.
        if 'oracle_alias' not in cols: self.execute("ALTER TABLE expenses ADD COLUMN oracle_alias TEXT DEFAULT ''")
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
        # One-time seed of known recurring vendors, so they exist (and autocomplete) even in a DB created
        # before this list. Guarded by a flag so a vendor the user later deletes stays gone — same contract
        # as the template seed above.
        if self.get_setting('seed_vendors_v1') is None:
            for vname in VENDOR_SEED:
                self.execute('INSERT OR IGNORE INTO vendors(name) VALUES(?)', (vname,))
            self.set_setting('seed_vendors_v1', '1')
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
    def upsert_contact(self, c):
        """Insert or top-up one contact. Never blanks a field you already filled in from a thinner paste."""
        row=self.row('SELECT * FROM contacts WHERE last_name=? AND first_name=? AND email=?',
                     (q(c.get('last_name')), q(c.get('first_name')), q(c.get('email'))))
        if not row and q(c.get('email')):
            row=self.row('SELECT * FROM contacts WHERE email=? AND email<>\'\'', (q(c.get('email')),))
        if row:
            for f in ('email','company','title','atn_type'):
                if q(c.get(f)).strip() and not q(row[f]).strip():
                    self.execute(f'UPDATE contacts SET {f}=? WHERE id=?', (q(c[f]).strip(), row['id']))
            return row['id'], False
        cur=self.execute('INSERT INTO contacts(last_name,first_name,email,company,title,atn_type) VALUES(?,?,?,?,?,?)',
                         (q(c.get('last_name')), q(c.get('first_name')), q(c.get('email')),
                          q(c.get('company')), q(c.get('title')), q(c.get('atn_type')) or 'SYSEMP'))
        return cur.lastrowid, True
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

def report_folder(rid):
    """<receipt root>/<user>/Reports/<report name> — where a report's receipts are gathered."""
    r=S.row('SELECT r.name, u.name AS user, u.receipt_root FROM reports r JOIN users u ON u.id=r.user_id WHERE r.id=?',(rid,))
    if not r: return None
    root=q(r['receipt_root']).strip() or get_local_setting('receipt_root')
    if not root: return None
    return Path(root)/safe_filename(r['user'])/'Reports'/safe_filename(r['name'])
def collect_report_receipts(rid):
    """COPY every file attached to a report's expenses into that report's folder.

    Copies, never moves: the originals stay filed under User\\Year exactly where the database expects them,
    so this can run automatically without ever being the reason an attachment goes missing. Returns
    (folder, copied, skipped_already_there, missing_files)."""
    folder=report_folder(rid)
    if not folder: return None, 0, 0, []
    copied=already=0; missing=[]
    rows=S.rows('SELECT * FROM expenses WHERE report_id=?',(rid,))
    for e in rows:
        for field in ('receipt_path','invoice_path'):
            stored=q(e[field]).strip()
            if not stored: continue
            src=resolve_attachment(stored)
            if not src or not Path(src).exists(): missing.append(f"{q(e['vendor'])} ({'receipt' if field=='receipt_path' else 'invoice'})"); continue
            dest=folder/Path(src).name
            try:
                if dest.exists() and dest.stat().st_size==Path(src).stat().st_size: already+=1; continue
                folder.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dest); copied+=1
            except Exception: missing.append(f"{q(e['vendor'])} (could not copy)")
    return folder, copied, already, missing

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
    'missing receipt acknowledgment form attached': 'missing_receipt_ack', 'missing receipt acknowledgment': 'missing_receipt_ack',
    'missing receipt affidavit': 'missing_receipt_ack', 'missing receipt declaration': 'missing_receipt_ack',
    'receipt': '_receipt_in_concur', 'receipt status': '_receipt_in_concur', 'report name': '_report_name',
}
# Concur greys these out on the expense form because they arrive from the card feed and nothing in Concur
# (or here) can edit them. The export is therefore the authoritative copy: a merge always takes them from
# the file. They are shown before you apply, but never asked about — there is nothing to decide.
CARD_FIELDS = ('transaction_date', 'vendor', 'amount', 'payment_type', 'currency')
# Everything else in an export IS editable in Concur, so a filled-in Concur Buddy value that disagrees is a
# real conflict and the user picks a side. Empty here = just take the export's value.
MERGE_FIELDS = ('business_purpose', 'business_name', 'city', 'state', 'country', 'comment', 'attendees',
                'is_vendor_invoice', 'personal_no_reimburse', 'missing_receipt_ack')
# The expense type's code and its name are two halves of ONE fact, so they are decided TOGETHER under the
# pseudo-field '_expense_type'. Deciding them apart would let "keep my code" + "take their name" mint a
# code/name pair that exists in neither system (one code carrying another type's official name).
TYPE_PAIR = ('expense_type_code', 'expense_type_label')
FIELD_LABELS = {'_expense_type': 'Expense type', 'transaction_date': 'Transaction date', 'vendor': 'Vendor name', 'amount': 'Amount',
                'payment_type': 'Payment type', 'currency': 'Currency', 'expense_type_code': 'Expense type code',
                'expense_type_label': 'Expense type', 'business_purpose': 'Business purpose', 'business_name': 'Business name',
                'city': 'City', 'state': 'State', 'country': 'Country', 'comment': 'Comment', 'attendees': 'Attendees',
                'is_vendor_invoice': 'Is a vendor invoice', 'personal_no_reimburse': 'Personal / no reimburse',
                'missing_receipt_ack': 'Missing receipt acknowledgment'}
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
            elif f in ('is_vendor_invoice', 'personal_no_reimburse', 'missing_receipt_ack'): row[f] = 1 if _yes(val) else 0
            elif f == '_expense_type':
                # "12345-SOME EXPENSE TYPE" -> ('12345', 'SOME EXPENSE TYPE'). Split on the FIRST
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
    matches what you typed ('EXAMPLE PARKING 123' vs 'Sample Parking'), so this only RANKS candidates."""
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
def plan_rows(row, exp):
    """EVERY field a merge would touch, as (field, staged_value, export_value, default, why) — one row per
    field, so the dialog and the writer share a single source of truth and nothing is applied off-screen.
    Only `why == 'both'` is a real decision; the other two are settled by the rules, not by taste:
      card  — Concur GREYS THE FIELD OUT because it comes from the card feed. Not negotiable: the export wins.
      both  — both sides have a value and they differ. This is the one the user picks, per field or in bulk.
      blank — nothing staged here, so the export just fills it in. No decision to make."""
    card, conflicts, fills = merge_plan(row, exp)
    return ([(f, o, n, 'concur', 'card') for f, o, n in card]
            + [(f, o, n, 'buddy', 'both') for f, o, n in conflicts]
            + [(f, o, n, 'concur', 'blank') for f, o, n in fills])
def apply_field(vals, row, field):
    """Copy one decided field out of an export row into a column->value dict, expanding the expense-type
    pseudo-field back into its two real columns."""
    if field == '_expense_type':
        for f in TYPE_PAIR: vals[f] = row.get(f, '')
    else: vals[field] = row[field]
    return vals

# --- Expense-type search engine -------------------------------------------------------------------
# The whole point of the app is FINDING the right expense type when you don't know its official name.
# Institutional expense-type names (e.g. a "Promotional Items" or "Uniforms" line) are never the words a
# person reaches for ("clothing", "swag", "giveaway"). So the search does three things a plain LIKE can't:
# multi-word matching across code+name+tags+notes (any order), everyday-English synonym expansion, and
# a typo-tolerant fallback. All of it only widens what a search FINDS — the value stored is still the
# verbatim Concur code/name. The synonym table is org-neutral English, safe for the public build.
PICKER_STOPWORDS = {'the','a','an','to','of','for','and','or','with','in','on','at','by','from',
                    'item','items','expense','type','stuff','thing','things','some','my','our',
                    'give','away','giving','need','buy','bought','paid','pay'}
FIELD_W = {'code':6,'name':5,'tags':3,'notes':2}  # which field a match landed in, most trustworthy first
# Everyday word <-> the institutional vocabulary official expense-type names use. Each group is matched by
# WHOLE WORD, so it deliberately carries BOTH the human word AND the official term: typing "clothing"
# reaches a type named with words like "promotion"/"publicity"/"uniforms" because those live in the same
# group. Whole-word matching is why short members ('cap','air') are safe — they no longer hit inside
# 'capitalized' or 'repairs'. Org-neutral English, safe to ship in the public build.
SYNONYM_GROUPS = [
    # branded clothing / promo giveaways -> promotion / publicity / uniforms type names
    {'clothing','clothes','apparel','uniform','uniforms','shirt','shirts','tshirt','tshirts','jacket','jackets',
     'hat','hats','cap','caps','swag','merch','merchandise','branded','giveaway','giveaways','promotional',
     'promotion','promotions','promo','publicity','wearable','wearables','logo','logos','tote','totes'},
    {'food','meal','meals','lunch','dinner','breakfast','catering','catered','restaurant','dining',
     'refreshments','snacks','snack','beverage','beverages','coffee','banquet'},
    {'taxi','cab','cabs','rideshare','uber','lyft','transportation','transport','local','metro','subway',
     'bus','train','rail','commute','mileage','parking'},
    {'flight','flights','airfare','airfares','plane','airline','airlines'},
    {'hotel','hotels','lodging','motel','accommodation','accommodations','stay','room','rooms','airbnb'},
    {'software','saas','subscription','subscriptions','app','apps','application','license','licence',
     'licenses','cloud','online','computing'},
    {'computer','computing','laptop','laptops','hardware','equipment','monitor','monitors','peripheral',
     'peripherals','device','devices','electronics','tech'},
    {'office','supplies','supply','stationery','stationary','paper','pens','toner','ink'},
    {'printing','print','prints','reproduction','copies','copying','flyer','flyers','brochure','brochures',
     'poster','posters','banner','banners','signage','binding'},
    {'event','events','conference','conferences','reception','receptions','venue','venues','meeting',
     'meetings','hosting','host','special','activity','activities'},
    {'registration','training','course','courses','seminar','seminars','workshop','workshops','tuition',
     'education','educational'},
    {'membership','memberships','dues','association','associations'},
    {'postage','shipping','mail','mailing','freight','courier','fedex','ups','usps','delivery'},
    {'phone','telephone','telecom','mobile','cell','cellular','internet','wifi','data','broadband'},
    {'gift','gifts','award','awards','recognition','prize','prizes','incentive','incentives'},
    {'car','vehicle','rental','fuel','gas','gasoline','toll','tolls'},
    {'entertainment','tickets','ticket','show','shows','sponsorship','sponsor'},
    {'marketing','advertising','advertisement','advertisements','ads','media','campaign','outreach','publicity'},
    {'consultant','consulting','contractor','contract','freelance','honorarium','speaker','professional'},
]
_SYN_INDEX = {}
for _g in SYNONYM_GROUPS:
    for _w in _g: _SYN_INDEX.setdefault(_w, set()).update(_g)
_SYN_VOCAB = sorted(_SYN_INDEX)
def _expand_word(word):
    """A query word plus every everyday-synonym in its group(s). A word we don't know is first snapped to
    the nearest known everyday word (so 'clohting' -> 'clothing'), giving typo tolerance where it matters."""
    word=word.lower()
    if word in _SYN_INDEX: return {word} | _SYN_INDEX[word]
    near=difflib.get_close_matches(word, _SYN_VOCAB, n=1, cutoff=0.84)
    if near: return {word, near[0]} | _SYN_INDEX[near[0]]
    return {word}
def picker_query_words(term):
    """Split a free-typed query into meaningful lowercase words (drop punctuation + stopwords)."""
    return [w for w in re.split(r'[^0-9a-z]+', (term or '').lower()) if w and w not in PICKER_STOPWORDS]
def _tokens(s):
    return {t for t in re.split(r'[^a-z0-9]+', (s or '').lower()) if t}
def _match_word(w, alts, field_toks, code_raw):
    """Best way query-word w hits a row: (score, field, matched_term, is_direct) or (0,...). The literal
    word may match a whole token, a token PREFIX (>=3 chars, so 'print'->'printing'), or a code substring;
    synonyms match a whole token only — that whole-word rule is what stops 'cap' hitting 'capitalized'."""
    best=(0,None,None,False)
    for field,toks in field_toks.items():
        weight=FIELD_W[field]
        if w in toks or (len(w)>=3 and any(t.startswith(w) for t in toks)) or (field=='code' and w in code_raw):
            if weight+2>best[0]: best=(weight+2, field, w, True)   # a literal hit beats a synonym hit
            continue
        hit=next((a for a in alts if a!=w and a in toks), None)
        if hit and weight>best[0]: best=(weight, field, hit, False)
    return best
def rank_expense_codes(rows, term, vendor_codes=(), limit=12):
    """Rank expense-code rows for a free-typed query. Multi-word AND across code/name/tags/notes with
    everyday-English synonym expansion and typo snapping. `vendor_codes` = codes already used with the
    vendor on the form (a ranking nudge, never a filter). Returns [(row_dict, why), ...] best first.
    Pure and GUI-free so it is unit-testable without Tk."""
    words=picker_query_words(term)
    if not words: return []
    vendor_codes={str(c) for c in vendor_codes if c}
    scored=[]
    for r in rows:
        code_raw=q(r['code']).lower()
        field_toks={'code':_tokens(r['code']),'name':_tokens(r['name']),'tags':_tokens(r['tags']),'notes':_tokens(r['notes'])}
        total=0; matched=0; reasons=[]
        for w in words:
            score,field,termhit,direct=_match_word(w, _expand_word(w), field_toks, code_raw)
            if score: total+=score; matched+=1; reasons.append((w,field,termhit,direct))
        if not matched: continue
        all_matched=matched==len(words); fav=int(r['favorite'] or 0); code=q(r['code'])
        # AND-complete matches float above partial ones; then raw score; favorites and this vendor's own
        # history are tie-breakers, so the code you actually use for this vendor tends to come up first.
        boost=(1000 if all_matched else 0)+total+(4 if fav else 0)+(20 if code in vendor_codes else 0)
        scored.append((boost, all_matched, fav, code, dict(r), reasons))
    if not scored:  # nothing matched even with synonyms -> typo-tolerant fallback on whole name/tag tokens
        for r in rows:
            toks=_tokens(f"{q(r['name'])} {q(r['tags'])}")
            ratio=max((difflib.SequenceMatcher(None,w,t).ratio() for w in words for t in toks), default=0)
            if ratio>=0.78:
                scored.append((ratio, False, int(r['favorite'] or 0), q(r['code']), dict(r), [('~','name',None,False)]))
    scored.sort(key=lambda t:(-t[0], -t[1], -t[2], t[3]))
    return [(r, _why_label(reasons)) for _,_,_,_,r,reasons in scored[:limit]]
def _why_label(reasons):
    """A short 'why this surfaced' note for a picker row — shown only when the reason isn't obvious from
    the visible name/code (i.e. it matched a hidden tag/note, or came in through a synonym / near-miss)."""
    if reasons and reasons[0][0]=='~': return 'closest match'
    bits=[]
    for w,field,term,direct in reasons:
        if direct and field in ('name','code'): continue       # self-evident from the row text
        if term and term!=w: bits.append(f'{w}→{term}')       # matched via a synonym (word -> term)
        elif field in ('tags','notes'): bits.append(term or w)  # matched a hidden field
    seen=set(); uniq=[b for b in bits if not (b in seen or seen.add(b))]
    return ('matches ' + ', '.join(uniq[:2])) if uniq else ''

class CodePicker(ttk.Frame):
    """Find-the-right-expense-type box. Type anything — a number, the official name, or the everyday word
    you actually think in ('clothing', 'swag', 'taxi', 'flight') — and it searches code + name + tags +
    notes at once, expands common words to the institutional terms (so 'clothing' also finds a promotion-
    or uniforms-type line), tolerates typos, and ranks the code you use for this vendor
    first. Each row shows WHY it surfaced when that isn't obvious. Empty box lists your ★ favorites. When
    you pick a code the plain word you searched is remembered on it, so next time it's a direct hit."""
    def __init__(self, master, code_var, label_var, vendor_var=None):
        super().__init__(master); self.code_var=code_var; self.label_var=label_var; self.vendor_var=vendor_var
        self._results=[]; self._last_words=[]
        top=ttk.Frame(self); top.pack(fill='x')
        self.entry=ttk.Entry(top, textvariable=code_var); self.entry.pack(side='left', fill='x', expand=True)
        ttk.Button(top, text='★ Favorites', width=11, command=self.show_favorites).pack(side='left')
        self.lb=tk.Listbox(self, height=7); self.lb.pack(fill='x'); self.lb.pack_forget()
        self.entry.bind('<KeyRelease>', self.search); self.entry.bind('<FocusIn>', self.search); self.lb.bind('<<ListboxSelect>>', self.pick)
    def show_favorites(self):
        self.code_var.set(''); self.entry.focus_set(); self.search()
    def _vendor_name(self):
        return q(self.vendor_var.get()).strip() if self.vendor_var is not None else ''
    def _vendor_history(self):
        """(code, label) pairs this vendor has actually been filed under, most-used first, with the vendor's
        remembered last code pinned to the top — so a known vendor's usual code is row 1 before any typing."""
        v=self._vendor_name()
        if not v: return []
        out=[(r['code'], q(r['name'])) for r in S.rows(
            "SELECT expense_type_code code, expense_type_label name, COUNT(*) n FROM expenses "
            "WHERE vendor=? AND expense_type_code<>'' GROUP BY expense_type_code, expense_type_label ORDER BY n DESC, code",(v,))]
        last=S.row('SELECT last_code, last_label FROM vendors WHERE name=?',(v,))
        if last and last['last_code']:
            pair=(last['last_code'], q(last['last_label']))
            out=[pair]+[p for p in out if p!=pair]
        return out
    def _vendor_codes(self):
        """Just the codes from this vendor's history — the ranking nudge passed to rank_expense_codes."""
        return [c for c,_ in self._vendor_history()]
    def search(self, e=None):
        term=self.code_var.get().strip(); self.lb.delete(0,'end'); self._results=[]
        self._last_words=picker_query_words(term)
        if term:
            allrows=S.rows("SELECT code,name,tags,favorite,notes FROM expense_codes")
            ranked=rank_expense_codes(allrows, term, vendor_codes=self._vendor_codes(), limit=12)
            for r,why in ranked:
                star='★ ' if r.get('favorite') else ''
                self.lb.insert('end', f"{star}{r['code']} — {r['name']}" + (f"    · {why}" if why else ''))
                self._results.append((r['code'], r['name']))
            if not ranked:
                self.lb.insert('end', '  (no match — try a plainer word, e.g. "clothing", "taxi", "food")')
        else:
            # Empty box: a known vendor's own codes first (row 1 = the usual answer), then your ★ favorites.
            shown=set()
            for code,name in self._vendor_history():
                r=S.row('SELECT code,name FROM expense_codes WHERE code=? AND name=?',(code,name)) or S.row('SELECT code,name FROM expense_codes WHERE code=? ORDER BY name LIMIT 1',(code,))
                key=(r['code'], r['name']) if r else (code, name)
                if key in shown: continue
                self.lb.insert('end', f"{key[0]} — {key[1]}    · used for {self._vendor_name()}")
                self._results.append(key); shown.add(key)
            for r in S.rows("SELECT code,name FROM expense_codes WHERE favorite=1 ORDER BY code"):
                key=(r['code'], r['name'])
                if key in shown: continue
                self.lb.insert('end', f"★ {key[0]} — {key[1]}"); self._results.append(key); shown.add(key)
        self.lb.pack(fill='x') if self.lb.size() else self.lb.pack_forget()
    def pick(self, e=None):
        sel=self.lb.curselection()
        if not sel or sel[0]>=len(self._results): return  # the 'no match' hint row isn't pickable
        code,name=self._results[sel[0]]  # kept alongside the list so the ' · why' suffix never confuses parsing
        r=S.row('SELECT * FROM expense_codes WHERE code=? AND name=?',(code,name)) or S.row('SELECT * FROM expense_codes WHERE code=?',(code,))
        if r:
            self.code_var.set(r['code']); self.label_var.set(r['name']); self._learn(r['code'], r['name'])
        self.lb.pack_forget()
    def _learn(self, code, name):
        """Remember the everyday word you searched as a tag on the code you chose, so the same search is a
        direct hit next time. Only genuinely-new words (>=3 letters, not already in the code/name/tags) are
        added — numbers and words it already matches are skipped, so this stays low-noise and self-correcting
        (edit or clear tags any time in More ▾ -> Glossaries -> Expense Codes)."""
        if not self._last_words: return
        row=S.row('SELECT tags FROM expense_codes WHERE code=? AND name=?',(code,name))
        if not row: return
        have=f"{code} {name} {q(row['tags'])}".lower()
        add=[w for w in dict.fromkeys(self._last_words) if len(w)>=3 and not w.isdigit() and w not in have]
        if not add: return
        tags=q(row['tags']).strip()
        S.execute('UPDATE expense_codes SET tags=? WHERE code=? AND name=?',
                  ((tags+', ' if tags else '')+', '.join(add), code, name))

class OraclePicker(ttk.Frame):
    """Autocomplete for org/oracle aliases. Clicking the box lists what you already have (favourites first) —
    it used to search only on a keystroke, which made every existing code unreachable unless you already knew
    how it started. New values still grow the list when the record is saved."""
    def __init__(self, master, value_var):
        super().__init__(master); self.value_var=value_var
        row=ttk.Frame(self); row.pack(fill='x')
        self.entry=ttk.Entry(row, textvariable=value_var); self.entry.pack(side='left', fill='x', expand=True)
        ttk.Button(row, text='▾', width=3, command=self.show_all).pack(side='left')
        self.lb=tk.Listbox(self, height=6); self.lb.pack(fill='x'); self.lb.pack_forget()
        self.entry.bind('<KeyRelease>', self.search); self.entry.bind('<FocusIn>', self.search)
        self.lb.bind('<<ListboxSelect>>', self.pick)
    def show_all(self):
        self.value_var.set(''); self.entry.focus_set(); self.search()
    def search(self, e=None):
        term=self.value_var.get().strip(); self.lb.delete(0,'end')
        rows=S.rows("SELECT code,name FROM oracle_codes WHERE code LIKE ? OR name LIKE ? ORDER BY favorite DESC, code, name LIMIT 20",
                    ('%'+term+'%', '%'+term+'%'))
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
                        ('Open Receipt',lambda:open_attachment(self.vars['receipt_path'].get())),('Open Receipt Loc',lambda:open_file_location(self.vars['receipt_path'].get())),
                        ('Open Invoice',lambda:open_attachment(self.vars['invoice_path'].get())),('Open Invoice Loc',lambda:open_file_location(self.vars['invoice_path'].get())),
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
        # Open (not-yet-filed) reports come first and already-filed ones are labelled as such — same reasoning as
        # AssignReportDialog: the reports you can still add to must not be lost among the ones already submitted.
        self.report_map={report_choice_label(r):r['id'] for r in
                         S.rows("SELECT id,name,status FROM reports WHERE user_id=? "
                                "ORDER BY (status='Filed'), created_at DESC",(self.user_map.get(user_name),))}
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
        add('Expense type code', CodePicker(frm, v('expense_type_code',''), v('expense_type_label',''), vendor_var=self.vars.get('vendor'))); add('Expense type label', ttk.Entry(frm, textvariable=self.vars['expense_type_label']))
        # Org alias per EXPENSE: most spend is coded and filed without ever becoming a report, and it is
        # usually the same alias with the occasional exception, so it belongs here as well as on the report.
        self.vars['oracle_alias']=tk.StringVar(value=q(data.get('oracle_alias','')))
        oa=ttk.Frame(frm); oa.columnconfigure(0,weight=1)
        OraclePicker(oa, self.vars['oracle_alias']).grid(row=0,column=0,sticky='ew')
        add('Org / Oracle alias', oa)
        for fld,label in [('business_purpose','Business purpose'),('business_name','Business name'),('city','City'),('state','State'),('country','Country')]: add(label, ttk.Entry(frm, textvariable=v(fld, 'US' if fld=='country' else '')))
        add('Payment type', ttk.Combobox(frm, textvariable=v('payment_type',PAYMENT_TYPES[0]), values=PAYMENT_TYPES))
        for fld,label in [('is_vendor_invoice','Is this a vendor invoice?'),('personal_no_reimburse','Personal expense / do not reimburse'),('missing_receipt_ack','Missing receipt acknowledgement attached')]: ttk.Checkbutton(frm,text=label,variable=b(fld)).grid(row=row,column=1,sticky='w',pady=1); row+=1
        # Comment / Loose notes / Attendees / OCR are multi-line: each gets its own scrollbar + word wrap, and
        # its row is weighted so the box grows when you enlarge the window (OCR text is longest, so it grows most).
        for fld,label,h,wt in [('comment','Comment',3,1),('loose_notes','Loose notes',3,1),('attendees','Attendees',4,1),('ocr_text','OCR / extracted text',6,2)]:
            ttk.Label(frm,text=label).grid(row=row,column=0,sticky='nw',pady=3)
            box,t=scrolled_text(frm,h); box.grid(row=row,column=1,sticky='nsew',pady=3); frm.rowconfigure(row, weight=wt)
            t.insert('1.0', q(data.get(fld,''))); self.vars[fld]=t; row+=1
            if fld=='attendees':  # address book button sits directly UNDER the attendees box, not out in a 3rd column (which widened the form)
                ab_btn=ttk.Button(frm,text='Address book…',command=self.pick_attendees); ab_btn.grid(row=row,column=1,sticky='w',pady=(0,4)); row+=1
                Tooltip(ab_btn, 'Pick attendees you have used before (or paste new ones in), then export the file '
                                'Concur\'s importer wants.')
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
        if amt>=receipt_limit() and not self.vars['receipt_path'].get().strip() and not self.vars['missing_receipt_ack'].get():
            msgs.append(f'receipt needed at or over ${receipt_limit():.0f}')
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
    def pick_attendees(self):
        """Open the address book pre-ticked with whoever is already on this expense, and write back a
        readable list. Concur only needs the names; the emails ride along because that is how you find an
        employee in Concur's own search."""
        def took(rows):
            lines=[f"{q(r['first_name'])} {q(r['last_name'])}".strip() + (f" <{q(r['email'])}>" if q(r['email']).strip() else '')
                   for r in rows]
            box=self.vars['attendees']; box.delete('1.0','end'); box.insert('1.0', '\n'.join(lines))
        AttendeeDialog(self, on_pick=took, preselect=self.vars['attendees'].get('1.0','end-1c'))

    def save_as_template(self):
        if not self.txt('vendor').strip(): messagebox.showwarning('Vendor required','A template must have a vendor — enter the vendor first.', parent=self); return
        SaveTemplateDialog(self, self.current_template_values(), on_saved=self.refresh_templates)
    def refresh_templates(self):
        self.tpl_cb['values']=[r['name'] for r in S.get_templates()]
    def txt(self, name):
        v=self.vars[name]; return v.get('1.0','end').strip() if isinstance(v, tk.Text) else v.get()
    def save(self, close=True):
        if not self.txt('vendor').strip(): messagebox.showwarning('Required','Vendor name is required.', parent=self); return False
        self._rename_organized_files()  # keep the organized receipt/invoice filename in sync with vendor/amount/date
        keys=['transaction_date','vendor','amount','amount_after_fee','currency','status','expense_type_code','expense_type_label','business_purpose','business_name','city','state','country','payment_type','oracle_alias','comment','loose_notes','attendees','receipt_path','invoice_path','invoice_number','ocr_text']
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
    def organized_root(self):
        """The folder an attachment for THIS expense files into: the user's own receipt folder\\Year, or the
        shared receipt root\\User\\Year. Filesystem path — machine-specific, and may not exist yet."""
        year=(self.vars['transaction_date'].get() or today())[:4]
        user_root=S.scalar('SELECT receipt_root FROM users WHERE id=?', (self.user_map[self.user_var.get()],))
        return (Path(user_root) / year) if user_root else (Path(get_local_setting('receipt_root')) / safe_filename(self.user_var.get()) / year)
    def organized_name(self, doc, ext, exclude=None):
        """Collision-safe organized path `date vendor amount [Receipt|Invoice].ext` from the CURRENT form
        fields. `exclude` (a path) is not counted as a collision, so a file can keep its own slot on rename."""
        root=self.organized_root(); root.mkdir(parents=True, exist_ok=True)
        base=f"{self.vars['transaction_date'].get()} {safe_filename(self.vars['vendor'].get())} {money_part(self.vars['amount'].get())} {doc}"
        ex=Path(exclude).resolve() if exclude else None
        dest=root/f"{base}{ext}"; i=2
        while dest.exists() and dest.resolve()!=ex: dest=root/f"{base} ({i}){ext}"; i+=1
        return dest
    def _rename_organized_files(self):
        """On save, re-derive each attached file's name from the current date/vendor/amount and rename it to
        match. This is the fix for attaching a receipt BEFORE the vendor/amount is known (common: the card feed
        supplies the real vendor a day or two later) — otherwise the file keeps its blank-vendor / 0.00 name and
        the OCR->autofill pipeline can't find it. Best-effort: any problem (unreachable receipt root, file
        already moved) leaves the file and stored path untouched and NEVER blocks the save."""
        for doc, fld in (('Receipt','receipt_path'), ('Invoice','invoice_path')):
            stored=self.vars[fld].get().strip()
            if not stored: continue
            try:
                cur=resolve_attachment(stored)
                if not (cur and os.path.exists(cur)): continue
                dest=self.organized_name(doc, Path(cur).suffix, exclude=cur)
                if Path(cur).resolve()==dest.resolve(): continue  # already named right
                shutil.move(cur, dest)
                self.vars[fld].set(str(dest))
            except Exception:
                pass  # a rename is a nicety; the DB save must always go through
    def attach(self, doc):
        _inbox=get_local_setting('inbox_path'); _initial=_inbox if (_inbox and Path(_inbox).exists()) else str(Path.home())
        p=filedialog.askopenfilename(initialdir=_initial, title=f'Select {doc.lower()}')
        if not p: return
        try:
            ext=Path(p).suffix; dest=self.organized_name(doc, ext)
            inbox=get_local_setting('inbox_path')
            if is_within(p, inbox):  # came from the inbox -> move it out entirely (copy+rename then drop original)
                shutil.move(p, dest)
            else:
                shutil.copy2(p, dest)
        except Exception as e:
            # Most often the Receipt root is unset/unreachable on THIS machine — tell the user where to fix it
            # instead of silently doing nothing (the old behavior swallowed the error in the button callback).
            messagebox.showerror('Could not save the receipt',
                f"Couldn't file the {doc.lower()} under:\n{self.organized_root()}\n\n{type(e).__name__}: {e}\n\n"
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
                    from PIL import Image, ImageOps
                    # Normalize before OCR: exif_transpose honors phone-photo rotation, and convert('L')
                    # grayscales it (better for OCR) AND clears PIL's detected .format. That format reset is
                    # the fix for phone JPEGs PIL reports as MPO (multi-picture) or other: pytesseract only
                    # accepts a fixed format whitelist and otherwise raises "Unsupported image format/type",
                    # even with Tesseract correctly installed.
                    text=pytesseract.image_to_string(ImageOps.exif_transpose(Image.open(p)).convert('L'))
                except ImportError:
                    text="OCR needs the pytesseract + Pillow packages installed (see README setup)."
                except Exception as e:
                    # A missing Tesseract binary arrives as pytesseract.TesseractNotFoundError; anything else
                    # (an image that still can't be read) lands here too. Report the real reason instead of
                    # telling the user to install something they already have.
                    if e.__class__.__name__=='TesseractNotFoundError':
                        text="OCR needs the Tesseract program. Install Tesseract for Windows (see README) and restart Concur Buddy."
                    else:
                        text=f"Couldn't read text from this file: {e}"
        except Exception as e: text=f"Text extraction failed: {e}"
        self.vars['ocr_text'].delete('1.0','end'); self.vars['ocr_text'].insert('1.0', text)
        self.detect_payment_card(text)
    def detect_payment_card(self, text):
        """After OCR, work out WHICH card paid — the difference between 'file it' and 'claim it back'."""
        cards=cards_for(self.user_map.get(self.user_var.get()))
        if not cards: return
        card, conf, why = detect_card(text, cards)
        if not card:
            self.hint_var.set(f'Card not identified: {why}'); return
        want=q(card['payment_type']).strip() or PAYMENT_TYPES[0]
        cur=self.vars['payment_type'].get()
        label=q(card['label']).strip() or f"{q(card['network'])} {q(card['last4'])}".strip()
        if cur==want:
            messagebox.showinfo('Card recognised', f'{why}.\n\nPayment type is already "{want}" — nothing to change.', parent=self); return
        personal = want==PAYMENT_TYPES[1]
        msg=(f'{why}.\n\nThat is your "{label}" card, which you have set up as:\n    {want}\n\n'
             + ('This looks like money you paid yourself and need claiming back.\n\n' if personal else '')
             + f'Change the payment type from "{cur}" to "{want}"?')
        if messagebox.askyesno('Card recognised', msg, parent=self):
            self.vars['payment_type'].set(want); self._concur_hints()

class ReportDialog(tk.Toplevel):
    def __init__(self, master, rep_id=None, default_user_id=None):
        super().__init__(master); self.rep_id=rep_id; self.title('Expense Report'); self.transient(master); self.grab_set()
        data=dict(S.row('SELECT * FROM reports WHERE id=?',(rep_id,))) if rep_id else {}; self.vars={}; self.user_map={r['name']:r['id'] for r in S.rows('SELECT * FROM users ORDER BY name')}
        user_name=next((n for n,i in self.user_map.items() if i==data.get('user_id')), None) or next((n for n,i in self.user_map.items() if i==default_user_id), None) or list(self.user_map)[0]
        self.user_var=tk.StringVar(value=user_name)
        btn=FlowBar(self,padding=4); btn.pack(side='bottom',fill='x'); btn.button('Save',self.save); btn.button('Oracle Codes',self.pick_oracle); btn.button('Close',self.close)
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
    def pick_oracle(self):
        """Open the oracle list as a PICKER and write the chosen code into this report.

        Two things this has to get right, both of which the old one-liner did not:
        it opened the glossary parented to the MAIN WINDOW with no way to return a value, so the codes
        were there to look at and impossible to actually use; and this dialog holds a modal grab, so a
        child window opened over it takes no clicks until the grab is handed across and back."""
        def done(vals):
            # (code, name, favorite, notes) — same convention as OraclePicker: the code, or the bare
            # name for an entry that has none (e.g. "Unknown").
            self.vars['oracle_alias'].set(q(vals[0]) or q(vals[1]))
        try: self.grab_release()
        except Exception: pass
        g=OracleGlossary(self, on_pick=done)
        g.transient(self)
        try: g.grab_set()
        except Exception: pass
        self.wait_window(g)
        try: self.grab_set()  # take the modal grab back, or the report dialog goes dead
        except Exception: pass
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
    """Shared editor for the generated lists (expense codes, oracle codes, vendors): edit notes, toggle favorite, delete.

    With `on_pick` it is also a PICKER: opened from a form that wants one of these values, the list has to
    hand a row back, not just let you curate it. Without it, double-click still means "edit notes" — the
    management use (More ▾ → Glossaries) has no form to answer to."""
    def __init__(self, master, title, columns, load_fn, key_index, fav_col, notes_col, delete_fn,
                 on_pick=None, pick_label='Use this'):
        super().__init__(master); self.title(title); self.geometry('780x520'); self.minsize(640, 400)
        self.load_fn=load_fn; self.key_index=key_index; self.fav_col=fav_col; self.notes_col=notes_col; self.delete_fn=delete_fn; self.cols=columns
        self.on_pick=on_pick
        top=ttk.Frame(self,padding=8); top.pack(fill='x')
        self.search=tk.StringVar(); ttk.Entry(top,textvariable=self.search).pack(side='left',fill='x',expand=True); ttk.Button(top,text='Search',command=self.refresh).pack(side='left')
        self.bar=FlowBar(self,padding=(8,2)); self.bar.pack(fill='x')
        # In pick mode the picking action leads the bar: it is why the window was opened.
        if on_pick: Tooltip(self.bar.button(pick_label,self.pick), 'Put the selected row into the form you came from. Double-click a row (or press Enter) does the same.')
        self.bar.button('Toggle Favorite',self.toggle_fav); self.bar.button('Edit Notes',self.edit_notes); self.bar.button('Delete',self.delete)
        self.tree=ttk.Treeview(self,columns=columns,show='headings')
        for c in columns: self.tree.heading(c,text=c.title()); self.tree.column(c,width=130)
        self.tree.pack(fill='both',expand=True)
        self.tree.bind('<Double-1>', (lambda e: self.pick()) if on_pick else (lambda e: self.edit_notes()))
        if on_pick:
            self.bind('<Return>', lambda e: self.pick())
            ttk.Label(self,text=f'Double-click a row (or select it and press {pick_label}) to put it into the form.',
                      padding=(8,2),foreground='#555').pack(anchor='w',side='bottom')
        self.bind('<Escape>', lambda e: self.destroy()); self.refresh()
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for vals in self.load_fn(self.search.get().strip()): self.tree.insert('', 'end', values=vals)
    def _sel(self):
        item=self.tree.focus(); vals=self.tree.item(item,'values'); return vals or None
    def pick(self):
        """Hand the selected row back to whatever opened this, then close."""
        vals=self._sel()
        if not vals:
            messagebox.showinfo('Pick a row','Select a row first.',parent=self); return
        self.on_pick(vals); self.destroy()
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
    def __init__(self, master, on_pick=None):
        super().__init__(master, 'Oracle / Account Codes', ('code','name','favorite','notes'),
                         load_fn=self._load, key_index=1, fav_col=2, notes_col=3, delete_fn=lambda vals: S.execute('DELETE FROM oracle_codes WHERE name=?',(vals[1],)),
                         on_pick=on_pick, pick_label='Use this code')
        self.bar.button('Add Oracle Code', self.add)
    def _load(self, term):
        t='%'+term+'%'; return [(r['code'],r['name'],'★' if r['favorite'] else '',r['notes']) for r in S.rows('SELECT * FROM oracle_codes WHERE code LIKE ? OR name LIKE ? ORDER BY favorite DESC, name',(t,t))]
    def _update(self, vals, favorite=None, notes=None):
        if favorite is not None: S.execute('UPDATE oracle_codes SET favorite=? WHERE name=?',(favorite,vals[1]))
        if notes is not None: S.execute('UPDATE oracle_codes SET notes=? WHERE name=?',(notes,vals[1]))
    def add(self): OracleEditor(self, self.refresh)

class OracleEditor(tk.Toplevel):
    """Add one org/oracle alias: code and name together in a single window.

    This replaced a chain of two prompt boxes opened off the glossary. Chained modals are how you end up
    with input grabbed by a window that is behind another one, which reads to the user as a freeze."""
    def __init__(self, master, on_saved):
        super().__init__(master); self.title('Add Oracle Code'); self.transient(master); self.grab_set()
        self.on_saved=on_saved
        frm=ttk.Frame(self,padding=10); frm.pack(fill='both',expand=True); frm.columnconfigure(1,weight=1)
        self.code=tk.StringVar(); self.name=tk.StringVar(); self.notes=tk.StringVar()
        for i,(label,var,hint) in enumerate([('Code',self.code,'the number Concur shows in brackets'),
                                             ('Name',self.name,'what it is called on the report'),
                                             ('Notes',self.notes,'when you use it (optional)')]):
            ttk.Label(frm,text=label).grid(row=i,column=0,sticky='w',pady=3)
            ttk.Entry(frm,textvariable=var).grid(row=i,column=1,sticky='ew')
            ttk.Label(frm,text=hint,foreground='#777').grid(row=i,column=2,sticky='w',padx=(6,0))
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        bar.button('Save',self.save); bar.button('Cancel',self.destroy)
        self.bind('<Escape>', lambda e: self.destroy()); self.bind('<Return>', lambda e: self.save())
        fit_to_screen(self, min_w=520, min_h=220)
    def save(self):
        code=self.code.get().strip(); name=self.name.get().strip()
        if not code and not name: messagebox.showerror('Add Oracle Code','Give it a code or a name.',parent=self); return
        try: S.execute('INSERT OR IGNORE INTO oracle_codes(code,name,notes) VALUES(?,?,?)',(code,name or code,self.notes.get().strip()))
        except Exception as e: messagebox.showerror('Add failed', str(e), parent=self); return
        self.on_saved(); self.destroy()

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
            if f=='expense_type_code': CodePicker(frm, self.vars['expense_type_code'], self.vars['expense_type_label'], vendor_var=self.vars.get('vendor')).grid(row=row,column=1,sticky='ew',pady=2)
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
    or a list of them (bulk assign from a multi-selection).

    The list is SPLIT: reports still open (not yet submitted in Concur) sit at the top, already-filed ones
    below a heading. A flat A-Z list buried the three reports he can actually still add to among a year of
    dead ones — the whole point of picking from a list is not having to remember which is which."""
    def __init__(self, master, exp_id, user_id, on_done):
        super().__init__(master); self.title('Assign to Report'); self.geometry('460x400'); self.minsize(380, 320)
        self.transient(master); self.grab_set(); self.on_done=on_done
        self.exp_ids=list(exp_id) if isinstance(exp_id,(list,tuple,set)) else [exp_id]
        self.reps=S.rows('SELECT id,name,status FROM reports WHERE user_id=? ORDER BY name',(user_id,))
        openr=[r for r in self.reps if q(r['status'])!='Filed']; filed=[r for r in self.reps if q(r['status'])=='Filed']
        n=len(self.exp_ids)
        ttk.Label(self,text=(f'Assign these {n} expenses to:' if n!=1 else 'Assign this expense to:'),padding=8).pack(anchor='w')
        self.lb=tk.Listbox(self); self.lb.pack(fill='both',expand=True,padx=8)
        # `entries` runs line-for-line with the listbox: a report id, None for <Unassigned>, or the HEADING
        # sentinel for the two grey divider lines (which are not a choice you can make).
        self.entries=[]
        def add_line(text, entry, grey=False):
            self.lb.insert('end', text); self.entries.append(entry)
            if grey: self.lb.itemconfig(self.lb.size()-1, foreground='#7a7a7a')
        add_line('<Unassigned>', None)
        if openr:
            add_line('— Open reports (not filed yet) —', self.HEADING, grey=True)
            for r in openr: add_line(r['name'], r['id'])
        if filed:
            add_line('— Already filed in Concur —', self.HEADING, grey=True)
            for r in filed: add_line(r['name'], r['id'], grey=True)
        if not self.reps: add_line('(no reports yet for this user)', self.HEADING, grey=True)
        ttk.Label(self,text='Open reports are the ones you can still add to. Filed ones are already submitted —\n'
                            'assigning to one is allowed, but it will not change what Concur already has.',
                  padding=(8,4),foreground='#555',justify='left').pack(anchor='w')
        btn=FlowBar(self,padding=8); btn.pack(side='bottom',fill='x'); btn.button('Assign',self.assign); btn.button('Cancel',self.destroy)
        self.lb.bind('<Double-1>', lambda e: self.assign()); self.bind('<Escape>', lambda e: self.destroy())
    HEADING = object()  # marks a divider line: shown, greyed, and never assignable
    def assign(self):
        sel=self.lb.curselection()
        if not sel: return
        rid=self.entries[sel[0]]
        if rid is self.HEADING:
            messagebox.showinfo('Assign to Report','That line is a heading — pick one of the reports under it.',parent=self); return
        for eid in self.exp_ids: S.execute('UPDATE expenses SET report_id=? WHERE id=?',(rid,eid))
        if rid:
            try: collect_report_receipts(rid)  # keep the report's folder of receipts current
            except Exception: pass
        self.on_done(); self.destroy()

def receipt_request_text(rows, include_total=True):
    """Plain text for chasing receipts by email. Deliberately dumb formatting: no tabs, no aligned columns,
    no markdown — those all fall apart when pasted into Gmail or Outlook. One expense per line, in the only
    three facts the other person needs to FIND the receipt (vendor, amount, date). Nothing about Concur."""
    def when(d):
        try: return datetime.strptime(q(d)[:10], '%Y-%m-%d').strftime('%b %-d, %Y')
        except Exception:
            try: return datetime.strptime(q(d)[:10], '%Y-%m-%d').strftime('%b %d, %Y').replace(' 0', ' ')
            except Exception: return q(d)
    lines=[f"• {q(r['vendor']).strip() or '(no vendor)'} — ${money_part(r['amount'])} — {when(r['transaction_date'])}"
           for r in rows]
    if include_total and rows:
        total=sum(abs(float(r['amount'] or 0)) for r in rows)
        lines.append('')
        lines.append(f"{len(rows)} receipt{'' if len(rows)==1 else 's'}, ${total:,.2f} total")
    return '\n'.join(lines)

class AttendeeDialog(tk.Toplevel):
    """Your own address book of attendees, and the bridge to Concur's import.

    Concur's version makes you search one person at a time and is fussy about what it will take. This one
    accepts a paste of almost anything, tidies it, and hands back a file Concur's importer will swallow."""
    TYPE_LABELS = dict(ATTENDEE_TYPES)
    def __init__(self, master, on_pick=None, preselect=''):
        super().__init__(master); self.title('Attendees'); self.transient(master); self.grab_set()
        self.on_pick=on_pick
        top=ttk.Frame(self,padding=(10,8,10,2)); top.pack(fill='x')
        ttk.Label(top,text='Search').pack(side='left')
        self.term=tk.StringVar(); e=ttk.Entry(top,textvariable=self.term,width=28); e.pack(side='left',padx=6)
        self.term.trace_add('write', lambda *a: self.refresh())
        ttk.Label(top,text='Ticked people go onto the expense and into the Concur file.',foreground='#555').pack(side='left',padx=8)
        self.tree=ttk.Treeview(self,columns=('name','email','type','company'),show='headings',selectmode='extended')
        for c,w,t in [('name',190,'Name'),('email',210,'Email'),('type',140,'Attendee type'),('company',150,'Company')]:
            self.tree.heading(c,text=t); self.tree.column(c,width=w)
        self.tree.pack(fill='both',expand=True,padx=10,pady=6)
        bar=FlowBar(self,padding=(10,4)); bar.pack(fill='x')
        pick=ttk.LabelFrame(bar,text=' Ticked people ',padding=(6,2))
        for txt,cmd,tip in [('Use on this expense',self.use,'Put the ticked people on the expense you came from.'),
                            ('Export for Concur…',self.export,'Write the Concur attendee-import spreadsheet for the ticked people, ready to upload under Add Attendees → Import Attendees.'),
                            ('Copy emails',self.copy_emails,'Copy their addresses, tidied to the domain Concur holds — for finding employees in Concur\'s own search.')]:
            b=ttk.Button(pick,text=txt,command=cmd); b.pack(side='left',padx=2); Tooltip(b,tip)
        bar.attach(pick)
        book=ttk.LabelFrame(bar,text=' Address book ',padding=(6,2))
        for txt,cmd in [('Paste people…',self.paste_import),('Edit…',self.edit),('Delete',self.delete),('Close',self.destroy)]:
            ttk.Button(book,text=txt,command=cmd).pack(side='left',padx=2)
        bar.attach(book)
        self.bind('<Escape>', lambda e: self.destroy()); self.tree.bind('<Double-1>', lambda e: self.edit())
        self.refresh()
        if preselect: self.preselect(preselect)
        fit_to_screen(self, min_w=760, min_h=420)
    def refresh(self):
        keep={self.tree.item(i,'values')[1] for i in self.tree.selection()}
        self.tree.delete(*self.tree.get_children())
        t=f"%{self.term.get().strip()}%"
        rows=S.rows('SELECT * FROM contacts WHERE last_name LIKE ? OR first_name LIKE ? OR email LIKE ? OR company LIKE ? '
                    'ORDER BY use_count DESC, last_name, first_name',(t,t,t,t))
        for r in rows:
            self.tree.insert('','end',iid=str(r['id']),
                             values=(f"{q(r['last_name'])}, {q(r['first_name'])}".strip(', '), q(r['email']),
                                     self.TYPE_LABELS.get(q(r['atn_type']), q(r['atn_type'])), q(r['company'])))
        if keep:
            back=[i for i in self.tree.get_children() if self.tree.item(i,'values')[1] in keep]
            if back: self.tree.selection_set(*back)
    def preselect(self, names):
        want={n.strip().lower() for n in q(names).replace(';','\n').split('\n') if n.strip()}
        hit=[i for i in self.tree.get_children()
             if self.tree.item(i,'values')[0].lower() in want or self.tree.item(i,'values')[1].lower() in want]
        if hit: self.tree.selection_set(*hit)
    def chosen(self):
        return [S.row('SELECT * FROM contacts WHERE id=?',(int(i),)) for i in self.tree.selection()]
    def paste_import(self): ContactPasteDialog(self, self.refresh)
    def edit(self):
        f=self.tree.focus()
        if not f: messagebox.showinfo('Attendees','Pick someone first.',parent=self); return
        ContactEditor(self, S.row('SELECT * FROM contacts WHERE id=?',(int(f),)), self.refresh)
    def delete(self):
        rows=self.chosen()
        if rows and messagebox.askyesno('Delete', f'Remove {len(rows)} person(s) from your address book?', parent=self):
            for r in rows: S.execute('DELETE FROM contacts WHERE id=?',(r['id'],))
            self.refresh()
    def _need(self):
        rows=self.chosen()
        if not rows: messagebox.showinfo('Attendees','Tick the people you want first.',parent=self)
        return rows
    def use(self):
        rows=self._need()
        if not rows: return
        for r in rows: S.execute('UPDATE contacts SET use_count=use_count+1 WHERE id=?',(r['id'],))
        if self.on_pick: self.on_pick(rows)
        self.destroy()
    def copy_emails(self):
        rows=self._need()
        if not rows: return
        body='; '.join(q(r['email']) for r in rows if q(r['email']).strip())
        self.clipboard_clear(); self.clipboard_append(body); self.update()
        messagebox.showinfo('Copied', f'{len(body.split(";")) if body else 0} address(es) copied.\n\n{body}', parent=self)
    def export(self):
        rows=self._need()
        if not rows: return
        p=filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile='ConcurAttendees.xlsx',
                                       filetypes=[('Excel','*.xlsx')], title='Save the Concur attendee import file', parent=self)
        if not p: return
        try: write_xlsx(p, attendee_rows(rows))
        except Exception as e: messagebox.showerror('Export failed', str(e), parent=self); return
        for r in rows: S.execute('UPDATE contacts SET use_count=use_count+1 WHERE id=?',(r['id'],))
        if messagebox.askyesno('Saved', f'{len(rows)} attendee(s) written to:\n{p}\n\nIn Concur: Add Attendees → '
                               'Import Attendees → Upload Spreadsheet.\n\nOpen the folder now?', parent=self):
            open_file_location(p)

class ContactPasteDialog(tk.Toplevel):
    """Paste people in from anywhere. Shows what it understood BEFORE anything is saved."""
    def __init__(self, master, on_saved):
        super().__init__(master); self.title('Paste people'); self.transient(master); self.grab_set()
        self.on_saved=on_saved
        ttk.Label(self,padding=(10,8,10,2),justify='left',text=
                  'Paste names, addresses, or both — one per line. All of these work:\n'
                  '    Chidi Okafor <c.okafor@mail.example.edu>   Whitfield, Dana; dwhitfield@example.edu; Example Co\n'
                  '    c.okafor@mail.example.edu                  Reyes⇥Marisol⇥mreyes@example.edu').pack(anchor='w')
        self.txt=tk.Text(self,height=9,width=68,wrap='none'); self.txt.pack(fill='both',expand=True,padx=10,pady=4)
        row=ttk.Frame(self,padding=(10,2)); row.pack(fill='x')
        ttk.Label(row,text='Treat them as').pack(side='left')
        self.kind=tk.StringVar(value='Employee')
        ttk.Combobox(row,textvariable=self.kind,values=[lbl for _,lbl in ATTENDEE_TYPES],state='readonly',width=22).pack(side='left',padx=6)
        self.preview=ttk.Label(self,text='',padding=(10,2),foreground='#555',justify='left',wraplength=560)
        self.preview.pack(anchor='w')
        self.txt.bind('<KeyRelease>', lambda e: self.show_preview())
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        bar.button('Add them',self.save); bar.button('Cancel',self.destroy)
        self.bind('<Escape>', lambda e: self.destroy()); self.txt.focus_set()
        fit_to_screen(self, min_w=600, min_h=380)
    def _code(self): return next((c for c,l in ATTENDEE_TYPES if l==self.kind.get()), 'SYSEMP')
    def parsed(self): return parse_contact_blob(self.txt.get('1.0','end-1c'), self._code())
    def show_preview(self):
        got=self.parsed()
        if not got: self.preview.configure(text=''); return
        shown=', '.join(f"{p['first_name']} {p['last_name']}".strip() + (f" <{p['email']}>" if p['email'] else '') for p in got[:4])
        self.preview.configure(text=f"Understood {len(got)} person(s): {shown}" + (' …' if len(got)>4 else ''))
    def save(self):
        got=self.parsed()
        if not got: messagebox.showinfo('Paste people','Nothing recognisable in there yet.',parent=self); return
        added=updated=0
        for c in got:
            _id, is_new = S.upsert_contact(c)
            added += 1 if is_new else 0; updated += 0 if is_new else 1
        self.on_saved()
        messagebox.showinfo('Added', f'{added} new, {updated} already known.', parent=self)
        self.destroy()

class ContactEditor(tk.Toplevel):
    """One person. Email is tidied to the domain Concur holds as you save."""
    def __init__(self, master, row, on_saved):
        super().__init__(master); self.title('Attendee'); self.transient(master); self.grab_set()
        self.row=row; self.on_saved=on_saved
        frm=ttk.Frame(self,padding=10); frm.pack(fill='both',expand=True); frm.columnconfigure(1,weight=1)
        self.v={}
        for i,(f,label) in enumerate([('last_name','Last name'),('first_name','First name'),('email','Email'),
                                      ('company','Company'),('title','Attendee title')]):
            ttk.Label(frm,text=label).grid(row=i,column=0,sticky='w',pady=3)
            self.v[f]=tk.StringVar(value=q(row[f]) if row else '')
            ttk.Entry(frm,textvariable=self.v[f]).grid(row=i,column=1,sticky='ew')
        ttk.Label(frm,text='Attendee type').grid(row=5,column=0,sticky='w',pady=3)
        cur=dict(ATTENDEE_TYPES).get(q(row['atn_type']) if row else 'SYSEMP','Employee')
        self.kind=tk.StringVar(value=cur)
        ttk.Combobox(frm,textvariable=self.kind,values=[l for _,l in ATTENDEE_TYPES],state='readonly').grid(row=5,column=1,sticky='ew')
        ttk.Label(frm,text='Only the type, last name and first name reach Concur\'s import file.',
                  foreground='#555').grid(row=6,column=1,sticky='w',pady=(4,0))
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        bar.button('Save',self.save); bar.button('Cancel',self.destroy)
        self.bind('<Escape>', lambda e: self.destroy()); fit_to_screen(self, min_w=480, min_h=300)
    def save(self):
        last=self.v['last_name'].get().strip(); first=self.v['first_name'].get().strip()
        if not last and not first: messagebox.showerror('Attendee','A name is required.',parent=self); return
        code=next((c for c,l in ATTENDEE_TYPES if l==self.kind.get()),'SYSEMP')
        vals=(last, first, normalize_email(self.v['email'].get()), self.v['company'].get().strip(),
              self.v['title'].get().strip(), code)
        if self.row: S.execute('UPDATE contacts SET last_name=?,first_name=?,email=?,company=?,title=?,atn_type=? WHERE id=?', vals+(self.row['id'],))
        else: S.upsert_contact(dict(zip(('last_name','first_name','email','company','title','atn_type'), vals)))
        self.on_saved(); self.destroy()

class CardsDialog(tk.Toplevel):
    """The cards you might pay with, so OCR can tell whose money it was.

    Each row is last four + network + which payment type that card means. Both halves matter: matching a
    receipt on the digits alone snags invoice numbers, on the network alone snags every payment footer."""
    def __init__(self, master, user_id=None):
        super().__init__(master); self.title('My cards'); self.transient(master); self.grab_set()
        self.user_id=user_id if user_id is not None else getattr(master,'current_user_id',lambda: None)()
        ttk.Label(self,text='Cards you pay with. When you OCR a receipt, Concur Buddy looks for BOTH the network\n'
                            'and the last four before deciding which card it was.',padding=(10,8,10,4)).pack(anchor='w')
        self.tree=ttk.Treeview(self,columns=('last4','network','pay','label','who'),show='headings',height=6)
        for c,w,t in [('last4',66,'Last 4'),('network',100,'Network'),('pay',240,'Means this payment type'),('label',140,'Your name for it'),('who',110,'Whose card')]:
            self.tree.heading(c,text=t); self.tree.column(c,width=w)
        self.tree.pack(fill='both',expand=True,padx=10,pady=4); self.tree.bind('<Double-1>', lambda e: self.edit())
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        for txt,cmd in [('Add…',self.add),('Edit…',self.edit),('Delete',self.delete),('Close',self.destroy)]: bar.button(txt,cmd)
        self.bind('<Escape>', lambda e: self.destroy())
        self.refresh(); fit_to_screen(self, min_w=620, min_h=320)
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        names={i:n for n,i in {r['name']:r['id'] for r in S.rows('SELECT id,name FROM users')}.items()}
        for r in cards_for(self.user_id):
            self.tree.insert('','end',iid=str(r['id']),values=(r['last4'],r['network'],r['payment_type'],r['label'],
                                                               names.get(r['user_id'],'everyone')))
    def _sel(self):
        f=self.tree.focus(); return S.row('SELECT * FROM cards WHERE id=?',(int(f),)) if f else None
    def add(self): CardEditor(self, None, self.refresh, self.user_id)
    def edit(self):
        r=self._sel()
        if r: CardEditor(self, r, self.refresh, self.user_id)
        else: messagebox.showinfo('My cards','Pick a card first.',parent=self)
    def delete(self):
        r=self._sel()
        if r and messagebox.askyesno('Delete card', f"Forget the card ending {r['last4']}?", parent=self):
            S.execute('DELETE FROM cards WHERE id=?',(r['id'],)); self.refresh()

class CardEditor(tk.Toplevel):
    """Add or change one card."""
    def __init__(self, master, row, on_saved, user_id=None):
        super().__init__(master); self.title('Card'); self.transient(master); self.grab_set()
        self.row=row; self.on_saved=on_saved
        self.users={r['name']:r['id'] for r in S.rows('SELECT id,name FROM users ORDER BY name')}
        owner=next((n for n,i in self.users.items() if i==(row['user_id'] if row else user_id)), 'everyone')
        frm=ttk.Frame(self,padding=10); frm.pack(fill='both',expand=True); frm.columnconfigure(1,weight=1)
        get=lambda k,d='': q(row[k]) if row else d
        self.last4=tk.StringVar(value=get('last4')); self.net=tk.StringVar(value=get('network','Visa'))
        self.pay=tk.StringVar(value=get('payment_type',PAYMENT_TYPES[1])); self.label=tk.StringVar(value=get('label'))
        ttk.Label(frm,text='Last 4 digits').grid(row=0,column=0,sticky='w',pady=3)
        ttk.Entry(frm,textvariable=self.last4,width=8).grid(row=0,column=1,sticky='w')
        ttk.Label(frm,text='Network').grid(row=1,column=0,sticky='w',pady=3)
        ttk.Combobox(frm,textvariable=self.net,values=list(CARD_NETWORKS),state='readonly',width=16).grid(row=1,column=1,sticky='w')
        ttk.Label(frm,text='Means this payment type').grid(row=2,column=0,sticky='w',pady=3)
        ttk.Combobox(frm,textvariable=self.pay,values=PAYMENT_TYPES,state='readonly').grid(row=2,column=1,sticky='ew')
        ttk.Label(frm,text='Your name for it').grid(row=3,column=0,sticky='w',pady=3)
        ttk.Entry(frm,textvariable=self.label).grid(row=3,column=1,sticky='ew')
        ttk.Label(frm,text='Whose card').grid(row=4,column=0,sticky='w',pady=3)
        self.owner=tk.StringVar(value=owner)
        ttk.Combobox(frm,textvariable=self.owner,values=['everyone']+list(self.users),state='readonly').grid(row=4,column=1,sticky='ew')
        ttk.Label(frm,text='e.g. "personal Visa" — only ever shown to you. Two people\'s cards are separate:\n'
                           'a card set to one user is never matched against another user\'s receipt.',
                  foreground='#555',justify='left').grid(row=5,column=1,sticky='w',pady=(4,0))
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        bar.button('Save',self.save); bar.button('Cancel',self.destroy)
        self.bind('<Escape>', lambda e: self.destroy()); fit_to_screen(self, min_w=460, min_h=260)
    def save(self):
        d=re.sub(r'\D','',self.last4.get())[-4:]
        if len(d)!=4: messagebox.showerror('Card','Enter the last 4 digits of the card.',parent=self); return
        uid=self.users.get(self.owner.get())  # None = everyone's
        if self.row: S.execute('UPDATE cards SET last4=?,network=?,payment_type=?,label=?,user_id=? WHERE id=?',
                               (d,self.net.get(),self.pay.get(),self.label.get().strip(),uid,self.row['id']))
        else:
            if S.row('SELECT 1 FROM cards WHERE last4=? AND network=? AND (user_id IS ? OR user_id=?)',(d,self.net.get(),uid,uid)):
                messagebox.showerror('Card',f'A {self.net.get()} ending {d} is already set up for {self.owner.get()}.',parent=self); return
            S.execute('INSERT INTO cards(last4,network,payment_type,label,user_id) VALUES(?,?,?,?,?)',
                      (d,self.net.get(),self.pay.get(),self.label.get().strip(),uid))
        self.on_saved(); self.destroy()

class ReceiptRequestDialog(tk.Toplevel):
    """The 'please send me these receipts' list, ready to paste into an email.

    Deliberately NOT an email client: it hands over plain text and gets out of the way, because the message
    around it is his to write and every mail app mangles rich formatting differently."""
    def __init__(self, master, rows):
        super().__init__(master); self.title('Receipt request'); self.transient(master); self.grab_set()
        ttk.Label(self,text=f"{len(rows)} expense{'' if len(rows)==1 else 's'} — plain text, safe to paste "
                            "straight into Gmail or Outlook.",padding=(10,8,10,2)).pack(anchor='w')
        self.txt=tk.Text(self,height=12,width=64,wrap='word'); self.txt.pack(fill='both',expand=True,padx=10,pady=4)
        self.txt.insert('1.0', receipt_request_text(rows))
        ttk.Label(self,text='Edit it here first if you like — Copy takes whatever is in the box.',
                  padding=(10,0),foreground='#555').pack(anchor='w')
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        Tooltip(bar.button('Copy to clipboard',self.copy),'Puts the text on the clipboard, ready to paste into an email.')
        bar.button('Select all',self.select_all); bar.button('Close',self.destroy)
        self.bind('<Escape>', lambda e: self.destroy())
        self.txt.focus_set(); fit_to_screen(self, min_w=520, min_h=340)
    def select_all(self):
        self.txt.tag_add('sel','1.0','end-1c'); self.txt.focus_set()
    def copy(self):
        body=self.txt.get('1.0','end-1c')
        self.clipboard_clear(); self.clipboard_append(body); self.update()
        try: self.master.flash_status('Receipt list copied to clipboard')
        except Exception: pass
        self.destroy()

class ImportRowDialog(tk.Toplevel):
    """Review ONE imported row: pick which staged expense it is, then choose a side for EVERY field that
    differs — one row per field, in a plain comparison table. Nothing is forced: the card-fed fields just
    default to the export (Concur greys them out, so its copy is the one that matches the card statement)."""
    WRAP = 165  # value-column width; every cell wraps inside it so the table can't overflow sideways
    def __init__(self, master, plan, on_done=None):
        super().__init__(master); self.title('Review imported expense'); self.transient(master); self.grab_set()
        self.plan=plan; self.on_done=on_done; self.row=plan['row']; self.choice_vars={}
        self.bold=tkfont.nametofont('TkDefaultFont').copy(); self.bold.configure(weight='bold')
        r=self.row
        bits=[q(r.get('transaction_date')), q(r.get('vendor')), money_part(r.get('amount'))]
        if r.get('expense_type_code'): bits.append(q(r.get('expense_type_code')))
        head=ttk.Frame(self,padding=(10,8,10,2)); head.pack(fill='x')
        ttk.Label(head,text='Expense from the export',foreground='#555').pack(anchor='w')
        ttk.Label(head,text='   '.join(bits),font=self.bold,wraplength=560).pack(anchor='w')
        pick=ttk.Frame(self,padding=(10,6,10,0)); pick.pack(fill='x')
        ttk.Label(pick,text='Update which expense?',font=self.bold).pack(anchor='w')
        self.lb=tk.Listbox(pick,height=4,exportselection=False); self.lb.pack(fill='x',pady=(2,0))
        self.lb.insert('end','None of these — add it as a new expense')
        for e,gap,sim in plan['candidates']:
            self.lb.insert('end', f"#{e['id']}   {q(e['transaction_date'])}   {q(e['vendor'])}   {money_part(e['amount'])}"
                                  f"   ({gap} day{'' if gap==1 else 's'} apart, {int(sim*100)}% name match)")
        self.lb.selection_set(0 if plan['action']!='merge' else 1+plan['cand_index'])
        self.lb.bind('<<ListboxSelect>>', lambda e: self.build_fields())
        self.fields=ttk.Frame(self,padding=(10,8)); self.fields.pack(fill='both',expand=True)
        bar=FlowBar(self,padding=8); bar.pack(side='bottom',fill='x')
        Tooltip(bar.button('Done',self.ok),'Keep these choices and go back to the list. Nothing is written until you press Apply there.')
        Tooltip(bar.button('Skip this row',self.skip),'Leave this export row out of the import entirely.')
        Tooltip(bar.button('Cancel',self.destroy),'Close without changing this row.')
        self.bind('<Escape>', lambda e: self.destroy())
        self.build_fields(); fit_to_screen(self, min_w=620, min_h=420)
    def _selected_candidate(self):
        sel=self.lb.curselection(); i=(sel[0] if sel else 0)-1
        return i if 0<=i<len(self.plan['candidates']) else None
    def build_fields(self):
        for w in self.fields.winfo_children(): w.destroy()
        self.choice_vars={}
        i=self._selected_candidate()
        if i is None:
            ttk.Label(self.fields,text='It will be added as a new expense, exactly as the export has it.').grid(row=0,column=0,sticky='w')
            return
        rows=plan_rows(self.row, self.plan['candidates'][i][0])
        if not rows:
            ttk.Label(self.fields,text='Every field already agrees — updating changes nothing, it just ties the two together.').grid(row=0,column=0,sticky='w')
            return
        decide=[r for r in rows if r[4]=='both']
        banner=tk.Label(self.fields,padx=8,pady=4,anchor='w',
                        text=(f"  {len(decide)} field{'' if len(decide)==1 else 's'} need{'s' if len(decide)==1 else ''} you to pick a winner"
                              if decide else '  Nothing to decide — every field follows the rules below'),
                        background='#fff3cd' if decide else '#e6f4ea', foreground='#4a3a00' if decide else '#14532d')
        banner.grid(row=0,column=0,columnspan=5,sticky='ew',pady=(0,6))
        for c,w in enumerate((62, 124, self.WRAP, self.WRAP, 122)): self.fields.columnconfigure(c,minsize=w)
        for c,t in enumerate(('','Field','In Concur Buddy','In this export','Use')):
            ttk.Label(self.fields,text=t,font=self.bold).grid(row=1,column=c,sticky='w',padx=(0,10),pady=(0,2))
        ttk.Separator(self.fields,orient='horizontal').grid(row=2,column=0,columnspan=5,sticky='ew',pady=(0,4))
        # One chip per row says what KIND of row it is, so "what do I actually have to do here" is a glance,
        # not a read. Only DECIDE rows get a control; the other two are already settled by the rules.
        chips={'both':('DECIDE','#fff3cd','#4a3a00'),'card':('LOCKED','#f1f3f4','#5f6368'),'blank':('FILL','#e8f0fe','#1a3d7c')}
        for n,(f,old_v,new_v,default,why) in enumerate(rows, start=3):
            label,bg,fg=chips[why]
            tk.Label(self.fields,text=label,background=bg,foreground=fg,font=self.bold,
                     padx=5,relief='solid',borderwidth=1).grid(row=n,column=0,sticky='nw',pady=3)
            ttk.Label(self.fields,text=FIELD_LABELS.get(f,f),wraplength=120,
                      font=self.bold if why=='both' else None).grid(row=n,column=1,sticky='nw',padx=(6,10),pady=3)
            ttk.Label(self.fields,text=old_v or '—',wraplength=self.WRAP,foreground='#444').grid(row=n,column=2,sticky='nw',padx=(0,10),pady=3)
            ttk.Label(self.fields,text=new_v or '—',wraplength=self.WRAP,foreground='#444').grid(row=n,column=3,sticky='nw',padx=(0,10),pady=3)
            if why!='both':
                ttk.Label(self.fields,text='the export',foreground='#777').grid(row=n,column=4,sticky='nw',pady=3)
                continue
            v=tk.StringVar(value=self.plan['choices'].get(f,default)); self.choice_vars[f]=v
            box=ttk.Frame(self.fields); box.grid(row=n,column=4,sticky='nw',pady=2)
            ttk.Radiobutton(box,text='Mine',variable=v,value='buddy').pack(side='left')
            ttk.Radiobutton(box,text='Export',variable=v,value='concur').pack(side='left',padx=(6,0))
        end=len(rows)+3
        ttk.Separator(self.fields,orient='horizontal').grid(row=end,column=0,columnspan=5,sticky='ew',pady=(6,4))
        if self.choice_vars:
            allbar=ttk.Frame(self.fields); allbar.grid(row=end+1,column=0,columnspan=5,sticky='w')
            n=len(self.choice_vars)
            ttk.Label(allbar,text=('This choice:' if n==1 else f'All {n} choices:')).pack(side='left',padx=(0,6))
            ttk.Button(allbar,text='Take the export',width=16,command=lambda:self.set_all('concur')).pack(side='left')
            ttk.Button(allbar,text='Keep mine',width=12,command=lambda:self.set_all('buddy')).pack(side='left',padx=4)
        note=('LOCKED = the card feed owns it; Concur lets neither side edit it, so the export wins.    '
              'FILL = blank here, so the export fills it in.    DECIDE = both sides have a value.')
        ttk.Label(self.fields,text=note,wraplength=560,foreground='#555').grid(row=end+2,column=0,columnspan=5,sticky='w',pady=(6,0))
    def set_all(self, side):
        for v in self.choice_vars.values(): v.set(side)
    def ok(self):
        i=self._selected_candidate()
        if i is None: self.plan['action']='new'; self.plan['cand_index']=None
        else:
            self.plan['action']='merge'; self.plan['cand_index']=i
            self.plan['choices']={f:v.get() for f,v in self.choice_vars.items()}
        self.plan['reviewed']=True  # ticks the ✓ back on the list so you can see what you've already been through
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
                  'action':'merge' if propose else 'new','choices':{},'reviewed':False}
            if propose: taken.add(cands[0][0]['id'])  # one staged expense can only absorb one export row
            self.plans.append(plan)
        top=ttk.Frame(self,padding=(8,6)); top.pack(fill='x')
        ttk.Label(top,text=f"{len(rows)} expenses in {Path(path).name}   ·   matching against "
                           f"{len(staged)} staged for this user").pack(anchor='w')
        # The report row. Two controls, because there are two different jobs: NAME the report (Concur's own
        # name is the canonical one, so it is prefilled from the export) and, optionally, say that this import
        # belongs to a report already staged here under a name he typed himself before Concur had one.
        rep=FlowBar(self,padding=(8,2)); rep.pack(fill='x')
        namebox=ttk.LabelFrame(rep,text=' Report name (from Concur) ',padding=(6,2))
        self.report_name=tk.StringVar(value=report_name)
        ttk.Entry(namebox,textvariable=self.report_name,width=30).pack(side='left')
        ttk.Label(namebox,text=' blank = leave them loose',foreground='#555').pack(side='left')
        rep.attach(namebox)
        mergebox=ttk.LabelFrame(rep,text=' Existing report to merge into ',padding=(6,2))
        self.existing=S.rows("SELECT id,name,status FROM reports WHERE user_id=? ORDER BY (status='Filed'), name",(user_id,))
        self.merge_choice=tk.StringVar(value=NEW_REPORT)
        self.merge_labels={report_choice_label(r):r['id'] for r in self.existing}
        cb=ttk.Combobox(mergebox,textvariable=self.merge_choice,width=26,state='readonly',
                        values=[NEW_REPORT]+list(self.merge_labels))
        cb.pack(side='left'); cb.bind('<<ComboboxSelected>>', lambda e: self._merge_note())
        Tooltip(cb,"Pick the report you already staged these under here. Its expenses stay put — this import "
                   "just joins it instead of creating a second report for the same trip.")
        self.rename_var=tk.IntVar(value=1)
        ttk.Checkbutton(mergebox,text="Rename it to Concur's name",variable=self.rename_var,
                        command=self._merge_note).pack(side='left',padx=(6,0))
        rep.attach(mergebox)
        # Retyping the name changes what the merge is about to do, so the note follows every keystroke.
        self.report_name.trace_add('write', lambda *a: self._merge_note())
        self.merge_note=tk.StringVar()
        ttk.Label(self,textvariable=self.merge_note,padding=(8,0),foreground='#1a3d7c',wraplength=880,justify='left').pack(anchor='w')
        self._merge_note()
        # Every button says what it acts on, and each carries a hover hint — the three middle ones only
        # retag the rows you have selected, which is not guessable from a bare verb.
        # Buttons live in LABELLED boxes, not behind hover text: which rows a button acts on has to be
        # readable at a glance, and a tooltip you have to discover isn't that. Tooltips stay as detail.
        bar=FlowBar(self,padding=(8,4)); bar.pack(fill='x')
        rowbox=ttk.LabelFrame(bar,text=' Selected row(s) ',padding=(6,2))
        for txt,cmd,tip in [('Review…',self.review,'Open the selected row: change which staged expense it updates, and settle any field the two sides disagree on. Double-clicking a row does the same.'),
                            ('Update existing',lambda:self.set_action('merge'),'The selected row is an expense you ALREADY have here — fold Concur\'s copy into it rather than ending up with two. Keeps your receipt, notes and status. No effect on a row with no amount match.'),
                            ('Add as new',lambda:self.set_action('new'),'You do NOT already have this one — create a fresh expense from it. Use this when the suggested match is wrong.'),
                            ('Ignore',lambda:self.set_action('skip'),'Leave the selected row(s) out entirely — nothing created, nothing changed.')]:
            b=ttk.Button(rowbox,text=txt,command=cmd); b.pack(side='left',padx=2); Tooltip(b,tip)
        bar.attach(rowbox)
        allbox=ttk.LabelFrame(bar,text=' Whole import ',padding=(6,2))
        for txt,cmd,tip in [('Apply',self.apply,'Carry out every row\'s "What will happen". This is the only step that writes anything.'),
                            ('Cancel',self.destroy,'Close and import nothing.')]:
            b=ttk.Button(allbox,text=txt,command=cmd); b.pack(side='left',padx=2); Tooltip(b,tip)
        bar.attach(allbox)
        self.tree=ttk.Treeview(self,columns=('ok','date','vendor','amount','type','action','detail'),show='headings',selectmode='extended')
        for c,w,t in [('ok',34,'✓'),('date',90,'Date'),('vendor',175,'Vendor (export)'),('amount',80,'Amount'),
                      ('type',185,'Expense type'),('action',180,'What will happen'),('detail',330,'Notes')]:
            self.tree.heading(c,text=t); self.tree.column(c,width=w,anchor='center' if c=='ok' else 'w')
        # Colour carries the one thing that matters at a glance: does this row still want something from you?
        self.tree.tag_configure('needs', background='#fff3cd')   # amber - a field is still undecided
        self.tree.tag_configure('ready', background='#e6f4ea')   # green - will merge, nothing left to settle
        self.tree.tag_configure('new',   background='#e8f0fe')   # blue  - becomes a new expense
        self.tree.tag_configure('skip',  background='#f1f3f4', foreground='#80868b')
        self.tree.pack(fill='both',expand=True,padx=8,pady=(6,2)); self.tree.bind('<Double-1>', lambda e: self.review())
        legend=ttk.Frame(self,padding=(8,0)); legend.pack(fill='x')
        for text,bg in [('needs a decision','#fff3cd'),('ready to merge','#e6f4ea'),('new expense','#e8f0fe'),('skipped','#f1f3f4')]:
            tk.Label(legend,text='   ',background=bg,relief='solid',borderwidth=1,width=3).pack(side='left',pady=1)
            ttk.Label(legend,text=f' {text}   ',foreground='#555').pack(side='left')
        hint=('Every row does one of three things:   UPDATE an expense you already have here (your receipt, notes '
              'and status stay) · ADD it as a new expense · IGNORE it.\n'
              'Nothing is written until you press Apply.   ✓ = you have been through that row.')
        if unknown: hint+=f"\nColumns with no field here are kept as notes on new expenses: {', '.join(unknown[:6])}"
        ttk.Label(self,text=hint,padding=(8,2),foreground='#555',wraplength=900).pack(anchor='w',side='bottom')
        self.bind('<Escape>', lambda e: self.destroy())
        self.refresh(); fit_to_screen(self, min_w=980, min_h=500)
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for i,p in enumerate(self.plans):
            r=p['row']; undecided=0
            # The Notes cell is the only place that says WHY a row proposes what it proposes, so it is written
            # as short sentences, not counters-and-jargon: "1 from card" told him nothing he could act on.
            plural=lambda n,one,many: f"{n} {one if n==1 else many}"
            if p['action']=='merge':
                e=p['candidates'][p['cand_index']][0]
                rows=plan_rows(r,e); undecided=sum(1 for f,_,_,_,why in rows if why=='both')
                act=f"Update #{e['id']} {q(e['vendor'])}"
                bits=[]
                if undecided: bits.append(plural(undecided,'field needs','fields need')+' your decision' if not p.get('reviewed')
                                          else 'you settled '+plural(undecided,'field','fields'))
                n_card=sum(1 for _,_,_,_,why in rows if why=='card'); n_fill=sum(1 for _,_,_,_,why in rows if why=='blank')
                if n_card: bits.append(plural(n_card,'field','fields')+" taken from Concur's card feed")
                if n_fill: bits.append(plural(n_fill,'blank','blanks')+' here will be filled in')
                if len(p['candidates'])>1: bits.append(f"{len(p['candidates'])} expenses here could be this one")
                if r.get('_receipt_in_concur') is False and (e['receipt_path'] or e['invoice_path']):
                    bits.append('you have the receipt, Concur does not')
                detail=' · '.join(bits) or 'both copies already agree — this just links them'
                tag='needs' if (undecided and not p.get('reviewed')) else 'ready'
            elif p['action']=='new':
                act='Add as a new expense'; tag='new'
                detail=('nothing staged here has this amount' if not p['candidates']
                        else plural(len(p['candidates']),'expense here has','expenses here have')+' this amount — Review to link one')
            else: act='Ignore — nothing happens'; detail='left out of this import'; tag='skip'
            etype=' '.join(x for x in (q(r.get('expense_type_code')), q(r.get('expense_type_label'))) if x)
            self.tree.insert('','end',iid=str(i),tags=(tag,),
                             values=('✓' if p.get('reviewed') else '', q(r.get('transaction_date')), q(r.get('vendor')),
                                     money_part(r.get('amount')), etype, act, detail))
    def _merge_target(self):
        """The existing report this import should join, or None when it should find/create one by name."""
        return self.merge_labels.get(self.merge_choice.get())
    def _merge_note(self):
        """Say in one sentence what the two report controls are about to do — a rename is destructive enough
        that it must never be a surprise discovered after Apply."""
        rid=self._merge_target(); name=self.report_name.get().strip()
        if rid is None:
            self.merge_note.set('' if name else 'These expenses will be left loose — no report.'); return
        cur=q(S.scalar('SELECT name FROM reports WHERE id=?',(rid,)))
        if name and self.rename_var.get() and name!=cur:
            note=f'Joining "{cur}" — and renaming it to "{name}", since Concur\'s name is the real one.'
            # Report names are not unique, so a rename CAN leave two reports called the same thing. Say so
            # rather than quietly making a pair he then has to tell apart in every picker.
            if S.row('SELECT 1 FROM reports WHERE user_id=? AND name=? AND id<>?',(self.user_id,name,rid)):
                note+=f'  ⚠ You already have another report called "{name}" — you would end up with two.'
            self.merge_note.set(note)
        else:
            self.merge_note.set(f'Joining "{cur}" — its name is kept.')
    def _sel(self): return [self.plans[int(i)] for i in self.tree.selection()]
    def set_action(self, action):
        for p in self._sel():
            if action=='merge' and not p['candidates']: continue  # nothing to merge into
            p['action']=action
            if action=='merge' and p['cand_index'] is None: p['cand_index']=0
            p['reviewed']=False  # retagging a row invalidates any earlier review of it
        self.refresh()
    def review(self):
        sel=self._sel()
        if not sel: messagebox.showinfo('Review','Select a row first.',parent=self); return
        ImportRowDialog(self, sel[0], on_done=self.refresh)
    def apply(self):
        name=self.report_name.get().strip(); rid=self._merge_target(); renamed=None
        if rid is not None:
            # Merging into a report already here. Concur's name is canonical, so it wins over whatever he
            # called it before the report existed in Concur — unless he unticks the rename box.
            was=q(S.scalar('SELECT name FROM reports WHERE id=?',(rid,)))
            if name and self.rename_var.get() and name!=was:
                S.execute('UPDATE reports SET name=? WHERE id=?',(name,rid)); renamed=(was,name)
            else: name=was
        elif name:
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
                vals={}
                for f,_,_,default,why in plan_rows(r,exp):
                    # Only a both-sides-filled row is the user's call; card-fed and blank rows follow the rule.
                    if why!='both' or p['choices'].get(f,default)=='concur': apply_field(vals, r, f)
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
        if renamed: msg+=f'\nRenamed report "{renamed[0]}" to "{renamed[1]}" (Concur\'s name).'
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
        ttk.Label(top,text='Filter').pack(side='left'); ttk.Combobox(top,textvariable=self.filter_col,values=['Any','vendor','amount','status','doc','type','purpose','date'],state='readonly',width=8).pack(side='left',padx=2); self.filter_col.trace_add('write',lambda *a:self.refresh())
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
               'Ask for Receipts':'Turn the ticked expenses (or a selected report) into a plain-text list you can paste into an email. With nothing selected it offers the ones marked NEEDED.',
               'More ▾':'Everything else: glossaries, copy helpers, folders, import/export, settings.'}
        toolbar=[('Quick Add',self.quick_add),('New Report',self.new_report),None,
                 ('Edit',self.edit),('Attach File',self.attach_selected),('Mark Filed',self.mark_filed),None,
                 ('Add to Report',self.assign_report),('Open Receipt/Invoice',self.open_doc),None,
                 ('Delete',self.delete),None,
                 ('Ask for Receipts',self.request_receipts),('Open Concur',lambda:webbrowser.open(CONCUR_URL))]
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
        self.count_lbl.bind('<Button-1>', lambda e: self.select_needed_receipts())
        Tooltip(self.count_lbl,'Open = every expense for this user that is not marked Filed.\n'
                               '"need a receipt" = nothing attached AND at or over the receipt limit in Settings.\n'
                               'Counts ignore the search box, so this is always the full pile.\n'
                               'Click here to tick every expense that still owes a receipt.')
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
        for c,w in [('id',55),('date',95),('vendor',190),('amount',85),('status',120),('doc',92),('code',170),('purpose',210),('ready',60)]: self.tree.heading(c,text=c.title()); self.tree.column(c,width=w)
        # The 'code' column shows the human-readable expense-type LABEL (not the bare GL number, which is
        # meaningless at a glance) — truncated by width is fine. Right-click-to-change-field is ROADMAP item 15.
        self.tree.heading('code', text='Expense Type')
        self.tree.pack(fill='both',expand=True,padx=4,pady=4); self.tree.bind('<Double-1>',self.on_double)
        # Multi-select (Ctrl/Shift-click) + two ways to file a batch into a report: right-click menu, or drag onto a report row.
        self.menu=tk.Menu(self, tearoff=0)
        self.menu.add_command(label='Add to report…', command=self.assign_report)
        self.menu.add_command(label='Copy expense', command=self.copy_expense)
        self.menu.add_command(label='Edit', command=self.edit)
        self.menu.add_separator()
        self.menu.add_command(label='Open receipt/invoice', command=self.open_doc)
        self.menu.add_command(label='Open file location', command=self.open_doc_location)
        self.menu.add_command(label='Open report folder…', command=self.report_folder_action)
        self.menu.add_command(label='Copy Concur summary', command=self.copy_summary)
        self.menu.add_command(label='Ask for these receipts…', command=self.request_receipts)
        self.menu.add_separator()
        self.menu.add_command(label='Move to user…', command=self.switch_user)
        self.menu.add_command(label='Mark filed', command=self.mark_filed)
        self.menu.add_command(label='Delete', command=self.delete)
        self.tree.bind('<Button-3>', self.popup_menu)
        self.tree.tag_configure('needreceipt', background='#fff3cd')  # amber = this one owes a receipt
        self.tree.tag_configure('reimburse', background='#fce8e6')    # red = you are owed money and the month is nearly out
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
        gloss.add_command(label='My Cards…', command=lambda:CardsDialog(self))
        gloss.add_command(label='Attendees…', command=lambda:AttendeeDialog(self))
        m.add_cascade(label='Glossaries', menu=gloss)
        m.add_separator()
        m.add_command(label='Copy Expense', command=self.copy_expense, accelerator='Ctrl+D')
        m.add_command(label='Move to User…', command=self.switch_user)
        m.add_command(label='Add User…', command=self.add_user)
        m.add_separator()
        m.add_command(label='Copy Concur Summary', command=self.copy_summary)
        m.add_command(label='Copy Attendees', command=self.copy_attendees)
        m.add_separator()
        m.add_command(label='Select the receipts I am owed', command=self.select_needed_receipts)
        m.add_command(label='Ask for Receipts…', command=self.request_receipts)
        m.add_separator()
        m.add_command(label='Open File Location', command=self.open_doc_location)
        m.add_command(label='Open Report Folder…', command=self.report_folder_action)
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
        elif col=='type': hay=(q(e['expense_type_code'])+' '+q(e['expense_type_label'])).lower()  # 'Expense Type' column: filter code AND label
        else: hay=q(e[field_map.get(col,'vendor')]).lower()
        return term in hay
    @staticmethod
    def _docs(e):
        # Both docs can coexist on one record. With NEITHER attached the cell answers the only question that
        # matters — does this one actually owe a receipt, or is it under the limit and none of your problem?
        got='+'.join(x for x,p in [('Receipt', e['receipt_path']),('Invoice', e['invoice_path'])] if p)
        return got or ('NEEDED' if needs_receipt(e) else 'not needed')
    def insert_exp(self,parent,e):
        self.tree.insert(parent,'end',iid=f"E{e['id']}", text='[Expense] '+e['vendor'], image=self.chk_img[False],
                         tags=('reimburse',) if reimbursement_due(e) else (('needreceipt',) if needs_receipt(e) else ()), values=(e['id'],e['transaction_date'],e['vendor'],money_part(e['amount']),e['status'],self._docs(e),(e['expense_type_label'] or e['expense_type_code'] or ''),e['business_purpose'] or '','✓' if e['status'] in ('Ready to file','Filed') else ''))
        self._chk_state[f"E{e['id']}"]=False
    def refresh_counter(self):
        """The home-screen counter. Deliberately NOT filtered by the search box / status dropdown — it
        answers "how much is still on my plate", which a filtered view would understate."""
        uid=self.current_user_id()
        if not uid: self.open_count.set(''); return
        n=lambda where: S.scalar(f"SELECT COUNT(*) FROM expenses WHERE user_id=? {where}",(uid,)) or 0
        openn=n("AND status<>'Filed'")
        if not openn: self.open_count.set('No open receipts — all caught up.'); return
        need=n("AND status<>'Filed' AND COALESCE(receipt_path,'')='' AND COALESCE(invoice_path,'')='' "
               "AND COALESCE(missing_receipt_ack,0)=0 AND ABS(COALESCE(amount,0))>=%f" % receipt_limit())
        ready=n("AND status='Ready to file'")
        owed=sum(1 for r in S.rows("SELECT * FROM expenses WHERE user_id=? AND status<>'Filed'",(uid,)) if reimbursement_due(r))
        parts=[f"{openn} open receipt{'' if openn==1 else 's'}"]
        if need: parts.append(f"{need} need{'s' if need==1 else ''} a receipt")
        if ready: parts.append(f"{ready} ready to file")
        if owed: parts.append(f"{owed} to claim back before the month ends")
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
        # Double-click opens whatever you clicked: an expense, or a report's header fields — the ones Concur
        # demands on Create Expense Report. Expanding a report to see what is inside it
        # stays on the +/- control, which is why the rows are tall.
        iid=self.tree.identify_row(e.y)
        if iid.startswith('R'): ReportDialog(self, int(iid[1:]))
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
    def needed_receipt_ids(self):
        """Every expense this user still owes a receipt for, newest first — the rows the list marks NEEDED."""
        return [r['id'] for r in S.rows('SELECT * FROM expenses WHERE user_id=? ORDER BY transaction_date DESC',
                                        (self.current_user_id(),)) if needs_receipt(r)]
    def select_needed_receipts(self):
        """Tick every row that owes a receipt, so a chase list is one click instead of hunting the amber rows."""
        ids=[f'E{i}' for i in self.needed_receipt_ids() if self.tree.exists(f'E{i}')]
        self.tree.selection_set(*ids) if ids else self.tree.selection_set(())
        self._sync_checks()
        self.flash_status(f"{len(ids)} receipt{'' if len(ids)==1 else 's'} selected" if ids else 'No receipts are owed')
        return ids
    def request_receipts(self):
        """Ticked expenses (or a selected report's expenses) -> a plain-text list to paste into an email.
        Nothing selected = offer the owed ones."""
        ids=self.scoped_expense_ids()
        if not ids:
            owed=self.needed_receipt_ids()
            if not owed: messagebox.showinfo('Ask for Receipts','Nothing is waiting on a receipt right now.',parent=self); return
            if not messagebox.askyesno('Ask for Receipts', f"Nothing is ticked. Use the {len(owed)} expense(s) "
                                       "marked NEEDED instead?", parent=self): return
            ids=owed; self.select_needed_receipts()
        rows=[S.row('SELECT * FROM expenses WHERE id=?',(i,)) for i in ids]
        rows=[r for r in rows if r]
        rows.sort(key=lambda r: q(r['transaction_date']))
        ReceiptRequestDialog(self, rows)
    def selected_expense_ids(self):
        """Every EXPENSE row currently selected (multi-select via Ctrl/Shift-click). Report rows are ignored."""
        return [int(iid[1:]) for iid in self.tree.selection() if iid.startswith('E')]
    def selected_report_ids(self):
        """Every REPORT row currently selected."""
        return [int(iid[1:]) for iid in self.tree.selection() if iid.startswith('R')]
    def scoped_expense_ids(self):
        """The expenses a batch action should touch, given what's selected: the ticked expense rows
        PLUS every expense inside any selected REPORT. Selecting a report therefore means 'act on
        everything in this report' — the action fans out to its expenses WITHOUT ticking each child
        box (that would just be a noisy select-all). De-duplicated; a report's members follow it."""
        ids=[]
        for iid in self.tree.selection():
            if iid.startswith('R'):
                ids += [r['id'] for r in S.rows('SELECT id FROM expenses WHERE report_id=? ORDER BY transaction_date DESC',(int(iid[1:]),))]
            elif iid.startswith('E'):
                ids.append(int(iid[1:]))
        seen=set(); out=[]
        for i in ids:
            if i not in seen: seen.add(i); out.append(i)
        return out
    def _report_scope_label(self):
        """A short ' from report …' / ' from N reports' suffix for feedback, when the action's scope
        came from selecting report row(s) rather than individual expenses. '' when no report is selected."""
        reps=self.selected_report_ids()
        if not reps: return ''
        if len(reps)==1: return f" from report {S.scalar('SELECT name FROM reports WHERE id=?',(reps[0],)) or ''}".rstrip()
        return f" from {len(reps)} reports"
    def assign_report(self):
        ids=self.selected_expense_ids()
        if not ids: messagebox.showinfo('Assign to Report','Select one or more expenses first.'); return
        AssignReportDialog(self, ids, self.current_user_id(), self.refresh)
    def _assign_expenses_to_report(self, ids, rid):
        for eid in ids: S.execute('UPDATE expenses SET report_id=? WHERE id=?',(rid,eid))
        if rid: self.gather_report_files(rid, quiet=True)  # keep the report's own folder of receipts current
        self.refresh()
    def gather_report_files(self, rid, quiet=False, reveal=False):
        """Copy the report's receipts into its own folder on disk. Copies — the originals stay filed under
        User\\Year where the database points, so this can run on every assignment without risking anything."""
        try: folder, copied, already, missing = collect_report_receipts(rid)
        except Exception as e:
            if not quiet: messagebox.showerror('Report folder', str(e), parent=self)
            return None
        if not folder:
            if not quiet: messagebox.showinfo('Report folder','Set a receipt root in Settings first.',parent=self)
            return None
        if copied and quiet: self.flash_status(f"{copied} receipt{'' if copied==1 else 's'} copied into the report folder")
        if not quiet:
            note=f'{folder}\n\nCopied {copied}, already there {already}.'
            if missing: note+='\n\nNo file found for:\n  '+'\n  '.join(missing[:10])
            note+='\n\nThese are COPIES — the originals stay where they were filed.'
            messagebox.showinfo('Report folder', note, parent=self)
        if reveal and folder.exists(): open_folder(str(folder), 'report folder')
        return folder
    def report_folder_action(self):
        """More ▾ / right-click: refresh the selected report's folder and open it."""
        typ,rid=self.selected()
        if typ!='R':
            typ2,eid=self.selected()
            row=S.row('SELECT report_id FROM expenses WHERE id=?',(eid,)) if typ2=='E' else None
            rid=row['report_id'] if row and row['report_id'] else None
            if not rid: messagebox.showinfo('Report folder','Select a report (or an expense inside one) first.',parent=self); return
        self.gather_report_files(rid, quiet=False, reveal=True)
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
        # Policy number, not a path -> DB setting, so it follows the synced database across machines.
        lim=tk.StringVar(value=f"{receipt_limit():.2f}"); vars_db={'receipt_limit': lim}
        ttk.Label(frm,text='Receipt needed at or above').grid(row=2,column=0,sticky='w')
        ttk.Entry(frm,textvariable=lim,width=12,validate='key',validatecommand=(self.register(_num_validate),'%P')).grid(row=2,column=1,sticky='w')
        ttk.Label(frm,text='$  — below this, the list shows "not needed" and stops nagging').grid(row=2,column=2,sticky='w')
        # --- Database location (move to a Drive-synced folder, or point at an existing DB on another device) ---
        ttk.Separator(frm,orient='horizontal').grid(row=3,column=0,columnspan=3,sticky='ew',pady=8)
        ttk.Label(frm,text='Database').grid(row=4,column=0,sticky='w')
        db_var=tk.StringVar(value=str(S.path)); ttk.Entry(frm,textvariable=db_var,state='readonly').grid(row=4,column=1,sticky='ew')
        dbbar=FlowBar(frm,padding=(0,2)); dbbar.grid(row=5,column=0,columnspan=3,sticky='ew')
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
        dbbar.button('My cards…', lambda: CardsDialog(self))
        def save():
            for k,v in vars.items(): set_local_setting(k, v.get())  # inbox_path/receipt_root are per-machine local
            for k,v in vars_db.items(): S.set_setting(k, clean_number(v.get()))  # policy numbers ride with the DB
            receipt_limit(reload=True); self.refresh()  # repaint the list against the new limit straight away
            win.destroy()
        actbar=FlowBar(frm,padding=(0,8)); actbar.grid(row=6,column=0,columnspan=3,sticky='ew')
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
        # Selecting a report exports everything inside it (scoped_expense_ids fans the report out to its
        # expenses); selecting expenses exports just those; selecting nothing falls back to Ready-to-file.
        ids=self.scoped_expense_ids()
        rows=[S.row('SELECT * FROM expenses WHERE id=?',(i,)) for i in ids] if ids \
            else S.rows("SELECT * FROM expenses WHERE user_id=? AND status='Ready to file' ORDER BY transaction_date",(self.current_user_id(),))
        rows=[r for r in rows if r]
        if not rows:
            messagebox.showinfo('Export for Autofill','Select one or more expenses, or a report, first (or mark some "Ready to file").'); return
        total=sum(float(r['amount'] or 0) for r in rows)
        p=filedialog.asksaveasfilename(defaultextension='.json', initialfile=autofill_filename(len(rows), total),
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
        self.flash_status(f'Exported {len(out)} expense(s){self._report_scope_label()} for autofill')
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
