"""
SQLite data-access layer for the shop-rent management bot.
All functions open a short-lived connection (simple & safe for a bot this size).
"""

import os
import sqlite3
import secrets
from datetime import datetime, date, timedelta
from contextlib import contextmanager

# Sentinel meaning "clear this field to NULL", distinct from Python's own
# None (which update_payment treats as "leave this field alone").
_CLEAR = object()
CLEAR = _CLEAR  # public name for other modules (webapp.py, bot.py) to import

# Always resolve the database next to this file, regardless of the current
# working directory the script happens to be launched from.
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rent.db")

# Floors are rows in a `floors` table (see init_db) managed entirely at
# runtime — see add_floor, update_floor_label, update_floor_lump_sum,
# move_floor, delete_floor below, and /addfloors + /editfloors in bot.py /
# the Floors page in webapp.py. A brand-new database starts with ZERO
# floors — nothing is pre-seeded — so the admin must add at least one floor
# (Add Floor) before any shop can be added. See the floors-exist guard in
# add_shop() below and the equivalent checks in bot.py's addshop flow and
# webapp.py's admin_shop_new route.


def _slugify_floor_key(label):
    """Best-effort short key derived from a floor's display label, e.g.
    'Fourth Floor' -> 'fourth_floor'. Only used when creating a NEW floor
    (add_floor) — existing keys are never recomputed from their label, so
    renaming a floor later never breaks the shops already on it."""
    slug = "".join(c if c.isalnum() else "_" for c in label.strip().lower())
    slug = "_".join(part for part in slug.split("_") if part)
    return slug or "floor"


def get_floors():
    """(key, label) pairs in building display order. Replaces the old
    hardcoded FLOORS list."""
    with get_conn() as conn:
        rows = conn.execute("SELECT key, label FROM floors ORDER BY sort_order").fetchall()
    return [(r["key"], r["label"]) for r in rows]


def get_floor_labels():
    """{key: label} dict. Replaces the old hardcoded FLOOR_LABELS."""
    return dict(get_floors())


def get_lump_sum_floors():
    """List of floor keys that pay in a lump sum. Replaces the old
    hardcoded LUMP_SUM_FLOORS."""
    with get_conn() as conn:
        rows = conn.execute("SELECT key FROM floors WHERE lump_sum = 1").fetchall()
    return [r["key"] for r in rows]


def get_floors_with_counts():
    """Every floor plus how many ACTIVE shops currently sit on it — for the
    /editfloors and webapp Floors admin screens (a floor with active shops
    on it can't be deleted). Deactivated shops don't count: they're the
    normal way an admin clears a floor for deletion without permanently
    erasing shop history, so they mustn't keep blocking it."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT f.key, f.label, f.lump_sum, "
            "(SELECT COUNT(*) FROM shops s WHERE s.floor = f.key AND s.active = 1) AS shop_count "
            "FROM floors f ORDER BY f.sort_order"
        ).fetchall()
    return [dict(r) for r in rows]


def get_floor(key):
    with get_conn() as conn:
        row = conn.execute("SELECT key, label, lump_sum FROM floors WHERE key = ?", (key,)).fetchone()
    return dict(row) if row else None


def floor_label(key):
    return get_floor_labels().get(key, key or "-")


def add_floor(label, lump_sum=False):
    """Adds a new floor at the bottom of the display order. The key is
    derived from the label and de-duplicated if needed. Returns the new
    floor's key."""
    label = label.strip()
    with get_conn() as conn:
        existing_keys = {r["key"] for r in conn.execute("SELECT key FROM floors")}
        base = _slugify_floor_key(label)
        key = base
        n = 2
        while key in existing_keys:
            key = f"{base}_{n}"
            n += 1
        next_order = conn.execute("SELECT COALESCE(MAX(sort_order), -1) + 1 FROM floors").fetchone()[0]
        conn.execute(
            "INSERT INTO floors (key, label, sort_order, lump_sum) VALUES (?, ?, ?, ?)",
            (key, label, next_order, 1 if lump_sum else 0),
        )
    return key


def update_floor_label(key, new_label):
    with get_conn() as conn:
        conn.execute("UPDATE floors SET label = ? WHERE key = ?", (new_label.strip(), key))


def update_floor_lump_sum(key, lump_sum):
    with get_conn() as conn:
        conn.execute("UPDATE floors SET lump_sum = ? WHERE key = ?", (1 if lump_sum else 0, key))


def move_floor(key, direction):
    """Swaps this floor's display position with its neighbor.
    direction: 'up' or 'down'. No-op at either end of the list."""
    with get_conn() as conn:
        rows = conn.execute("SELECT key, sort_order FROM floors ORDER BY sort_order").fetchall()
        keys = [r["key"] for r in rows]
        if key not in keys:
            return
        i = keys.index(key)
        j = i - 1 if direction == "up" else i + 1
        if j < 0 or j >= len(keys):
            return
        a, b = rows[i], rows[j]
        conn.execute("UPDATE floors SET sort_order = ? WHERE key = ?", (a["sort_order"], b["key"]))
        conn.execute("UPDATE floors SET sort_order = ? WHERE key = ?", (b["sort_order"], a["key"]))


def delete_floor(key):
    """Deletes a floor. Raises ValueError (with the count) if any ACTIVE
    shop is still assigned to it — reassign or deactivate those shops
    first. A deactivated shop still tagged with this floor key does NOT
    block deletion (it's already off the active roll; its floor tag is
    just history), matching what get_floors_with_counts shows the admin."""
    with get_conn() as conn:
        count = conn.execute(
            "SELECT COUNT(*) FROM shops WHERE floor = ? AND active = 1", (key,)
        ).fetchone()[0]
        if count:
            raise ValueError(count)
        conn.execute("DELETE FROM floors WHERE key = ?", (key,))


# ---------------------------------------------------------------------------
# Ethiopian calendar conversions.
#
# Every date in this app is still keyed and queried in Gregorian ('YYYY-MM-DD',
# sortable, unambiguous for range queries), but every table that stores a date
# also stores that date's Ethiopian-calendar equivalent alongside it (in a
# '<column>_ec' column, as 'EY-EM-ED' using Ethiopian year/month/day numbers)
# so the Ethiopian date is a real stored fact, not something recomputed only
# for display.
#
# Algorithm verified against known reference points (Ethiopian New Year 2018 =
# 2025-09-11 G.C., 2016 = 2023-09-12 G.C.) and a 20,000-sample round-trip check.
# ---------------------------------------------------------------------------

_ETH_EPOCH_JDN = 1723856  # JDN of Ethiopian year 0, Meskerem 1

ETHIOPIAN_MONTHS = {
    1: "መስከረም", 2: "ጥቅምት", 3: "ኅዳር", 4: "ታኅሳስ", 5: "ጥር", 6: "የካቲት",
    7: "መጋቢት", 8: "ሚያዝያ", 9: "ግንቦት", 10: "ሰኔ", 11: "ሐምሌ", 12: "ነሐሴ", 13: "ጳጉሜ",
}


def _gregorian_to_jdn(year, month, day):
    a = (14 - month) // 12
    y = year + 4800 - a
    m = month + 12 * a - 3
    return day + (153 * m + 2) // 5 + 365 * y + y // 4 - y // 100 + y // 400 - 32045


def _ethiopian_new_year_jdn(year):
    return _ETH_EPOCH_JDN + 365 * year + year // 4


def _ethiopian_to_jdn(year, month, day):
    return _ethiopian_new_year_jdn(year) + (month - 1) * 30 + (day - 1)


def _jdn_to_ethiopian(jdn):
    r = (jdn - _ETH_EPOCH_JDN) % 1461
    n = (r % 365) + 365 * (r // 1460)
    year = 4 * ((jdn - _ETH_EPOCH_JDN) // 1461) + (r // 365) - (r // 1460)
    month = n // 30 + 1
    day = n % 30 + 1
    return year, month, day


def gregorian_to_ethiopian(year, month, day):
    return _jdn_to_ethiopian(_gregorian_to_jdn(year, month, day))


def _jdn_to_gregorian(jdn):
    a = jdn + 32044
    b = (4 * a + 3) // 146097
    c = a - (146097 * b) // 4
    d = (4 * c + 3) // 1461
    e = c - (1461 * d) // 4
    m = (5 * e + 2) // 153
    day = e - (153 * m + 2) // 5 + 1
    month = m + 3 - 12 * (m // 10)
    year = 100 * b + d - 4800 + m // 10
    return year, month, day


def ethiopian_to_gregorian(year, month, day):
    return _jdn_to_gregorian(_ethiopian_to_jdn(year, month, day))


def add_ethiopian_months(year, month, delta):
    """Add `delta` Ethiopian months (13-month calendar, Pagume counts as one)
    to (year, month), returning the resulting (year, month). Used to step
    through rent periods, which are keyed by Ethiopian month (see
    apply_monthly_charge / add_payment) rather than Gregorian month —
    the two don't line up, since every Ethiopian month starts partway
    through a Gregorian one (most sharply around Meskerem 1 / Ethiopian
    New Year, ~Sept 11, which falls in the middle of Gregorian September)."""
    total = (year * 13 + (month - 1)) + delta
    return total // 13, total % 13 + 1


def format_ethiopian_date(year, month, day):
    return f"{day} {ETHIOPIAN_MONTHS[month]} {year}"


def gc_iso_to_ec_label(gc_iso_date):
    """'2026-09-13' -> '3 መስከረም 2019' (Ethiopian calendar)."""
    if not gc_iso_date:
        return None
    y, m, d = (int(x) for x in gc_iso_date.split("-"))
    ey, em, ed = gregorian_to_ethiopian(y, m, d)
    return format_ethiopian_date(ey, em, ed)


def gc_iso_to_ec_iso(gc_iso_date):
    """'2026-09-13' -> '2019-01-03' — the Ethiopian-calendar equivalent, kept
    in the same zero-padded 'YYYY-MM-DD' shape as the Gregorian column it sits
    next to, for storage rather than display."""
    if not gc_iso_date:
        return None
    y, m, d = (int(x) for x in gc_iso_date.split("-"))
    ey, em, ed = gregorian_to_ethiopian(y, m, d)
    return f"{ey:04d}-{em:02d}-{ed:02d}"


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS shops (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shop_no TEXT UNIQUE NOT NULL,
                tenant_name TEXT,
                phone TEXT,
                monthly_rent REAL NOT NULL,
                floor TEXT,
                purpose TEXT,
                national_id TEXT,
                tin_number TEXT,
                document_file_id TEXT,
                document_kind TEXT,
                start_date TEXT,
                start_date_ec TEXT,
                lease_end_date TEXT,
                lease_end_date_ec TEXT,
                active INTEGER DEFAULT 1,
                is_rented INTEGER DEFAULT 0,
                area_sqm REAL,
                telegram_id INTEGER UNIQUE,
                link_code TEXT UNIQUE
            );

            CREATE TABLE IF NOT EXISTS charges (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shop_id INTEGER NOT NULL REFERENCES shops(id),
                amount REAL NOT NULL,
                charge_date TEXT NOT NULL,
                charge_date_ec TEXT,
                period TEXT NOT NULL,        -- 'YYYY-MM', one rent charge per shop per period
                description TEXT,
                UNIQUE(shop_id, period)
            );

            CREATE TABLE IF NOT EXISTS payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shop_id INTEGER NOT NULL REFERENCES shops(id),
                amount REAL NOT NULL,
                payment_date TEXT NOT NULL,
                payment_date_ec TEXT,
                note TEXT,
                period TEXT,              -- 'YYYY-MM' the payment is being applied to
                bank TEXT,                -- 'CBE Birr' / 'Telebirr' / 'Cash' / 'Other'
                reference_no TEXT,        -- bank/transaction reference number
                receipt_file_id TEXT      -- Telegram file_id of the attached receipt photo
            );

            CREATE TABLE IF NOT EXISTS recurring_expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                category TEXT UNIQUE NOT NULL,  -- e.g. 'Electricity Bill', 'Water Bill', 'Salary'
                amount REAL NOT NULL,
                active INTEGER DEFAULT 1
            );

            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                description TEXT NOT NULL,
                amount REAL NOT NULL,
                category TEXT,
                expense_date TEXT NOT NULL,
                expense_date_ec TEXT,
                receipt_file_id TEXT,     -- Telegram file_id of the attached receipt photo
                period TEXT,              -- 'YYYY-MM' E.C., set when auto-applied (see recurring_expenses)
                recurring_expense_id INTEGER REFERENCES recurring_expenses(id)
            );

            CREATE TABLE IF NOT EXISTS shop_documents (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shop_id INTEGER NOT NULL REFERENCES shops(id),
                file_id TEXT NOT NULL,
                kind TEXT NOT NULL,          -- 'photo' or 'document'
                label TEXT,                  -- e.g. 'Lease document', 'ID', 'TIN certificate'
                uploaded_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL,              -- 'admin' or 'tenant'
                shop_id INTEGER REFERENCES shops(id),
                full_name TEXT,
                active INTEGER DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS floors (
                key TEXT PRIMARY KEY,
                label TEXT NOT NULL,
                sort_order INTEGER NOT NULL,
                lump_sum INTEGER NOT NULL DEFAULT 0   -- 1 = whole lease paid up front, not monthly
            );

            CREATE TABLE IF NOT EXISTS rent_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shop_id INTEGER NOT NULL REFERENCES shops(id),
                effective_period TEXT NOT NULL,  -- Ethiopian 'YYYY-MM' — first period this rate applies from
                rent REAL NOT NULL,
                created_date TEXT NOT NULL,       -- Gregorian ISO date this change was recorded, for audit
                UNIQUE(shop_id, effective_period)
            );
            """
        )

        # No floor seeding: a fresh database intentionally starts with zero
        # floors. The admin adds floors explicitly (Add Floor), and shops
        # can't be created until at least one floor exists.

        # Migration: add newer payment columns for databases created before they existed.
        existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(payments)")}
        for col in ("period", "bank", "reference_no", "receipt_file_id", "payment_date_ec"):
            if col not in existing_cols:
                conn.execute(f"ALTER TABLE payments ADD COLUMN {col} TEXT")

        # Migration: add newer expense columns (recurring/"permanent" expenses
        # like Electricity Bill, Water Bill, Salary) for databases created
        # before they existed.
        existing_expense_cols = {row["name"] for row in conn.execute("PRAGMA table_info(expenses)")}
        for col, col_type in (("period", "TEXT"), ("recurring_expense_id", "INTEGER")):
            if col not in existing_expense_cols:
                conn.execute(f"ALTER TABLE expenses ADD COLUMN {col} {col_type}")
        # One auto-applied expense per recurring definition per period (mirrors
        # the UNIQUE(shop_id, period) constraint on charges) — partial index
        # since manually-logged expenses (recurring_expense_id IS NULL) aren't
        # limited this way.
        conn.execute(
            "CREATE UNIQUE INDEX IF NOT EXISTS idx_expenses_recurring_period "
            "ON expenses(recurring_expense_id, period) WHERE recurring_expense_id IS NOT NULL"
        )

        # Migration: add columns for databases created before they existed.
        # (sqm/price_per_sqm are deliberately NOT in this list any more — the
        # per-m2 pricing model was replaced by a flat monthly rent, and the
        # rebuild below drops those two columns for good.)
        existing_shop_cols = {row["name"] for row in conn.execute("PRAGMA table_info(shops)")}
        for col, col_type in (("floor", "TEXT"), ("purpose", "TEXT"), ("national_id", "TEXT"),
                              ("tin_number", "TEXT"), ("document_file_id", "TEXT"),
                              ("document_kind", "TEXT"), ("deactivated_date", "TEXT"),
                              ("deactivated_date_ec", "TEXT"),
                              ("is_rented", "INTEGER DEFAULT 0"), ("area_sqm", "REAL"),
                              ("lease_end_date", "TEXT"), ("start_date_ec", "TEXT"),
                              ("lease_end_date_ec", "TEXT")):
            if col not in existing_shop_cols:
                conn.execute(f"ALTER TABLE shops ADD COLUMN {col} {col_type}")

        # Migration: add the Ethiopian-date column for charges/expenses too.
        existing_charge_cols = {row["name"] for row in conn.execute("PRAGMA table_info(charges)")}
        if "charge_date_ec" not in existing_charge_cols:
            conn.execute("ALTER TABLE charges ADD COLUMN charge_date_ec TEXT")
        existing_expense_cols = {row["name"] for row in conn.execute("PRAGMA table_info(expenses)")}
        if "expense_date_ec" not in existing_expense_cols:
            conn.execute("ALTER TABLE expenses ADD COLUMN expense_date_ec TEXT")
        if "receipt_file_id" not in existing_expense_cols:
            conn.execute("ALTER TABLE expenses ADD COLUMN receipt_file_id TEXT")

        # Migration: shops used to require a tenant at creation time (tenant_name
        # and start_date were NOT NULL) and rent was derived from sqm x price/m2.
        # Shops can now be added vacant and priced with a flat monthly rent, so
        # rebuild the table without those constraints/columns, preserving any
        # existing rows. Any shop that already had a tenant_name is treated as
        # rented.
        needs_rebuild = False
        for row in conn.execute("PRAGMA table_info(shops)"):
            if row["name"] in ("tenant_name", "start_date") and row["notnull"]:
                needs_rebuild = True
                break
        if needs_rebuild:
            conn.execute(
                """
                CREATE TABLE shops_new (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    shop_no TEXT UNIQUE NOT NULL,
                    tenant_name TEXT,
                    phone TEXT,
                    monthly_rent REAL NOT NULL,
                    floor TEXT,
                    purpose TEXT,
                    national_id TEXT,
                    tin_number TEXT,
                    document_file_id TEXT,
                    document_kind TEXT,
                    start_date TEXT,
                    start_date_ec TEXT,
                    lease_end_date TEXT,
                    lease_end_date_ec TEXT,
                    active INTEGER DEFAULT 1,
                    is_rented INTEGER DEFAULT 0,
                    deactivated_date TEXT,
                    deactivated_date_ec TEXT,
                    area_sqm REAL,
                    telegram_id INTEGER UNIQUE,
                    link_code TEXT UNIQUE
                )
                """
            )
            conn.execute(
                "INSERT INTO shops_new (id, shop_no, tenant_name, phone, monthly_rent, "
                "floor, purpose, national_id, tin_number, document_file_id, document_kind, "
                "start_date, start_date_ec, lease_end_date, lease_end_date_ec, active, is_rented, "
                "deactivated_date, deactivated_date_ec, area_sqm, telegram_id, link_code) "
                "SELECT id, shop_no, tenant_name, phone, monthly_rent, floor, purpose, "
                "national_id, tin_number, document_file_id, document_kind, start_date, "
                "start_date_ec, lease_end_date, lease_end_date_ec, active, "
                "CASE WHEN tenant_name IS NOT NULL AND tenant_name != '' THEN 1 ELSE 0 END, "
                "deactivated_date, deactivated_date_ec, area_sqm, telegram_id, link_code FROM shops"
            )
            conn.execute("DROP TABLE shops")
            conn.execute("ALTER TABLE shops_new RENAME TO shops")

        # Backfill: any row whose Gregorian date was stored before the _ec
        # columns existed won't have an Ethiopian equivalent yet — compute it
        # once so old records get the same "store both" treatment as new ones.
        for row in conn.execute(
            "SELECT id, start_date, start_date_ec, lease_end_date, lease_end_date_ec, "
            "deactivated_date, deactivated_date_ec FROM shops"
        ):
            updates, params = [], []
            if row["start_date"] and not row["start_date_ec"]:
                updates.append("start_date_ec = ?")
                params.append(gc_iso_to_ec_iso(row["start_date"]))
            if row["lease_end_date"] and not row["lease_end_date_ec"]:
                updates.append("lease_end_date_ec = ?")
                params.append(gc_iso_to_ec_iso(row["lease_end_date"]))
            if row["deactivated_date"] and not row["deactivated_date_ec"]:
                updates.append("deactivated_date_ec = ?")
                params.append(gc_iso_to_ec_iso(row["deactivated_date"]))
            if updates:
                params.append(row["id"])
                conn.execute(f"UPDATE shops SET {', '.join(updates)} WHERE id = ?", params)
        for row in conn.execute("SELECT id, charge_date, charge_date_ec FROM charges"):
            if row["charge_date"] and not row["charge_date_ec"]:
                conn.execute(
                    "UPDATE charges SET charge_date_ec = ? WHERE id = ?",
                    (gc_iso_to_ec_iso(row["charge_date"]), row["id"]),
                )
        for row in conn.execute("SELECT id, payment_date, payment_date_ec FROM payments"):
            if row["payment_date"] and not row["payment_date_ec"]:
                conn.execute(
                    "UPDATE payments SET payment_date_ec = ? WHERE id = ?",
                    (gc_iso_to_ec_iso(row["payment_date"]), row["id"]),
                )
        for row in conn.execute("SELECT id, expense_date, expense_date_ec FROM expenses"):
            if row["expense_date"] and not row["expense_date_ec"]:
                conn.execute(
                    "UPDATE expenses SET expense_date_ec = ? WHERE id = ?",
                    (gc_iso_to_ec_iso(row["expense_date"]), row["id"]),
                )


# ---------- Shops ----------

def _new_link_code():
    return secrets.token_hex(3).upper()  # e.g. 'A1B2C3'


def add_shop(shop_no, monthly_rent, floor=None, area_sqm=None):
    """Adds a new shop with no tenant yet — vacant, with a flat monthly rent.
    Use assign_tenant() afterwards once/if the shop is rented out.

    Raises ValueError if no floors exist yet — a floor must be added first
    (see add_floor) before any shop can be created. This is enforced here,
    not just in the bot/webapp UI, so it holds no matter how add_shop is
    called."""
    with get_conn() as conn:
        if conn.execute("SELECT COUNT(*) FROM floors").fetchone()[0] == 0:
            raise ValueError("no_floors")
        conn.execute(
            "INSERT INTO shops (shop_no, monthly_rent, floor, area_sqm, is_rented) "
            "VALUES (?, ?, ?, ?, 0)",
            (shop_no, monthly_rent, floor, area_sqm),
        )


def assign_tenant(shop_no, tenant_name, phone=None, purpose=None, start_date=None):
    """Marks a vacant shop as rented: records the tenant's details and issues
    a link code for them to register with. Returns the link code."""
    start_date = start_date or date.today().isoformat()
    code = _new_link_code()
    with get_conn() as conn:
        conn.execute(
            "UPDATE shops SET tenant_name = ?, phone = ?, purpose = ?, start_date = ?, "
            "start_date_ec = ?, is_rented = 1, link_code = ? WHERE shop_no = ?",
            (tenant_name, phone, purpose, start_date, gc_iso_to_ec_iso(start_date), code, shop_no),
        )
    return code


def mark_vacant(shop_no):
    """Clears a shop's tenant info (e.g. they moved out) and marks it vacant
    again. Keeps the shop's floor and monthly rent."""
    with get_conn() as conn:
        shop = conn.execute("SELECT id FROM shops WHERE shop_no = ?", (shop_no,)).fetchone()
        conn.execute(
            "UPDATE shops SET is_rented = 0, tenant_name = NULL, phone = NULL, "
            "purpose = NULL, start_date = NULL, start_date_ec = NULL, lease_end_date = NULL, "
            "lease_end_date_ec = NULL, telegram_id = NULL, "
            "link_code = NULL, national_id = NULL, tin_number = NULL, document_file_id = NULL, "
            "document_kind = NULL WHERE shop_no = ?",
            (shop_no,),
        )
        if shop:
            conn.execute("DELETE FROM shop_documents WHERE shop_id = ?", (shop["id"],))


def update_national_id(shop_no, national_id):
    with get_conn() as conn:
        conn.execute("UPDATE shops SET national_id = ? WHERE shop_no = ?", (national_id, shop_no))


def update_tin(shop_no, tin_number):
    with get_conn() as conn:
        conn.execute("UPDATE shops SET tin_number = ? WHERE shop_no = ?", (tin_number, shop_no))


def update_document(shop_no, file_id, kind):
    with get_conn() as conn:
        conn.execute(
            "UPDATE shops SET document_file_id = ?, document_kind = ? WHERE shop_no = ?",
            (file_id, kind, shop_no),
        )


def set_lease_end(shop_no, lease_end_date):
    """Sets (or clears, if lease_end_date is None) the end of a tenant's
    leasing period. Stored as 'YYYY-MM-DD' (Gregorian), same convention as
    start_date, alongside its Ethiopian-calendar equivalent in
    lease_end_date_ec."""
    with get_conn() as conn:
        conn.execute(
            "UPDATE shops SET lease_end_date = ?, lease_end_date_ec = ? WHERE shop_no = ?",
            (lease_end_date, gc_iso_to_ec_iso(lease_end_date), shop_no),
        )


def add_shop_document(shop_no, file_id, kind, label=None):
    """Attaches one more document to a shop without discarding earlier ones —
    a shop can have any number (lease document, ID, TIN certificate, etc.)."""
    shop = get_shop(shop_no)
    if not shop:
        return False
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO shop_documents (shop_id, file_id, kind, label, uploaded_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (shop["id"], file_id, kind, label, datetime.now().isoformat(timespec="seconds")),
        )
    return True


def get_shop_documents(shop_no):
    """All documents attached to a shop, oldest first."""
    shop = get_shop(shop_no)
    if not shop:
        return []
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT id, file_id, kind, label, uploaded_at FROM shop_documents "
            "WHERE shop_id = ? ORDER BY uploaded_at",
            (shop["id"],),
        ).fetchall()
        return [dict(r) for r in rows]


def get_shop_document(doc_id):
    """One document row (including its shop_id, so callers can check who
    it belongs to), or None."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT id, shop_id, file_id, kind, label, uploaded_at "
            "FROM shop_documents WHERE id = ?", (doc_id,),
        ).fetchone()
        return dict(row) if row else None


def payment_status(shop_no):
    """Returns ('on_time' | 'overdue' | 'unknown', balance)."""
    balance = get_balance(shop_no)
    if balance is None:
        return "unknown", None
    return ("on_time" if balance <= 0 else "overdue"), balance


def regenerate_link_code(shop_no):
    code = _new_link_code()
    with get_conn() as conn:
        conn.execute("UPDATE shops SET link_code = ? WHERE shop_no = ?", (code, shop_no))
    return code


def get_shop_by_link_code(code):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM shops WHERE link_code = ?", (code.upper().strip(),)
        ).fetchone()
        return dict(row) if row else None


def get_shop_by_telegram_id(telegram_id):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM shops WHERE telegram_id = ?", (telegram_id,)
        ).fetchone()
        return dict(row) if row else None


def link_telegram_to_shop(shop_no, telegram_id):
    with get_conn() as conn:
        conn.execute(
            "UPDATE shops SET telegram_id = ? WHERE shop_no = ?", (telegram_id, shop_no)
        )


def unlink_shop(shop_no):
    with get_conn() as conn:
        conn.execute(
            "UPDATE shops SET telegram_id = NULL WHERE shop_no = ?", (shop_no,)
        )


def get_shop(shop_no):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM shops WHERE shop_no = ?", (shop_no,)
        ).fetchone()
        return dict(row) if row else None


def get_shop_by_id(shop_id):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM shops WHERE id = ?", (shop_id,)).fetchone()
        return dict(row) if row else None


def get_all_shops(active_only=True, rented_only=None):
    """rented_only=True limits to shops with a tenant; False limits to vacant
    shops; None (default) returns both."""
    with get_conn() as conn:
        q = "SELECT * FROM shops"
        clauses = []
        if active_only:
            clauses.append("active = 1")
        if rented_only is True:
            clauses.append("is_rented = 1")
        elif rented_only is False:
            clauses.append("is_rented = 0")
        if clauses:
            q += " WHERE " + " AND ".join(clauses)
        q += " ORDER BY CAST(shop_no AS INTEGER), shop_no"
        return [dict(r) for r in conn.execute(q).fetchall()]


def get_shops_grouped_by_floor(active_only=True):
    """Shops grouped by floor, in building order (Underground -> Third Floor).
    Any shop with no floor set yet is returned last, under 'Unassigned'."""
    shops = get_all_shops(active_only=active_only)
    by_floor = {}
    for s in shops:
        by_floor.setdefault(s["floor"], []).append(s)
    grouped = [(key, label, by_floor[key]) for key, label in get_floors() if key in by_floor]
    if None in by_floor:
        grouped.append((None, "Unassigned", by_floor[None]))
    return grouped


def update_rent(shop_no, new_rent, effective_period=None):
    """Change a shop's rent, effective from `effective_period` (an Ethiopian
    'YYYY-MM' string) onward — defaults to the current Ethiopian month if
    not given.

    Every change is logged to rent_history with its effective period, so
    apply_monthly_charge() (and, through it, every report) can always look
    up exactly what a shop's rent was for a given month, even after it's
    been changed again since — that's what keeps past charges and past
    reports from drifting when a rent is edited.

    shops.monthly_rent (used everywhere "current rent" is displayed) is
    only updated immediately when the effective period is this month or
    earlier. A rate scheduled for a future month is recorded but
    monthly_rent is left alone until that month actually arrives — the
    scheduled monthly charge run (or /chargenow) picks it up automatically
    at that point, see apply_monthly_charge()."""
    shop = get_shop(shop_no)
    if not shop:
        return False
    t = date.today()
    cur_ey, cur_em, _ = gregorian_to_ethiopian(t.year, t.month, t.day)
    current_period_str = f"{cur_ey:04d}-{cur_em:02d}"
    if effective_period is None:
        effective_period = current_period_str
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO rent_history (shop_id, effective_period, rent, created_date) "
            "VALUES (?, ?, ?, ?) "
            "ON CONFLICT(shop_id, effective_period) DO UPDATE SET "
            "rent = excluded.rent, created_date = excluded.created_date",
            (shop["id"], effective_period, new_rent, t.isoformat()),
        )
        if effective_period <= current_period_str:
            conn.execute(
                "UPDATE shops SET monthly_rent = ? WHERE shop_no = ?", (new_rent, shop_no)
            )
    return True


def get_rent_history(shop_no):
    """Every recorded rent change for a shop, most recent effective_period
    first — for showing a shop's rent history (e.g. on its detail page)."""
    shop = get_shop(shop_no)
    if not shop:
        return []
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT effective_period, rent, created_date FROM rent_history "
            "WHERE shop_id = ? ORDER BY effective_period DESC",
            (shop["id"],),
        ).fetchall()
    return [dict(r) for r in rows]


def rent_for_period(shop_id, period):
    """The rent rate that applies for a given Ethiopian 'YYYY-MM' period —
    the most recent rent_history entry with effective_period <= period, or
    the shop's current monthly_rent if it has no rent history at all yet
    (covers shops whose rent has never been changed since this table was
    added). Used when posting a new charge, so a rent change scheduled for
    a future month is honored on exactly the right month — never early,
    never late."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT rent FROM rent_history WHERE shop_id = ? AND effective_period <= ? "
            "ORDER BY effective_period DESC LIMIT 1",
            (shop_id, period),
        ).fetchone()
    if row:
        return row["rent"]
    shop = get_shop_by_id(shop_id)
    return shop["monthly_rent"] if shop else None


def update_shop_no(old_shop_no, new_shop_no):
    with get_conn() as conn:
        conn.execute(
            "UPDATE shops SET shop_no = ? WHERE shop_no = ?", (new_shop_no, old_shop_no)
        )


def update_floor(shop_no, floor):
    with get_conn() as conn:
        conn.execute("UPDATE shops SET floor = ? WHERE shop_no = ?", (floor, shop_no))


def update_area(shop_no, area_sqm):
    with get_conn() as conn:
        conn.execute("UPDATE shops SET area_sqm = ? WHERE shop_no = ?", (area_sqm, shop_no))


def update_tenant_name(shop_no, tenant_name):
    with get_conn() as conn:
        conn.execute("UPDATE shops SET tenant_name = ? WHERE shop_no = ?", (tenant_name, shop_no))


def update_phone(shop_no, phone):
    with get_conn() as conn:
        conn.execute("UPDATE shops SET phone = ? WHERE shop_no = ?", (phone, shop_no))


def update_purpose(shop_no, purpose):
    with get_conn() as conn:
        conn.execute("UPDATE shops SET purpose = ? WHERE shop_no = ?", (purpose, shop_no))


def update_start_date(shop_no, start_date):
    with get_conn() as conn:
        conn.execute(
            "UPDATE shops SET start_date = ?, start_date_ec = ? WHERE shop_no = ?",
            (start_date, gc_iso_to_ec_iso(start_date), shop_no),
        )


def set_shop_active(shop_no, active: bool):
    """Activating clears deactivated_date; deactivating stamps today's date so
    historical reports can tell whether a shop was rented during a past month."""
    with get_conn() as conn:
        if active:
            conn.execute(
                "UPDATE shops SET active = 1, deactivated_date = NULL, "
                "deactivated_date_ec = NULL WHERE shop_no = ?",
                (shop_no,),
            )
        else:
            today_gc = date.today().isoformat()
            conn.execute(
                "UPDATE shops SET active = 0, deactivated_date = ?, "
                "deactivated_date_ec = ? WHERE shop_no = ?",
                (today_gc, gc_iso_to_ec_iso(today_gc), shop_no),
            )


def shops_active_in_range(start_date, end_date):
    """Shops that were active/rented at some point during [start_date, end_date)
    (both 'YYYY-MM-DD'): started before the period ended, and either still
    active or only deactivated on/after the period started."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM shops WHERE start_date < ? "
            "AND (deactivated_date IS NULL OR deactivated_date >= ?) "
            "ORDER BY CAST(shop_no AS INTEGER), shop_no",
            (end_date, start_date),
        ).fetchall()
        return [dict(r) for r in rows]


def delete_shop(shop_no):
    with get_conn() as conn:
        shop = conn.execute(
            "SELECT id FROM shops WHERE shop_no = ?", (shop_no,)
        ).fetchone()
        if not shop:
            return False
        sid = shop["id"]
        conn.execute("DELETE FROM payments WHERE shop_id = ?", (sid,))
        conn.execute("DELETE FROM charges WHERE shop_id = ?", (sid,))
        conn.execute("DELETE FROM shop_documents WHERE shop_id = ?", (sid,))
        conn.execute("DELETE FROM shops WHERE id = ?", (sid,))
        return True


# ---------- Charges (monthly rent) ----------

def apply_monthly_charge(shop, period=None):
    """Insert this month's rent charge for one shop, if not already applied.
    `period` is the Ethiopian 'EY-EM' month key (e.g. '2019-01' for
    Meskerem 2019) — the same tagging bot.py's Telegram interface uses, so
    a charge/payment recorded from either front end lands in the same
    period. Defaults to the Ethiopian month we're in today.

    The amount charged comes from rent_for_period(), not shop['monthly_rent']
    directly — that's what makes a rent change scheduled for a future
    month actually take effect on the right month. Once a charge is
    posted, shops.monthly_rent is synced to match if it doesn't already,
    so "current rent" displays catch up automatically the moment a
    scheduled change's month arrives — no separate job needed."""
    if period is None:
        t = date.today()
        ey, em, _ = gregorian_to_ethiopian(t.year, t.month, t.day)
        period = f"{ey:04d}-{em:02d}"
    rent = rent_for_period(shop["id"], period)
    today_gc = date.today().isoformat()
    with get_conn() as conn:
        try:
            conn.execute(
                "INSERT INTO charges (shop_id, amount, charge_date, charge_date_ec, period, "
                "description) VALUES (?, ?, ?, ?, ?, ?)",
                (shop["id"], rent, today_gc, gc_iso_to_ec_iso(today_gc), period,
                 f"Rent for {period}"),
            )
            if rent != shop["monthly_rent"]:
                conn.execute(
                    "UPDATE shops SET monthly_rent = ? WHERE id = ?", (rent, shop["id"])
                )
            return True
        except sqlite3.IntegrityError:
            return False  # already charged this period


def apply_monthly_charges_to_all(period=None):
    """Only rented shops accrue rent — a vacant shop has no tenant to charge."""
    applied = []
    for shop in get_all_shops(active_only=True, rented_only=True):
        if apply_monthly_charge(shop, period):
            applied.append(shop["shop_no"])
    return applied


# ---------- Payments ----------

def add_payment(shop_no, amount, note=None, payment_date=None, period=None,
                 bank=None, reference_no=None, receipt_file_id=None):
    shop = get_shop(shop_no)
    if not shop:
        return False
    payment_date = payment_date or date.today().isoformat()
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO payments (shop_id, amount, payment_date, payment_date_ec, note, "
            "period, bank, reference_no, receipt_file_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (shop["id"], amount, payment_date, gc_iso_to_ec_iso(payment_date), note, period,
             bank, reference_no, receipt_file_id),
        )
    return True


# ---------- Balances / ledger ----------

def get_balance(shop_no):
    shop = get_shop(shop_no)
    if not shop:
        return None
    with get_conn() as conn:
        charged = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS t FROM charges WHERE shop_id = ?",
            (shop["id"],),
        ).fetchone()["t"]
        paid = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS t FROM payments WHERE shop_id = ?",
            (shop["id"],),
        ).fetchone()["t"]
        return round(charged - paid, 2)


def get_all_balances(active_only=True):
    result = []
    for shop in get_all_shops(active_only=active_only):
        result.append({**shop, "balance": get_balance(shop["shop_no"])})
    return result


def get_ledger(shop_no, limit=10):
    shop = get_shop(shop_no)
    if not shop:
        return []
    with get_conn() as conn:
        charges = conn.execute(
            "SELECT charge_date AS d, charge_date_ec AS d_ec, amount, description, period "
            "FROM charges WHERE shop_id = ? ORDER BY charge_date DESC LIMIT ?",
            (shop["id"], limit),
        ).fetchall()
        payments = conn.execute(
            "SELECT id, payment_date AS d, payment_date_ec AS d_ec, amount, note, period, bank, "
            "reference_no, receipt_file_id FROM payments WHERE shop_id = ? ORDER BY payment_date DESC LIMIT ?",
            (shop["id"], limit),
        ).fetchall()
    already_paid_periods = paid_periods(shop["id"])
    entries = [{"date": r["d"], "date_ec": r["d_ec"], "type": "charge", "amount": r["amount"],
                "note": r["description"], "period": r["period"],
                # A charge's own period showing up again on its own row would just
                # repeat the month a payment row below/above it already shows —
                # so once that period has a payment on record, flag it instead of
                # doubling up the same month label on two rows.
                "already_paid": bool(r["period"]) and r["period"] in already_paid_periods} for r in charges]
    entries += [{"id": r["id"], "date": r["d"], "date_ec": r["d_ec"], "type": "payment", "amount": r["amount"],
                 "note": r["note"], "period": r["period"], "bank": r["bank"],
                 "reference_no": r["reference_no"], "receipt_file_id": r["receipt_file_id"],
                 "already_paid": False} for r in payments]
    entries.sort(key=lambda e: e["date"], reverse=True)
    return entries[:limit]


def get_payment(payment_id):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM payments WHERE id = ?", (payment_id,)).fetchone()
        return dict(row) if row else None


def get_recent_payments(shop_no, limit=10):
    shop = get_shop(shop_no)
    if not shop:
        return []
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM payments WHERE shop_id = ? ORDER BY payment_date DESC, id DESC LIMIT ?",
            (shop["id"], limit),
        ).fetchall()
        return [dict(r) for r in rows]


def update_payment(payment_id, amount=None, payment_date=None, period=None,
                    bank=None, reference_no=None, note=None, receipt_file_id=None):
    """Updates only the fields explicitly passed (None means "leave alone",
    except note/reference_no/bank/period/receipt_file_id which use the
    sentinel _CLEAR to let a caller blank them out on purpose). Passing a
    new receipt_file_id replaces the receipt; the old file (if it was a web
    upload) is left on disk, same as delete_payment does."""
    payment = get_payment(payment_id)
    if not payment:
        return False
    fields = {
        "amount": amount, "payment_date": payment_date, "period": period,
        "bank": bank, "reference_no": reference_no, "note": note,
        "receipt_file_id": receipt_file_id,
    }
    updates, params = [], []
    for col, value in fields.items():
        if value is _CLEAR:
            updates.append(f"{col} = ?")
            params.append(None)
        elif value is not None:
            updates.append(f"{col} = ?")
            params.append(value)
    if payment_date is not None and payment_date is not _CLEAR:
        updates.append("payment_date_ec = ?")
        params.append(gc_iso_to_ec_iso(payment_date))
    if not updates:
        return False
    params.append(payment_id)
    with get_conn() as conn:
        conn.execute(f"UPDATE payments SET {', '.join(updates)} WHERE id = ?", params)
    return True


def delete_payment(payment_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM payments WHERE id = ?", (payment_id,))


def available_charge_years():
    """Distinct Ethiopian years ('YYYY') that have any rent charges on
    record, newest first — for populating a year picker in reports."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT DISTINCT substr(period, 1, 4) AS y FROM charges "
            "WHERE period IS NOT NULL ORDER BY y DESC"
        ).fetchall()
        return [r["y"] for r in rows]


def export_rows_for_year(ec_year):
    """One row per shop for the classic spreadsheet-style rent report:
    shop info plus one column per Ethiopian month (1..13) with that
    month's charged rent amount, read from the charges table's
    'EY-EM' period tag. `ec_year` is an int or numeric string, e.g. 2018."""
    ec_year = int(ec_year)
    with get_conn() as conn:
        shops = conn.execute("SELECT * FROM shops ORDER BY shop_no").fetchall()
        charge_rows = conn.execute(
            "SELECT shop_id, period, amount FROM charges WHERE period LIKE ?",
            (f"{ec_year:04d}-%",),
        ).fetchall()
    charges_by_shop = {}
    for c in charge_rows:
        charges_by_shop.setdefault(c["shop_id"], {})[c["period"]] = c["amount"]
    rows = []
    for s in shops:
        start_ec = s["start_date_ec"]
        if start_ec:
            y, m, d = start_ec.split("-")
            start_display = f"{int(d)}/{int(m)}/{y}"
        else:
            start_display = None
        lease_end_ec = s["lease_end_date_ec"]
        if lease_end_ec:
            y, m, d = lease_end_ec.split("-")
            lease_end_display = f"{int(d)}/{int(m)}/{y}"
        else:
            lease_end_display = None
        months_for_shop = charges_by_shop.get(s["id"], {})
        rows.append({
            "shop_no": s["shop_no"],
            "phone": s["phone"],
            "tenant_name": s["tenant_name"],
            "start_date_ec": start_display,
            "lease_end_date_ec": lease_end_display,
            "monthly_rent": s["monthly_rent"],
            "months": [months_for_shop.get(f"{ec_year:04d}-{m:02d}") for m in range(1, 13)],
        })
    return rows


def payment_punctuality(shop_no, grace_days=7):
    """Grades a tenant on how on-time their payments have been, by comparing
    each charged period's charge_date (the start of that rent period) to the
    earliest payment recorded against that same period.

    Returns None if the shop doesn't exist, otherwise a dict with the grade,
    counts, and the on-time percentage."""
    shop = get_shop(shop_no)
    if not shop:
        return None
    with get_conn() as conn:
        charges = conn.execute(
            "SELECT period, charge_date FROM charges WHERE shop_id = ? ORDER BY period",
            (shop["id"],),
        ).fetchall()
        payments = conn.execute(
            "SELECT period, payment_date FROM payments WHERE shop_id = ? AND period IS NOT NULL",
            (shop["id"],),
        ).fetchall()
    earliest_payment = {}
    for p in payments:
        if not p["payment_date"]:
            continue
        cur = earliest_payment.get(p["period"])
        if cur is None or p["payment_date"] < cur:
            earliest_payment[p["period"]] = p["payment_date"]

    today_iso = date.today().isoformat()
    on_time = late = unpaid = 0
    for c in charges:
        try:
            due = date.fromisoformat(c["charge_date"])
        except (TypeError, ValueError):
            continue
        grace = (due + timedelta(days=grace_days)).isoformat()
        paid_on = earliest_payment.get(c["period"])
        if paid_on:
            if paid_on <= grace:
                on_time += 1
            else:
                late += 1
        elif c["charge_date"] < today_iso:
            unpaid += 1  # charged, due date long passed, still nothing on record
        # else: period hasn't reached its due date yet — not counted either way

    total = on_time + late + unpaid
    late_total = late + unpaid
    if total == 0:
        return {"grade": "New", "on_time": 0, "late": 0, "total": 0, "pct_on_time": None}
    pct = round(100 * on_time / total)
    if pct >= 90:
        grade = "A - Excellent"
    elif pct >= 75:
        grade = "B - Good"
    elif pct >= 50:
        grade = "C - Fair"
    else:
        grade = "D - Poor"
    return {"grade": grade, "on_time": on_time, "late": late_total, "total": total, "pct_on_time": pct}


# ---------- Expenses ----------

PERMANENT_EXPENSE_CATEGORIES = ["Electricity Bill", "Water Bill", "Salary"]


def get_recurring_expenses():
    """All recurring/'permanent' expense definitions (Electricity Bill, Water
    Bill, Salary, ...) — each with the amount that gets applied automatically
    once per Ethiopian month, the same way rent charges do."""
    with get_conn() as conn:
        return conn.execute("SELECT * FROM recurring_expenses ORDER BY category").fetchall()


def get_recurring_expense(category):
    with get_conn() as conn:
        return conn.execute(
            "SELECT * FROM recurring_expenses WHERE category = ?", (category,)
        ).fetchone()


def set_recurring_expense(category, amount, active=True):
    """Create or update a recurring/permanent expense's monthly amount."""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO recurring_expenses (category, amount, active) VALUES (?, ?, ?) "
            "ON CONFLICT(category) DO UPDATE SET amount = excluded.amount, active = excluded.active",
            (category, amount, 1 if active else 0),
        )


def set_recurring_expense_active(category, active: bool):
    with get_conn() as conn:
        conn.execute(
            "UPDATE recurring_expenses SET active = ? WHERE category = ?",
            (1 if active else 0, category),
        )


def apply_recurring_expense(rec, period):
    """Insert this period's amount for one recurring expense definition, if
    not already applied this period. `period` is the Ethiopian 'EY-EM' month
    key, same tagging as apply_monthly_charge."""
    today_gc = date.today().isoformat()
    with get_conn() as conn:
        try:
            conn.execute(
                "INSERT INTO expenses (description, amount, category, expense_date, "
                "expense_date_ec, period, recurring_expense_id) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (rec["category"], rec["amount"], rec["category"], today_gc,
                 gc_iso_to_ec_iso(today_gc), period, rec["id"]),
            )
            return True
        except sqlite3.IntegrityError:
            return False  # already applied this period


def apply_recurring_expenses_to_all(period=None):
    """Log this period's amount for every ACTIVE permanent expense
    (Electricity Bill, Water Bill, Salary, ...) that hasn't been applied yet
    this period — the expense-side equivalent of apply_monthly_charges_to_all.
    Defaults to the Ethiopian month we're in today."""
    if period is None:
        t = date.today()
        ey, em, _ = gregorian_to_ethiopian(t.year, t.month, t.day)
        period = f"{ey:04d}-{em:02d}"
    applied = []
    for rec in get_recurring_expenses():
        if rec["active"] and apply_recurring_expense(rec, period):
            applied.append(rec["category"])
    return applied


def add_expense(description, amount, category=None, expense_date=None, receipt_file_id=None):
    expense_date = expense_date or date.today().isoformat()
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO expenses (description, amount, category, expense_date, expense_date_ec, "
            "receipt_file_id) VALUES (?, ?, ?, ?, ?, ?)",
            (description, amount, category, expense_date, gc_iso_to_ec_iso(expense_date),
             receipt_file_id),
        )


def get_expense(expense_id):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM expenses WHERE id = ?", (expense_id,)).fetchone()
        return dict(row) if row else None


def update_expense(expense_id, description=None, amount=None, category=None,
                   expense_date=None, receipt_file_id=None):
    """Updates only the fields explicitly passed (None means "leave alone").
    category and receipt_file_id also accept the _CLEAR sentinel to blank
    them on purpose; description, amount and expense_date can't be blank.
    Changing expense_date (Gregorian 'YYYY-MM-DD') also refreshes the stored
    Ethiopian date. Returns True if an expense was updated."""
    for name, value in (("description", description), ("amount", amount),
                        ("expense_date", expense_date)):
        if value is _CLEAR:
            raise ValueError(f"{name} can't be cleared")
    if get_expense(expense_id) is None:
        return False
    fields = {
        "description": description, "amount": amount, "category": category,
        "expense_date": expense_date, "receipt_file_id": receipt_file_id,
    }
    updates, params = [], []
    for col, value in fields.items():
        if value is _CLEAR:
            updates.append(f"{col} = ?")
            params.append(None)
        elif value is not None:
            updates.append(f"{col} = ?")
            params.append(value)
    if expense_date is not None:
        updates.append("expense_date_ec = ?")
        params.append(gc_iso_to_ec_iso(expense_date))
    if not updates:
        return False
    params.append(expense_id)
    with get_conn() as conn:
        conn.execute(f"UPDATE expenses SET {', '.join(updates)} WHERE id = ?", params)
    return True


def get_expenses(year, month):
    prefix = f"{year:04d}-{month:02d}"
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM expenses WHERE expense_date LIKE ? ORDER BY expense_date",
            (f"{prefix}%",),
        ).fetchall()
        return [dict(r) for r in rows]


def get_expenses_in_range(start_date, end_date):
    """All expenses with expense_date in the Gregorian [start_date, end_date)
    window — the range-based counterpart to get_expenses(year, month)."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT * FROM expenses WHERE expense_date >= ? AND expense_date < ? "
            "ORDER BY expense_date",
            (start_date, end_date),
        ).fetchall()
        return [dict(r) for r in rows]


def expenses_by_category(start_date, end_date):
    """Expenses in [start_date, end_date) grouped by category — manual
    entries logged with no category are grouped under 'Other'."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT COALESCE(NULLIF(category, ''), 'Other') AS category, "
            "COALESCE(SUM(amount), 0) AS total FROM expenses "
            "WHERE expense_date >= ? AND expense_date < ? "
            "GROUP BY COALESCE(NULLIF(category, ''), 'Other') ORDER BY total DESC",
            (start_date, end_date),
        ).fetchall()
        return [dict(r) for r in rows]


def period_bounds(period):
    """Gregorian [start_date, end_date) span (end exclusive) covered by an
    Ethiopian 'EY-EM' period string — for turning a rent period into real
    calendar dates that shops_active_in_range() and friends understand."""
    ey, em = (int(x) for x in period.split("-"))
    next_ey, next_em = add_ethiopian_months(ey, em, 1)
    sy, sm, sd = ethiopian_to_gregorian(ey, em, 1)
    ny, nm, nd = ethiopian_to_gregorian(next_ey, next_em, 1)
    return date(sy, sm, sd).isoformat(), date(ny, nm, nd).isoformat()


def _periods_covering(start_date, end_date):
    """The list of Ethiopian 'YYYY-MM' periods whose bounds fall inside
    [start_date, end_date) — turns an arbitrary Gregorian window (a
    calendar month or a full Ethiopian year, as used by
    monthly_report_range/yearly_report) into the period keys
    rent_for_period() understands, so "expected" can be computed one
    period at a time even when the caller only has a date range."""
    sy, sm, sd = (int(x) for x in start_date.split("-"))
    ey, em, _ = gregorian_to_ethiopian(sy, sm, sd)
    periods = []
    p_start, _ = period_bounds(f"{ey:04d}-{em:02d}")
    while p_start < end_date:
        periods.append(f"{ey:04d}-{em:02d}")
        ey, em = add_ethiopian_months(ey, em, 1)
        p_start, _ = period_bounds(f"{ey:04d}-{em:02d}")
    return periods


def expected_rent_by_shop_for_period(period):
    """{shop_id: rent} for every shop that was actively rented at some
    point during the Ethiopian 'YYYY-MM' period, using rent_for_period()
    — the rent on record for that period from rent_history — rather than
    whatever has already been posted to `charges`. This is what lets a
    report's "rent expected" be correct as soon as the period starts,
    instead of reading 0 (or last period's total) until the monthly
    charge job actually runs."""
    start_date, end_date = period_bounds(period)
    shops = [s for s in shops_active_in_range(start_date, end_date) if s["is_rented"]]
    result = {}
    for s in shops:
        rent = rent_for_period(s["id"], period)
        if rent is not None:
            result[s["id"]] = rent
    return result


def expected_rent_by_shop_for_range(start_date, end_date):
    """{shop_id: total expected rent} summed across every Ethiopian period
    that falls within [start_date, end_date) — the range-based
    counterpart of expected_rent_by_shop_for_period, for a report window
    that may cover one month or a whole year."""
    totals = {}
    for period in _periods_covering(start_date, end_date):
        for shop_id, rent in expected_rent_by_shop_for_period(period).items():
            totals[shop_id] = totals.get(shop_id, 0.0) + rent
    return totals


def expected_rent_for_period(period, exclude_floors=None):
    """Total rent owed for a single Ethiopian 'YYYY-MM' period — the
    single-period convenience wrapper around expected_rent_by_shop_for_period."""
    by_shop = expected_rent_by_shop_for_period(period)
    if exclude_floors:
        shops_by_id = {s["id"]: s for s in get_all_shops(active_only=False)}
        by_shop = {
            sid: rent for sid, rent in by_shop.items()
            if not (shops_by_id.get(sid) and shops_by_id[sid]["floor"] in exclude_floors)
        }
    return round(sum(by_shop.values()), 2)


# ---------- Reports ----------

def _shop_exclusion_filter(exclude_floors=None, exclude_shop_ids=None):
    """SQL fragment + params that leave whole floors and/or individual shops
    out of a report query. Expects the shops table to be aliased `s`.
    Returns ("", []) when nothing is excluded."""
    clauses, params = [], []
    if exclude_floors:
        ph = ",".join("?" * len(exclude_floors))
        clauses.append(f"AND (s.floor IS NULL OR s.floor NOT IN ({ph}))")
        params += list(exclude_floors)
    if exclude_shop_ids:
        ph = ",".join("?" * len(exclude_shop_ids))
        clauses.append(f"AND s.id NOT IN ({ph})")
        params += list(exclude_shop_ids)
    return " ".join(clauses), params


def monthly_report(year, month, exclude_floors=None):
    """Cash-basis report for a Gregorian calendar month (year, month) — used
    by the admin's Gregorian month/year report picker. Rent periods
    elsewhere are tagged by Ethiopian month, which doesn't line up with a
    Gregorian month, so 'no payment yet' here is based on whether a shop
    made any payment during this Gregorian month at all (not on the
    Ethiopian period tag).

    exclude_floors: optional list of floor keys to leave out entirely —
    e.g. LUMP_SUM_FLOORS (the "special"/ETHIO TERIT and "tsedeybank"/Tsedey
    Bank floors), which pay in a single lump sum for the whole lease
    rather than monthly, so folding their full monthly_rent into a single
    month's "rent expected" would badly skew that figure."""
    prefix = f"{year:04d}-{month:02d}"
    if exclude_floors:
        placeholders = ",".join("?" * len(exclude_floors))
        floor_filter = f"AND (s.floor IS NULL OR s.floor NOT IN ({placeholders}))"
        floor_params = list(exclude_floors)
    else:
        floor_filter = ""
        floor_params = []
    with get_conn() as conn:
        collected = conn.execute(
            f"SELECT COALESCE(SUM(p.amount), 0) AS t FROM payments p "
            f"JOIN shops s ON s.id = p.shop_id "
            f"WHERE p.payment_date LIKE ? {floor_filter}",
            (f"{prefix}%", *floor_params),
        ).fetchone()["t"]
        expenses = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS t FROM expenses WHERE expense_date LIKE ?",
            (f"{prefix}%",),
        ).fetchone()["t"]
        paid_shop_ids = {
            row["shop_id"] for row in conn.execute(
                f"SELECT DISTINCT p.shop_id FROM payments p "
                f"JOIN shops s ON s.id = p.shop_id "
                f"WHERE p.payment_date LIKE ? {floor_filter}",
                (f"{prefix}%", *floor_params),
            ).fetchall()
        }
    rented_shops = get_all_shops(active_only=True, rented_only=True)
    if exclude_floors:
        rented_shops = [s for s in rented_shops if s["floor"] not in exclude_floors]
    # "Expected" is calculated live from each shop's rent on record for
    # whichever Ethiopian period(s) overlap this Gregorian month (via
    # rent_for_period/rent_history) — not from what's already been posted
    # to `charges`. That's what makes this correct as soon as the month
    # starts, rather than reading 0 until the monthly charge job runs.
    month_start = date(year, month, 1).isoformat()
    next_year, next_month = (year + 1, 1) if month == 12 else (year, month + 1)
    month_end = date(next_year, next_month, 1).isoformat()
    rent_expected = expected_rent_for_range(month_start, month_end, exclude_floors=exclude_floors)
    unpaid_count = len([s for s in rented_shops if s["id"] not in paid_shop_ids])
    return {
        "period": prefix,
        "rent_expected": rent_expected,
        "rent_collected": round(collected, 2),
        "expenses": round(expenses, 2),
        "net": round(collected - expenses, 2),
        "unpaid_shops_this_period": unpaid_count,
    }


def payments_for_period_range(start_date, end_date, exclude_floors=None, exclude_shop_ids=None):
    """The individual payment rows counted toward monthly_report_range's
    'rent_collected' for [start_date, end_date) — same matching rules
    (period tag, or payment_date fallback for untagged rows) — so an admin
    can see exactly which payments make up that total instead of just the
    sum. Newest first."""
    floor_filter, floor_params = _shop_exclusion_filter(exclude_floors, exclude_shop_ids)
    periods = _periods_covering(start_date, end_date)
    period_placeholders = ",".join("?" * len(periods)) if periods else "NULL"
    with get_conn() as conn:
        rows = conn.execute(
            f"SELECT s.shop_no, s.tenant_name, s.floor, p.amount, p.payment_date, "
            f"p.payment_date_ec, p.period, p.bank, p.reference_no, p.note "
            f"FROM payments p JOIN shops s ON s.id = p.shop_id "
            f"WHERE (p.period IN ({period_placeholders}) "
            f"OR (p.period IS NULL AND p.payment_date >= ? AND p.payment_date < ?)) {floor_filter} "
            f"ORDER BY s.shop_no, p.payment_date DESC",
            (*periods, start_date, end_date, *floor_params),
        ).fetchall()
    return [dict(r) for r in rows]


def monthly_report_range(start_date, end_date, label, exclude_floors=None, exclude_shop_ids=None):
    floor_filter, floor_params = _shop_exclusion_filter(exclude_floors, exclude_shop_ids)
    # Match payments/charges to this window by their *period* tag (which
    # month's rent they're actually for) rather than the date they were
    # physically recorded on — otherwise a payment made today to catch up
    # on a past month's rent would show up in *today's* report instead of
    # the month it was tagged as covering. Payments with no period tag at
    # all fall back to their payment_date, same as before.
    periods = _periods_covering(start_date, end_date)
    period_placeholders = ",".join("?" * len(periods)) if periods else "NULL"
    with get_conn() as conn:
        collected = conn.execute(
            f"SELECT COALESCE(SUM(p.amount), 0) AS t FROM payments p "
            f"JOIN shops s ON s.id = p.shop_id "
            f"WHERE (p.period IN ({period_placeholders}) "
            f"OR (p.period IS NULL AND p.payment_date >= ? AND p.payment_date < ?)) {floor_filter}",
            (*periods, start_date, end_date, *floor_params),
        ).fetchone()["t"]
        charged = conn.execute(
            f"SELECT COALESCE(SUM(c.amount), 0) AS t FROM charges c "
            f"JOIN shops s ON s.id = c.shop_id "
            f"WHERE (c.period IN ({period_placeholders}) "
            f"OR (c.period IS NULL AND c.charge_date >= ? AND c.charge_date < ?)) {floor_filter}",
            (*periods, start_date, end_date, *floor_params),
        ).fetchone()["t"]
        expenses = conn.execute(
            "SELECT COALESCE(SUM(amount), 0) AS t FROM expenses "
            "WHERE expense_date >= ? AND expense_date < ?",
            (start_date, end_date),
        ).fetchone()["t"]
    balances = get_all_balances()
    if exclude_floors:
        balances = [b for b in balances if b["floor"] not in exclude_floors]
    if exclude_shop_ids:
        balances = [b for b in balances if b["id"] not in exclude_shop_ids]
    total_outstanding = sum(b["balance"] for b in balances)
    rent_expected = expected_rent_for_range(
        start_date, end_date, exclude_floors=exclude_floors, exclude_shop_ids=exclude_shop_ids,
    )
    return {
        "period": label,
        "rent_expected": rent_expected,
        "rent_charged": round(charged, 2),
        "rent_collected": round(collected, 2),
        "rent_not_collected": round(rent_expected - collected, 2),
        "expenses": round(expenses, 2),
        "net": round(collected - expenses, 2),
        "total_outstanding_all_shops": round(total_outstanding, 2),
    }


def ethiopian_year_bounds(ec_year):
    """Gregorian [start_iso, end_iso) bounds for one whole Ethiopian year —
    Meskerem 1 through the following Meskerem 1 (so Pagume is included)."""
    start_gy, start_gm, start_gd = ethiopian_to_gregorian(ec_year, 1, 1)
    end_gy, end_gm, end_gd = ethiopian_to_gregorian(ec_year + 1, 1, 1)
    return date(start_gy, start_gm, start_gd).isoformat(), date(end_gy, end_gm, end_gd).isoformat()


def yearly_report(ec_year, exclude_floors=None, exclude_shop_ids=None):
    """Same shape as monthly_report_range, but for an entire Ethiopian year,
    plus a 'months' breakdown (one monthly_report_range dict per Ethiopian
    month, Meskerem through Pagume) for a month-by-month table."""
    start_iso, end_iso = ethiopian_year_bounds(ec_year)
    label = f"{ec_year} E.C."
    r = monthly_report_range(
        start_iso, end_iso, label, exclude_floors=exclude_floors, exclude_shop_ids=exclude_shop_ids,
    )
    months = []
    for m in range(1, 14):
        m_sy, m_sm, m_sd = ethiopian_to_gregorian(ec_year, m, 1)
        next_ey, next_em = add_ethiopian_months(ec_year, m, 1)
        m_ey, m_em, m_ed = ethiopian_to_gregorian(next_ey, next_em, 1)
        m_start_iso = date(m_sy, m_sm, m_sd).isoformat()
        m_end_iso = date(m_ey, m_em, m_ed).isoformat()
        months.append(monthly_report_range(
            m_start_iso, m_end_iso, f"{ETHIOPIAN_MONTHS[m]} {ec_year}",
            exclude_floors=exclude_floors, exclude_shop_ids=exclude_shop_ids,
        ))
    r["months"] = months
    return r


def report_by_floor(start_date, end_date, period=None, exclude_shop_ids=None):
    """Per-floor breakdown of active shops, rent expected, and rent
    collected for the Gregorian [start_date, end_date) window (the same
    window monthly_report_range / yearly_report use). Every floor is
    included, lump-sum floors too — the point of this report is seeing each
    floor's own figures side by side.

    If `period` (an Ethiopian 'YYYY-MM' string, e.g. from current_period())
    is given, each floor's dict also gets 'paid_count' / 'unpaid_count' —
    how many of that floor's currently-rented shops have/haven't got a
    payment tagged for that period, using the same period-tag matching as
    shops_unpaid_for_period (not payment_date, so a late payment still
    counts as paid for the period it was recorded against). Vacant shops
    (is_rented = 0) aren't counted either way, since no rent is due from
    them. Without a period, those two keys are left off entirely — a
    year-wide window has no single period to check shops against."""
    shops = shops_active_in_range(start_date, end_date)
    if exclude_shop_ids:
        shops = [s for s in shops if s["id"] not in exclude_shop_ids]
    # "Expected" is calculated live from each shop's rent on record for
    # every Ethiopian period inside this window (via rent_for_period),
    # not from whatever's already been posted to `charges` — so it's
    # right even before the monthly charge job has run for the period.
    expected_by_shop = expected_rent_by_shop_for_range(start_date, end_date)
    # "Collected" is matched by each payment's *period* tag, not the date
    # it was physically recorded — a payment made today to catch up on an
    # earlier month's rent should count toward that earlier month, not
    # today's. Untagged payments fall back to payment_date.
    periods = _periods_covering(start_date, end_date)
    period_placeholders = ",".join("?" * len(periods)) if periods else "NULL"
    with get_conn() as conn:
        collected_rows = conn.execute(
            f"SELECT s.id AS shop_id, COALESCE(SUM(p.amount), 0) AS collected "
            f"FROM shops s LEFT JOIN payments p ON p.shop_id = s.id "
            f"AND (p.period IN ({period_placeholders}) "
            f"OR (p.period IS NULL AND p.payment_date >= ? AND p.payment_date < ?)) "
            f"GROUP BY s.id",
            (*periods, start_date, end_date),
        ).fetchall()
        paid_ids = None
        if period:
            paid_ids = {
                row["shop_id"]
                for row in conn.execute(
                    "SELECT DISTINCT shop_id FROM payments WHERE period = ?", (period,)
                ).fetchall()
            }
    collected_by_shop = {row["shop_id"]: row["collected"] for row in collected_rows}
    by_floor = {}
    for s in shops:
        key = s["floor"]
        b = by_floor.setdefault(key, {
            "floor": key, "label": floor_label(key), "shop_count": 0,
            "rent_expected": 0.0, "rent_collected": 0.0,
        })
        if paid_ids is not None:
            b.setdefault("paid_count", 0)
            b.setdefault("unpaid_count", 0)
        b["shop_count"] += 1
        b["rent_expected"] += expected_by_shop.get(s["id"], 0.0)
        b["rent_collected"] += collected_by_shop.get(s["id"], 0.0)
        if paid_ids is not None and s["is_rented"]:
            if s["id"] in paid_ids:
                b["paid_count"] += 1
            else:
                b["unpaid_count"] += 1
    ordered = [by_floor[key] for key, _ in get_floors() if key in by_floor]
    if None in by_floor:
        ordered.append(by_floor[None])
    for b in ordered:
        b["rent_expected"] = round(b["rent_expected"], 2)
        b["rent_collected"] = round(b["rent_collected"], 2)
    return ordered


def expected_rent_for_range(start_date, end_date, exclude_floors=None, exclude_shop_ids=None):
    """Total rent owed across [start_date, end_date), computed live from
    each shop's rent history (via expected_rent_by_shop_for_range) rather
    than from what's already been posted to `charges` — so it's right
    even before the monthly charge job has run for the period(s) in this
    window."""
    shops_by_id = {s["id"]: s for s in get_all_shops(active_only=False)}
    total = 0.0
    for shop_id, rent in expected_rent_by_shop_for_range(start_date, end_date).items():
        shop = shops_by_id.get(shop_id)
        if exclude_floors and shop and shop["floor"] in exclude_floors:
            continue
        if exclude_shop_ids and shop_id in exclude_shop_ids:
            continue
        total += rent
    return round(total, 2)


def has_paid_for_period(shop_no, period, exclude_payment_id=None):
    """True if this shop has at least one payment recorded for the given
    Ethiopian 'YYYY-MM' period (the same tagging used by shops_unpaid_for_period).

    exclude_payment_id: when checking from the payment-edit form, pass the
    id of the payment being edited so it doesn't count as a duplicate of
    itself."""
    shop = get_shop(shop_no)
    if not shop:
        return False
    with get_conn() as conn:
        if exclude_payment_id is None:
            row = conn.execute(
                "SELECT 1 FROM payments WHERE shop_id = ? AND period = ? LIMIT 1",
                (shop["id"], period),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT 1 FROM payments WHERE shop_id = ? AND period = ? AND id != ? LIMIT 1",
                (shop["id"], period, exclude_payment_id),
            ).fetchone()
    return row is not None


def paid_total_for_period(shop_no, period, exclude_payment_id=None):
    """Sum of payment amounts already recorded for a shop's given Ethiopian
    'YYYY-MM' period — used to tell a legitimate installment (period not
    yet fully covered) apart from a redundant duplicate (period already
    fully paid). exclude_payment_id lets the payment-edit form leave out
    the row it's currently editing."""
    shop = get_shop(shop_no)
    if not shop:
        return 0.0
    with get_conn() as conn:
        if exclude_payment_id is None:
            row = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) AS t FROM payments WHERE shop_id = ? AND period = ?",
                (shop["id"], period),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT COALESCE(SUM(amount), 0) AS t FROM payments WHERE shop_id = ? "
                "AND period = ? AND id != ?",
                (shop["id"], period, exclude_payment_id),
            ).fetchone()
    return row["t"] or 0.0


def charges_by_shop_for_period(period):
    """{shop_id: amount actually charged} for the given Ethiopian 'YYYY-MM'
    period — what a shop owed for that specific period, using the amount
    recorded at charge time rather than its current monthly_rent."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT shop_id, COALESCE(SUM(amount), 0) AS t FROM charges "
            "WHERE period = ? GROUP BY shop_id",
            (period,),
        ).fetchall()
    return {r["shop_id"]: r["t"] for r in rows}


def shops_unpaid_for_period(period, start_date=None, end_date=None):
    """Shops with no payment recorded whose `period` (the Ethiopian 'YYYY-MM'
    the admin tagged the payment as being for) matches the given period.
    If start_date/end_date are given, only shops that were actually
    active/rented during that span are considered (so a shop that hadn't
    started yet, or was already deactivated, isn't flagged as unpaid)."""
    if start_date and end_date:
        shops = shops_active_in_range(start_date, end_date)
    else:
        shops = get_all_shops(active_only=True)
    with get_conn() as conn:
        paid_ids = {
            row["shop_id"]
            for row in conn.execute(
                "SELECT DISTINCT shop_id FROM payments WHERE period = ?", (period,)
            ).fetchall()
        }
    return [s for s in shops if s["id"] not in paid_ids]


# ---------- Web app users (separate from the Telegram link-code system) ----------

def any_admin_exists():
    with get_conn() as conn:
        row = conn.execute("SELECT 1 FROM users WHERE role = 'admin' LIMIT 1").fetchone()
        return row is not None


def create_user(username, password_hash, role, shop_id=None, full_name=None):
    """role is 'admin' or 'tenant'. Returns the new user id, or None if the
    username is already taken."""
    with get_conn() as conn:
        try:
            cur = conn.execute(
                "INSERT INTO users (username, password_hash, role, shop_id, full_name, "
                "active, created_at) VALUES (?, ?, ?, ?, ?, 1, ?)",
                (username.strip().lower(), password_hash, role, shop_id, full_name,
                 datetime.now().isoformat(timespec="seconds")),
            )
            return cur.lastrowid
        except sqlite3.IntegrityError:
            return None


def get_user_by_username(username):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE username = ?", (username.strip().lower(),)
        ).fetchone()
        return dict(row) if row else None


def get_user_by_id(user_id):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return dict(row) if row else None


def get_user_by_shop(shop_id):
    with get_conn() as conn:
        row = conn.execute(
            "SELECT * FROM users WHERE shop_id = ? AND role = 'tenant'", (shop_id,)
        ).fetchone()
        return dict(row) if row else None


def list_users(role=None):
    with get_conn() as conn:
        if role:
            rows = conn.execute(
                "SELECT users.*, shops.shop_no AS shop_no FROM users "
                "LEFT JOIN shops ON shops.id = users.shop_id "
                "WHERE role = ? ORDER BY username", (role,),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT users.*, shops.shop_no AS shop_no FROM users "
                "LEFT JOIN shops ON shops.id = users.shop_id ORDER BY role, username"
            ).fetchall()
        return [dict(r) for r in rows]


def set_user_password(user_id, password_hash):
    with get_conn() as conn:
        conn.execute("UPDATE users SET password_hash = ? WHERE id = ?", (password_hash, user_id))


def set_user_active(user_id, active: bool):
    with get_conn() as conn:
        conn.execute("UPDATE users SET active = ? WHERE id = ?", (1 if active else 0, user_id))


def delete_user(user_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM users WHERE id = ?", (user_id,))


def rename_user(user_id, new_username):
    with get_conn() as conn:
        try:
            conn.execute(
                "UPDATE users SET username = ? WHERE id = ?",
                (new_username.strip().lower(), user_id),
            )
            return True
        except sqlite3.IntegrityError:
            return False


def paid_periods(shop_id):
    """Set of Ethiopian 'YYYY-MM' periods this shop has at least one payment
    tagged for."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT DISTINCT period FROM payments WHERE shop_id = ? AND period IS NOT NULL",
            (shop_id,),
        ).fetchall()
        return {row["period"] for row in rows}
