"""
Web app for the shop-rent management system (same rent.db as bot.py).

Run:
    pip install -r requirements.txt
    python webapp.py

Or just double-click webapp.py (or "Start Rent Web App.bat") — see
_bootstrap() below, which behaves the same way bot.py's does: creates a
.env if you don't have one yet, installs whatever's missing, and keeps
the console window open on an early error so you can read it.

This app and bot.py can run at the same time and share the same rent.db —
the web app is a second front door (with full username/password logins
for both the admin and tenants), the Telegram bot keeps sending its
scheduled notifications. Nothing about bot.py needs to change.
"""

import sys
import os
import subprocess
import sqlite3
import tempfile
import zipfile
import hashlib
import math
import re
import json
import threading
import urllib.error
import urllib.parse
import urllib.request
from io import BytesIO


def _bootstrap():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(script_dir, ".env")

    if not os.path.exists(env_path):
        try:
            open(env_path, "a").close()
        except OSError as e:
            print(f"Could not create .env: {e}")
            input("\nPress Enter to exit...")
            sys.exit(1)

    # Make sure a WEB_SECRET_KEY line exists (used to sign login sessions).
    with open(env_path, "r", encoding="utf-8") as f:
        env_text = f.read()
    if "WEB_SECRET_KEY" not in env_text:
        import secrets
        with open(env_path, "a", encoding="utf-8") as f:
            if env_text and not env_text.endswith("\n"):
                f.write("\n")
            f.write(f"WEB_SECRET_KEY={secrets.token_hex(32)}\n")

    try:
        import flask  # noqa: F401
        import dotenv  # noqa: F401
    except ImportError:
        req_path = os.path.join(script_dir, "requirements.txt")
        print("Installing required packages (first run only, this may take a minute)...")
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", "-q", "-r", req_path]
            )
        except subprocess.CalledProcessError as e:
            print(f"Could not install dependencies automatically: {e}")
            print(f"Try running manually: {sys.executable} -m pip install -r requirements.txt")
            input("\nPress Enter to exit...")
            sys.exit(1)
        except FileNotFoundError:
            print("Could not find Python. Install it from https://www.python.org/downloads/")
            input("\nPress Enter to exit...")
            sys.exit(1)
        print("Packages installed.\n")
        import importlib
        importlib.invalidate_caches()


if __name__ == "__main__":
    try:
        _bootstrap()
    except SystemExit:
        raise
    except Exception:
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)

import functools
from datetime import date, datetime

from dotenv import load_dotenv
from flask import (
    Flask, render_template, request, redirect, url_for, session, flash,
    send_from_directory, send_file, abort, g, jsonify,
)
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename

import database as db
import excel_export
from database import (
    gc_iso_to_ec_label, floor_label, ETHIOPIAN_MONTHS,
    gregorian_to_ethiopian, ethiopian_to_gregorian,
)

load_dotenv()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)
ALLOWED_UPLOAD_EXT = {"png", "jpg", "jpeg", "pdf", "webp"}

app = Flask(__name__)
app.secret_key = os.getenv("WEB_SECRET_KEY", "dev-only-insecure-key")
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024  # 12 MB uploads


def money(x):
    if x is None:
        return "-"
    return f"{x:,.2f}"


app.jinja_env.filters["money"] = money
app.jinja_env.filters["ec"] = lambda d: gc_iso_to_ec_label(d) if d else "-"


def period_label(period):
    """'2019-01' -> 'መስከረም 2019' — the human-readable Ethiopian month/year
    for a stored 'EY-EM' period key (see current_period() below for why
    periods are keyed by Ethiopian month). Falls back to the raw value for
    anything that doesn't parse as one — e.g. a handful of payments/charges
    recorded before this app tagged periods by Ethiopian month, which may
    still carry an old Gregorian 'YYYY-MM' key."""
    if not period:
        return "-"
    try:
        ey, em = (int(x) for x in period.split("-"))
        return f"{ETHIOPIAN_MONTHS[em]} {ey}"
    except (ValueError, KeyError):
        return period


app.jinja_env.filters["period_label"] = period_label
app.jinja_env.globals["today"] = lambda: date.today().isoformat()


def current_period():
    """The Ethiopian 'EY-EM' period key for the Ethiopian month we're in
    today — e.g. '2019-01' for Meskerem 2019. This is the same key
    apply_monthly_charge()/add_payment() tag charges and payments with (and
    the same scheme bot.py's Telegram interface already used), so a period
    picked here always lines up with the underlying data. It deliberately
    does NOT use date.today().strftime('%Y-%m'): a single Gregorian month
    can span two different Ethiopian months (most obviously across
    Ethiopian New Year, ~Sept 11 — e.g. Sept 10, 2026 is still Nehase 2018,
    while Sept 15, 2026 is already Meskerem 2019, even though both are
    '2026-09' in the Gregorian calendar)."""
    t = date.today()
    ey, em, _ = gregorian_to_ethiopian(t.year, t.month, t.day)
    return f"{ey:04d}-{em:02d}"


def period_bounds(period):
    """Gregorian [start_iso, end_iso) span (end exclusive) covered by an
    Ethiopian 'EY-EM' period, for filtering by actual calendar dates
    (e.g. which shops were active during that Ethiopian month)."""
    ey, em = (int(x) for x in period.split("-"))
    next_ey, next_em = db.add_ethiopian_months(ey, em, 1)
    sy, sm, sd = ethiopian_to_gregorian(ey, em, 1)
    ny, nm, nd = ethiopian_to_gregorian(next_ey, next_em, 1)
    return date(sy, sm, sd).isoformat(), date(ny, nm, nd).isoformat()


def period_options(months_back=12, months_fwd=12):
    """(value, label) pairs for a month picker, one per Ethiopian month.
    value is the Ethiopian 'EY-EM' period key — the same key used
    everywhere charges/payments are tagged — and label is the Ethiopian
    month/year shown to the admin, e.g. 'መስከረም 2019'."""
    ey0, em0 = (int(x) for x in current_period().split("-"))
    options = []
    for i in range(-months_back, months_fwd + 1):
        ey, em = db.add_ethiopian_months(ey0, em0, i)
        value = f"{ey:04d}-{em:02d}"
        options.append((value, f"{ETHIOPIAN_MONTHS[em]} {ey}"))
    return options


app.jinja_env.globals["period_options"] = period_options
app.jinja_env.globals["current_period"] = current_period
app.jinja_env.globals["ethiopian_months"] = ETHIOPIAN_MONTHS

# The only bank / payment-method choices offered when recording or editing a
# payment. Change the list here and both forms follow.
PAYMENT_METHODS = ["CBE", "TSEDEYBANK", "TELEBIRR", "AWASH", "BOA", "CASH"]
app.jinja_env.globals["payment_methods"] = PAYMENT_METHODS


def tenancy_start_period(shop):
    """The Ethiopian 'EY-EM' period the tenant's lease actually started in,
    from the shop's start_date_ec ('EY-EM-ED'). None if the shop has no
    start date on file (e.g. a vacant shop, or one never assigned a tenant),
    in which case there's nothing to check a payment's period against."""
    start_ec = shop.get("start_date_ec")
    if not start_ec:
        return None
    ey, em, _ = start_ec.split("-")
    return f"{ey}-{em}"


def previous_period(period):
    """The Ethiopian 'EY-EM' period immediately before the given one, e.g.
    '2019-01' -> '2018-13' (Pagume). Used to flag an unpaid prior month
    when a payment is being recorded for a later one."""
    ey, em = (int(x) for x in period.split("-"))
    prev_ey, prev_em = db.add_ethiopian_months(ey, em, -1)
    return f"{prev_ey:04d}-{prev_em:02d}"


def unpaid_periods_before(shop_no, period, start_period, cap=36):
    """Ethiopian 'EY-EM' periods strictly before `period`, oldest first,
    that have no payment on record for this shop — so an admin recording
    a payment for one month can be told about every earlier gap at once,
    not just the single month right before it.

    Stops at the tenant's start_period (nothing owed before a lease
    started), or after `cap` months back if start_period is unknown, so a
    shop with no start date on file can't turn this into an unbounded
    scan."""
    ey, em = (int(x) for x in period.split("-"))
    candidates = []
    cursor_ey, cursor_em = ey, em
    for _ in range(cap):
        cursor_ey, cursor_em = db.add_ethiopian_months(cursor_ey, cursor_em, -1)
        cursor = f"{cursor_ey:04d}-{cursor_em:02d}"
        if start_period and cursor < start_period:
            break
        candidates.append(cursor)
    candidates.reverse()
    return [p for p in candidates if not db.has_paid_for_period(shop_no, p)]


def today_ec():
    t = date.today()
    ey, em, ed = gregorian_to_ethiopian(t.year, t.month, t.day)
    return {"year": ey, "month": em, "day": ed}


app.jinja_env.globals["today_ec"] = today_ec


def payment_status(shop_no, period):
    """'paid', 'pending', or 'late' for a shop's given Ethiopian 'EY-EM'
    period (see current_period() for why it's Ethiopian, not Gregorian).

    - 'paid': a payment is already on record for the period.
    - 'pending': no payment yet, but we're still within the grace window —
      from the start of the Ethiopian month up to (not including) day 30,
      the rent due date.
    - 'late': no payment yet and day 30 (or later) has arrived, or the
      period in question has already fully elapsed.
    """
    if db.has_paid_for_period(shop_no, period):
        return "paid"
    current = current_period()
    if period < current:
        return "late"
    if period > current:
        return "pending"
    return "pending" if today_ec()["day"] < 30 else "late"


app.jinja_env.globals["payment_status"] = payment_status
app.jinja_env.globals["payment_punctuality"] = db.payment_punctuality


def to_ec(date_iso):
    if not date_iso:
        return None
    y, m, d = (int(x) for x in date_iso.split("-"))
    ey, em, ed = gregorian_to_ethiopian(y, m, d)
    return {"year": ey, "month": em, "day": ed}


app.jinja_env.globals["to_ec"] = to_ec


def ec_form_to_gregorian_iso(prefix, form):
    """Reads <prefix>_year/_month/_day from a submitted form (Ethiopian
    calendar) and returns the equivalent Gregorian 'YYYY-MM-DD', or None if
    any part is missing/invalid."""
    y = form.get(f"{prefix}_year", "").strip()
    m = form.get(f"{prefix}_month", "").strip()
    d = form.get(f"{prefix}_day", "").strip()
    if not (y and m and d):
        return None
    try:
        gy, gm, gd = ethiopian_to_gregorian(int(y), int(m), int(d))
        return f"{gy:04d}-{gm:02d}-{gd:02d}"
    except (ValueError, TypeError):
        return None


def ec_form_to_gregorian_iso_strict(prefix, form):
    """Like ec_form_to_gregorian_iso, but rejects dates that don't exist on
    the Ethiopian calendar (day 31, month 14, Pagume 7, ...) instead of
    quietly rolling them into the next month: the date must convert to
    Gregorian and back to exactly what was typed."""
    iso = ec_form_to_gregorian_iso(prefix, form)
    if not iso:
        return None
    try:
        typed = tuple(int(form.get(f"{prefix}_{k}", "").strip()) for k in ("year", "month", "day"))
        gy, gm, gd = (int(x) for x in iso.split("-"))
        if tuple(gregorian_to_ethiopian(gy, gm, gd)) != typed:
            return None
    except (ValueError, TypeError):
        return None
    return iso


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------

def current_user():
    if "user_id" not in session:
        return None
    if not hasattr(g, "_user"):
        g._user = db.get_user_by_id(session["user_id"])
    return g._user


app.jinja_env.globals["current_user"] = current_user


def login_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        user = current_user()
        if not user or not user["active"]:
            session.clear()
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @functools.wraps(view)
    @login_required
    def wrapped(*args, **kwargs):
        if current_user()["role"] != "admin":
            abort(403)
        return view(*args, **kwargs)
    return wrapped


def tenant_required(view):
    @functools.wraps(view)
    @login_required
    def wrapped(*args, **kwargs):
        if current_user()["role"] != "tenant":
            abort(403)
        return view(*args, **kwargs)
    return wrapped


def save_upload(file_storage, subdir):
    """Saves an uploaded file under uploads/<subdir>/ with a collision-safe
    name; returns the relative path to store in the DB, or None."""
    if not file_storage or not file_storage.filename:
        return None
    ext = file_storage.filename.rsplit(".", 1)[-1].lower() if "." in file_storage.filename else ""
    if ext not in ALLOWED_UPLOAD_EXT:
        flash(f"Skipped file '{file_storage.filename}': unsupported type.", "warning")
        return None
    folder = os.path.join(UPLOAD_DIR, subdir)
    os.makedirs(folder, exist_ok=True)
    safe_name = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}_{secure_filename(file_storage.filename)}"
    file_storage.save(os.path.join(folder, safe_name))
    return f"{subdir}/{safe_name}"


# ---------------------------------------------------------------------------
# First-run setup + auth routes
# ---------------------------------------------------------------------------

@app.route("/setup", methods=["GET", "POST"])
def setup():
    if db.any_admin_exists():
        return redirect(url_for("login"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm", "")
        if not username or not password:
            flash("Username and password are required.", "danger")
        elif password != confirm:
            flash("Passwords don't match.", "danger")
        elif len(password) < 6:
            flash("Password should be at least 6 characters.", "danger")
        else:
            uid = db.create_user(username, generate_password_hash(password), "admin",
                                  full_name="Administrator")
            if uid is None:
                flash("That username is taken.", "danger")
            else:
                session["user_id"] = uid
                flash("Admin account created. Welcome!", "success")
                return redirect(url_for("admin_dashboard"))
    return render_template("setup.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if not db.any_admin_exists():
        return redirect(url_for("setup"))
    if current_user():
        return redirect(url_for("index"))
    if request.method == "POST":
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        user = db.get_user_by_username(username)
        if not user or not check_password_hash(user["password_hash"], password):
            flash("Incorrect username or password.", "danger")
        elif not user["active"]:
            flash("This account has been deactivated. Contact the admin.", "danger")
        else:
            session.clear()
            session["user_id"] = user["id"]
            nxt = request.args.get("next")
            return redirect(nxt or url_for("index"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def index():
    user = current_user()
    if not user:
        return redirect(url_for("login"))
    return redirect(url_for("admin_dashboard") if user["role"] == "admin" else url_for("tenant_dashboard"))


@app.route("/files/<path:relpath>")
@login_required
def serve_file(relpath):
    """Serves uploaded receipts/documents. Tenants may only see files that
    belong to their own shop (subfolder is named after the shop_no)."""
    user = current_user()
    if user["role"] == "tenant":
        allowed_prefix = f"shop_{user['shop_id']}/"
        if not relpath.startswith(allowed_prefix):
            abort(403)
    return send_from_directory(UPLOAD_DIR, relpath)


# Files sent through the Telegram bot (payment receipts, expense receipts,
# shop documents) are stored as Telegram file_ids, not files on this server.
# To show them here, the web app asks Telegram for the file once, using the
# bot token from the environment, and keeps a copy under uploads/tg_cache/ so
# every later view is served locally (and the copy is included in the uploads
# backup). Files uploaded through the web app are stored as
# 'shop_<id>/...' or 'expenses/...' paths and are served straight from disk.
TG_CACHE_DIR = os.path.join(UPLOAD_DIR, "tg_cache")
TG_FILE_MAX_BYTES = 20 * 1024 * 1024  # Telegram's own getFile limit
INLINE_SAFE_EXT = ALLOWED_UPLOAD_EXT  # types the browser may display inline


def is_web_upload(file_ref):
    """Web uploads are stored as '<folder>/<name>'; a Telegram file_id never
    contains a slash."""
    return "/" in (file_ref or "")


def fetch_telegram_file(file_id):
    """Local path of the file behind a Telegram file_id, downloading and
    caching it on first use. Returns None if it can't be fetched (no bot
    token configured, Telegram unreachable, file no longer available).
    Never logs or returns URLs, since they contain the bot token."""
    key = hashlib.sha256(file_id.encode("utf-8")).hexdigest()[:32]
    os.makedirs(TG_CACHE_DIR, exist_ok=True)
    for name in os.listdir(TG_CACHE_DIR):
        if name.startswith(key + "."):
            return os.path.join(TG_CACHE_DIR, name)

    token = os.getenv("TELEGRAM_BOT_TOKEN")
    if not token:
        app.logger.warning("Telegram file requested but TELEGRAM_BOT_TOKEN is not set.")
        return None
    try:
        meta_url = (f"https://api.telegram.org/bot{token}/getFile?"
                    + urllib.parse.urlencode({"file_id": file_id}))
        with urllib.request.urlopen(meta_url, timeout=15) as resp:
            meta = json.load(resp)
        if not meta.get("ok"):
            return None
        tg_path = meta["result"]["file_path"]
        ext = os.path.splitext(tg_path)[1].lower()
        if not re.fullmatch(r"\.[a-z0-9]{1,5}", ext):
            ext = ".bin"
        file_url = (f"https://api.telegram.org/file/bot{token}/"
                    + urllib.parse.quote(tg_path))
        with urllib.request.urlopen(file_url, timeout=30) as resp:
            data = resp.read(TG_FILE_MAX_BYTES + 1)
        if not data or len(data) > TG_FILE_MAX_BYTES:
            return None
        final = os.path.join(TG_CACHE_DIR, key + ext)
        tmp = os.path.join(TG_CACHE_DIR, f"tmp_{key}_{os.getpid()}_{threading.get_ident()}")
        with open(tmp, "wb") as fh:
            fh.write(data)
        os.replace(tmp, final)  # atomic, so a half-written file is never served
        return final
    except Exception as exc:  # network, JSON, disk — all mean "couldn't load it"
        app.logger.warning("Could not fetch Telegram file: %s", type(exc).__name__)
        return None


def serve_stored_file(file_ref):
    """Serves whatever a receipt/document column points at, web upload or
    Telegram file_id alike. Only image/PDF types are shown inline; anything
    else (e.g. a .docx or .html sent to the bot as a document) is forced to
    download so it can never run as a page on this site."""
    if is_web_upload(file_ref):
        return send_from_directory(UPLOAD_DIR, file_ref)
    path = fetch_telegram_file(file_ref)
    if not path:
        return render_template(
            "error.html", code=502,
            message="Couldn't load this file from Telegram right now. Try again in a moment.",
        ), 502
    folder, name = os.path.split(path)
    inline_ok = name.rsplit(".", 1)[-1].lower() in INLINE_SAFE_EXT
    if inline_ok:
        return send_from_directory(folder, name)
    return send_from_directory(folder, name, as_attachment=True,
                               mimetype="application/octet-stream")


@app.route("/payments/<int:payment_id>/receipt")
@login_required
def payment_receipt(payment_id):
    """Opens a payment's receipt whether it was uploaded on the web or sent
    to the Telegram bot. Tenants may only open receipts for their own shop."""
    payment = db.get_payment(payment_id)
    if not payment or not payment.get("receipt_file_id"):
        abort(404)
    user = current_user()
    if user["role"] == "tenant" and user["shop_id"] != payment["shop_id"]:
        abort(403)
    return serve_stored_file(payment["receipt_file_id"])


@app.route("/admin/expenses/<int:expense_id>/receipt")
@admin_required
def expense_receipt(expense_id):
    expense = db.get_expense(expense_id)
    if not expense or not expense.get("receipt_file_id"):
        abort(404)
    return serve_stored_file(expense["receipt_file_id"])


@app.route("/documents/<int:doc_id>/file")
@login_required
def shop_document_file(doc_id):
    """Opens a shop document (lease, ID, ...) whether it was uploaded on the
    web or sent to the Telegram bot. Tenants may only open their own shop's."""
    doc = db.get_shop_document(doc_id)
    if not doc:
        abort(404)
    user = current_user()
    if user["role"] == "tenant" and user["shop_id"] != doc["shop_id"]:
        abort(403)
    return serve_stored_file(doc["file_id"])


# ---------------------------------------------------------------------------
# Admin: dashboard
# ---------------------------------------------------------------------------

@app.route("/admin")
@admin_required
def admin_dashboard():
    shops = db.get_all_shops(active_only=True)
    ethio_terit = [s for s in shops if s["floor"] == "special"]
    tsedey_bank = [s for s in shops if s["floor"] == "tsedeybank"]
    other = [s for s in shops if s["floor"] not in db.get_lump_sum_floors()]
    rented = [s for s in other if s["is_rented"]]
    vacant = [s for s in other if not s["is_rented"]]
    y, m = date.today().year, date.today().month
    # LUMP_SUM_FLOORS (ETHIO TERIT and Tsedey Bank) pay their whole lease in
    # one lump sum up front rather than monthly, so they're excluded here —
    # otherwise their full monthly_rent would inflate "rent expected" for
    # whichever single month/period they happened to pay in.
    report = db.monthly_report(y, m, exclude_floors=db.get_lump_sum_floors())
    period = current_period()
    start_iso, end_iso = period_bounds(period)
    unpaid = [
        s for s in db.shops_unpaid_for_period(period, start_iso, end_iso)
        if s["is_rented"] and s["floor"] not in db.get_lump_sum_floors()
    ]
    late_count = sum(1 for s in unpaid if payment_status(s["shop_no"], period) == "late")
    return render_template(
        "admin_dashboard.html",
        shops=shops, vacant_count=len(vacant),
        ethio_terit_count=len(ethio_terit),
        tsedey_bank_count=len(tsedey_bank),
        report=report, unpaid_count=len(unpaid), late_count=late_count, shop_count=len(rented),
    )


# ---------------------------------------------------------------------------
# Admin: shops
# ---------------------------------------------------------------------------

@app.route("/admin/shops")
@admin_required
def admin_shops():
    show_inactive = request.args.get("inactive") == "1"
    grouped = db.get_shops_grouped_by_floor(active_only=not show_inactive)
    return render_template("admin_shops.html", grouped=grouped, show_inactive=show_inactive)


@app.route("/admin/shops/new", methods=["GET", "POST"])
@admin_required
def admin_shop_new():
    if not db.get_floors():
        flash("Add a floor first before adding a shop.", "warning")
        return redirect(url_for("admin_floors"))
    if request.method == "POST":
        shop_no = request.form.get("shop_no", "").strip()
        rent = request.form.get("monthly_rent", "").strip()
        floor = request.form.get("floor") or None
        area = request.form.get("area_sqm", "").strip()
        if not shop_no or not rent:
            flash("Shop number and monthly rent are required.", "danger")
        elif db.get_shop(shop_no):
            flash(f"Shop {shop_no} already exists.", "danger")
        else:
            try:
                rent_val = float(rent)
                area_val = float(area) if area else None
            except ValueError:
                flash("Rent and area must be numbers.", "danger")
                return render_template("admin_shop_new.html", floors=db.get_floors())
            db.add_shop(shop_no, rent_val, floor, area_val)
            flash(f"Shop {shop_no} added (vacant).", "success")
            return redirect(url_for("admin_shop_detail", shop_no=shop_no))
    return render_template("admin_shop_new.html", floors=db.get_floors())


@app.route("/admin/shops/<shop_no>")
@admin_required
def admin_shop_detail(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    ledger = db.get_ledger(shop_no, limit=25)
    documents = db.get_shop_documents(shop_no) if hasattr(db, "get_shop_documents") else []
    account = db.get_user_by_shop(shop["id"])
    period = current_period()
    status = payment_status(shop_no, period)
    grade = db.payment_punctuality(shop_no)
    rent_history = db.get_rent_history(shop_no)
    # What each selectable period's rent actually is, so the payment form can
    # show it next to the amount field as the admin picks a different month.
    rent_by_period = {p: db.rent_for_period(shop["id"], p) for p, _ in period_options()}
    return render_template(
        "admin_shop_detail.html", shop=shop, ledger=ledger, status=status,
        documents=documents, account=account, floors=db.get_floors(), grade=grade,
        rent_history=rent_history, rent_by_period=rent_by_period,
    )


@app.route("/admin/shops/<shop_no>/assign-tenant", methods=["POST"])
@admin_required
def admin_assign_tenant(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    name = request.form.get("tenant_name", "").strip()
    phone = request.form.get("phone", "").strip() or None
    purpose = request.form.get("purpose", "").strip() or None
    start_date = ec_form_to_gregorian_iso("start", request.form)
    if not name:
        flash("Tenant name is required.", "danger")
    else:
        db.assign_tenant(shop_no, name, phone, purpose, start_date)
        flash(f"{name} is now the tenant of shop {shop_no}.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


@app.route("/admin/shops/<shop_no>/vacate", methods=["POST"])
@admin_required
def admin_vacate(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    account = db.get_user_by_shop(shop["id"])
    if account:
        db.delete_user(account["id"])
    db.mark_vacant(shop_no)
    flash(f"Shop {shop_no} is now marked vacant.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


@app.route("/admin/shops/<shop_no>/edit", methods=["POST"])
@admin_required
def admin_shop_edit(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    new_no = request.form.get("shop_no", "").strip()
    rent = request.form.get("monthly_rent", "").strip()
    rent_effective_period = request.form.get("rent_effective_period") or None
    floor = request.form.get("floor") or None
    area = request.form.get("area_sqm", "").strip()
    try:
        if rent:
            rent_val = float(rent)
            # Only log a rent change (and its effective month) if the rent
            # actually changed — this form always resubmits the current
            # rent, so without this check every shop-detail save would
            # write a redundant rent_history row.
            if rent_val != shop["monthly_rent"]:
                db.update_rent(shop_no, rent_val, rent_effective_period)
        if area:
            db.update_area(shop_no, float(area))
    except ValueError:
        flash("Rent and area must be numbers.", "danger")
        return redirect(url_for("admin_shop_detail", shop_no=shop_no))
    db.update_floor(shop_no, floor)
    if new_no and new_no != shop_no:
        db.update_shop_no(shop_no, new_no)
        shop_no = new_no
    flash("Shop details updated.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


@app.route("/admin/shops/<shop_no>/edit-tenant", methods=["POST"])
@admin_required
def admin_tenant_edit(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    name = request.form.get("tenant_name", "").strip()
    phone = request.form.get("phone", "").strip()
    purpose = request.form.get("purpose", "").strip()
    national_id = request.form.get("national_id", "").strip()
    tin_number = request.form.get("tin_number", "").strip()
    start_date = ec_form_to_gregorian_iso("start", request.form)
    lease_end = ec_form_to_gregorian_iso("lease_end", request.form)
    if name:
        db.update_tenant_name(shop_no, name)
    db.update_phone(shop_no, phone or None)
    db.update_purpose(shop_no, purpose or None)
    db.update_national_id(shop_no, national_id or None)
    db.update_tin(shop_no, tin_number or None)
    if start_date:
        db.update_start_date(shop_no, start_date)
    db.set_lease_end(shop_no, lease_end or None)
    flash("Tenant details updated.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


@app.route("/admin/shops/<shop_no>/set-active", methods=["POST"])
@admin_required
def admin_shop_set_active(shop_no):
    active = request.form.get("active") == "1"
    db.set_shop_active(shop_no, active)
    flash(f"Shop {shop_no} {'reactivated' if active else 'deactivated'}.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


@app.route("/admin/shops/<shop_no>/delete", methods=["POST"])
@admin_required
def admin_shop_delete(shop_no):
    account = db.get_user_by_shop(db.get_shop(shop_no)["id"]) if db.get_shop(shop_no) else None
    if account:
        db.delete_user(account["id"])
    db.delete_shop(shop_no)
    flash(f"Shop {shop_no} deleted permanently.", "warning")
    return redirect(url_for("admin_shops"))


@app.route("/admin/shops/<shop_no>/document", methods=["POST"])
@admin_required
def admin_shop_document(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    label = request.form.get("label", "").strip() or "Document"
    f = request.files.get("file")
    rel = save_upload(f, f"shop_{shop['id']}")
    if rel:
        kind = "photo" if rel.rsplit(".", 1)[-1].lower() in ("png", "jpg", "jpeg", "webp") else "document"
        if hasattr(db, "add_shop_document"):
            db.add_shop_document(shop_no, rel, kind, label)
        else:
            db.update_document(shop_no, rel, kind)
        flash("Document uploaded.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


# ---------------------------------------------------------------------------
# Admin: tenant login account for a shop
# ---------------------------------------------------------------------------

@app.route("/admin/shops/<shop_no>/account", methods=["POST"])
@admin_required
def admin_shop_account(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    existing = db.get_user_by_shop(shop["id"])
    if not username or not password:
        flash("Username and password are required.", "danger")
        return redirect(url_for("admin_shop_detail", shop_no=shop_no))
    if len(password) < 6:
        flash("Password should be at least 6 characters.", "danger")
        return redirect(url_for("admin_shop_detail", shop_no=shop_no))
    pw_hash = generate_password_hash(password)
    if existing:
        db.set_user_password(existing["id"], pw_hash)
        if username != existing["username"]:
            if not db.rename_user(existing["id"], username):
                flash("That username is taken.", "danger")
                return redirect(url_for("admin_shop_detail", shop_no=shop_no))
        flash("Tenant login updated.", "success")
    else:
        uid = db.create_user(username, pw_hash, "tenant", shop_id=shop["id"],
                              full_name=shop["tenant_name"])
        if uid is None:
            flash("That username is taken.", "danger")
        else:
            flash("Tenant login created.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


@app.route("/admin/shops/<shop_no>/account/delete", methods=["POST"])
@admin_required
def admin_shop_account_delete(shop_no):
    shop = db.get_shop(shop_no)
    account = db.get_user_by_shop(shop["id"]) if shop else None
    if account:
        db.delete_user(account["id"])
        flash("Tenant login removed.", "success")
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


# ---------------------------------------------------------------------------
# Admin: payments
# ---------------------------------------------------------------------------

def duplicate_payment_reason(shop_no, period, expected, already_paid_total):
    """None if this isn't a duplicate; otherwise the message explaining why
    — shared between the pre-submit AJAX check and the actual POST handler
    so the two never disagree about what counts as 'already fully paid'."""
    if not period or already_paid_total <= 0:
        return None
    if expected is not None and already_paid_total < expected:
        return None  # a legitimate installment, not a duplicate
    return (
        f"{period_label(period)} is already fully paid for shop {shop_no} "
        f"({money(already_paid_total)} recorded)."
    )


def payment_confirm_reasons(shop, period, expected, already_paid_total):
    """Every reason this payment is unusual enough to need an explicit
    'are you sure' before saving: a period before the tenant's lease
    started, a period that's already fully paid, and/or an earlier month
    that still has no payment on record at all (catching the case an admin
    means to pay this month but an earlier one is quietly still owed).
    Empty list means it can save straight through with no prompt. Shared
    by the pre-submit AJAX check and both POST handlers (record + edit) so
    all three always agree on what needs confirming."""
    if not period:
        return []
    reasons = []
    start_period = tenancy_start_period(shop)
    if start_period and period < start_period:
        reasons.append(
            f"Shop {shop['shop_no']}'s current tenant started on "
            f"{period_label(start_period)} — {period_label(period)} is before that."
        )
    dup = duplicate_payment_reason(shop["shop_no"], period, expected, already_paid_total)
    if dup:
        reasons.append(dup)
    gaps = unpaid_periods_before(shop["shop_no"], period, start_period)
    if gaps:
        if len(gaps) == 1:
            reasons.append(
                f"{period_label(gaps[0])} doesn't have a payment on record yet "
                f"for shop {shop['shop_no']}."
            )
        else:
            gap_list = ", ".join(period_label(p) for p in gaps)
            reasons.append(
                f"Shop {shop['shop_no']} has no payment on record for: {gap_list}."
            )
    return reasons


@app.route("/admin/shops/<shop_no>/pay/check_duplicate")
@admin_required
def admin_pay_check_duplicate(shop_no):
    """Used by the payment form's JS to ask, before submitting, whether the
    period picked needs confirmation (before the tenant's start date, or
    already fully paid) — so the admin can be shown the reason and asked
    to confirm instead of just having the save rejected after the fact."""
    shop = db.get_shop(shop_no)
    if not shop:
        return jsonify(duplicate=False)
    period = request.args.get("period", "").strip()
    exclude_id = request.args.get("exclude_payment_id", type=int)
    if not period:
        return jsonify(duplicate=False)
    already_paid_total = db.paid_total_for_period(shop_no, period, exclude_payment_id=exclude_id)
    expected = db.rent_for_period(shop["id"], period)
    reasons = payment_confirm_reasons(shop, period, expected, already_paid_total)
    return jsonify(duplicate=bool(reasons), reason=" ".join(reasons))


@app.route("/admin/shops/<shop_no>/pay", methods=["POST"])
@admin_required
def admin_pay(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        abort(404)
    amount = request.form.get("amount", "").strip()
    period = request.form.get("period", "").strip() or None
    bank = request.form.get("bank") or None
    reference_no = request.form.get("reference_no", "").strip() or None
    payment_date = request.form.get("payment_date") or None
    note = request.form.get("note", "").strip() or None
    confirmed = request.form.get("confirm_duplicate") == "1"
    try:
        amount_val = float(amount)
    except ValueError:
        flash("Enter a valid payment amount.", "danger")
        return redirect(url_for("admin_shop_detail", shop_no=shop_no))
    already_paid_total = db.paid_total_for_period(shop_no, period) if period else 0.0
    expected = db.rent_for_period(shop["id"], period) if period else None
    # Block a payment that needs a second look — before the tenant's lease
    # started, the period's already fully paid, or an earlier month is
    # still unpaid — UNLESS the admin has already been shown the reason
    # (via the form's confirm dialog) and chosen to continue anyway.
    # Anything less than the expected rent for THIS period is a legitimate
    # installment (splitting one month's rent into two or more payments),
    # so that alone never triggers a prompt.
    reasons = payment_confirm_reasons(shop, period, expected, already_paid_total)
    if reasons and not confirmed:
        flash(
            " ".join(reasons) + " If you meant to record it anyway, confirm "
            "and submit again.", "danger",
        )
        return redirect(url_for("admin_shop_detail", shop_no=shop_no))
    receipt_rel = save_upload(request.files.get("receipt"), f"shop_{shop['id']}/receipts")
    db.add_payment(shop_no, amount_val, note=note, payment_date=payment_date, period=period,
                    bank=bank, reference_no=reference_no, receipt_file_id=receipt_rel)
    flash(f"Payment of {money(amount_val)} recorded for shop {shop_no}.", "success")
    if period and expected is not None:
        new_total = round(already_paid_total + amount_val, 2)
        diff = round(new_total - expected, 2)
        if diff == 0:
            flash(f"{period_label(period)} is now fully paid.", "info")
        elif diff > 0:
            flash(
                f"Note: total paid for {period_label(period)} is now {money(new_total)} "
                f"— {money(diff)} over the {money(expected)} rent.", "warning",
            )
        else:
            flash(
                f"Note: {money(abs(diff))} still remaining for {period_label(period)} "
                f"(paid {money(new_total)} of {money(expected)}).", "warning",
            )
    return redirect(url_for("admin_shop_detail", shop_no=shop_no))


@app.route("/admin/payments/<int:payment_id>/edit", methods=["GET", "POST"])
@admin_required
def admin_payment_edit(payment_id):
    payment = db.get_payment(payment_id)
    if not payment:
        abort(404)
    shop = db.get_shop_by_id(payment["shop_id"])
    if request.method == "POST":
        amount = request.form.get("amount", "").strip()
        payment_date = request.form.get("payment_date") or None
        period = request.form.get("period", "").strip()
        bank = request.form.get("bank", "").strip()
        reference_no = request.form.get("reference_no", "").strip()
        note = request.form.get("note", "").strip()
        confirmed = request.form.get("confirm_duplicate") == "1"
        try:
            amount_val = float(amount)
        except ValueError:
            flash("Enter a valid payment amount.", "danger")
            return redirect(url_for("admin_payment_edit", payment_id=payment_id))
        if period:
            other_total = db.paid_total_for_period(shop["shop_no"], period, exclude_payment_id=payment_id)
            expected = db.rent_for_period(shop["id"], period)
            # Same rule as recording a new payment: only block when there's a
            # reason to double-check (before the tenant's start, or the
            # period already fully covered by OTHER payments) AND it hasn't
            # been confirmed yet — so moving a payment into an installment
            # plan still works without any prompt.
            reasons = payment_confirm_reasons(shop, period, expected, other_total)
            if reasons and not confirmed:
                flash(
                    " ".join(reasons) + " Choose a different period, adjust "
                    "things, or confirm and save again if this is intentional.",
                    "danger",
                )
                return redirect(url_for("admin_payment_edit", payment_id=payment_id))
        # Receipt: a newly chosen file replaces (or adds) the receipt; the
        # "remove" box clears it; otherwise it's left exactly as it was.
        # None below means "leave alone" to db.update_payment.
        receipt_value = None
        new_receipt = request.files.get("receipt")
        if new_receipt and new_receipt.filename:
            receipt_value = save_upload(new_receipt, f"shop_{shop['id']}/receipts")
            if receipt_value is None:
                # save_upload already flashed why (unsupported file type);
                # nothing has been saved yet, so let them pick another file.
                return redirect(url_for("admin_payment_edit", payment_id=payment_id))
        elif request.form.get("remove_receipt") == "1":
            receipt_value = db.CLEAR
        db.update_payment(
            payment_id,
            amount=amount_val,
            payment_date=payment_date,
            period=period or db.CLEAR,
            bank=bank or db.CLEAR,
            reference_no=reference_no or db.CLEAR,
            note=note or db.CLEAR,
            receipt_file_id=receipt_value,
        )
        flash("Payment updated.", "success")
        return redirect(url_for("admin_shop_detail", shop_no=shop["shop_no"]))
    return render_template("admin_payment_edit.html", payment=payment, shop=shop)


@app.route("/admin/payments/<int:payment_id>/delete", methods=["POST"])
@admin_required
def admin_payment_delete(payment_id):
    payment = db.get_payment(payment_id)
    if not payment:
        abort(404)
    shop = db.get_shop_by_id(payment["shop_id"])
    db.delete_payment(payment_id)
    flash("Payment deleted.", "warning")
    return redirect(url_for("admin_shop_detail", shop_no=shop["shop_no"]))


@app.route("/admin/late-payments")
@admin_required
def admin_late_payments():
    period = current_period()
    start_iso, end_iso = period_bounds(period)
    unpaid = [s for s in db.shops_unpaid_for_period(period, start_iso, end_iso) if s["is_rented"]]
    rows = [dict(s, status=payment_status(s["shop_no"], period)) for s in unpaid]
    # Late shops first, then pending.
    rows.sort(key=lambda r: 0 if r["status"] == "late" else 1)
    return render_template("admin_late.html", shops=rows, period=period)


# ---------------------------------------------------------------------------
# Admin: expenses
# ---------------------------------------------------------------------------

@app.route("/admin/expenses", methods=["GET", "POST"])
@admin_required
def admin_expenses():
    if request.method == "POST":
        description = request.form.get("description", "").strip()
        amount = request.form.get("amount", "").strip()
        category = request.form.get("category", "").strip() or None
        expense_date = ec_form_to_gregorian_iso("expense", request.form)
        if not description or not amount:
            flash("Description and amount are required.", "danger")
        else:
            try:
                amount_val = float(amount)
            except ValueError:
                flash("Amount must be a number.", "danger")
                return redirect(url_for("admin_expenses"))
            receipt_rel = save_upload(request.files.get("receipt"), "expenses")
            db.add_expense(description, amount_val, category, expense_date, receipt_rel)
            flash("Expense logged.", "success")
        return redirect(url_for("admin_expenses"))
    today = date.today()
    year = int(request.args.get("year", today.year))
    month = int(request.args.get("month", today.month))
    expenses = db.get_expenses(year, month)
    total = sum(e["amount"] for e in expenses)
    return render_template(
        "admin_expenses.html", expenses=expenses, year=year, month=month, total=total,
        recurring_expenses=db.get_recurring_expenses(),
    )


@app.route("/admin/expenses/<int:expense_id>/edit", methods=["GET", "POST"])
@admin_required
def admin_expense_edit(expense_id):
    expense = db.get_expense(expense_id)
    if not expense:
        abort(404)
    back = url_for("admin_expense_edit", expense_id=expense_id)
    if request.method == "POST":
        description = request.form.get("description", "").strip()
        amount = request.form.get("amount", "").strip()
        category = request.form.get("category", "").strip()
        if not description:
            flash("Description is required.", "danger")
            return redirect(back)
        try:
            amount_val = float(amount)
            if not math.isfinite(amount_val):
                raise ValueError
        except ValueError:
            flash("Enter a valid amount.", "danger")
            return redirect(back)
        expense_date = ec_form_to_gregorian_iso_strict("expense", request.form)
        if not expense_date:
            flash("That isn't a valid Ethiopian date — check the day, month and year.", "danger")
            return redirect(back)
        # Receipt: a newly chosen file replaces (or adds) it; the "remove"
        # box clears it; otherwise it's left exactly as it was. None below
        # means "leave alone" to db.update_expense.
        receipt_value = None
        new_receipt = request.files.get("receipt")
        if new_receipt and new_receipt.filename:
            receipt_value = save_upload(new_receipt, "expenses")
            if receipt_value is None:
                # save_upload already flashed why (unsupported file type);
                # nothing has been saved yet, so let them pick another file.
                return redirect(back)
        elif request.form.get("remove_receipt") == "1":
            receipt_value = db.CLEAR
        db.update_expense(
            expense_id,
            description=description,
            amount=amount_val,
            category=category or db.CLEAR,
            expense_date=expense_date,
            receipt_file_id=receipt_value,
        )
        flash("Expense updated.", "success")
        return redirect(url_for("admin_expenses", year=int(expense_date[:4]),
                                month=int(expense_date[5:7])))
    gy, gm, gd = (int(x) for x in expense["expense_date"].split("-"))
    ey, em, ed = gregorian_to_ethiopian(gy, gm, gd)
    recurring = None
    if expense.get("recurring_expense_id"):
        recurring = expense.get("category") or expense.get("description")
    return render_template(
        "admin_expense_edit.html", expense=expense,
        ec_date={"year": ey, "month": em, "day": ed}, recurring=recurring,
    )


@app.route("/editexpense")
@app.route("/editexpense/<int:expense_id>")
@admin_required
def editexpense_shortcut(expense_id=None):
    """Short URL for the same thing as the bot's /editexpense: with an id it
    opens that expense's edit page, without one it goes to the expenses list
    where each row has an Edit button."""
    if expense_id is not None:
        return redirect(url_for("admin_expense_edit", expense_id=expense_id))
    flash("Pick the expense you want to change and press Edit.", "info")
    return redirect(url_for("admin_expenses"))


@app.route("/admin/expenses/permanent", methods=["GET", "POST"])
@admin_required
def admin_permanent_expenses():
    """Manage recurring/'permanent' expenses (Electricity Bill, Water Bill,
    Salary, ...) — each with a monthly amount that gets applied
    automatically on the 1st of every Ethiopian month, the same way rent
    charges are (see database.apply_recurring_expenses_to_all, and the
    monthly_charge_job / 'Charge Rent Now' hooks in bot.py)."""
    if request.method == "POST":
        category = request.form.get("category", "").strip()
        amount = request.form.get("amount", "").strip()
        active = request.form.get("active") == "on"
        if not category or not amount:
            flash("Category and amount are required.", "danger")
        else:
            try:
                amount_val = float(amount)
            except ValueError:
                flash("Amount must be a number.", "danger")
                return redirect(url_for("admin_permanent_expenses"))
            db.set_recurring_expense(category, amount_val, active=active)
            flash(f"{category} set to {money(amount_val)}/month.", "success")
        return redirect(url_for("admin_permanent_expenses"))
    return render_template(
        "admin_permanent_expenses.html",
        recurring=db.get_recurring_expenses(),
        suggested_categories=db.PERMANENT_EXPENSE_CATEGORIES,
    )


# ---------------------------------------------------------------------------
# Admin: reports
# ---------------------------------------------------------------------------

def _excluded_shop_ids():
    """Shops the admin has unticked in the report's shop picker.

    Two ways in, both giving the same answer:
      * the picker form itself submits every shop it listed (`listed`) and
        the ones still ticked (`shop`) — the difference is what's excluded.
        Working from `listed` means a shop that wasn't on the page when the
        form was built (e.g. a different month's list) is never dropped by
        accident;
      * links (Month/Year tabs, Excel download) carry the result forward as
        repeated `exclude=<shop id>` params.
    """
    listed = request.args.getlist("listed", type=int)
    if listed:
        selected = set(request.args.getlist("shop", type=int))
        return sorted(set(listed) - selected)
    return sorted(set(request.args.getlist("exclude", type=int)))


def _report_shop_groups(start_iso, end_iso, exclude_floors, excluded):
    """Shops for the report's tick-list, grouped by floor in building order.
    Shops on a floor already left out by the 'Without Special Shops' view
    aren't listed — they're out of the report either way."""
    shops = db.shops_active_in_range(start_iso, end_iso)
    if exclude_floors:
        shops = [x for x in shops if x["floor"] not in exclude_floors]
    by_floor = {}
    for shop in shops:
        by_floor.setdefault(shop["floor"], []).append(shop)
    groups = [(label, by_floor[key]) for key, label in db.get_floors() if key in by_floor]
    if None in by_floor:
        groups.append(("Unassigned", by_floor[None]))
    included = len([x for x in shops if x["id"] not in set(excluded)])
    return groups, len(shops), included


@app.route("/admin/report")
@admin_required
def admin_report():
    period_type = request.args.get("period_type", "month")
    scope = request.args.get("scope", "all")
    exclude_floors = db.get_lump_sum_floors() if scope == "no_special" else None
    excluded = _excluded_shop_ids()

    if period_type == "year":
        cur_ey, _ = (int(x) for x in current_period().split("-"))
        ey = int(request.args.get("year", cur_ey))
        start_iso, end_iso = db.ethiopian_year_bounds(ey)
        report = db.yearly_report(ey, exclude_floors=exclude_floors, exclude_shop_ids=excluded)
        shop_groups, shop_total, shop_included = _report_shop_groups(
            start_iso, end_iso, exclude_floors, excluded)
        return render_template(
            "admin_report.html", report=report, year=ey, month=None,
            period_type="year", scope=scope, payments=None,
            shop_groups=shop_groups, shop_total=shop_total,
            shop_included=shop_included, excluded=excluded,
        )

    cur_ey, cur_em = (int(x) for x in current_period().split("-"))
    ey = int(request.args.get("year", cur_ey))
    em = int(request.args.get("month", cur_em))
    period = f"{ey:04d}-{em:02d}"
    start_iso, end_iso = period_bounds(period)
    label = f"{ETHIOPIAN_MONTHS[em]} {ey}"
    report = db.monthly_report_range(
        start_iso, end_iso, label, exclude_floors=exclude_floors, exclude_shop_ids=excluded,
    )
    unpaid = db.shops_unpaid_for_period(period, start_iso, end_iso)
    if exclude_floors:
        unpaid = [s for s in unpaid if s["floor"] not in exclude_floors]
    if excluded:
        unpaid = [s for s in unpaid if s["id"] not in excluded]
    report["unpaid_shops_this_period"] = len([s for s in unpaid if s["is_rented"]])
    payments = db.payments_for_period_range(
        start_iso, end_iso, exclude_floors=exclude_floors, exclude_shop_ids=excluded,
    )
    shop_groups, shop_total, shop_included = _report_shop_groups(
        start_iso, end_iso, exclude_floors, excluded)
    return render_template(
        "admin_report.html", report=report, year=ey, month=em,
        period_type="month", scope=scope, payments=payments,
        shop_groups=shop_groups, shop_total=shop_total,
        shop_included=shop_included, excluded=excluded,
    )


@app.route("/admin/report/export.xlsx")
@admin_required
def admin_report_export():
    """The full-year, spreadsheet-style ledger table (shop info + one column
    per Ethiopian month) as a downloadable .xlsx — the same layout as the
    building's original manual ledger. See admin_report_export_financial for
    the collections-vs-expenses report instead."""
    cur_ey, _ = (int(x) for x in current_period().split("-"))
    ec_year = int(request.args.get("year", cur_ey))
    buf = excel_export.build_workbook(ec_year)
    return send_file(
        buf, as_attachment=True,
        download_name=f"rent_{ec_year}_EC.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/admin/report/export_financial.xlsx")
@admin_required
def admin_report_export_financial():
    """The /admin/report screen itself (collections vs expenses, by floor,
    by month for a year) as a downloadable .xlsx."""
    period_type = request.args.get("period_type", "month")
    scope = request.args.get("scope", "all")
    exclude_floors = db.get_lump_sum_floors() if scope == "no_special" else None
    cur_ey, cur_em = (int(x) for x in current_period().split("-"))
    if period_type == "year":
        ey = int(request.args.get("year", cur_ey))
        kind, key = "year", ey
    else:
        ey = int(request.args.get("year", cur_ey))
        em = int(request.args.get("month", cur_em))
        kind, key = "month", f"{ey:04d}-{em:02d}"
    buf = excel_export.build_report_workbook(
        kind, key, exclude_floors=exclude_floors, exclude_shop_ids=_excluded_shop_ids(),
    )
    return send_file(
        buf, as_attachment=True,
        download_name=f"rent_report_{key}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Admin: backups — download a copy of the live data. Use this before any
# redeploy, host migration, or container rebuild that isn't guaranteed to
# preserve rent.db / uploads/ on a persistent volume, so you always have a
# fallback if the running container's filesystem doesn't survive.
# ---------------------------------------------------------------------------

@app.route("/admin/backup/database")
@admin_required
def admin_backup_database():
    """Download a point-in-time copy of rent.db. Uses SQLite's own backup
    API rather than just copying the file, so a write happening at the
    same moment (a payment being saved, say) can't produce a corrupted
    copy."""
    tmp = tempfile.NamedTemporaryFile(suffix=".db", delete=False)
    tmp_path = tmp.name
    tmp.close()
    src = sqlite3.connect(db.DB_PATH)
    dst = sqlite3.connect(tmp_path)
    try:
        src.backup(dst)
    finally:
        dst.close()
        src.close()
    stamp = date.today().isoformat()
    response = send_file(
        tmp_path, as_attachment=True,
        download_name=f"rent_backup_{stamp}.db",
        mimetype="application/octet-stream",
    )
    response.call_on_close(lambda: os.path.exists(tmp_path) and os.remove(tmp_path))
    return response


@app.route("/admin/backup/uploads.zip")
@admin_required
def admin_backup_uploads():
    """Download every uploaded receipt/document (the uploads/ folder) as
    one .zip — same idea as the database backup above."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(UPLOAD_DIR):
            for fname in files:
                full = os.path.join(root, fname)
                rel = os.path.relpath(full, UPLOAD_DIR)
                zf.write(full, rel)
    buf.seek(0)
    stamp = date.today().isoformat()
    return send_file(
        buf, as_attachment=True,
        download_name=f"uploads_backup_{stamp}.zip",
        mimetype="application/zip",
    )


# ---------------------------------------------------------------------------
# Admin: users (web logins)
# ---------------------------------------------------------------------------

@app.route("/admin/users")
@admin_required
def admin_users():
    users = db.list_users()
    return render_template("admin_users.html", users=users)


@app.route("/admin/users/new", methods=["POST"])
@admin_required
def admin_users_new():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    role = request.form.get("role")
    if role != "admin":
        flash("Use a shop's page to create tenant logins.", "danger")
        return redirect(url_for("admin_users"))
    if not username or len(password) < 6:
        flash("Username and a password of at least 6 characters are required.", "danger")
    else:
        uid = db.create_user(username, generate_password_hash(password), "admin")
        flash("Admin account created." if uid else "That username is taken.",
              "success" if uid else "danger")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:user_id>/toggle", methods=["POST"])
@admin_required
def admin_users_toggle(user_id):
    if user_id == current_user()["id"]:
        flash("You can't deactivate your own account.", "danger")
        return redirect(url_for("admin_users"))
    user = db.get_user_by_id(user_id)
    if user:
        db.set_user_active(user_id, not user["active"])
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:user_id>/delete", methods=["POST"])
@admin_required
def admin_users_delete(user_id):
    if user_id == current_user()["id"]:
        flash("You can't delete your own account.", "danger")
        return redirect(url_for("admin_users"))
    db.delete_user(user_id)
    flash("Account deleted.", "success")
    return redirect(url_for("admin_users"))


# ---------------------------------------------------------------------------
# Admin: floors
# ---------------------------------------------------------------------------

@app.route("/admin/floors")
@admin_required
def admin_floors():
    return render_template("admin_floors.html", floors=db.get_floors_with_counts())


@app.route("/admin/floors/new", methods=["POST"])
@admin_required
def admin_floors_new():
    label = request.form.get("label", "").strip()
    lump_sum = request.form.get("lump_sum") == "1"
    if not label:
        flash("Floor name is required.", "danger")
    elif label.lower() in {l.lower() for _, l in db.get_floors()}:
        flash(f"A floor called '{label}' already exists.", "danger")
    else:
        db.add_floor(label, lump_sum=lump_sum)
        flash(f"Floor '{label}' added.", "success")
    return redirect(url_for("admin_floors"))


@app.route("/admin/floors/<key>/edit", methods=["POST"])
@admin_required
def admin_floors_edit(key):
    f = db.get_floor(key)
    if not f:
        abort(404)
    label = request.form.get("label", "").strip()
    lump_sum = request.form.get("lump_sum") == "1"
    if not label:
        flash("Floor name is required.", "danger")
    elif label.lower() != f["label"].lower() and label.lower() in {l.lower() for k, l in db.get_floors() if k != key}:
        flash(f"A floor called '{label}' already exists.", "danger")
    else:
        db.update_floor_label(key, label)
        db.update_floor_lump_sum(key, lump_sum)
        flash(f"Floor '{label}' updated.", "success")
    return redirect(url_for("admin_floors"))


@app.route("/admin/floors/<key>/move", methods=["POST"])
@admin_required
def admin_floors_move(key):
    direction = request.form.get("direction")
    if direction in ("up", "down"):
        db.move_floor(key, direction)
    return redirect(url_for("admin_floors"))


@app.route("/admin/floors/<key>/delete", methods=["POST"])
@admin_required
def admin_floors_delete(key):
    f = db.get_floor(key)
    if not f:
        abort(404)
    try:
        db.delete_floor(key)
        flash(f"Floor '{f['label']}' deleted.", "success")
    except ValueError as e:
        flash(f"Can't delete '{f['label']}' — {e.args[0]} shop(s) are still on it. Move or deactivate them first.", "danger")
    return redirect(url_for("admin_floors"))


@app.route("/admin/floors/delete-all", methods=["POST"])
@admin_required
def admin_floors_delete_all():
    deleted, kept = [], []
    for f in db.get_floors_with_counts():
        if f["shop_count"]:
            kept.append(f["label"])
        else:
            db.delete_floor(f["key"])
            deleted.append(f["label"])
    if deleted:
        flash(f"Deleted: {', '.join(deleted)}.", "success")
    if kept:
        flash(f"Kept (still has shops): {', '.join(kept)}.", "warning")
    if not deleted and not kept:
        flash("No floors to delete.", "info")
    return redirect(url_for("admin_floors"))


@app.route("/account/password", methods=["POST"])
@login_required
def account_password():
    user = current_user()
    current = request.form.get("current_password", "")
    new = request.form.get("new_password", "")
    confirm = request.form.get("confirm_password", "")
    if not check_password_hash(user["password_hash"], current):
        flash("Current password is incorrect.", "danger")
    elif len(new) < 6:
        flash("New password should be at least 6 characters.", "danger")
    elif new != confirm:
        flash("New passwords don't match.", "danger")
    else:
        db.set_user_password(user["id"], generate_password_hash(new))
        flash("Password updated.", "success")
    return redirect(url_for("admin_dashboard") if user["role"] == "admin" else url_for("tenant_dashboard"))


# ---------------------------------------------------------------------------
# Tenant views
# ---------------------------------------------------------------------------

@app.route("/me")
@tenant_required
def tenant_dashboard():
    user = current_user()
    shop = db.get_shop_by_id(user["shop_id"]) if user["shop_id"] else None
    if not shop:
        flash("Your account isn't linked to a shop. Contact the admin.", "warning")
        return render_template("tenant_dashboard.html", shop=None, status=None, ledger=[])
    period = current_period()
    status = payment_status(shop["shop_no"], period)
    ledger = db.get_ledger(shop["shop_no"], limit=10)
    return render_template("tenant_dashboard.html", shop=shop, status=status, ledger=ledger)


@app.route("/me/ledger")
@tenant_required
def tenant_ledger():
    user = current_user()
    shop = db.get_shop_by_id(user["shop_id"]) if user["shop_id"] else None
    if not shop:
        abort(404)
    ledger = db.get_ledger(shop["shop_no"], limit=200)
    return render_template("tenant_ledger.html", shop=shop, ledger=ledger)


@app.route("/me/shop")
@tenant_required
def tenant_shop():
    user = current_user()
    shop = db.get_shop_by_id(user["shop_id"]) if user["shop_id"] else None
    if not shop:
        abort(404)
    documents = db.get_shop_documents(shop["shop_no"]) if hasattr(db, "get_shop_documents") else []
    return render_template("tenant_shop.html", shop=shop, documents=documents)


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------

@app.errorhandler(403)
def forbidden(e):
    return render_template("error.html", code=403, message="You don't have access to that page."), 403


@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", code=404, message="That page doesn't exist."), 404


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    db.init_db()
    host = os.getenv("WEB_HOST", "0.0.0.0")
    port = int(os.getenv("PORT") or os.getenv("WEB_PORT", "5000"))

    url = f"http://localhost:{port}"
    print(f"Rent Manager web app starting at {url}")
    print("Leave this window open while you want the web app running. Closing it stops it.")

    try:
        import webbrowser
        import threading
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    except Exception:
        pass

    try:
        from waitress import serve
        serve(app, host=host, port=port)
    except ImportError:
        app.run(host=host, port=port, debug=False)


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (0, None):
            print(f"\n{e}" if str(e) else "")
            input("Press Enter to exit...")
        raise
    except KeyboardInterrupt:
        pass
    except Exception:
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)
