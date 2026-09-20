"""
Telegram bot for a 30-shop building rent management system.

Everything is menu-driven:
  - A persistent button menu at the bottom of the chat (role-specific).
  - Telegram's native "/" command menu (also role-specific).
  - Multi-step guided flows: tap a button, then answer one question at a
    time instead of typing a long command with arguments.

Two roles:
  - ADMIN  : full control (defined by TELEGRAM_ADMIN_IDS in .env)
  - TENANT : a shop owner who linked their Telegram account with /register <code>

Run:
    pip install -r requirements.txt
    cp .env.example .env   # fill in your values
    python bot.py

You can also just double-click this file. See _bootstrap() below - on first
run it creates .env for you, installs whatever's missing from
requirements.txt, and keeps the window open if something goes wrong so you
can read the error instead of it flashing shut.
"""

import sys
import os
import subprocess


def _bootstrap():
    """Runs before anything else when this file is executed directly
    (e.g. double-clicked). Makes sure .env exists and dependencies are
    installed, and keeps the console window open on an early failure
    instead of letting it vanish."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    env_path = os.path.join(script_dir, ".env")
    example_path = os.path.join(script_dir, ".env.example")

    if not os.path.exists(env_path):
        try:
            if os.path.exists(example_path):
                import shutil
                shutil.copy(example_path, env_path)
            else:
                open(env_path, "a").close()
        except OSError as e:
            print(f"Could not create .env: {e}")
            input("\nPress Enter to exit...")
            sys.exit(1)

        print("A .env file was just created for you in this folder.")
        print("Open it, fill in TELEGRAM_BOT_TOKEN and TELEGRAM_ADMIN_IDS, save it,")
        print("then run this file again.")
        if os.name == "nt":
            try:
                subprocess.Popen(["notepad.exe", env_path])
            except OSError:
                pass
        input("\nPress Enter to exit...")
        sys.exit(0)

    try:
        import dotenv  # noqa: F401
        import telegram  # noqa: F401
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
            print('During install, check "Add python.exe to PATH".')
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

import asyncio
import logging
import io
import re
from datetime import date, datetime
from functools import wraps

from dotenv import load_dotenv
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    BotCommand,
    BotCommandScopeChat,
    InputFile,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    ContextTypes,
    ExtBot,
    filters,
)
from telegram.error import TelegramError, TimedOut, NetworkError
from telegram.request import HTTPXRequest

import database as db
import excel_export
from database import (
    ETHIOPIAN_MONTHS,
    gregorian_to_ethiopian,
    ethiopian_to_gregorian,
    format_ethiopian_date,
    gc_iso_to_ec_label,
    gc_iso_to_ec_iso,
)

load_dotenv()

try:
    import shutil as _shutil
    import pytesseract
    from PIL import Image
    # pytesseract (the pip package) can import fine even when the actual
    # `tesseract` binary isn't installed on the system (e.g. a plain Python
    # buildpack that only runs `pip install`, with no apt/system-package
    # step). Calling pytesseract.image_to_string() in that case doesn't
    # crash — it's already wrapped in a try/except below — but it wastes a
    # full image decode + subprocess-spawn attempt on every receipt photo
    # before failing. Checking for the binary up front skips that wasted
    # work and makes OCR_AVAILABLE reflect reality: OCR only reports
    # "available" when it can actually run.
    OCR_AVAILABLE = _shutil.which("tesseract") is not None
except ImportError:
    OCR_AVAILABLE = False

BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
ADMIN_IDS = {
    int(x) for x in os.getenv("TELEGRAM_ADMIN_IDS", "").replace(" ", "").split(",") if x
}

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def money(x: float) -> str:
    return f"{x:,.2f}"


def _shop_no_key(shop_no):
    """Numeric-first sort key for a shop number, so '2' sorts before '10'
    instead of a plain string sort putting '10' first. Shop numbers that
    aren't purely numeric fall back to sorting by their text, after every
    numeric one."""
    s = str(shop_no)
    return (0, int(s)) if s.isdigit() else (1, s)


def _field(row, key, default=None):
    """Safe lookup for a DB row that may not have this column yet (e.g. on a
    database.py that hasn't been migrated for a newly added field)."""
    try:
        value = row[key]
        return default if value is None else value
    except (IndexError, KeyError, TypeError):
        return default


def _shop_documents(shop_no):
    """List of {'file_id', 'kind', 'label'} dicts for a shop. Uses
    db.get_shop_documents if database.py has a multi-document table;
    otherwise falls back to the single legacy document_file_id/kind field
    so nothing breaks pre-migration."""
    try:
        docs = db.get_shop_documents(shop_no)
        if docs:
            return list(docs)
    except AttributeError:
        pass
    shop = db.get_shop(shop_no)
    if shop and shop["document_file_id"]:
        return [{"file_id": shop["document_file_id"], "kind": shop["document_kind"], "label": "Document"}]
    return []


# Files uploaded through the web app are stored as 'shop_<id>/<name>' paths
# under uploads/ (next to this script), whereas files sent to this bot are
# stored as Telegram file_ids. Telegram file_ids never contain a slash.
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")


def _local_upload_path(file_ref):
    """Absolute path of a web-uploaded file if it exists inside uploads/,
    else None (also None for anything that would escape that folder)."""
    try:
        root = os.path.realpath(UPLOAD_DIR)
        full = os.path.realpath(os.path.join(root, file_ref))
        if os.path.commonpath([root, full]) != root or not os.path.isfile(full):
            return None
    except (ValueError, OSError):
        return None
    return full


async def _send_stored_document(message, doc, caption):
    """Replies with a shop document, whether it was sent to the bot (a
    Telegram file_id) or uploaded on the web app (a file on this server)."""
    ref, kind = doc["file_id"], doc.get("kind")
    if "/" not in ref:  # already on Telegram's servers
        if kind == "photo":
            await message.reply_photo(ref, caption=caption)
        else:
            await message.reply_document(ref, caption=caption)
        return
    path = _local_upload_path(ref)
    if not path:
        await message.reply_text(f"{caption}\n(The uploaded file is missing on the server.)")
        return
    if kind == "photo":
        try:
            with open(path, "rb") as fh:
                await message.reply_photo(fh, caption=caption)
            return
        except TelegramError as e:
            logger.warning("Sending %s as a photo failed (%s); sending as a file instead.",
                           os.path.basename(path), e)
    with open(path, "rb") as fh:
        await message.reply_document(fh, caption=caption, filename=os.path.basename(path))


def _save_shop_document(shop_no, file_id, kind, label=None):
    """Adds one more document for a shop without discarding earlier ones.
    Uses db.add_shop_document if available (multi-document schema);
    otherwise falls back to db.update_document, which only keeps the single
    most recent document — see the note added to shop_detail_text/profile
    about migrating database.py for full multi-document support."""
    try:
        db.add_shop_document(shop_no, file_id, kind, label)
    except AttributeError:
        db.update_document(shop_no, file_id, kind)


# ---------------------------------------------------------------------------
# "One screen at a time": the chat never accumulates old messages. Every
# message the bot sends to a chat deletes whatever it sent there last, and
# every message the user sends gets deleted right after it's handled — so
# at any moment there's only ever the bot's current screen on the tenant's
# or admin's phone, nothing older.
#
# This is deliberately a "delete the previous one, always" policy: a few
# spots in this bot send more than one message on purpose (e.g. a payment
# receipt photo followed by "Back to menu", or notifying a tenant with two
# messages). Under this policy, whichever of those goes out first will be
# replaced by the one right after it almost immediately — that trade-off
# was a deliberate choice, not an oversight.
#
# Telegram allows both halves of this: a bot can delete its own messages,
# and can delete a user's incoming messages, in a private chat, as long as
# the message is less than 48 hours old. Any delete that fails for another
# reason (already gone, chat migrated, etc.) is just ignored.
# ---------------------------------------------------------------------------

class AutoCleanBot(ExtBot):
    """Bot subclass that keeps each chat down to one message. Every
    send_message / send_photo / send_document call deletes whatever this
    bot last sent to that chat, right after the new one goes out. Editing
    an existing message in place (inline-keyboard menus updated via
    edit_message_text) needs no extra handling — the message id doesn't
    change, so the tracked "last message" is still correct."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._last_screen_msg: dict[int, int] = {}

    @staticmethod
    def _chat_id_of(args, kwargs):
        return kwargs.get("chat_id", args[0] if args else None)

    async def _send_and_replace(self, chat_id, send_coro):
        msg = await send_coro
        old_id = self._last_screen_msg.get(chat_id)
        if old_id and old_id != msg.message_id:
            try:
                await self.delete_message(chat_id, old_id)
            except TelegramError:
                pass  # already deleted, too old (>48h), or otherwise gone
        if chat_id is not None:
            self._last_screen_msg[chat_id] = msg.message_id
        return msg

    async def send_message(self, *args, **kwargs):
        chat_id = self._chat_id_of(args, kwargs)
        return await self._send_and_replace(chat_id, super().send_message(*args, **kwargs))

    async def send_photo(self, *args, **kwargs):
        chat_id = self._chat_id_of(args, kwargs)
        return await self._send_and_replace(chat_id, super().send_photo(*args, **kwargs))

    async def send_document(self, *args, **kwargs):
        chat_id = self._chat_id_of(args, kwargs)
        return await self._send_and_replace(chat_id, super().send_document(*args, **kwargs))


async def clean_incoming(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Deletes the user's own message after everything else has had a
    chance to handle it. Registered in a late handler group (see main()),
    so it runs after the normal handlers — deleting a message from Telegram
    doesn't remove it from this Update object, so nothing upstream is
    affected by it disappearing right after."""
    msg = update.message
    if msg is None:
        return
    try:
        await msg.delete()
    except TelegramError:
        pass  # already gone, too old, etc. — nothing to do


async def send(update: Update, text: str, **kwargs):
    """Reply whether the update came from a typed message or a tapped button."""
    if update.message:
        await update.message.reply_text(text, **kwargs)
    else:
        try:
            await update.callback_query.answer()
        except Exception:
            pass  # already answered (e.g. re-used after a switch-flow confirmation) — harmless
        await update.callback_query.message.reply_text(text, **kwargs)


def admin_only(func):
    @wraps(func)
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        if not is_admin(update.effective_user.id):
            await send(update, "This is for building admins only.")
            return ConversationHandler.END
        return await func(update, context)
    return wrapper


def tenant_only(func):
    @wraps(func)
    async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
        shop = db.get_shop_by_telegram_id(update.effective_user.id)
        if not shop:
            await update.message.reply_text(
                "Your Telegram account isn't linked to a shop yet. Tap Register below, "
                "or send /register <code>.",
                reply_markup=unregistered_menu_kb(),
            )
            return
        return await func(update, context, shop)
    return wrapper


# ---------------------------------------------------------------------------
# Interrupting a flow that's already in progress
#
# Each guided admin flow below is tagged with a short label. When a flow is
# active and the admin taps/sends a DIFFERENT flow's trigger, we don't just
# silently swallow it or barge in over half-entered data — we ask whether
# they want to abandon what they were doing. flow_entry() records which flow
# just started (and that it started cleanly, as an admin); the actual
# interrupt prompt is wired up around admin_conv further down.
# ---------------------------------------------------------------------------

FLOW_LABELS = {
    "addshop": "adding a shop",
    "addtenant": "adding a tenant",
    "vacate": "marking a shop vacant",
    "pay": "recording a payment",
    "expense": "logging an expense",
    "editrent": "editing a shop's rent",
    "editshop": "editing a shop",
    "edittenant": "editing a tenant",
    "linkcode": "looking up a link code",
    "deactivate": "deactivating a shop",
    "activate": "activating a shop",
    "notify": "messaging a tenant",
    "broadcast": "sending a broadcast",
    "view": "viewing a shop",
    "editpay": "editing a payment",
    "editexp": "editing an expense",
    "addfloors": "adding a floor",
    "editfloors": "editing a floor",
}


def flow_entry(key):
    """Like @admin_only, but also remembers which guided flow is now active
    so an interrupting command can name it in its confirmation prompt."""
    def decorator(func):
        @wraps(func)
        async def wrapper(update: Update, context: ContextTypes.DEFAULT_TYPE):
            if not is_admin(update.effective_user.id):
                await send(update, "This is for building admins only.")
                return ConversationHandler.END
            context.user_data["flow_key"] = key
            return await func(update, context)
        return wrapper
    return decorator


# ---------------------------------------------------------------------------
# Menus
# ---------------------------------------------------------------------------

def admin_menu_kb():
    """Ordered by how often each action gets used day-to-day: payments and
    balance checks first, periodic admin tasks in the middle, rare setup
    actions (and the More/Help catch-alls) last."""
    return ReplyKeyboardMarkup(
        [
            ["💰 Record Payment", "🏬 Shops & Balances"],
            ["📉 Outstanding Dues", "⏰ Late Payments"],
            ["📊 Monthly Report", "🧾 Add Expense"],
            ["➕ Add Shop", "❓ Help"],
            ["⚙️ More"],
        ],
        resize_keyboard=True,
    )


def tenant_menu_kb():
    return ReplyKeyboardMarkup(
        [["💰 My Balance", "📜 My Ledger"], ["🏬 My Shop", "🪪 My Profile"], ["❓ Help"]],
        resize_keyboard=True,
    )


def unregistered_menu_kb():
    return ReplyKeyboardMarkup([["🔗 Register"]], resize_keyboard=True)


def menu_for(user_id: int):
    if is_admin(user_id):
        return admin_menu_kb()
    if db.get_shop_by_telegram_id(user_id):
        return tenant_menu_kb()
    return unregistered_menu_kb()


def more_inline_kb():
    """Ordered by how often each action gets used: quick lookups and
    tenant communication first, occasional edits and lifecycle changes
    (adding/vacating/deactivating a shop) toward the bottom, Close last."""
    rows = [
        [
            InlineKeyboardButton("🔍 View Shop", callback_data="more:view"),
            InlineKeyboardButton("🔗 Link Code", callback_data="more:link"),
        ],
        [
            InlineKeyboardButton("✉️ Notify Tenant", callback_data="more:notify"),
            InlineKeyboardButton("📢 Broadcast", callback_data="more:broadcast"),
        ],
        [
            InlineKeyboardButton("✏️ Edit Rent", callback_data="more:rent"),
            InlineKeyboardButton("🏗 Edit Shop", callback_data="more:editshop"),
        ],
        [
            InlineKeyboardButton("🧑‍💼 Edit Tenant", callback_data="more:edittenant"),
            InlineKeyboardButton("🧾 Edit Payment", callback_data="more:editpay"),
        ],
        [
            InlineKeyboardButton("📅 Month's Expenses", callback_data="more:exp"),
            InlineKeyboardButton("⚡ Charge Rent Now", callback_data="more:charge"),
        ],
        [
            InlineKeyboardButton("💡 Permanent Expenses", callback_data="more:permexp"),
            InlineKeyboardButton("✏️ Edit Expense", callback_data="more:editexp"),
        ],
        [
            InlineKeyboardButton("🧑 Add Tenant", callback_data="more:addtenant"),
            InlineKeyboardButton("🏚 Mark Vacant", callback_data="more:vacate"),
        ],
        [
            InlineKeyboardButton("⏸ Deactivate Shop", callback_data="more:deact"),
            InlineKeyboardButton("▶️ Activate Shop", callback_data="more:act"),
        ],
        [
            InlineKeyboardButton("🏢 Add Floor", callback_data="more:addfloors"),
            InlineKeyboardButton("🏢 Edit Floors", callback_data="more:editfloors"),
        ],
        [
            InlineKeyboardButton("✖ Close", callback_data="more:close"),
        ],
    ]
    return InlineKeyboardMarkup(rows)


def floor_kb():
    rows = [[InlineKeyboardButton(label, callback_data=f"floor:{key}")] for key, label in db.get_floors()]
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="floor:CANCEL")])
    return InlineKeyboardMarkup(rows)


def floor_and_all_kb(prefix):
    """Floor picker used to narrow a shop list down before showing it —
    includes an 'All Floors' shortcut plus the individual floors."""
    rows = [[InlineKeyboardButton("🏢 All Floors", callback_data=f"{prefix}:ALL")]]
    rows += [[InlineKeyboardButton(label, callback_data=f"{prefix}:{key}")] for key, label in db.get_floors()]
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data=f"{prefix}:CANCEL")])
    return InlineKeyboardMarkup(rows)


def editfloor_select_kb():
    rows = [
        [InlineKeyboardButton(
            f"{f['label']} ({f['shop_count']} shop{'s' if f['shop_count'] != 1 else ''})",
            callback_data=f"editfloor:{f['key']}",
        )]
        for f in db.get_floors_with_counts()
    ]
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="editfloor:CANCEL")])
    return InlineKeyboardMarkup(rows)


def _lumpsum_kb(prefix):
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("📅 Monthly rent (normal)", callback_data=f"{prefix}:no"),
          InlineKeyboardButton("💰 Lump sum (whole lease up front)", callback_data=f"{prefix}:yes")]]
    )


def shop_picker_kb(prefix, shops=None, active_only=True):
    if shops is None:
        shops = db.get_all_shops(active_only=active_only)
    buttons = [
        InlineKeyboardButton(
            f"{s['shop_no']} · {s['tenant_name'] or 'Vacant'}", callback_data=f"{prefix}:{s['shop_no']}"
        )
        for s in shops
    ]
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data=f"{prefix}:CANCEL")])
    return InlineKeyboardMarkup(rows), shops


# ---------------------------------------------------------------------------
# Ethiopian calendar (Amete Mihret) conversion now lives in database.py
# (gregorian_to_ethiopian, ethiopian_to_gregorian, format_ethiopian_date,
# gc_iso_to_ec_label, gc_iso_to_ec_iso, ETHIOPIAN_MONTHS — imported at the top
# of this file) so the storage layer can stamp every date's Ethiopian
# equivalent itself, not just the bot's display code.
# ---------------------------------------------------------------------------

PAY_MONTH_PAGE_SIZE = 4
PAY_MONTH_BEHIND = 4  # months behind "today" that are always reachable


def _add_eth_months(year, month, delta):
    """Add `delta` Ethiopian months (13-month calendar) to (year, month)."""
    total = (year * 13 + (month - 1)) + delta
    return total // 13, total % 13 + 1


def pay_month_kb(page=0, paid_periods=None):
    """Paged month picker: page 0 starts 4 months behind today and pages
    forward through today and indefinitely into future months, since tenants
    sometimes pay months in advance and sometimes are behind.

    `paid_periods` is the set of 'YYYY-MM' Ethiopian periods this shop
    already has a payment tagged for — those months get a ✅ so the admin
    doesn't record the same month twice."""
    paid_periods = paid_periods or set()
    today = date.today()
    ey, em, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    start_offset = -PAY_MONTH_BEHIND + page * PAY_MONTH_PAGE_SIZE
    buttons = []
    for i in range(PAY_MONTH_PAGE_SIZE):
        yr, mo = _add_eth_months(ey, em, start_offset + i)
        period = f"{yr}-{mo:02d}"
        label = f"{ETHIOPIAN_MONTHS[mo]} {yr}"
        if period in paid_periods:
            label = f"✅ {label}"
        buttons.append(InlineKeyboardButton(
            label, callback_data=f"paymonth:SELECT:{period}"
        ))
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("◀ Back", callback_data=f"paymonth:PAGE:{page - 1}"))
    nav.append(InlineKeyboardButton("Next ▶", callback_data=f"paymonth:PAGE:{page + 1}"))
    rows.append(nav)
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="paymonth:CANCEL")])
    return InlineKeyboardMarkup(rows)


REPORT_MONTH_PAGE_SIZE = 4
REPORT_MONTH_BEHIND = 4

def report_period_kb():
    rows = [
        [InlineKeyboardButton("📅 Monthly Report", callback_data="reportperiod:month")],
        [InlineKeyboardButton("🗓 Yearly Report", callback_data="reportperiod:year")],
        [InlineKeyboardButton("✖ Cancel", callback_data="reportperiod:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


REPORT_YEAR_COUNT = 5  # current Ethiopian year plus this many prior years


def report_year_kb():
    today = date.today()
    cur_ey, _, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    rows = [
        [InlineKeyboardButton(f"{cur_ey - i} E.C.", callback_data=f"reportyear:SELECT:{cur_ey - i}")]
        for i in range(REPORT_YEAR_COUNT)
    ]
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="reportyear:CANCEL")])
    return InlineKeyboardMarkup(rows)


def report_scope_kb(kind, key):
    """kind is 'month' or 'year'; key is that month's period ('2019-01') or
    that year (as a string, e.g. '2019') — folded into the SCOPE callback so
    the final step (report_month_select / report_year_select) knows what to
    compute."""
    prefix = f"report{kind}"
    rows = [
        [InlineKeyboardButton("📊 All Floors (incl. Special & Bank Shops)",
                               callback_data=f"{prefix}:SCOPE:{key}:all")],
        [InlineKeyboardButton("📊 Without Special & Bank Shops",
                               callback_data=f"{prefix}:SCOPE:{key}:nospecial")],
        [InlineKeyboardButton("✖ Cancel", callback_data=f"{prefix}:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


def _floor_breakdown_lines(start_iso, end_iso):
    lines = ["", "By floor:"]
    for f in db.report_by_floor(start_iso, end_iso):
        lines.append(
            f"  {f['label']}: {f['shop_count']} shop(s) — expected {money(f['rent_expected'])}/mo, "
            f"collected {money(f['rent_collected'])}"
        )
    return lines
def report_month_kb(page=0):
    """Paged month picker for /report, in Ethiopian months — since payments are
    recorded against Ethiopian months, the report is broken down the same way.
    Page 0 starts 4 months behind today; Next pages further back, Back pages forward."""
    today = date.today()
    ey, em, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    start_offset = -REPORT_MONTH_BEHIND + page * REPORT_MONTH_PAGE_SIZE
    buttons = []
    for i in range(REPORT_MONTH_PAGE_SIZE):
        yr, mo = _add_eth_months(ey, em, start_offset + i)
        buttons.append(InlineKeyboardButton(
            f"{ETHIOPIAN_MONTHS[mo]} {yr}", callback_data=f"reportmonth:SELECT:{yr}-{mo:02d}"
        ))
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("◀ Back", callback_data=f"reportmonth:PAGE:{page - 1}"))
    nav.append(InlineKeyboardButton("Next ▶", callback_data=f"reportmonth:PAGE:{page + 1}"))
    rows.append(nav)
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="reportmonth:CANCEL")])
    return InlineKeyboardMarkup(rows)


EXPENSE_MONTH_PAGE_SIZE = 4
EXPENSE_MONTH_BEHIND = 4


def expense_month_kb(page=0):
    """Paged month picker for /expense, in Ethiopian months — same shape as
    pay_month_kb/report_month_kb. Page 0 starts 4 months behind today, since
    an expense (a bill, a repair invoice) is sometimes logged after the fact."""
    today = date.today()
    ey, em, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    start_offset = -EXPENSE_MONTH_BEHIND + page * EXPENSE_MONTH_PAGE_SIZE
    buttons = []
    for i in range(EXPENSE_MONTH_PAGE_SIZE):
        yr, mo = _add_eth_months(ey, em, start_offset + i)
        buttons.append(InlineKeyboardButton(
            f"{ETHIOPIAN_MONTHS[mo]} {yr}", callback_data=f"expmonth:SELECT:{yr}-{mo:02d}"
        ))
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("◀ Back", callback_data=f"expmonth:PAGE:{page - 1}"))
    nav.append(InlineKeyboardButton("Next ▶", callback_data=f"expmonth:PAGE:{page + 1}"))
    rows.append(nav)
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="expmonth:CANCEL")])
    return InlineKeyboardMarkup(rows)


RENT_MONTH_PAGE_SIZE = 4
RENT_MONTH_BEHIND = 1  # allow a small amount of backdating


def rent_month_kb(page=0):
    """Paged month picker asking which month a rent change should start
    applying from — same shape as pay_month_kb/report_month_kb. Page 0
    starts 1 month behind today (a little backdating room) through the
    next several months, since rent changes are usually scheduled ahead
    ('starting next month') rather than backdated."""
    today = date.today()
    ey, em, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    cur_period = f"{ey:04d}-{em:02d}"
    start_offset = -RENT_MONTH_BEHIND + page * RENT_MONTH_PAGE_SIZE
    buttons = []
    for i in range(RENT_MONTH_PAGE_SIZE):
        yr, mo = _add_eth_months(ey, em, start_offset + i)
        period = f"{yr:04d}-{mo:02d}"
        label = f"{ETHIOPIAN_MONTHS[mo]} {yr}"
        if period == cur_period:
            label = f"{label} (this month)"
        buttons.append(InlineKeyboardButton(label, callback_data=f"rentmonth:SELECT:{period}"))
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton("◀ Back", callback_data=f"rentmonth:PAGE:{page - 1}"))
    nav.append(InlineKeyboardButton("Next ▶", callback_data=f"rentmonth:PAGE:{page + 1}"))
    rows.append(nav)
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="rentmonth:CANCEL")])
    return InlineKeyboardMarkup(rows)


BANK_KEYWORDS = [
    (re.compile(r"telebirr", re.I), "Telebirr"),
    (re.compile(r"commercial\s*bank|\bcbe\b", re.I), "CBE Birr"),
    (re.compile(r"dashen", re.I), "Dashen Bank"),
    (re.compile(r"awash", re.I), "Awash Bank"),
    (re.compile(r"abyssinia|\bboa\b", re.I), "Bank of Abyssinia"),
    (re.compile(r"wegagen", re.I), "Wegagen Bank"),
    (re.compile(r"cooperative bank of oromia|\bcbo\b", re.I), "Coop Bank of Oromia"),
]
AMOUNT_RE = re.compile(r"(?:amount|total|birr|etb)\D{0,6}([0-9][0-9,]*\.[0-9]{2}|[0-9][0-9,]*)", re.I)
DATE_RE = re.compile(r"(\d{4}-\d{2}-\d{2})|(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})")
REF_RE = re.compile(
    r"(?:reference|ref\.?|transaction|txn|receipt)\s*(?:no\.?|na|number|id)?\s*[:\-]?\s*([A-Za-z0-9]{6,})",
    re.I,
)
REF_FALLBACK_RE = re.compile(r"\bFT[A-Z0-9]{6,}\b", re.I)


def _normalize_date(raw):
    raw = raw.strip().rstrip(",.")
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def extract_receipt_info(image_bytes):
    """Best-effort OCR read of a receipt screenshot. Any field it can't
    confidently read is left as None, so the admin is only asked for that one."""
    result = {"bank": None, "amount": None, "payment_date": None, "reference_no": None}
    if not OCR_AVAILABLE:
        return result
    try:
        img = Image.open(io.BytesIO(image_bytes))
        text = pytesseract.image_to_string(img)
    except Exception as e:
        logger.warning("Receipt OCR failed: %s", e)
        return result

    for pattern, label in BANK_KEYWORDS:
        if pattern.search(text):
            result["bank"] = label
            break

    m = AMOUNT_RE.search(text)
    if m:
        try:
            result["amount"] = float(m.group(1).replace(",", ""))
        except ValueError:
            pass

    m = DATE_RE.search(text)
    if m:
        result["payment_date"] = _normalize_date(m.group(0))

    m = REF_RE.search(text)
    if not m:
        m = REF_FALLBACK_RE.search(text)
        if m:
            result["reference_no"] = m.group(0)
    elif m.group(1):
        result["reference_no"] = m.group(1)

    return result


def pay_bank_kb():
    rows = [
        [InlineKeyboardButton("🏦 CBE Birr", callback_data="paybank:CBE Birr"),
         InlineKeyboardButton("📱 Telebirr", callback_data="paybank:Telebirr")],
        [InlineKeyboardButton("💵 Cash", callback_data="paybank:Cash"),
         InlineKeyboardButton("🏛 Other Bank", callback_data="paybank:Other")],
        [InlineKeyboardButton("✖ Cancel", callback_data="paybank:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


def _current_eth_month():
    """Returns (ey, em) for today, in the Ethiopian calendar."""
    today = date.today()
    ey, em, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    return ey, em


def _eth_month_gregorian_bounds(ey, em):
    """Gregorian [start_iso, end_iso) covering one Ethiopian month."""
    next_y, next_m = _add_eth_months(ey, em, 1)
    sgy, sgm, sgd = ethiopian_to_gregorian(ey, em, 1)
    egy, egm, egd = ethiopian_to_gregorian(next_y, next_m, 1)
    return date(sgy, sgm, sgd).isoformat(), date(egy, egm, egd).isoformat()


def _expenses_for_eth_month(ey, em):
    """This Ethiopian month can span two Gregorian months, so pull both
    Gregorian months the range touches and filter down to the exact range."""
    start_iso, end_iso = _eth_month_gregorian_bounds(ey, em)
    y1, m1 = int(start_iso[:4]), int(start_iso[5:7])
    y2, m2 = int(end_iso[:4]), int(end_iso[5:7])
    rows = list(db.get_expenses(y1, m1))
    if (y2, m2) != (y1, m1):
        rows += list(db.get_expenses(y2, m2))
    return [r for r in rows if start_iso <= r["expense_date"] < end_iso]


def shop_detail_text(shop_no):
    shop = db.get_shop(shop_no)
    if not shop:
        return "No such shop."
    balance = db.get_balance(shop_no)
    ledger = db.get_ledger(shop_no, limit=8)
    lines = [
        f"Shop {shop['shop_no']} — {shop['tenant_name'] or 'Vacant'}",
        f"Occupancy: {'Rented' if shop['is_rented'] else 'Vacant'}",
        f"Floor: {db.floor_label(shop['floor'])}",
    ]
    if shop["area_sqm"]:
        lines.append(f"Area: {shop['area_sqm']:g} m²")
    lines += [
        f"Monthly rent: {money(shop['monthly_rent'])}",
        f"Status: {'active' if shop['active'] else 'inactive'}",
    ]
    if shop["is_rented"]:
        lines += [
            f"Purpose: {shop['purpose'] or '-'}",
            f"Phone: {shop['phone'] or '-'}",
            f"ID number: {shop['national_id'] or '-'}",
            f"TIN number: {shop['tin_number'] or '-'}",
            f"Documents on file: {len(_shop_documents(shop['shop_no']))}",
            f"Start date: {gc_iso_to_ec_label(shop['start_date'])} E.C.",
        ]
    lease_end = _field(shop, "lease_end_date")
    if lease_end:
        lines.append(f"Lease end: {gc_iso_to_ec_label(lease_end)} E.C.")
    if shop["is_rented"]:
        lines.append(f"Linked to Telegram: {'yes' if shop['telegram_id'] else 'no'}")
    lines.append(f"Balance due: {money(balance)}")
    if shop["is_rented"]:
        lines.append(f"Payment status: {profile_status_line(shop_no)}")
        grade = db.payment_punctuality(shop_no)
        if grade and grade["total"]:
            lines.append(
                f"Payment grade: {grade['grade']} "
                f"({grade['on_time']}/{grade['total']} months on time)"
            )
    lines += [
        "",
        "Recent activity:",
    ]
    for e in ledger:
        sign = "+" if e["type"] == "charge" else "-"
        lines.append(f"  {gc_iso_to_ec_label(e['date'])}  {e['type']}: {sign}{money(e['amount'])}")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Conversation states (one shared numbering for the whole admin conversation)
# ---------------------------------------------------------------------------

(
    ADDSHOP_FLOOR, ADDSHOP_NO, ADDSHOP_AREA, ADDSHOP_RENT, ADDSHOP_CONFIRM,
    PAY_FLOOR, PAY_SELECT, PAY_MONTH, PAY_RECEIPT, PAY_BANK, PAY_AMOUNT, PAY_DATE, PAY_REF, PAY_CONFIRM,
    EXPENSE_MONTH, EXPENSE_REASON, EXPENSE_AMOUNT, EXPENSE_DESC, EXPENSE_RECEIPT,
    RENT_SELECT, RENT_NEW, RENT_EFFECTIVE,
    NOTIFY_SELECT, NOTIFY_MSG,
    BROADCAST_MSG, BROADCAST_CONFIRM,
    VIEW_SELECT,
    LINK_SELECT,
    DEACT_SELECT,
    ACT_SELECT,
    ADDTENANT_SELECT, ADDTENANT_NAME, ADDTENANT_PURPOSE, ADDTENANT_PHONE, ADDTENANT_START, ADDTENANT_CONFIRM,
    ADDTENANT_FLOOR,
    ADDTENANT_LEASE_END, ADDTENANT_DOC,
    ADDTENANT_DOC_MORE,
    VACATE_SELECT, VACATE_CONFIRM,
    EDITSHOP_SELECT, EDITSHOP_FIELD, EDITSHOP_VALUE, EDITSHOP_CONFIRM,
    EDITTENANT_SELECT, EDITTENANT_FIELD, EDITTENANT_VALUE, EDITTENANT_CONFIRM,
    EXPENSE_CONFIRM,
    RENT_CONFIRM,
    DEACT_CONFIRM,
    ACT_CONFIRM,
    NOTIFY_CONFIRM,
    EDITPAY_SELECT, EDITPAY_PAYMENT, EDITPAY_FIELD, EDITPAY_VALUE, EDITPAY_CONFIRM,
    EDITEXP_MONTH, EDITEXP_PICK, EDITEXP_FIELD, EDITEXP_VALUE, EDITEXP_CONFIRM,
    ADDFLOOR_LABEL, ADDFLOOR_LUMPSUM, ADDFLOOR_CONFIRM,
    EDITFLOOR_SELECT, EDITFLOOR_FIELD, EDITFLOOR_VALUE, EDITFLOOR_CONFIRM,
) = range(72)

REGISTER_CODE = 0  # separate, small conversation for tenants
PROFILE_ID, PROFILE_TIN, PROFILE_DOC, PROFILE_DOC_MORE = range(4)  # separate, small conversation for tenant profile edits


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("Cancelled.", reply_markup=menu_for(update.effective_user.id))
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Shared / entry commands
# ---------------------------------------------------------------------------

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if is_admin(uid):
        await update.message.reply_text(
            "Welcome, admin. Use the buttons below, or /help for the full list.",
            reply_markup=admin_menu_kb(),
        )
        return
    shop = db.get_shop_by_telegram_id(uid)
    if shop:
        await update.message.reply_text(
            f"Welcome back, Shop {shop['shop_no']}.", reply_markup=tenant_menu_kb()
        )
    else:
        await update.message.reply_text(
            "Welcome! If you're a shop tenant, tap Register and enter the code "
            "the building admin gave you.",
            reply_markup=unregistered_menu_kb(),
        )


async def menu_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Here's your menu:", reply_markup=menu_for(update.effective_user.id))


async def help_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    if is_admin(uid):
        await update.message.reply_text(
            "Tap a button below, or use these commands (each walks you through it "
            "step by step):\n\n"
            "/pay — record a payment\n"
            "/shops — list shops & balances\n"
            "/dues — outstanding balances\n"
            "/latepayments — shops with unpaid months, marked by how overdue\n"
            "/report — monthly or yearly collections vs expenses, by floor\n"
            "/exportexcel [year] — download the full-year rent table as Excel\n"
            "/expense — log an expense\n"
            "/permanentexpenses — manage recurring expenses (electricity/water/salary)\n"
            "/addshop — add a shop (vacant, with a monthly rent)\n"
            "/notify — message one tenant\n"
            "/broadcast — message every linked tenant\n"
            "/linkcode — get a shop's tenant link code\n"
            "/editrent — change a shop's rent\n"
            "/editshop — edit a shop's number, floor, area or rent\n"
            "/edittenant — edit a tenant's name, phone, purpose, dates, ID or TIN\n"
            "/editpayment — edit an existing payment's amount, date, bank, period, "
            "reference, note or receipt\n"
            "/editexpense — edit a logged expense's description, amount, category, "
            "date or receipt\n"
            "/expenses — this month's expenses\n"
            "/chargenow — apply this month's rent immediately\n"
            "/addtenant — rent out a vacant shop\n"
            "/vacate — mark a shop vacant (tenant moved out)\n"
            "/deactivate, /activate — toggle a shop\n"
            "/cancel — stop whatever you're in the middle of",
            reply_markup=admin_menu_kb(),
        )
        return
    shop = db.get_shop_by_telegram_id(uid)
    if shop:
        await update.message.reply_text(
            "Your commands:\n"
            "/mybalance — current amount due\n"
            "/myledger — recent charges & payments\n"
            "/myshop — your shop details\n"
            "/myprofile — your profile (ID, TIN, payment status, document)\n"
            "/registercode — view your shop's link code again",
            reply_markup=tenant_menu_kb(),
        )
    else:
        await update.message.reply_text(
            "Tap Register below and enter the code your admin gave you.",
            reply_markup=unregistered_menu_kb(),
        )


# ---------------------------------------------------------------------------
# Tenant registration (its own small conversation)
# ---------------------------------------------------------------------------

async def register_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if context.args:
        return await register_apply(update, context, context.args[0])
    await update.message.reply_text(
        "Send the link code your admin gave you.", reply_markup=ReplyKeyboardRemove()
    )
    return REGISTER_CODE


async def register_code_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    return await register_apply(update, context, update.message.text.strip())


async def register_apply(update: Update, context: ContextTypes.DEFAULT_TYPE, code: str):
    shop = db.get_shop_by_link_code(code)
    if not shop:
        await update.message.reply_text("That code isn't valid. Double-check with the admin, or /cancel.")
        return REGISTER_CODE
    if shop["telegram_id"] and shop["telegram_id"] != update.effective_user.id:
        await update.message.reply_text(
            "That shop is already linked to a different Telegram account. Ask the admin for help."
        )
        return ConversationHandler.END
    db.link_telegram_to_shop(shop["shop_no"], update.effective_user.id)
    await update.message.reply_text(
        f"You're linked to Shop {shop['shop_no']} ({shop['tenant_name']}).\n"
        f"Your code was: {code}\n"
        "You can pull it up again anytime with /registercode.",
        reply_markup=tenant_menu_kb(),
    )
    return ConversationHandler.END


register_conv = ConversationHandler(
    entry_points=[
        CommandHandler("register", register_start),
        MessageHandler(filters.Regex("^🔗 Register$"), register_start),
    ],
    states={REGISTER_CODE: [MessageHandler(filters.TEXT & ~filters.COMMAND, register_code_msg)]},
    fallbacks=[CommandHandler("cancel", cancel)],
)


# ---------------------------------------------------------------------------
# Tenant-facing view commands (no input needed)
# ---------------------------------------------------------------------------

@tenant_only
async def my_balance(update: Update, context: ContextTypes.DEFAULT_TYPE, shop):
    balance = db.get_balance(shop["shop_no"])
    if balance > 0:
        text = f"Shop {shop['shop_no']}: you owe {money(balance)}."
    elif balance < 0:
        text = f"Shop {shop['shop_no']}: you're paid ahead by {money(-balance)}."
    else:
        text = f"Shop {shop['shop_no']}: you're all settled up. \U00002705"
    await update.message.reply_text(text)


@tenant_only
async def my_ledger(update: Update, context: ContextTypes.DEFAULT_TYPE, shop):
    entries = db.get_ledger(shop["shop_no"], limit=10)
    if not entries:
        await update.message.reply_text("No transactions yet.")
        return
    lines = [f"Recent activity — Shop {shop['shop_no']}:"]
    for e in entries:
        sign = "+" if e["type"] == "charge" else "-"
        label = "Rent charge" if e["type"] == "charge" else "Payment"
        lines.append(f"{gc_iso_to_ec_label(e['date'])}  {label}: {sign}{money(e['amount'])}")
    await update.message.reply_text("\n".join(lines))


@tenant_only
async def my_shop(update: Update, context: ContextTypes.DEFAULT_TYPE, shop):
    await update.message.reply_text(
        f"Shop {shop['shop_no']}\n"
        + (f"Purpose: {shop['purpose']}\n" if shop['purpose'] else "")
        + f"Floor: {db.floor_label(shop['floor'])}\n"
        + (f"Area: {shop['area_sqm']:g} m²\n" if shop['area_sqm'] else "")
        + f"Tenant: {shop['tenant_name']}\n"
        f"Phone: {shop['phone'] or '-'}\n"
        f"Monthly rent: {money(shop['monthly_rent'])}\n"
        f"Tenancy start: {gc_iso_to_ec_label(shop['start_date'])} E.C."
        + (f"\nLease end: {gc_iso_to_ec_label(_field(shop, 'lease_end_date'))} E.C."
           if _field(shop, "lease_end_date") else "")
    )


@tenant_only
async def my_registercode(update: Update, context: ContextTypes.DEFAULT_TYPE, shop):
    code = shop["link_code"] or db.regenerate_link_code(shop["shop_no"])
    await update.message.reply_text(
        f"Shop {shop['shop_no']} link code: {code}\n"
        "Give this to anyone who needs to link a Telegram account to your shop."
    )


# ---------------------------------------------------------------------------
# Tenant: profile (view + self-service ID / TIN / document)
# ---------------------------------------------------------------------------

def profile_status_line(shop_no):
    status, balance = db.payment_status(shop_no)
    if status == "on_time":
        return "✅ Paid on time"
    if status == "overdue":
        return f"⚠️ Overdue — owes {money(balance)}"
    return "Unknown"


def profile_text(shop):
    return (
        f"Shop {shop['shop_no']} — {db.floor_label(shop['floor'])}\n"
        f"Purpose: {shop['purpose'] or '-'}\n\n"
        f"Tenant: {shop['tenant_name']}\n"
        f"Phone: {shop['phone'] or '-'}\n"
        f"ID number: {shop['national_id'] or 'not set'}\n"
        f"TIN number: {shop['tin_number'] or 'not set'}\n"
        f"Documents on file: {len(_shop_documents(shop['shop_no']))}\n\n"
        f"Payment status: {profile_status_line(shop['shop_no'])}"
    )


def profile_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🆔 Set ID Number", callback_data="profile:id"),
         InlineKeyboardButton("🧾 Set TIN Number", callback_data="profile:tin")],
        [InlineKeyboardButton("📎 Add Document", callback_data="profile:doc")],
        [InlineKeyboardButton("✖ Close", callback_data="profile:close")],
    ])


@tenant_only
async def my_profile(update: Update, context: ContextTypes.DEFAULT_TYPE, shop):
    await update.message.reply_text(profile_text(shop), reply_markup=profile_kb())


async def profile_close(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    await query.edit_message_reply_markup(reply_markup=None)


async def profile_id_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop = db.get_shop_by_telegram_id(update.effective_user.id)
    if not shop:
        return ConversationHandler.END
    await query.message.reply_text("Send your ID number. /cancel to stop.")
    return PROFILE_ID


async def profile_id_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    shop = db.get_shop_by_telegram_id(update.effective_user.id)
    if not shop:
        return ConversationHandler.END
    db.update_national_id(shop["shop_no"], update.message.text.strip())
    await update.message.reply_text("ID number saved.", reply_markup=tenant_menu_kb())
    return ConversationHandler.END


async def profile_tin_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop = db.get_shop_by_telegram_id(update.effective_user.id)
    if not shop:
        return ConversationHandler.END
    await query.message.reply_text("Send your TIN number. /cancel to stop.")
    return PROFILE_TIN


async def profile_tin_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    shop = db.get_shop_by_telegram_id(update.effective_user.id)
    if not shop:
        return ConversationHandler.END
    db.update_tin(shop["shop_no"], update.message.text.strip())
    await update.message.reply_text("TIN number saved.", reply_markup=tenant_menu_kb())
    return ConversationHandler.END


async def profile_doc_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop = db.get_shop_by_telegram_id(update.effective_user.id)
    if not shop:
        return ConversationHandler.END
    await query.message.reply_text(
        "Send a photo or file of your document (ID, TIN certificate, license, etc.). "
        "/cancel to stop."
    )
    return PROFILE_DOC


async def profile_doc_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    shop = db.get_shop_by_telegram_id(update.effective_user.id)
    if not shop:
        return ConversationHandler.END
    if update.message.photo:
        file_id, kind = update.message.photo[-1].file_id, "photo"
    elif update.message.document:
        file_id, kind = update.message.document.file_id, "document"
    else:
        await update.message.reply_text("Please send a photo or a file, or /cancel.")
        return PROFILE_DOC
    _save_shop_document(shop["shop_no"], file_id, kind)
    await update.message.reply_text(
        "Document saved. Add another document?",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("➕ Add Another", callback_data="profiledoc:more"),
              InlineKeyboardButton("✅ Done", callback_data="profiledoc:done")]]
        ),
    )
    return PROFILE_DOC_MORE


async def profile_doc_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "profiledoc:done":
        await query.edit_message_reply_markup(reply_markup=None)
        await query.message.reply_text("All set.", reply_markup=tenant_menu_kb())
        return ConversationHandler.END
    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.reply_text("Send the next photo or file. /cancel to stop.")
    return PROFILE_DOC


profile_conv = ConversationHandler(
    entry_points=[
        CallbackQueryHandler(profile_id_start, pattern="^profile:id$"),
        CallbackQueryHandler(profile_tin_start, pattern="^profile:tin$"),
        CallbackQueryHandler(profile_doc_start, pattern="^profile:doc$"),
    ],
    states={
        PROFILE_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, profile_id_msg)],
        PROFILE_TIN: [MessageHandler(filters.TEXT & ~filters.COMMAND, profile_tin_msg)],
        PROFILE_DOC: [MessageHandler((filters.PHOTO | filters.Document.ALL) & ~filters.COMMAND, profile_doc_msg)],
        PROFILE_DOC_MORE: [CallbackQueryHandler(profile_doc_more, pattern="^profiledoc:")],
    },
    fallbacks=[CommandHandler("cancel", cancel)],
)


# ---------------------------------------------------------------------------
# Admin: view commands (no input needed)
# ---------------------------------------------------------------------------

def shops_list_kb(floor_key=None):
    """floor_key=None means 'All Floors' (every shop, grouped by floor as before).
    A specific floor key narrows the list down to just that floor."""
    shops = db.get_all_balances(active_only=True)
    if floor_key is not None:
        shops = [s for s in shops if s["floor"] == floor_key]
    balance_by_no = {s["shop_no"]: s["balance"] for s in shops}
    if floor_key is not None:
        grouped = [(floor_key, db.floor_label(floor_key), shops)] if shops else []
    else:
        grouped = db.get_shops_grouped_by_floor(active_only=True)
    rows = []
    for _key, label, floor_shops in grouped:
        rows.append([InlineKeyboardButton(f"— {label} —", callback_data="listshop:NOOP")])
        for s in floor_shops:
            balance = balance_by_no.get(s["shop_no"], 0)
            purpose_tag = f" ({s['purpose']})" if s["purpose"] else ""
            tenant_tag = s["tenant_name"] if s["is_rented"] else "Vacant"
            rows.append([InlineKeyboardButton(
                f"{s['shop_no']} · {tenant_tag}{purpose_tag} — {money(balance)}",
                callback_data=f"listshop:{s['shop_no']}",
            )])
    rows.append([InlineKeyboardButton("« Back to floors", callback_data="listshop:BACKFLOOR")])
    rows.append([InlineKeyboardButton("✖ Close", callback_data="listshop:CLOSE")])
    return InlineKeyboardMarkup(rows), shops


@admin_only
async def list_shops(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not db.get_all_shops(active_only=True):
        await update.message.reply_text("No shops added yet. Tap Add Shop to get started.")
        return
    await update.message.reply_text("Which floor?", reply_markup=floor_and_all_kb("sbfloor"))


async def sb_floor_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_admin(update.effective_user.id):
        return
    key = query.data.split(":", 1)[1]
    if key == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return
    floor_key = None if key == "ALL" else key
    context.user_data["sb_floor"] = floor_key
    kb, shops = shops_list_kb(floor_key)
    label = "All Floors" if floor_key is None else db.floor_label(floor_key)
    if not shops:
        await query.edit_message_text(f"No shops on {label}.")
        return
    await query.edit_message_text(f"{label} — tap a shop to see its full status:", reply_markup=kb)


async def list_shop_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_admin(update.effective_user.id):
        return
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "NOOP":
        return
    if shop_no == "CLOSE":
        await query.edit_message_reply_markup(reply_markup=None)
        return
    if shop_no == "BACKFLOOR":
        await query.edit_message_text("Which floor?", reply_markup=floor_and_all_kb("sbfloor"))
        return
    if shop_no == "BACK":
        floor_key = context.user_data.get("sb_floor")
        kb, shops = shops_list_kb(floor_key)
        label = "All Floors" if floor_key is None else db.floor_label(floor_key)
        if not shops:
            await query.edit_message_text(f"No shops on {label}.")
            return
        await query.edit_message_text(f"{label} — tap a shop to see its full status:", reply_markup=kb)
        return
    shop = db.get_shop(shop_no)
    rows = []
    doc_count = len(_shop_documents(shop_no)) if shop else 0
    if doc_count:
        label = "📎 View Document" if doc_count == 1 else f"📎 View Documents ({doc_count})"
        rows.append([InlineKeyboardButton(label, callback_data=f"viewdoc:{shop_no}")])
    rows.append([InlineKeyboardButton("« Back to list", callback_data="listshop:BACK")])
    await query.edit_message_text(
        shop_detail_text(shop_no),
        reply_markup=InlineKeyboardMarkup(rows),
    )


async def view_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_admin(update.effective_user.id):
        return
    shop_no = query.data.split(":", 1)[1]
    shop = db.get_shop(shop_no)
    docs = _shop_documents(shop_no)
    if not shop or not docs:
        await query.message.reply_text("No document on file for that shop.")
        return
    tenant_tag = shop["tenant_name"] or "Vacant"
    for i, doc in enumerate(docs, start=1):
        caption = f"{_field(doc, 'label', 'Document')} — Shop {shop_no} ({tenant_tag}) [{i}/{len(docs)}]"
        await _send_stored_document(query.message, doc, caption)


@admin_only
async def dues(update: Update, context: ContextTypes.DEFAULT_TYPE):
    shops = [s for s in db.get_all_balances(active_only=True) if s["balance"] > 0]
    if not shops:
        await update.message.reply_text("No outstanding dues. Everyone is settled up.")
        return
    shops.sort(key=lambda s: -s["balance"])
    lines = ["Shops with outstanding balance:"]
    for s in shops:
        lines.append(f"Shop {s['shop_no']} ({s['tenant_name']}): {money(s['balance'])}")
    await update.message.reply_text("\n".join(lines))


def _late_mark(days_late):
    """Colour mark for how overdue a shop's oldest unpaid month is.

    10+ days late -> 🟢, 20+ days late -> 🟡, more than a month (30+ days)
    late -> 🔴. Under 10 days gets no mark — thresholds are cumulative, so
    the mark shown is always the highest one reached."""
    if days_late > 30:
        return "🔴"
    if days_late >= 20:
        return "🟡"
    if days_late >= 10:
        return "🟢"
    return ""


@admin_only
async def late_payments(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Lists each active shop's unpaid Ethiopian months — every month from
    when they started renting up through the current month that has no
    payment tagged for it — marked by how overdue the oldest unpaid month
    is: 🟢 10+ days late, 🟡 20+ days late, 🔴 more than a month late."""
    today = date.today()
    ey, em, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    entries = []
    for shop in db.get_all_shops(active_only=True, rented_only=True):
        sy, sm, sd = (int(x) for x in shop["start_date"].split("-"))
        y, m, _ = gregorian_to_ethiopian(sy, sm, sd)
        paid = db.paid_periods(shop["id"])
        unpaid_labels = []
        oldest_unpaid = None
        while (y, m) <= (ey, em):
            if f"{y}-{m:02d}" not in paid:
                unpaid_labels.append(f"{ETHIOPIAN_MONTHS[m]} {y}")
                if oldest_unpaid is None:
                    oldest_unpaid = (y, m)
            y, m = _add_eth_months(y, m, 1)
        if unpaid_labels:
            due_gy, due_gm, due_gd = ethiopian_to_gregorian(*oldest_unpaid, 1)
            days_late = (today - date(due_gy, due_gm, due_gd)).days
            entries.append((shop, unpaid_labels, days_late))

    if not entries:
        await update.message.reply_text("All shops are up to date — no unpaid months.")
        return

    # Worst-first: shops furthest behind (most days late) at the top.
    entries.sort(key=lambda e: e[2], reverse=True)

    # Show the most recent MAX_SHOWN unpaid months per shop, with a count of
    # any older ones — a shop with no period-tagged payments at all could
    # otherwise list 100+ months and blow past Telegram's message-length limit.
    MAX_SHOWN = 12
    lines = ["Late payments:"]
    for shop, unpaid_labels, days_late in entries:
        mark = _late_mark(days_late)
        prefix = f"{mark} " if mark else ""
        if len(unpaid_labels) > MAX_SHOWN:
            shown = ", ".join(unpaid_labels[-MAX_SHOWN:])
            note = f" (+{len(unpaid_labels) - MAX_SHOWN} earlier months)"
        else:
            shown = ", ".join(unpaid_labels)
            note = ""
        lines.append(
            f"{prefix}{shop['shop_no']} — {shop['tenant_name']} "
            f"({len(unpaid_labels)} unpaid, oldest {days_late}d late): {shown}{note}"
        )

    # Chunk into multiple messages if the list is still long.
    chunk = ""
    for line in lines:
        candidate = f"{chunk}\n{line}" if chunk else line
        if len(candidate) > 3500:
            await update.message.reply_text(chunk)
            chunk = line
        else:
            chunk = candidate
    if chunk:
        await update.message.reply_text(chunk)


@admin_only
async def list_expenses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    ey, em = _current_eth_month()
    rows = _expenses_for_eth_month(ey, em)
    label = f"{ETHIOPIAN_MONTHS[em]} {ey}"
    if not rows:
        await update.message.reply_text(f"No expenses recorded for {label}.")
        return
    total = sum(r["amount"] for r in rows)
    lines = [f"Expenses for {label}:"]
    for r in rows:
        lines.append(f"{gc_iso_to_ec_label(r['expense_date'])}  {money(r['amount'])}  {r['description']}")
    lines.append(f"\nTotal: {money(total)}")
    await update.message.reply_text("\n".join(lines))


@admin_only
async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Monthly or yearly report?", reply_markup=report_period_kb())


@admin_only
async def report_period_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    value = query.data.split(":", 1)[1]
    if value == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return
    if value == "month":
        await query.edit_message_text("Which month?", reply_markup=report_month_kb())
    else:
        await query.edit_message_text("Which year?", reply_markup=report_year_kb())


@admin_only
async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """The other option next to /report — the full-year, spreadsheet-style
    table (shop info + one column per Ethiopian month) as a downloadable
    .xlsx, same layout as the building's original manual ledger."""
    if context.args:
        try:
            ec_year = int(context.args[0])
        except ValueError:
            await update.message.reply_text("Usage: /exportexcel [year], e.g. /exportexcel 2018")
            return
    else:
        t = date.today()
        ec_year, _, _ = db.gregorian_to_ethiopian(t.year, t.month, t.day)
    buf = excel_export.build_workbook(ec_year)
    await update.message.reply_document(
        document=InputFile(buf, filename=f"rent_{ec_year}_EC.xlsx"),
        caption=f"Rent report for {ec_year} E.C.",
    )


async def report_month_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_admin(update.effective_user.id):
        return
    parts = query.data.split(":", 3)
    action = parts[1]
    if action == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return
    if action == "PAGE":
        page = int(parts[2])
        await query.edit_message_reply_markup(reply_markup=report_month_kb(page))
        return
    if action == "SELECT":
        period = parts[2]
        yr, mo = (int(x) for x in period.split("-"))
        label = f"{ETHIOPIAN_MONTHS[mo]} {yr}"
        await query.edit_message_text(
            f"Report for {label} — which shops?", reply_markup=report_scope_kb("month", period)
        )
        return

    period_arg = parts[2]
    scope = parts[3]
    exclude_floors = db.get_lump_sum_floors() if scope == "nospecial" else None
    yr, mo = (int(x) for x in period_arg.split("-"))
    next_yr, next_mo = _add_eth_months(yr, mo, 1)
    start_gy, start_gm, start_gd = ethiopian_to_gregorian(yr, mo, 1)
    end_gy, end_gm, end_gd = ethiopian_to_gregorian(next_yr, next_mo, 1)
    start_iso = date(start_gy, start_gm, start_gd).isoformat()
    end_iso = date(end_gy, end_gm, end_gd).isoformat()
    label = f"{ETHIOPIAN_MONTHS[mo]} {yr}"
    period = f"{yr}-{mo:02d}"

    r = db.monthly_report_range(start_iso, end_iso, label, exclude_floors=exclude_floors)
    active_shops = db.shops_active_in_range(start_iso, end_iso)
    if exclude_floors:
        active_shops = [s for s in active_shops if s["floor"] not in exclude_floors]
    # "Rent expected" is calculated live from each shop's rent on record
    # for this period (via rent_for_period/rent_history), not from
    # whatever's already been posted to `charges` — so it's right even
    # before the monthly charge job has run for this period.
    expected = r["rent_expected"]
    unpaid = db.shops_unpaid_for_period(period, start_iso, end_iso)
    if exclude_floors:
        unpaid = [s for s in unpaid if s["floor"] not in exclude_floors]
    charged_this_period = db.charges_by_shop_for_period(period)
    unpaid_total = round(
        sum(charged_this_period.get(s["id"], s["monthly_rent"]) for s in unpaid), 2
    )

    scope_label = "Without Special & Bank Shops" if exclude_floors else "All Floors (incl. Special & Bank Shops)"
    lines = [
        f"Report for {label} — {scope_label}",
        f"Active shops: {len(active_shops)}",
        f"Rent expected: {money(expected)}",
        f"Rent collected: {money(r['rent_collected'])}",
        f"Expenses: {money(r['expenses'])}" + (" (building-wide)" if exclude_floors else ""),
        f"Net (collected - expenses): {money(r['net'])}",
        f"Unpaid rent this month: {money(unpaid_total)}",
    ]
    lines += _floor_breakdown_lines(start_iso, end_iso)
    kb = InlineKeyboardMarkup([[InlineKeyboardButton(
        "📥 Export to Excel", callback_data=f"reportexport:month:{period}:{scope}"
    )]])
    await query.edit_message_text("\n".join(lines), reply_markup=kb)


@admin_only
async def report_year_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":", 3)
    action = parts[1]
    if action == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return
    if action == "SELECT":
        yr = int(parts[2])
        await query.edit_message_text(
            f"Report for {yr} E.C. — which shops?", reply_markup=report_scope_kb("year", str(yr))
        )
        return

    yr = int(parts[2])
    scope = parts[3]
    exclude_floors = db.get_lump_sum_floors() if scope == "nospecial" else None
    r = db.yearly_report(yr, exclude_floors=exclude_floors)
    start_iso, end_iso = db.ethiopian_year_bounds(yr)
    scope_label = "Without Special & Bank Shops" if exclude_floors else "All Floors (incl. Special & Bank Shops)"
    lines = [
        f"Report for {yr} E.C. — {scope_label}",
        f"Rent charged: {money(r['rent_charged'])}",
        f"Rent collected: {money(r['rent_collected'])}",
        f"Expenses: {money(r['expenses'])} (building-wide)",
        f"Net (collected - expenses): {money(r['net'])}",
        f"Total outstanding ({'excl.' if exclude_floors else 'all'} shops, all time): "
        f"{money(r['total_outstanding_all_shops'])}",
        "",
        "By month:",
    ]
    for m in r["months"]:
        lines.append(f"  {m['period']}: collected {money(m['rent_collected'])}, expenses {money(m['expenses'])}")
    lines += _floor_breakdown_lines(start_iso, end_iso)
    kb = InlineKeyboardMarkup([[InlineKeyboardButton(
        "📥 Export to Excel", callback_data=f"reportexport:year:{yr}:{scope}"
    )]])
    await query.edit_message_text("\n".join(lines), reply_markup=kb)


@admin_only
async def report_export(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    _, kind, key, scope = query.data.split(":", 3)
    exclude_floors = db.get_lump_sum_floors() if scope == "nospecial" else None
    buf = excel_export.build_report_workbook(kind, key, exclude_floors=exclude_floors)
    label = f"{key} E.C." if kind == "year" else key
    await query.message.reply_document(
        document=InputFile(buf, filename=f"rent_report_{key}.xlsx"),
        caption=f"Rent report for {label}",
    )


def _chargenow_confirm_kb():
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Confirm", callback_data="chargenow:go"),
          InlineKeyboardButton("✖ Cancel", callback_data="chargenow:cancel")]]
    )


@admin_only
async def charge_now(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Apply this month's rent charge to every active shop that hasn't been charged yet?",
        reply_markup=_chargenow_confirm_kb(),
    )


@admin_only
async def charge_now_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "chargenow:cancel":
        await query.edit_message_text("Cancelled.")
        return
    applied = db.apply_monthly_charges_to_all()
    ey, em, _ = gregorian_to_ethiopian(date.today().year, date.today().month, date.today().day)
    applied_expenses = db.apply_recurring_expenses_to_all(f"{ey:04d}-{em:02d}")
    lines = []
    if applied:
        lines.append(f"Applied this month's rent to {len(applied)} shop(s): {', '.join(applied)}")
    else:
        lines.append("Rent: nothing to apply — this month's rent was already charged.")
    if applied_expenses:
        lines.append(f"Applied permanent expenses: {', '.join(applied_expenses)}")
    await query.edit_message_text("\n".join(lines))


@admin_only
async def more_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("More options:", reply_markup=more_inline_kb())


async def more_router(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles the 'More' menu buttons that are simple one-tap actions."""
    query = update.callback_query
    if not is_admin(update.effective_user.id):
        await query.answer("Admins only.", show_alert=True)
        return
    await query.answer()
    if query.data == "more:close":
        await query.edit_message_reply_markup(reply_markup=None)
        return
    if query.data == "more:exp":
        ey, em = _current_eth_month()
        rows = _expenses_for_eth_month(ey, em)
        label = f"{ETHIOPIAN_MONTHS[em]} {ey}"
        if not rows:
            await query.message.reply_text(f"No expenses recorded for {label}.")
        else:
            total = sum(r["amount"] for r in rows)
            lines = [f"Expenses for {label}:"]
            for r in rows:
                lines.append(f"{gc_iso_to_ec_label(r['expense_date'])}  {money(r['amount'])}  {r['description']}")
            lines.append(f"\nTotal: {money(total)}")
            await query.message.reply_text("\n".join(lines))
        return
    if query.data == "more:charge":
        await query.message.reply_text(
            "Apply this month's rent charge to every active shop that hasn't been charged yet?",
            reply_markup=_chargenow_confirm_kb(),
        )
        return


# ---------------------------------------------------------------------------
# Admin flow: Add Shop
# ---------------------------------------------------------------------------

@flow_entry("addshop")
async def addshop_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not db.get_floors():
        await send(
            update,
            "No floors yet — add a floor first (Add Floor) before adding a shop.",
            reply_markup=admin_menu_kb(),
        )
        return ConversationHandler.END
    context.user_data["new_shop"] = {}
    await send(
        update,
        "Let's add a shop. Which floor is it on?\nSend /cancel anytime to stop.",
        reply_markup=floor_kb(),
    )
    return ADDSHOP_FLOOR


async def addshop_floor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    key = query.data.split(":", 1)[1]
    if key == "CANCEL":
        context.user_data.pop("new_shop", None)
        await query.edit_message_text("Cancelled.")
        await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
        return ConversationHandler.END
    context.user_data["new_shop"]["floor"] = key
    await query.edit_message_text(f"Floor: {db.floor_label(key)}")
    await query.message.reply_text("Send the shop number (e.g. 12).")
    return ADDSHOP_NO


async def addshop_no(update: Update, context: ContextTypes.DEFAULT_TYPE):
    shop_no = update.message.text.strip()
    if db.get_shop(shop_no):
        await update.message.reply_text(f"Shop {shop_no} already exists. Send a different number, or /cancel.")
        return ADDSHOP_NO
    context.user_data["new_shop"]["shop_no"] = shop_no
    await update.message.reply_text("Shop area in m²? (e.g. 24, or send 'skip' if unknown)")
    return ADDSHOP_AREA


async def addshop_area(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() == "skip":
        context.user_data["new_shop"]["area_sqm"] = None
    else:
        try:
            area = float(text)
            if area <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("That's not a valid area. Send a number, e.g. 24, or 'skip'.")
            return ADDSHOP_AREA
        context.user_data["new_shop"]["area_sqm"] = area
    await update.message.reply_text("Monthly rent for this shop? (e.g. 6000)")
    return ADDSHOP_RENT


async def addshop_rent(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    try:
        rent = float(text)
        if rent <= 0:
            raise ValueError
    except ValueError:
        await update.message.reply_text("That's not a valid amount. Send a number, e.g. 6000")
        return ADDSHOP_RENT
    d = context.user_data["new_shop"]
    d["monthly_rent"] = rent
    area_line = f"Area: {d['area_sqm']:g} m²\n" if d.get("area_sqm") else ""
    await update.message.reply_text(
        "Please confirm:\n"
        f"Shop: {d['shop_no']}\n"
        f"Floor: {db.floor_label(d['floor'])}\n"
        f"{area_line}"
        f"Monthly rent: {money(rent)}\n"
        "Status: Vacant (no tenant yet)",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Save", callback_data="addshop:save"),
              InlineKeyboardButton("✖ Cancel", callback_data="addshop:cancel")]]
        ),
    )
    return ADDSHOP_CONFIRM


async def addshop_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    d = context.user_data.pop("new_shop", None)
    if query.data == "addshop:cancel" or not d:
        await query.edit_message_text("Cancelled.")
        await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
        return ConversationHandler.END
    db.add_shop(d["shop_no"], d["monthly_rent"], floor=d["floor"], area_sqm=d.get("area_sqm"))
    await query.edit_message_text(
        f"Added Shop {d['shop_no']} — vacant, rent {money(d['monthly_rent'])}.\n"
        "Use Add Tenant (in ⚙️ More) once it's rented out."
    )
    await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Add Tenant (rent out a vacant shop)
# ---------------------------------------------------------------------------

@flow_entry("addtenant")
async def addtenant_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not db.get_all_shops(active_only=True, rented_only=False):
        await send(update, "No vacant shops — every shop already has a tenant.")
        return ConversationHandler.END
    await send(update, "Which floor is the shop on?", reply_markup=floor_and_all_kb("addtenantfloor"))
    return ADDTENANT_FLOOR


async def addtenant_floor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    key = query.data.split(":", 1)[1]
    if key == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    floor_key = None if key == "ALL" else key
    vacant = db.get_all_shops(active_only=True, rented_only=False)
    if floor_key is not None:
        vacant = [s for s in vacant if s["floor"] == floor_key]
    kb, shops = shop_picker_kb("addtenantshop", shops=vacant)
    if not shops:
        await query.edit_message_text(f"No vacant shops on {db.floor_label(floor_key)}.")
        return ConversationHandler.END
    label = "All Floors" if floor_key is None else db.floor_label(floor_key)
    await query.edit_message_text(f"{label} — which shop are you renting out?", reply_markup=kb)
    return ADDTENANT_SELECT


async def addtenant_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["new_tenant"] = {"shop_no": shop_no}
    await query.edit_message_text(f"Shop {shop_no} — tenant's name?")
    return ADDTENANT_NAME


async def addtenant_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_tenant"]["tenant_name"] = update.message.text.strip()
    await update.message.reply_text(
        "What's the shop for? (e.g. Café, Pharmacy, Clothing, Electronics)"
    )
    return ADDTENANT_PURPOSE


async def addtenant_purpose(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_tenant"]["purpose"] = update.message.text.strip()
    await update.message.reply_text("Tenant's phone number? (send 'skip' if none)")
    return ADDTENANT_PHONE


async def addtenant_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    context.user_data["new_tenant"]["phone"] = None if text.lower() == "skip" else text
    await update.message.reply_text("Tenancy start date? Send as YYYY-MM-DD, or 'skip' for today.")
    return ADDTENANT_START


async def addtenant_start_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() == "skip":
        context.user_data["new_tenant"]["start_date"] = None
    else:
        try:
            datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            await update.message.reply_text("Please use YYYY-MM-DD format, or send 'skip'.")
            return ADDTENANT_START
        context.user_data["new_tenant"]["start_date"] = text
    await update.message.reply_text(
        "Leasing period end date? Send as YYYY-MM-DD (G.C.), or 'skip' for an open-ended lease."
    )
    return ADDTENANT_LEASE_END


async def addtenant_lease_end(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if text.lower() == "skip":
        context.user_data["new_tenant"]["lease_end_date"] = None
    else:
        try:
            datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            await update.message.reply_text("Please use YYYY-MM-DD format, or send 'skip'.")
            return ADDTENANT_LEASE_END
        context.user_data["new_tenant"]["lease_end_date"] = text
    await update.message.reply_text(
        "Send a photo or file of the signed lease document, or tap Skip.",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("⏭ Skip", callback_data="addtenantdoc:skip")]]
        ),
    )
    return ADDTENANT_DOC


async def addtenant_doc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    d = context.user_data["new_tenant"]
    if update.message.photo:
        file_id, kind = update.message.photo[-1].file_id, "photo"
    elif update.message.document:
        file_id, kind = update.message.document.file_id, "document"
    else:
        await update.message.reply_text(
            "Please send a photo or a file, or tap Skip.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("⏭ Skip", callback_data="addtenantdoc:skip")]]
            ),
        )
        return ADDTENANT_DOC
    d.setdefault("lease_docs", []).append({"file_id": file_id, "kind": kind, "label": "Lease document"})
    await update.message.reply_text(
        f"Document {len(d['lease_docs'])} saved. Add another document?",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("➕ Add Another Document", callback_data="addtenantdoc:more"),
              InlineKeyboardButton("✅ Done", callback_data="addtenantdoc:done")]]
        ),
    )
    return ADDTENANT_DOC_MORE


async def addtenant_doc_more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "addtenantdoc:done":
        await query.edit_message_reply_markup(reply_markup=None)
        return await _addtenant_show_confirm(update, context)
    await query.edit_message_reply_markup(reply_markup=None)
    await query.message.reply_text("Send the next document, or /cancel.")
    return ADDTENANT_DOC


async def addtenant_doc_skip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    return await _addtenant_show_confirm(update, context)


async def _addtenant_show_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    d = context.user_data["new_tenant"]
    shop = db.get_shop(d["shop_no"])
    start_label = f"{gc_iso_to_ec_label(d['start_date'])} E.C." if d["start_date"] else "today"
    end_label = f"{gc_iso_to_ec_label(d['lease_end_date'])} E.C." if d["lease_end_date"] else "open-ended"
    doc_count = len(d.get("lease_docs", []))
    doc_label = f"{doc_count} attached" if doc_count else "none attached"
    text = (
        "Please confirm:\n"
        f"Shop: {d['shop_no']}\n"
        f"Tenant: {d['tenant_name']}\n"
        f"Purpose: {d['purpose']}\n"
        f"Phone: {d['phone'] or '-'}\n"
        f"Rent: {money(shop['monthly_rent'])}\n"
        f"Lease start: {start_label}\n"
        f"Lease end: {end_label}\n"
        f"Lease documents: {doc_label}"
    )
    kb = InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Save", callback_data="addtenant:save"),
          InlineKeyboardButton("✖ Cancel", callback_data="addtenant:cancel")]]
    )
    if update.callback_query:
        await update.callback_query.message.reply_text(text, reply_markup=kb)
    else:
        await update.message.reply_text(text, reply_markup=kb)
    return ADDTENANT_CONFIRM


async def addtenant_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    d = context.user_data.pop("new_tenant", None)
    if query.data == "addtenant:cancel" or not d:
        await query.edit_message_text("Cancelled.")
        await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
        return ConversationHandler.END
    code = db.assign_tenant(
        d["shop_no"], d["tenant_name"], phone=d["phone"], purpose=d["purpose"],
        start_date=d["start_date"],
    )
    if d.get("lease_end_date"):
        db.set_lease_end(d["shop_no"], d["lease_end_date"])
    for doc in d.get("lease_docs", []):
        _save_shop_document(d["shop_no"], doc["file_id"], doc["kind"], doc["label"])
    await query.edit_message_text(
        f"Shop {d['shop_no']} is now rented to {d['tenant_name']}.\n"
        f"Link code for the tenant: {code}\nThey send: /register {code}"
    )
    await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Mark Vacant (tenant moved out)
# ---------------------------------------------------------------------------

@flow_entry("vacate")
async def vacate_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rented = db.get_all_shops(active_only=True, rented_only=True)
    kb, shops = shop_picker_kb("vacateshop", shops=rented)
    if not shops:
        await send(update, "No rented shops right now.")
        return ConversationHandler.END
    await send(update, "Which shop's tenant moved out?", reply_markup=kb)
    return VACATE_SELECT


async def vacate_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["vacate_shop_no"] = shop_no
    shop = db.get_shop(shop_no)
    await query.edit_message_text(
        f"Mark Shop {shop_no} ({shop['tenant_name']}) vacant? This clears the tenant's "
        "info and unlinks them from Telegram. Balance/ledger history is kept.",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Confirm", callback_data="vacate:save"),
              InlineKeyboardButton("✖ Cancel", callback_data="vacate:cancel")]]
        ),
    )
    return VACATE_CONFIRM


async def vacate_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = context.user_data.pop("vacate_shop_no", None)
    if query.data == "vacate:cancel" or not shop_no:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    db.mark_vacant(shop_no)
    await query.edit_message_text(f"Shop {shop_no} is now vacant.")
    await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Record Payment
# ---------------------------------------------------------------------------

@flow_entry("pay")
async def pay_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not db.get_all_shops(active_only=True, rented_only=True):
        await send(update, "No rented shops yet. Add a tenant first.")
        return ConversationHandler.END
    await send(update, "Which floor is the shop on?", reply_markup=floor_and_all_kb("payfloor"))
    return PAY_FLOOR


async def pay_floor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    key = query.data.split(":", 1)[1]
    if key == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    floor_key = None if key == "ALL" else key
    shops = db.get_all_shops(active_only=True, rented_only=True)
    if floor_key is not None:
        shops = [s for s in shops if s["floor"] == floor_key]
    kb, shops = shop_picker_kb("payshop", shops=shops)
    if not shops:
        await query.edit_message_text(f"No shops on {db.floor_label(floor_key)}.")
        return ConversationHandler.END
    label = "All Floors" if floor_key is None else db.floor_label(floor_key)
    await query.edit_message_text(f"{label} — which shop is this payment for?", reply_markup=kb)
    return PAY_SELECT


async def pay_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["pay"] = {"shop_no": shop_no, "resolved": set(), "auto": set()}
    shop = db.get_shop(shop_no)
    paid = db.paid_periods(shop["id"]) if shop else set()
    context.user_data["pay"]["paid_periods"] = paid
    await query.edit_message_text(
        f"Recording a payment for Shop {shop_no}.\nWhich month is this payment for?",
        reply_markup=pay_month_kb(paid_periods=paid),
    )
    return PAY_MONTH


async def pay_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":", 2)
    action = parts[1]
    if action == "CANCEL" or "pay" not in context.user_data:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    if action == "PAGE":
        page = int(parts[2])
        paid = context.user_data["pay"].get("paid_periods", set())
        await query.edit_message_reply_markup(reply_markup=pay_month_kb(page, paid_periods=paid))
        return PAY_MONTH
    value = parts[2]  # action == "SELECT"
    yr, mo = value.split("-")
    context.user_data["pay"]["period"] = value
    context.user_data["pay"]["period_label"] = f"{ETHIOPIAN_MONTHS[int(mo)]} {yr}"
    # Recorded up front (not just at the amount step) so it's also
    # available if the amount instead gets filled in from a receipt photo.
    shop = db.get_shop(context.user_data["pay"]["shop_no"])
    context.user_data["pay"]["expected_rent"] = (
        db.rent_for_period(shop["id"], value) if shop else None
    )
    await query.edit_message_text(
        f"Month: {context.user_data['pay']['period_label']}\n\n"
        "Please send a photo of the receipt — I'll read the bank, amount, date and "
        "reference number off it myself. Tap Skip if you don't have one.",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("⏭ Skip", callback_data="payreceipt:skip")]]
        ),
    )
    return PAY_RECEIPT


async def pay_receipt_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "pay" not in context.user_data:
        return ConversationHandler.END
    d = context.user_data["pay"]
    photo = update.message.photo[-1]
    d["receipt_file_id"] = photo.file_id

    tg_file = await context.bot.get_file(photo.file_id)
    image_bytes = bytes(await tg_file.download_as_bytearray())
    guesses = extract_receipt_info(image_bytes)

    found = []
    for key, label in (("bank", "bank"), ("amount", "amount"), ("payment_date", "date"),
                        ("reference_no", "reference number")):
        value = guesses.get(key)
        if value is not None:
            d[key] = value
            if key == "payment_date":
                d["payment_date_ec"] = gc_iso_to_ec_label(value)
            d["resolved"].add(key)
            d["auto"].add(key)
            found.append(label)

    if found:
        await update.message.reply_text("From the receipt, I picked up the " + ", ".join(found) + ".")
    elif OCR_AVAILABLE:
        await update.message.reply_text("Couldn't confidently read the receipt — I'll ask for the details.")
    else:
        await update.message.reply_text("Receipt attached.")

    return await pay_advance(update, context)


async def pay_receipt_skip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if "pay" not in context.user_data:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["pay"]["receipt_file_id"] = None
    return await pay_advance(update, context)


async def pay_advance(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Ask for the next field the receipt (or a previous answer) hasn't already
    given us, or show the final approval summary once everything is known."""
    d = context.user_data["pay"]
    resolved = d["resolved"]

    if "bank" not in resolved:
        await _pay_reply(update, "Which bank/method was this paid through?", pay_bank_kb())
        return PAY_BANK
    if "amount" not in resolved:
        expected = d.get("expected_rent")
        hint = f" (rent for {d.get('period_label', 'this month')} is {money(expected)})" if expected is not None else ""
        await _pay_reply(update, f"Method: {d['bank']}\n\nHow much was paid?{hint}")
        return PAY_AMOUNT
    if "payment_date" not in resolved:
        await _pay_reply(update, "Payment date? Send as YYYY-MM-DD, or 'skip' for today.")
        return PAY_DATE
    if "reference_no" not in resolved:
        await _pay_reply(update, "Reference/transaction number? (send 'skip' if none)")
        return PAY_REF

    await _pay_reply(update, *_pay_confirm_message(d))
    return PAY_CONFIRM


async def _pay_reply(update: Update, text: str, reply_markup=None):
    if update.callback_query:
        await update.callback_query.edit_message_text(text, reply_markup=reply_markup)
    else:
        await update.message.reply_text(text, reply_markup=reply_markup)


def _pay_confirm_message(d):
    def line(key, label, value):
        tag = " (from receipt)" if key in d.get("auto", set()) else ""
        return f"{label}: {value}{tag}"

    heading = "I filtered these details from the receipt — please approve:" \
        if d.get("auto") else "Please confirm this payment:"
    date_display = f"{d.get('payment_date') or 'today'} G.C. / {d.get('payment_date_ec', '-')} E.C."
    lines = [
        heading,
        f"Shop: {d['shop_no']}",
        f"Month: {d.get('period_label', '-')}",
        line("bank", "Method", d.get("bank") or "-"),
        line("amount", "Amount", money(d["amount"])),
        line("payment_date", "Date", date_display),
        line("reference_no", "Reference", d.get("reference_no") or "-"),
        f"Receipt: {'attached' if d.get('receipt_file_id') else 'none'}",
    ]
    expected = d.get("expected_rent")
    if expected is not None and round(d["amount"], 2) != round(expected, 2):
        diff = round(d["amount"] - expected, 2)
        word = "over" if diff > 0 else "short of"
        lines.append(
            f"⚠️ Rent for {d.get('period_label', 'this period')} is {money(expected)} — "
            f"this is {money(abs(diff))} {word} that. You can still save it if that's correct."
        )
    kb = InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Approve & Save", callback_data="payconfirm:save"),
          InlineKeyboardButton("✖ Cancel", callback_data="payconfirm:cancel")]]
    )
    return "\n".join(lines), kb


async def pay_bank(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    value = query.data.split(":", 1)[1]
    if value == "CANCEL" or "pay" not in context.user_data:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    d = context.user_data["pay"]
    d["bank"] = value
    d["resolved"].add("bank")
    return await pay_advance(update, context)


async def pay_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("Enter a number, e.g. 3000")
        return PAY_AMOUNT
    d = context.user_data["pay"]
    d["amount"] = amount
    d["resolved"].add("amount")
    return await pay_advance(update, context)


async def pay_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    d = context.user_data["pay"]
    if text.lower() == "skip":
        gc_date = date.today().isoformat()
    else:
        try:
            datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            await update.message.reply_text("Please use YYYY-MM-DD format, or send 'skip' for today.")
            return PAY_DATE
        gc_date = text
    d["payment_date"] = gc_date
    d["payment_date_ec"] = gc_iso_to_ec_label(gc_date)
    d["resolved"].add("payment_date")
    await update.message.reply_text(f"📅 {gc_date} G.C. — {d['payment_date_ec']} E.C.")
    return await pay_advance(update, context)


async def pay_ref(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    d = context.user_data["pay"]
    d["reference_no"] = None if text.lower() == "skip" else text
    d["resolved"].add("reference_no")
    return await pay_advance(update, context)


async def pay_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    d = context.user_data.pop("pay", None)
    if query.data == "payconfirm:cancel" or not d:
        await query.edit_message_text("Cancelled.")
        await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
        return ConversationHandler.END

    shop_no = d["shop_no"]
    note_bits = [d["bank"]] if d.get("bank") else []
    if d.get("reference_no"):
        note_bits.append(f"Ref: {d['reference_no']}")
    note = " · ".join(note_bits) or None

    db.add_payment(
        shop_no, d["amount"], note,
        payment_date=d.get("payment_date"),
        period=d.get("period"),
        bank=d.get("bank"),
        reference_no=d.get("reference_no"),
        receipt_file_id=d.get("receipt_file_id"),
    )
    balance = db.get_balance(shop_no)
    await query.edit_message_text(
        f"Recorded {money(d['amount'])} for Shop {shop_no} ({d.get('period_label', '-')}).\n"
        f"New balance: {money(balance)}."
    )
    if d.get("receipt_file_id"):
        try:
            await context.bot.send_photo(
                query.message.chat_id, d["receipt_file_id"], caption="Receipt on file."
            )
        except Exception as e:
            logger.warning("Could not resend receipt to admin: %s", e)
    await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())

    shop = db.get_shop(shop_no)
    if shop and shop["telegram_id"]:
        try:
            await context.bot.send_message(
                shop["telegram_id"],
                f"Payment of {money(d['amount'])} received for Shop {shop_no} "
                f"({d.get('period_label', '-')}). Your balance is now {money(balance)}.",
            )
            if d.get("receipt_file_id"):
                await context.bot.send_photo(
                    shop["telegram_id"], d["receipt_file_id"], caption="Your receipt."
                )
        except Exception as e:
            logger.warning("Could not notify tenant %s: %s", shop_no, e)
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Add Expense
# ---------------------------------------------------------------------------

COMMON_EXPENSE_REASONS = [
    ("⚡ Electricity Bill", "Electricity Bill"),
    ("🚰 Water Bill", "Water Bill"),
    ("💼 Salary", "Salary"),
]


def expense_reason_kb():
    rows = [[InlineKeyboardButton(label, callback_data=f"expreason:{value}")]
            for label, value in COMMON_EXPENSE_REASONS]
    rows.append([InlineKeyboardButton("✏️ Other", callback_data="expreason:OTHER")])
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="expreason:CANCEL")])
    return InlineKeyboardMarkup(rows)


@flow_entry("expense")
async def expense_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send(update, "Which month was this expense for?", reply_markup=expense_month_kb())
    return EXPENSE_MONTH


async def expense_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":", 2)
    action = parts[1]
    if action == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    if action == "PAGE":
        page = int(parts[2])
        await query.edit_message_reply_markup(reply_markup=expense_month_kb(page))
        return EXPENSE_MONTH
    yr, mo = (int(x) for x in parts[2].split("-"))  # action == "SELECT"
    gy, gm, gd = ethiopian_to_gregorian(yr, mo, 1)
    context.user_data["exp_month_label"] = f"{ETHIOPIAN_MONTHS[mo]} {yr}"
    context.user_data["exp_date"] = date(gy, gm, gd).isoformat()
    await query.edit_message_text(
        f"Month: {context.user_data['exp_month_label']}\n\nWhat was it for?",
        reply_markup=expense_reason_kb(),
    )
    return EXPENSE_REASON


async def expense_reason(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    value = query.data.split(":", 1)[1]
    if value == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    if value == "OTHER":
        await query.edit_message_text(
            f"Month: {context.user_data.get('exp_month_label', '-')}\n\nWhat was it for?"
        )
        return EXPENSE_DESC
    context.user_data["exp_desc"] = value
    context.user_data["exp_category"] = value
    await query.edit_message_text(
        f"Month: {context.user_data.get('exp_month_label', '-')}\n"
        f"Reason: {value}\n\nHow much was the expense?"
    )
    return EXPENSE_AMOUNT


async def expense_desc(update: Update, context: ContextTypes.DEFAULT_TYPE):
    desc = update.message.text.strip()
    context.user_data["exp_desc"] = desc
    await update.message.reply_text("How much was the expense?")
    return EXPENSE_AMOUNT


async def expense_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("Enter a number, e.g. 1200")
        return EXPENSE_AMOUNT
    context.user_data["exp_amount"] = amount
    await update.message.reply_text(
        "Send a photo of the receipt, or tap Skip if you don't have one.",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("⏭ Skip", callback_data="expreceipt:skip")]]
        ),
    )
    return EXPENSE_RECEIPT


async def expense_receipt_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["exp_receipt_file_id"] = update.message.photo[-1].file_id
    await expense_show_confirm(update, context)
    return EXPENSE_CONFIRM


async def expense_receipt_skip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    context.user_data["exp_receipt_file_id"] = None
    await expense_show_confirm(update, context)
    return EXPENSE_CONFIRM


async def expense_show_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    d = context.user_data
    text = (
        f"Log this expense?\n\n"
        f"Month: {d.get('exp_month_label', '-')}\n"
        f"Reason: {d.get('exp_desc', '-')}\n"
        f"Amount: {money(d.get('exp_amount', 0))}\n"
        f"Receipt: {'attached' if d.get('exp_receipt_file_id') else 'none'}"
    )
    kb = InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Save", callback_data="expconfirm:save"),
          InlineKeyboardButton("✖ Cancel", callback_data="expconfirm:cancel")]]
    )
    if update.callback_query:
        await update.callback_query.edit_message_text(text, reply_markup=kb)
    else:
        await update.message.reply_text(text, reply_markup=kb)


async def expense_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    amount = context.user_data.pop("exp_amount", None)
    desc = context.user_data.pop("exp_desc", None)
    expense_date = context.user_data.pop("exp_date", None)
    context.user_data.pop("exp_month_label", None)
    receipt_file_id = context.user_data.pop("exp_receipt_file_id", None)
    category = context.user_data.pop("exp_category", None)
    if query.data == "expconfirm:cancel" or amount is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    db.add_expense(desc, amount, category=category, expense_date=expense_date, receipt_file_id=receipt_file_id)
    await query.edit_message_text(f"Recorded expense: {money(amount)} — {desc}")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Edit Expense (pick a month, pick one of its expenses, pick a
# field, send the new value — same confirm-before-save pattern as Edit Payment)
# ---------------------------------------------------------------------------

EDITEXP_MONTHS_PER_PAGE = 4
EDITEXP_MAX_LISTED = 40  # Telegram allows at most 100 inline buttons per message


def editexp_month_kb(page=0):
    """Paged Ethiopian-month picker, newest first: page 0 is the current
    month plus the three before it, and 'Older' steps further back."""
    today = date.today()
    ey, em, _ = gregorian_to_ethiopian(today.year, today.month, today.day)
    start_offset = -(EDITEXP_MONTHS_PER_PAGE - 1) - page * EDITEXP_MONTHS_PER_PAGE
    buttons = []
    for i in range(EDITEXP_MONTHS_PER_PAGE):
        yr, mo = _add_eth_months(ey, em, start_offset + i)
        buttons.append(InlineKeyboardButton(
            f"{ETHIOPIAN_MONTHS[mo]} {yr}", callback_data=f"editexpmonth:SELECT:{yr}-{mo:02d}"
        ))
    rows = [buttons[i:i + 2] for i in range(0, len(buttons), 2)]
    nav = [InlineKeyboardButton("◀ Older", callback_data=f"editexpmonth:PAGE:{page + 1}")]
    if page > 0:
        nav.append(InlineKeyboardButton("Newer ▶", callback_data=f"editexpmonth:PAGE:{page - 1}"))
    rows.append(nav)
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="editexpmonth:CANCEL")])
    return InlineKeyboardMarkup(rows)


def editexp_pick_kb(expenses):
    rows = []
    for r in expenses[:EDITEXP_MAX_LISTED]:
        label = f"{gc_iso_to_ec_label(r['expense_date'])} · {money(r['amount'])} · {r['description']}"
        if len(label) > 55:
            label = label[:54] + "…"
        rows.append([InlineKeyboardButton(label, callback_data=f"editexppick:{r['id']}")])
    rows.append([InlineKeyboardButton("« Other month", callback_data="editexppick:BACK"),
                 InlineKeyboardButton("✖ Cancel", callback_data="editexppick:CANCEL")])
    return InlineKeyboardMarkup(rows)


def editexp_field_kb():
    rows = [
        [InlineKeyboardButton("📝 Description", callback_data="editexpfield:description"),
         InlineKeyboardButton("💵 Amount", callback_data="editexpfield:amount")],
        [InlineKeyboardButton("🏷 Category", callback_data="editexpfield:category"),
         InlineKeyboardButton("📅 Date", callback_data="editexpfield:expense_date")],
        [InlineKeyboardButton("🧾 Receipt", callback_data="editexpfield:receipt_file_id")],
        [InlineKeyboardButton("✖ Cancel", callback_data="editexpfield:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


def _editexp_confirm_kb():
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Confirm", callback_data="editexpconfirm:save"),
          InlineKeyboardButton("✖ Cancel", callback_data="editexpconfirm:cancel")]]
    )


def _clear_editexp_state(context):
    for key in ("editexp_id", "editexp_field", "editexp_pending", "editexp_month_label"):
        context.user_data.pop(key, None)


def _parse_ec_date(text):
    """'2019-02-15' (Ethiopian calendar) -> Gregorian 'YYYY-MM-DD', or None
    if it isn't a real Ethiopian date (day 31, month 14, Pagume 7, ...)."""
    try:
        y, m, d = (int(x) for x in text.strip().split("-"))
        gy, gm, gd = ethiopian_to_gregorian(y, m, d)
        if tuple(gregorian_to_ethiopian(gy, gm, gd)) != (y, m, d):
            return None
        return date(gy, gm, gd).isoformat()
    except (ValueError, OverflowError):
        return None


@flow_entry("editexp")
async def editexp_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send(update, "Which month's expense do you want to edit?", reply_markup=editexp_month_kb())
    return EDITEXP_MONTH


async def editexp_month(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":", 2)
    action = parts[1]
    if action == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    if action == "PAGE":
        await query.edit_message_reply_markup(reply_markup=editexp_month_kb(int(parts[2])))
        return EDITEXP_MONTH
    yr, mo = (int(x) for x in parts[2].split("-"))  # action == "SELECT"
    label = f"{ETHIOPIAN_MONTHS[mo]} {yr}"
    rows = _expenses_for_eth_month(yr, mo)
    if not rows:
        await query.edit_message_text(
            f"No expenses recorded for {label}.\n\nPick another month:",
            reply_markup=editexp_month_kb(),
        )
        return EDITEXP_MONTH
    context.user_data["editexp_month_label"] = label
    note = ""
    if len(rows) > EDITEXP_MAX_LISTED:
        note = f"\n(Showing the first {EDITEXP_MAX_LISTED} of {len(rows)}.)"
    await query.edit_message_text(
        f"{label} — which expense do you want to edit?{note}",
        reply_markup=editexp_pick_kb(rows),
    )
    return EDITEXP_PICK


async def editexp_pick(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    value = query.data.split(":", 1)[1]
    if value == "CANCEL":
        _clear_editexp_state(context)
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    if value == "BACK":
        await query.edit_message_text(
            "Which month's expense do you want to edit?", reply_markup=editexp_month_kb()
        )
        return EDITEXP_MONTH
    expense = db.get_expense(int(value))
    if not expense:
        await query.edit_message_text("That expense no longer exists.")
        return ConversationHandler.END
    context.user_data["editexp_id"] = expense["id"]
    lines = [
        f"Expense on {gc_iso_to_ec_label(expense['expense_date'])} ({expense['expense_date']} G.C.):",
        f"Description: {expense['description']}",
        f"Amount: {money(expense['amount'])}",
        f"Category: {expense['category'] or '-'}",
        f"Receipt: {'attached' if expense.get('receipt_file_id') else 'none'}",
    ]
    if expense.get("recurring_expense_id"):
        lines.append("ℹ️ Monthly recurring expense — edits here change only this month's entry.")
    lines.append("\nWhat do you want to change?")
    await query.edit_message_text("\n".join(lines), reply_markup=editexp_field_kb())
    return EDITEXP_FIELD


async def editexp_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    field = query.data.split(":", 1)[1]
    if field == "CANCEL":
        _clear_editexp_state(context)
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["editexp_field"] = field
    if field == "receipt_file_id":
        expense = db.get_expense(context.user_data.get("editexp_id"))
        if expense and expense.get("receipt_file_id"):
            prompt = ("Send a photo of the new receipt to replace the current one, "
                      "or type 'remove' to delete it.")
        else:
            prompt = "Send a photo of the receipt to attach it to this expense."
        await query.edit_message_text(prompt)
        return EDITEXP_VALUE
    prompts = {
        "description": "Send the new description.",
        "amount": "Send the new amount, e.g. 1200.",
        "category": "Send the new category, e.g. Repairs, or 'skip' to clear it.",
        "expense_date": "Send the new date in the Ethiopian calendar as YYYY-MM-DD, e.g. 2019-02-15.",
    }
    await query.edit_message_text(prompts[field])
    return EDITEXP_VALUE


async def editexp_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    field = context.user_data.get("editexp_field")
    text = update.message.text.strip()
    if field == "description":
        if not text:
            await update.message.reply_text("The description can't be empty.")
            return EDITEXP_VALUE
        value = text
        preview = f"Change the description to: {value}?"
    elif field == "amount":
        try:
            value = float(text)
            if value <= 0 or value != value or value == float("inf"):
                raise ValueError
        except ValueError:
            await update.message.reply_text("That's not a valid amount. Send a number, e.g. 1200")
            return EDITEXP_VALUE
        preview = f"Change the amount to {money(value)}?"
    elif field == "category":
        value = None if text.lower() == "skip" else text
        preview = "Clear this expense's category?" if value is None else f"Change the category to {value}?"
    elif field == "expense_date":
        value = _parse_ec_date(text)
        if not value:
            await update.message.reply_text(
                "That isn't a valid Ethiopian date. Use YYYY-MM-DD, e.g. 2019-02-15."
            )
            return EDITEXP_VALUE
        preview = f"Change the date to {gc_iso_to_ec_label(value)} ({value} G.C.)?"
    elif field == "receipt_file_id":
        if text.lower() not in ("remove", "skip"):
            await update.message.reply_text(
                "Please send a photo of the receipt, or type 'remove' to delete the current one."
            )
            return EDITEXP_VALUE
        expense = db.get_expense(context.user_data.get("editexp_id"))
        if not expense or not expense.get("receipt_file_id"):
            await update.message.reply_text(
                "This expense has no receipt to remove. Send a photo to attach one."
            )
            return EDITEXP_VALUE
        value = None
        preview = "Remove this expense's receipt?"
    else:
        await update.message.reply_text("Nothing changed.", reply_markup=admin_menu_kb())
        _clear_editexp_state(context)
        return ConversationHandler.END
    context.user_data["editexp_pending"] = (field, value)
    await update.message.reply_text(preview, reply_markup=_editexp_confirm_kb())
    return EDITEXP_CONFIRM


async def editexp_receipt_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """A photo arrived while editing an expense: it's the new receipt only if
    the admin picked the Receipt field; otherwise gently point them back."""
    if context.user_data.get("editexp_field") != "receipt_file_id":
        await update.message.reply_text(
            "I'm not expecting a photo for that field — send the new value as text."
        )
        return EDITEXP_VALUE
    expense = db.get_expense(context.user_data.get("editexp_id"))
    context.user_data["editexp_pending"] = ("receipt_file_id", update.message.photo[-1].file_id)
    if expense and expense.get("receipt_file_id"):
        preview = "Replace this expense's receipt with the photo you just sent?"
    else:
        preview = "Attach the photo you just sent as this expense's receipt?"
    await update.message.reply_text(preview, reply_markup=_editexp_confirm_kb())
    return EDITEXP_CONFIRM


async def editexp_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    expense_id = context.user_data.get("editexp_id")
    pending = context.user_data.get("editexp_pending")
    _clear_editexp_state(context)
    if query.data == "editexpconfirm:cancel" or not pending or expense_id is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    field, value = pending
    updated = db.update_expense(expense_id, **{field: (db.CLEAR if value is None else value)})
    if not updated:
        await query.edit_message_text("That expense no longer exists.")
        return ConversationHandler.END
    await query.edit_message_text("Expense updated.")
    if field == "receipt_file_id" and value:
        try:
            await context.bot.send_photo(query.message.chat_id, value, caption="Receipt on file.")
        except Exception as e:
            logger.warning("Could not resend receipt to admin: %s", e)
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Permanent (recurring) Expenses — Electricity Bill, Water Bill,
# Salary, or any other category the admin sets an amount for. Once set, the
# amount is applied automatically once per Ethiopian month, the same way
# rent charges are (see monthly_charge_job and the "Charge Rent Now" button).
# Standalone small ConversationHandler, same pattern as profile_conv — not
# part of the big shared admin_conv, so it doesn't need flow_entry/FLOW_LABELS.
# ---------------------------------------------------------------------------

PERMEXP_SELECT, PERMEXP_AMOUNT, PERMEXP_CONFIRM = range(3)


def permexp_list_kb():
    rows = []
    for cat in db.PERMANENT_EXPENSE_CATEGORIES:
        rec = db.get_recurring_expense(cat)
        if rec and rec["active"]:
            label = f"{cat} — {money(rec['amount'])}/mo ✅"
        else:
            label = f"{cat} — not set"
        rows.append([InlineKeyboardButton(label, callback_data=f"permexpsel:{cat}")])
    rows.append([InlineKeyboardButton("✖ Close", callback_data="permexpsel:CLOSE")])
    return InlineKeyboardMarkup(rows)


async def permexp_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "Permanent (recurring) expenses — these get logged automatically on "
        "the 1st of every Ethiopian month, the same way rent is charged, so "
        "you don't have to add them by hand each time.\n\n"
        "Tap one to set or change its monthly amount, or set it to 0 to stop "
        "auto-applying it."
    )
    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text(text, reply_markup=permexp_list_kb())
    else:
        await update.message.reply_text(text, reply_markup=permexp_list_kb())
    return PERMEXP_SELECT


async def permexp_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    value = query.data.split(":", 1)[1]
    if value == "CLOSE":
        await query.edit_message_text("Closed.")
        return ConversationHandler.END
    context.user_data["permexp_cat"] = value
    rec = db.get_recurring_expense(value)
    current = f" (currently {money(rec['amount'])}/mo)" if rec and rec["active"] else ""
    await query.edit_message_text(
        f"{value}{current}\n\nSend the new monthly amount, or 0 to stop auto-applying it."
    )
    return PERMEXP_AMOUNT


async def permexp_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        amount = float(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("Enter a number, e.g. 15000, or 0 to disable.")
        return PERMEXP_AMOUNT
    context.user_data["permexp_amount"] = amount
    cat = context.user_data.get("permexp_cat")
    if amount <= 0:
        text = f"Stop auto-applying {cat} every month?"
    else:
        text = f"Set {cat} to {money(amount)}/month, applied automatically each month?"
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Confirm", callback_data="permexpconfirm:save"),
        InlineKeyboardButton("✖ Cancel", callback_data="permexpconfirm:cancel"),
    ]])
    await update.message.reply_text(text, reply_markup=kb)
    return PERMEXP_CONFIRM


async def permexp_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    cat = context.user_data.pop("permexp_cat", None)
    amount = context.user_data.pop("permexp_amount", None)
    if query.data == "permexpconfirm:cancel" or cat is None or amount is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    if amount <= 0:
        db.set_recurring_expense_active(cat, False)
        await query.edit_message_text(f"{cat} will no longer be applied automatically.")
    else:
        db.set_recurring_expense(cat, amount, active=True)
        await query.edit_message_text(
            f"{cat} set to {money(amount)}/month — it'll be applied automatically on the 1st "
            "of each Ethiopian month, alongside rent charges."
        )
    return ConversationHandler.END


permexp_conv = ConversationHandler(
    entry_points=[
        CommandHandler("permanentexpenses", permexp_start),
        CallbackQueryHandler(permexp_start, pattern="^more:permexp$"),
    ],
    states={
        PERMEXP_SELECT: [CallbackQueryHandler(permexp_select, pattern="^permexpsel:")],
        PERMEXP_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, permexp_amount)],
        PERMEXP_CONFIRM: [CallbackQueryHandler(permexp_confirm, pattern="^permexpconfirm:")],
    },
    fallbacks=[CommandHandler("cancel", cancel)],
)


# ---------------------------------------------------------------------------
# Admin flow: Edit Rent
# ---------------------------------------------------------------------------

@flow_entry("editrent")
async def editrent_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb, shops = shop_picker_kb("rentshop")
    if not shops:
        await send(update, "No shops yet.")
        return ConversationHandler.END
    await send(update, "Which shop's rent do you want to change?", reply_markup=kb)
    return RENT_SELECT


async def rent_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["rent_shop_no"] = shop_no
    shop = db.get_shop(shop_no)
    await query.edit_message_text(
        f"Shop {shop_no}'s current rent is {money(shop['monthly_rent'])}. What should the new rent be?"
    )
    return RENT_NEW


async def rent_new(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        rent = float(update.message.text.strip())
    except ValueError:
        await update.message.reply_text("Enter a number, e.g. 6000")
        return RENT_NEW
    context.user_data["rent_new_value"] = rent
    await update.message.reply_text(
        "From which month should this rent apply?", reply_markup=rent_month_kb()
    )
    return RENT_EFFECTIVE


async def rent_effective_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    parts = query.data.split(":")
    action = parts[1]
    if action == "CANCEL":
        await query.edit_message_text("Cancelled.")
        context.user_data.pop("rent_shop_no", None)
        context.user_data.pop("rent_new_value", None)
        return ConversationHandler.END
    if action == "PAGE":
        await query.edit_message_reply_markup(reply_markup=rent_month_kb(int(parts[2])))
        return RENT_EFFECTIVE
    # action == "SELECT"
    period = parts[2]
    yr, mo = (int(x) for x in period.split("-"))
    context.user_data["rent_effective_period"] = period
    shop_no = context.user_data["rent_shop_no"]
    rent = context.user_data["rent_new_value"]
    await query.edit_message_text(
        f"Update Shop {shop_no}'s rent to {money(rent)}, starting {ETHIOPIAN_MONTHS[mo]} {yr}?",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Confirm", callback_data="rentconfirm:save"),
              InlineKeyboardButton("✖ Cancel", callback_data="rentconfirm:cancel")]]
        ),
    )
    return RENT_CONFIRM


async def rent_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = context.user_data.pop("rent_shop_no", None)
    rent = context.user_data.pop("rent_new_value", None)
    effective_period = context.user_data.pop("rent_effective_period", None)
    if query.data == "rentconfirm:cancel" or shop_no is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    db.update_rent(shop_no, rent, effective_period)
    yr, mo = (int(x) for x in effective_period.split("-"))
    await query.edit_message_text(
        f"Shop {shop_no} rent updated to {money(rent)}, starting {ETHIOPIAN_MONTHS[mo]} {yr}."
    )
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Edit Shop (number / floor / area / rent — pick a field, then
# send the new value; each field is saved as soon as it's entered)
# ---------------------------------------------------------------------------

def editshop_field_kb():
    rows = [
        [InlineKeyboardButton("🔢 Shop Number", callback_data="editshopfield:shop_no"),
         InlineKeyboardButton("🏢 Floor", callback_data="editshopfield:floor")],
        [InlineKeyboardButton("📐 Area (m²)", callback_data="editshopfield:area"),
         InlineKeyboardButton("💵 Monthly Rent", callback_data="editshopfield:rent")],
        [InlineKeyboardButton("✖ Cancel", callback_data="editshopfield:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


@flow_entry("editshop")
async def editshop_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb, shops = shop_picker_kb("editshop")
    if not shops:
        await send(update, "No shops yet.")
        return ConversationHandler.END
    await send(update, "Which shop do you want to edit?", reply_markup=kb)
    return EDITSHOP_SELECT


async def editshop_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    shop = db.get_shop(shop_no)
    context.user_data["editshop_no"] = shop_no
    area_line = f"{shop['area_sqm']:g} m²" if shop["area_sqm"] else "-"
    await query.edit_message_text(
        f"Shop {shop_no}\n"
        f"Floor: {db.floor_label(shop['floor'])}\n"
        f"Area: {area_line}\n"
        f"Monthly rent: {money(shop['monthly_rent'])}\n\n"
        "What do you want to change?",
        reply_markup=editshop_field_kb(),
    )
    return EDITSHOP_FIELD


async def editshop_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    field = query.data.split(":", 1)[1]
    shop_no = context.user_data.get("editshop_no")
    if field == "CANCEL":
        await query.edit_message_text("Cancelled.")
        context.user_data.pop("editshop_no", None)
        return ConversationHandler.END
    context.user_data["editshop_field"] = field
    if field == "floor":
        await query.edit_message_text(f"Shop {shop_no} — pick the new floor.", reply_markup=floor_kb())
        return EDITSHOP_VALUE
    prompts = {
        "shop_no": "Send the new shop number.",
        "area": "Send the new area in m² (or 'skip' to clear it).",
        "rent": "Send the new monthly rent (e.g. 6000).",
    }
    await query.edit_message_text(f"Shop {shop_no} — {prompts[field]}")
    return EDITSHOP_VALUE


def _editshop_confirm_kb():
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Confirm", callback_data="editshopconfirm:save"),
          InlineKeyboardButton("✖ Cancel", callback_data="editshopconfirm:cancel")]]
    )


async def editshop_value_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    field = context.user_data.get("editshop_field")
    shop_no = context.user_data.get("editshop_no")
    text = update.message.text.strip()
    if field == "shop_no":
        if text != shop_no and db.get_shop(text):
            await update.message.reply_text(f"Shop {text} already exists. Send a different number, or /cancel.")
            return EDITSHOP_VALUE
        context.user_data["editshop_pending"] = ("shop_no", text)
        preview = f"Renumber Shop {shop_no} to {text}?"
    elif field == "area":
        if text.lower() == "skip":
            context.user_data["editshop_pending"] = ("area", None)
            preview = f"Clear Shop {shop_no}'s area?"
        else:
            try:
                area = float(text)
                if area <= 0:
                    raise ValueError
            except ValueError:
                await update.message.reply_text("That's not a valid area. Send a number, e.g. 24, or 'skip'.")
                return EDITSHOP_VALUE
            context.user_data["editshop_pending"] = ("area", area)
            preview = f"Update Shop {shop_no}'s area to {area:g} m²?"
    elif field == "rent":
        try:
            rent = float(text)
            if rent <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("That's not a valid amount. Send a number, e.g. 6000")
            return EDITSHOP_VALUE
        # A rent change always goes through the same "which month does this
        # start from" step as /editrent, so it's logged to rent_history the
        # same way no matter which flow the admin used to get here.
        context.user_data["rent_shop_no"] = shop_no
        context.user_data["rent_new_value"] = rent
        context.user_data.pop("editshop_no", None)
        context.user_data.pop("editshop_field", None)
        await update.message.reply_text(
            "From which month should this rent apply?", reply_markup=rent_month_kb()
        )
        return RENT_EFFECTIVE
    else:
        await update.message.reply_text("Nothing changed.", reply_markup=admin_menu_kb())
        context.user_data.pop("editshop_no", None)
        context.user_data.pop("editshop_field", None)
        return ConversationHandler.END
    await update.message.reply_text(preview, reply_markup=_editshop_confirm_kb())
    return EDITSHOP_CONFIRM


async def editshop_value_floor(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    key = query.data.split(":", 1)[1]
    shop_no = context.user_data.get("editshop_no")
    if key == "CANCEL":
        context.user_data.pop("editshop_no", None)
        context.user_data.pop("editshop_field", None)
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["editshop_pending"] = ("floor", key)
    await query.edit_message_text(
        f"Update Shop {shop_no}'s floor to {db.floor_label(key)}?",
        reply_markup=_editshop_confirm_kb(),
    )
    return EDITSHOP_CONFIRM


async def editshop_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = context.user_data.pop("editshop_no", None)
    pending = context.user_data.pop("editshop_pending", None)
    context.user_data.pop("editshop_field", None)
    if query.data == "editshopconfirm:cancel" or not pending or shop_no is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    field, value = pending
    if field == "shop_no":
        db.update_shop_no(shop_no, value)
        msg = f"Shop {shop_no} renumbered to {value}."
    elif field == "area":
        db.update_area(shop_no, value)
        msg = f"Shop {shop_no} area cleared." if value is None else f"Shop {shop_no} area updated to {value:g} m²."
    elif field == "floor":
        db.update_floor(shop_no, value)
        msg = f"Shop {shop_no} floor updated to {db.floor_label(value)}."
    else:
        msg = "Nothing changed."
    await query.edit_message_text(msg)
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Edit Tenant (name / phone / purpose / dates / ID / TIN for a
# shop that already has a tenant — pick a field, send the new value)
# ---------------------------------------------------------------------------

def edittenant_field_kb():
    rows = [
        [InlineKeyboardButton("👤 Name", callback_data="edittenantfield:name"),
         InlineKeyboardButton("📞 Phone", callback_data="edittenantfield:phone")],
        [InlineKeyboardButton("🏷 Purpose", callback_data="edittenantfield:purpose"),
         InlineKeyboardButton("📅 Start Date", callback_data="edittenantfield:start_date")],
        [InlineKeyboardButton("📅 Lease End", callback_data="edittenantfield:lease_end"),
         InlineKeyboardButton("🪪 National ID", callback_data="edittenantfield:national_id")],
        [InlineKeyboardButton("🧾 TIN Number", callback_data="edittenantfield:tin"),
         InlineKeyboardButton("✖ Cancel", callback_data="edittenantfield:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


@flow_entry("edittenant")
async def edittenant_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    rented = db.get_all_shops(active_only=True, rented_only=True)
    if not rented:
        await send(update, "No rented shops yet — use Add Tenant first.")
        return ConversationHandler.END
    kb, shops = shop_picker_kb("edittenantshop", shops=rented)
    await send(update, "Which tenant's details do you want to edit?", reply_markup=kb)
    return EDITTENANT_SELECT


async def edittenant_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    shop = db.get_shop(shop_no)
    context.user_data["edittenant_no"] = shop_no
    start_label = f"{gc_iso_to_ec_label(shop['start_date'])} E.C." if shop["start_date"] else "-"
    end_label = f"{gc_iso_to_ec_label(shop['lease_end_date'])} E.C." if shop["lease_end_date"] else "open-ended"
    await query.edit_message_text(
        f"Shop {shop_no} — {shop['tenant_name']}\n"
        f"Phone: {shop['phone'] or '-'}\n"
        f"Purpose: {shop['purpose'] or '-'}\n"
        f"Start: {start_label}\n"
        f"Lease end: {end_label}\n"
        f"National ID: {shop['national_id'] or '-'}\n"
        f"TIN: {shop['tin_number'] or '-'}\n\n"
        "What do you want to change?",
        reply_markup=edittenant_field_kb(),
    )
    return EDITTENANT_FIELD


async def edittenant_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    field = query.data.split(":", 1)[1]
    shop_no = context.user_data.get("edittenant_no")
    if field == "CANCEL":
        await query.edit_message_text("Cancelled.")
        context.user_data.pop("edittenant_no", None)
        return ConversationHandler.END
    context.user_data["edittenant_field"] = field
    prompts = {
        "name": "Send the tenant's new name.",
        "phone": "Send the new phone number (or 'skip' to clear it).",
        "purpose": "Send what the shop is now used for.",
        "start_date": "Send the new tenancy start date as YYYY-MM-DD.",
        "lease_end": "Send the new lease end date as YYYY-MM-DD (or 'skip' for open-ended).",
        "national_id": "Send the tenant's national ID number.",
        "tin": "Send the tenant's TIN number.",
    }
    await query.edit_message_text(f"Shop {shop_no} — {prompts[field]}")
    return EDITTENANT_VALUE


_EDITTENANT_LABELS = {
    "name": "name",
    "phone": "phone",
    "purpose": "purpose",
    "start_date": "start date",
    "lease_end": "lease end",
    "national_id": "national ID",
    "tin": "TIN",
}


async def edittenant_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    field = context.user_data.get("edittenant_field")
    shop_no = context.user_data.get("edittenant_no")
    text = update.message.text.strip()
    if field in ("start_date", "lease_end") and text.lower() != "skip":
        try:
            datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            hint = ", or 'skip'." if field == "lease_end" else "."
            await update.message.reply_text(f"Please use YYYY-MM-DD format{hint}")
            return EDITTENANT_VALUE

    if field == "name" and not text:
        await update.message.reply_text("Name can't be empty.")
        return EDITTENANT_VALUE

    if field in ("phone", "lease_end"):
        value = None if text.lower() == "skip" else text
    else:
        value = text

    context.user_data["edittenant_pending"] = (field, value)
    label = _EDITTENANT_LABELS.get(field, field)
    if value is None:
        preview = f"Clear Shop {shop_no}'s {label}?"
    else:
        preview = f"Update Shop {shop_no}'s {label} to \"{value}\"?"
    await update.message.reply_text(
        preview,
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Confirm", callback_data="edittenantconfirm:save"),
              InlineKeyboardButton("✖ Cancel", callback_data="edittenantconfirm:cancel")]]
        ),
    )
    return EDITTENANT_CONFIRM


async def edittenant_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = context.user_data.pop("edittenant_no", None)
    pending = context.user_data.pop("edittenant_pending", None)
    context.user_data.pop("edittenant_field", None)
    if query.data == "edittenantconfirm:cancel" or not pending or shop_no is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    field, value = pending
    if field == "name":
        db.update_tenant_name(shop_no, value)
        msg = f"Shop {shop_no} tenant name updated to {value}."
    elif field == "phone":
        db.update_phone(shop_no, value)
        msg = f"Shop {shop_no} phone {'cleared' if value is None else 'updated'}."
    elif field == "purpose":
        db.update_purpose(shop_no, value)
        msg = f"Shop {shop_no} purpose updated to {value}."
    elif field == "start_date":
        db.update_start_date(shop_no, value)
        msg = f"Shop {shop_no} start date updated."
    elif field == "lease_end":
        db.set_lease_end(shop_no, value)
        msg = f"Shop {shop_no} lease end {'cleared (open-ended)' if value is None else 'updated'}."
    elif field == "national_id":
        db.update_national_id(shop_no, value)
        msg = f"Shop {shop_no} national ID updated."
    elif field == "tin":
        db.update_tin(shop_no, value)
        msg = f"Shop {shop_no} TIN updated."
    else:
        msg = "Nothing changed."
    await query.edit_message_text(msg)
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Edit Payment (pick a shop, pick one of its recent payments,
# pick a field, send the new value — same confirm-before-save pattern as
# every other guided flow)
# ---------------------------------------------------------------------------

def editpay_payment_kb(shop_no, payments):
    rows = []
    for p in payments:
        label = f"{p['payment_date']} · {money(p['amount'])}"
        if p.get("bank"):
            label += f" · {p['bank']}"
        rows.append([InlineKeyboardButton(label, callback_data=f"editpaypmt:{p['id']}")])
    rows.append([InlineKeyboardButton("✖ Cancel", callback_data="editpaypmt:CANCEL")])
    return InlineKeyboardMarkup(rows)


def editpay_field_kb():
    rows = [
        [InlineKeyboardButton("💵 Amount", callback_data="editpayfield:amount"),
         InlineKeyboardButton("📅 Date", callback_data="editpayfield:payment_date")],
        [InlineKeyboardButton("🏦 Bank", callback_data="editpayfield:bank"),
         InlineKeyboardButton("🗓 Period", callback_data="editpayfield:period")],
        [InlineKeyboardButton("🔖 Reference No.", callback_data="editpayfield:reference_no"),
         InlineKeyboardButton("📝 Note", callback_data="editpayfield:note")],
        [InlineKeyboardButton("🧾 Receipt", callback_data="editpayfield:receipt_file_id")],
        [InlineKeyboardButton("✖ Cancel", callback_data="editpayfield:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


def editpay_bank_kb():
    rows = [
        [InlineKeyboardButton("🏦 CBE Birr", callback_data="editpaybank:CBE Birr"),
         InlineKeyboardButton("📱 Telebirr", callback_data="editpaybank:Telebirr")],
        [InlineKeyboardButton("💵 Cash", callback_data="editpaybank:Cash"),
         InlineKeyboardButton("🏛 Other Bank", callback_data="editpaybank:Other")],
        [InlineKeyboardButton("✖ Cancel", callback_data="editpaybank:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


@flow_entry("editpay")
async def editpay_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb, shops = shop_picker_kb("editpayshop")
    if not shops:
        await send(update, "No shops yet.")
        return ConversationHandler.END
    await send(update, "Which shop's payment do you want to edit?", reply_markup=kb)
    return EDITPAY_SELECT


async def editpay_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    payments = db.get_recent_payments(shop_no, limit=10)
    if not payments:
        await query.edit_message_text(f"Shop {shop_no} has no payments on record yet.")
        return ConversationHandler.END
    context.user_data["editpay_shop_no"] = shop_no
    await query.edit_message_text(
        f"Shop {shop_no} — which payment do you want to edit?",
        reply_markup=editpay_payment_kb(shop_no, payments),
    )
    return EDITPAY_PAYMENT


async def editpay_payment(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    value = query.data.split(":", 1)[1]
    if value == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    payment = db.get_payment(int(value))
    if not payment:
        await query.edit_message_text("That payment no longer exists.")
        return ConversationHandler.END
    context.user_data["editpay_id"] = payment["id"]
    shop_no = context.user_data.get("editpay_shop_no", "-")
    await query.edit_message_text(
        f"Shop {shop_no} — payment on {payment['payment_date']}:\n"
        f"Amount: {money(payment['amount'])}\n"
        f"Bank: {payment['bank'] or '-'}\n"
        f"Period: {payment['period'] or '-'}\n"
        f"Reference: {payment['reference_no'] or '-'}\n"
        f"Note: {payment['note'] or '-'}\n"
        f"Receipt: {'attached' if payment.get('receipt_file_id') else 'none'}\n\n"
        "What do you want to change?",
        reply_markup=editpay_field_kb(),
    )
    return EDITPAY_FIELD


async def editpay_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    field = query.data.split(":", 1)[1]
    if field == "CANCEL":
        await query.edit_message_text("Cancelled.")
        context.user_data.pop("editpay_id", None)
        context.user_data.pop("editpay_shop_no", None)
        return ConversationHandler.END
    context.user_data["editpay_field"] = field
    if field == "bank":
        await query.edit_message_text("Pick the new bank (or Other/Cash).", reply_markup=editpay_bank_kb())
        return EDITPAY_VALUE
    if field == "receipt_file_id":
        payment = db.get_payment(context.user_data.get("editpay_id"))
        if payment and payment.get("receipt_file_id"):
            prompt = ("Send a photo of the new receipt to replace the current one, "
                      "or type 'remove' to delete it.")
        else:
            prompt = "Send a photo of the receipt to attach it to this payment."
        await query.edit_message_text(prompt)
        return EDITPAY_VALUE
    prompts = {
        "amount": "Send the new amount, e.g. 6000.",
        "payment_date": "Send the new date as YYYY-MM-DD.",
        "period": "Send the new period as EY-EM (e.g. 2018-01), or 'skip' to clear it.",
        "reference_no": "Send the new reference number, or 'skip' to clear it.",
        "note": "Send the new note, or 'skip' to clear it.",
    }
    await query.edit_message_text(prompts[field])
    return EDITPAY_VALUE


def _editpay_confirm_kb():
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Confirm", callback_data="editpayconfirm:save"),
          InlineKeyboardButton("✖ Cancel", callback_data="editpayconfirm:cancel")]]
    )


async def editpay_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    field = context.user_data.get("editpay_field")
    text = update.message.text.strip()
    if field == "amount":
        try:
            value = float(text)
            if value <= 0:
                raise ValueError
        except ValueError:
            await update.message.reply_text("That's not a valid amount. Send a number, e.g. 6000")
            return EDITPAY_VALUE
        preview = f"Update this payment's amount to {money(value)}?"
    elif field == "payment_date":
        try:
            datetime.strptime(text, "%Y-%m-%d")
        except ValueError:
            await update.message.reply_text("Please use YYYY-MM-DD format.")
            return EDITPAY_VALUE
        value = text
        preview = f"Update this payment's date to {value}?"
    elif field == "period":
        value = None if text.lower() == "skip" else text
        preview = "Clear this payment's period?" if value is None else f"Update this payment's period to {value}?"
    elif field == "reference_no":
        value = None if text.lower() == "skip" else text
        preview = "Clear this payment's reference number?" if value is None else f"Update the reference number to {value}?"
    elif field == "note":
        value = None if text.lower() == "skip" else text
        preview = "Clear this payment's note?" if value is None else f"Update the note to {value}?"
    elif field == "receipt_file_id":
        if text.lower() not in ("remove", "skip"):
            await update.message.reply_text(
                "Please send a photo of the receipt, or type 'remove' to delete the current one."
            )
            return EDITPAY_VALUE
        payment = db.get_payment(context.user_data.get("editpay_id"))
        if not payment or not payment.get("receipt_file_id"):
            await update.message.reply_text(
                "This payment has no receipt to remove. Send a photo to attach one."
            )
            return EDITPAY_VALUE
        value = None
        preview = "Remove this payment's receipt?"
    else:
        await update.message.reply_text("Nothing changed.", reply_markup=admin_menu_kb())
        return ConversationHandler.END
    context.user_data["editpay_pending"] = (field, value)
    await update.message.reply_text(preview, reply_markup=_editpay_confirm_kb())
    return EDITPAY_CONFIRM


async def editpay_receipt_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """A photo arrived while editing a payment: it's the new receipt only if
    the admin picked the Receipt field; otherwise gently point them back."""
    if context.user_data.get("editpay_field") != "receipt_file_id":
        await update.message.reply_text(
            "I'm not expecting a photo for that field — send the new value as text."
        )
        return EDITPAY_VALUE
    payment = db.get_payment(context.user_data.get("editpay_id"))
    file_id = update.message.photo[-1].file_id
    context.user_data["editpay_pending"] = ("receipt_file_id", file_id)
    if payment and payment.get("receipt_file_id"):
        preview = "Replace this payment's receipt with the photo you just sent?"
    else:
        preview = "Attach the photo you just sent as this payment's receipt?"
    await update.message.reply_text(preview, reply_markup=_editpay_confirm_kb())
    return EDITPAY_CONFIRM


async def editpay_bank_value(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    value = query.data.split(":", 1)[1]
    if value == "CANCEL":
        await query.edit_message_text("Cancelled.")
        context.user_data.pop("editpay_id", None)
        context.user_data.pop("editpay_shop_no", None)
        return ConversationHandler.END
    context.user_data["editpay_pending"] = ("bank", value)
    await query.edit_message_text(
        f"Update this payment's bank to {value}?", reply_markup=_editpay_confirm_kb()
    )
    return EDITPAY_CONFIRM


async def editpay_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    payment_id = context.user_data.pop("editpay_id", None)
    shop_no = context.user_data.pop("editpay_shop_no", None)
    pending = context.user_data.pop("editpay_pending", None)
    context.user_data.pop("editpay_field", None)
    if query.data == "editpayconfirm:cancel" or not pending or payment_id is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    field, value = pending
    kwargs = {field: (db.CLEAR if value is None else value)}
    db.update_payment(payment_id, **kwargs)
    await query.edit_message_text(f"Payment updated for Shop {shop_no}.")
    if field == "receipt_file_id" and value:
        try:
            await context.bot.send_photo(
                query.message.chat_id, value, caption="Receipt on file."
            )
        except Exception as e:
            logger.warning("Could not resend receipt to admin: %s", e)
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Add Floor
# ---------------------------------------------------------------------------

@flow_entry("addfloors")
async def addfloors_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["new_floor"] = {}
    await send(
        update,
        "Let's add a floor. What should it be called? (e.g. 'Fourth Floor')\n"
        "Send /cancel anytime to stop.",
    )
    return ADDFLOOR_LABEL


async def addfloors_label(update: Update, context: ContextTypes.DEFAULT_TYPE):
    label = update.message.text.strip()
    if not label:
        await update.message.reply_text("Send a name for the floor, e.g. 'Fourth Floor'.")
        return ADDFLOOR_LABEL
    if label.lower() in {l.lower() for _, l in db.get_floors()}:
        await update.message.reply_text(f"A floor called '{label}' already exists. Send a different name, or /cancel.")
        return ADDFLOOR_LABEL
    context.user_data["new_floor"]["label"] = label
    await update.message.reply_text(
        f"Does '{label}' pay monthly rent like most floors, or the whole lease as one lump sum "
        "(like ETHIO TERIT / Tsedey Bank)?",
        reply_markup=_lumpsum_kb("addfloorlump"),
    )
    return ADDFLOOR_LUMPSUM


async def addfloors_lumpsum(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    lump_sum = query.data.split(":", 1)[1] == "yes"
    d = context.user_data.get("new_floor")
    if not d:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    d["lump_sum"] = lump_sum
    billing = "Lump sum (whole lease up front)" if lump_sum else "Monthly rent (normal)"
    await query.edit_message_text(
        "Please confirm:\n"
        f"Floor name: {d['label']}\n"
        f"Billing: {billing}",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Save", callback_data="addfloor:save"),
              InlineKeyboardButton("✖ Cancel", callback_data="addfloor:cancel")]]
        ),
    )
    return ADDFLOOR_CONFIRM


async def addfloors_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    d = context.user_data.pop("new_floor", None)
    if query.data == "addfloor:cancel" or not d:
        await query.edit_message_text("Cancelled.")
        await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
        return ConversationHandler.END
    key = db.add_floor(d["label"], lump_sum=d.get("lump_sum", False))
    await query.edit_message_text(f"Added floor '{d['label']}'. It'll show up wherever floors are picked.")
    await query.message.reply_text("Back to menu.", reply_markup=admin_menu_kb())
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Edit Floors (rename / change billing type / reorder / delete —
# pick a floor, pick what to change; reordering applies right away since
# it's freely reversible, everything else asks for confirmation first)
# ---------------------------------------------------------------------------

def editfloor_field_kb(key):
    f = db.get_floor(key)
    rows = [
        [InlineKeyboardButton("✏️ Rename", callback_data="editfloorfield:rename"),
         InlineKeyboardButton("💳 Billing Type", callback_data="editfloorfield:billing")],
        [InlineKeyboardButton("⬆️ Move Up", callback_data="editfloorfield:moveup"),
         InlineKeyboardButton("⬇️ Move Down", callback_data="editfloorfield:movedown")],
        [InlineKeyboardButton("🗑 Delete", callback_data="editfloorfield:delete")],
        [InlineKeyboardButton("✖ Done", callback_data="editfloorfield:CANCEL")],
    ]
    return InlineKeyboardMarkup(rows)


def _editfloor_summary(key):
    f = db.get_floor(key)
    if not f:
        return "This floor no longer exists."
    billing = "Lump sum" if f["lump_sum"] else "Monthly rent"
    shop_count = next((x["shop_count"] for x in db.get_floors_with_counts() if x["key"] == f["key"]), 0)
    return (
        f"{f['label']}\n"
        f"Billing: {billing}\n"
        f"Shops on this floor: {shop_count}\n\n"
        "What do you want to change?"
    )


@flow_entry("editfloors")
async def editfloors_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    floors = db.get_floors_with_counts()
    if not floors:
        await send(update, "No floors yet — use Add Floor first.")
        return ConversationHandler.END
    await send(update, "Which floor do you want to edit?", reply_markup=editfloor_select_kb())
    return EDITFLOOR_SELECT


async def editfloors_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    key = query.data.split(":", 1)[1]
    if key == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["editfloor_key"] = key
    await query.edit_message_text(_editfloor_summary(key), reply_markup=editfloor_field_kb(key))
    return EDITFLOOR_FIELD


async def editfloors_field(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    field = query.data.split(":", 1)[1]
    key = context.user_data.get("editfloor_key")
    if field == "CANCEL" or not key:
        await query.edit_message_text("Done.")
        context.user_data.pop("editfloor_key", None)
        return ConversationHandler.END
    if field in ("moveup", "movedown"):
        db.move_floor(key, "up" if field == "moveup" else "down")
        await query.edit_message_text(_editfloor_summary(key), reply_markup=editfloor_field_kb(key))
        return EDITFLOOR_FIELD
    if field == "delete":
        f = db.get_floor(key)
        shop_count = next((x["shop_count"] for x in db.get_floors_with_counts() if x["key"] == key), 0)
        if shop_count:
            await query.edit_message_text(
                f"Can't delete '{f['label']}' — {shop_count} shop(s) are still on it. "
                "Move or deactivate them first.",
                reply_markup=editfloor_field_kb(key),
            )
            return EDITFLOOR_FIELD
        context.user_data["editfloor_pending"] = ("delete", None)
        await query.edit_message_text(
            f"Delete floor '{f['label']}'? This can't be undone.",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("✅ Confirm", callback_data="editfloorconfirm:save"),
                  InlineKeyboardButton("✖ Cancel", callback_data="editfloorconfirm:cancel")]]
            ),
        )
        return EDITFLOOR_CONFIRM
    context.user_data["editfloor_field"] = field
    if field == "billing":
        await query.edit_message_text(
            "Pick the new billing type.", reply_markup=_lumpsum_kb("editfloorbilling")
        )
        return EDITFLOOR_VALUE
    if field == "rename":
        await query.edit_message_text("Send the new name for this floor.")
        return EDITFLOOR_VALUE
    await query.edit_message_text(_editfloor_summary(key), reply_markup=editfloor_field_kb(key))
    return EDITFLOOR_FIELD


def _editfloor_confirm_kb():
    return InlineKeyboardMarkup(
        [[InlineKeyboardButton("✅ Confirm", callback_data="editfloorconfirm:save"),
          InlineKeyboardButton("✖ Cancel", callback_data="editfloorconfirm:cancel")]]
    )


async def editfloors_value_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    key = context.user_data.get("editfloor_key")
    field = context.user_data.get("editfloor_field")
    if field != "rename":
        return EDITFLOOR_VALUE
    new_label = update.message.text.strip()
    if not new_label:
        await update.message.reply_text("Send a name for the floor.")
        return EDITFLOOR_VALUE
    if new_label.lower() in {l.lower() for k, l in db.get_floors() if k != key}:
        await update.message.reply_text(f"A floor called '{new_label}' already exists. Send a different name, or /cancel.")
        return EDITFLOOR_VALUE
    context.user_data["editfloor_pending"] = ("rename", new_label)
    await update.message.reply_text(f"Rename this floor to '{new_label}'?", reply_markup=_editfloor_confirm_kb())
    return EDITFLOOR_CONFIRM


async def editfloors_value_billing(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    lump_sum = query.data.split(":", 1)[1] == "yes"
    billing = "Lump sum (whole lease up front)" if lump_sum else "Monthly rent (normal)"
    context.user_data["editfloor_pending"] = ("billing", lump_sum)
    await query.edit_message_text(f"Set billing type to '{billing}'?", reply_markup=_editfloor_confirm_kb())
    return EDITFLOOR_CONFIRM


async def editfloors_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    key = context.user_data.pop("editfloor_key", None)
    pending = context.user_data.pop("editfloor_pending", None)
    context.user_data.pop("editfloor_field", None)
    if query.data == "editfloorconfirm:cancel" or not pending or key is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    field, value = pending
    if field == "rename":
        db.update_floor_label(key, value)
        await query.edit_message_text(f"Floor renamed to '{value}'.")
    elif field == "billing":
        db.update_floor_lump_sum(key, value)
        await query.edit_message_text("Billing type updated.")
    elif field == "delete":
        try:
            db.delete_floor(key)
            await query.edit_message_text("Floor deleted.")
        except ValueError:
            await query.edit_message_text("Couldn't delete — shops were added to it in the meantime.")
    else:
        await query.edit_message_text("Nothing changed.")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: View Shop
# ---------------------------------------------------------------------------

@flow_entry("view")
async def view_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb, shops = shop_picker_kb("viewshop")
    if not shops:
        await send(update, "No shops yet.")
        return ConversationHandler.END
    await send(update, "Which shop do you want to view?", reply_markup=kb)
    return VIEW_SELECT


async def view_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    await query.edit_message_text(shop_detail_text(shop_no))
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Link Code
# ---------------------------------------------------------------------------

@flow_entry("linkcode")
async def linkcode_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb, shops = shop_picker_kb("linkshop")
    if not shops:
        await send(update, "No shops yet.")
        return ConversationHandler.END
    await send(update, "Which shop's link code do you need?", reply_markup=kb)
    return LINK_SELECT


async def link_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    shop = db.get_shop(shop_no)
    code = shop["link_code"] or db.regenerate_link_code(shop_no)
    await query.edit_message_text(f"Link code for Shop {shop_no}: {code}\nTenant sends: /register {code}")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Deactivate / Activate
# ---------------------------------------------------------------------------

@flow_entry("deactivate")
async def deactivate_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    kb, shops = shop_picker_kb("deactshop", active_only=True)
    if not shops:
        await send(update, "No active shops.")
        return ConversationHandler.END
    await send(update, "Which shop should be marked inactive?", reply_markup=kb)
    return DEACT_SELECT


async def deact_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["deact_shop_no"] = shop_no
    await query.edit_message_text(
        f"Mark Shop {shop_no} inactive? It will stop receiving monthly rent charges.",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Confirm", callback_data="deactconfirm:save"),
              InlineKeyboardButton("✖ Cancel", callback_data="deactconfirm:cancel")]]
        ),
    )
    return DEACT_CONFIRM


async def deact_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = context.user_data.pop("deact_shop_no", None)
    if query.data == "deactconfirm:cancel" or shop_no is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    db.set_shop_active(shop_no, False)
    await query.edit_message_text(f"Shop {shop_no} marked inactive (no more monthly charges).")
    return ConversationHandler.END


@flow_entry("activate")
async def activate_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    inactive = [s for s in db.get_all_shops(active_only=False) if not s["active"]]
    kb, shops = shop_picker_kb("actshop", shops=inactive)
    if not shops:
        await send(update, "All shops are already active.")
        return ConversationHandler.END
    await send(update, "Which shop should be reactivated?", reply_markup=kb)
    return ACT_SELECT


async def act_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["act_shop_no"] = shop_no
    await query.edit_message_text(
        f"Reactivate Shop {shop_no}? It will resume receiving monthly rent charges.",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Confirm", callback_data="actconfirm:save"),
              InlineKeyboardButton("✖ Cancel", callback_data="actconfirm:cancel")]]
        ),
    )
    return ACT_CONFIRM


async def act_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = context.user_data.pop("act_shop_no", None)
    if query.data == "actconfirm:cancel" or shop_no is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    db.set_shop_active(shop_no, True)
    await query.edit_message_text(f"Shop {shop_no} reactivated.")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Notify one tenant
# ---------------------------------------------------------------------------

@flow_entry("notify")
async def notify_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    linked = [s for s in db.get_all_shops(active_only=True) if s["telegram_id"]]
    kb, shops = shop_picker_kb("notifyshop", shops=linked)
    if not shops:
        await send(update, "No tenants are linked yet.")
        return ConversationHandler.END
    await send(update, "Message which tenant?", reply_markup=kb)
    return NOTIFY_SELECT


async def notify_select(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = query.data.split(":", 1)[1]
    if shop_no == "CANCEL":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    context.user_data["notify_shop_no"] = shop_no
    await query.edit_message_text(f"What message should I send to Shop {shop_no}?")
    return NOTIFY_MSG


async def notify_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    shop_no = context.user_data["notify_shop_no"]
    context.user_data["notify_text"] = update.message.text
    await update.message.reply_text(
        f"Send this to Shop {shop_no}?\n\n{update.message.text}",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Send", callback_data="notifyconfirm:send"),
              InlineKeyboardButton("✖ Cancel", callback_data="notifyconfirm:cancel")]]
        ),
    )
    return NOTIFY_CONFIRM


async def notify_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    shop_no = context.user_data.pop("notify_shop_no", None)
    message = context.user_data.pop("notify_text", None)
    if query.data == "notifyconfirm:cancel" or shop_no is None:
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    shop = db.get_shop(shop_no)
    try:
        await context.bot.send_message(shop["telegram_id"], f"Message from building admin:\n{message}")
        await query.edit_message_text("Sent.")
    except Exception as e:
        logger.warning("Could not notify tenant %s: %s", shop_no, e)
        await query.edit_message_text("Couldn't deliver that message.")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Admin flow: Broadcast to all tenants
# ---------------------------------------------------------------------------

@flow_entry("broadcast")
async def broadcast_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send(update, "What message should go to every linked tenant?", reply_markup=ReplyKeyboardRemove())
    return BROADCAST_MSG


async def broadcast_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["broadcast_msg"] = update.message.text
    await update.message.reply_text(
        f"Send this to all linked tenants?\n\n{update.message.text}",
        reply_markup=InlineKeyboardMarkup(
            [[InlineKeyboardButton("✅ Send", callback_data="bc:send"),
              InlineKeyboardButton("✖ Cancel", callback_data="bc:cancel")]]
        ),
    )
    return BROADCAST_CONFIRM


async def broadcast_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    message = context.user_data.pop("broadcast_msg", "")
    if query.data == "bc:cancel":
        await query.edit_message_text("Cancelled.")
        return ConversationHandler.END
    sent = 0
    for shop in db.get_all_shops(active_only=True):
        if shop["telegram_id"]:
            try:
                await context.bot.send_message(shop["telegram_id"], f"Message from building admin:\n{message}")
                sent += 1
            except Exception as e:
                logger.warning("Broadcast failed for shop %s: %s", shop["shop_no"], e)
    await query.edit_message_text(f"Broadcast sent to {sent} tenant(s).")
    return ConversationHandler.END


# ---------------------------------------------------------------------------
# Interrupting a flow that's already in progress (continued from flow_entry
# above): every trigger that can start one of the guided flows below also
# gets a "guard" version. While a flow is active, the guard intercepts any
# OTHER flow's trigger and asks for confirmation before discarding the
# in-progress one. Answering "No" leaves the original flow's state
# untouched (the guard returns None, which — per ConversationHandler's
# rules — means "don't change the state").
# ---------------------------------------------------------------------------

FLOW_STARTERS = {
    "addshop": addshop_start,
    "addtenant": addtenant_start,
    "vacate": vacate_start,
    "pay": pay_start,
    "expense": expense_start,
    "editrent": editrent_start,
    "editshop": editshop_start,
    "edittenant": edittenant_start,
    "linkcode": linkcode_start,
    "deactivate": deactivate_start,
    "activate": activate_start,
    "notify": notify_start,
    "broadcast": broadcast_start,
    "view": view_start,
    "editpay": editpay_start,
    "editexp": editexp_start,
    "addfloors": addfloors_start,
    "editfloors": editfloors_start,
}

FLOW_TRIGGERS = {
    "addshop": [("command", "addshop"), ("regex", "^➕ Add Shop$")],
    "addtenant": [("command", "addtenant"), ("callback", "^more:addtenant$")],
    "vacate": [("command", "vacate"), ("callback", "^more:vacate$")],
    "pay": [("command", "pay"), ("regex", "^💰 Record Payment$")],
    "expense": [("command", "expense"), ("regex", "^🧾 Add Expense$")],
    "editrent": [("command", "editrent"), ("callback", "^more:rent$")],
    "editshop": [("command", "editshop"), ("callback", "^more:editshop$")],
    "edittenant": [("command", "edittenant"), ("callback", "^more:edittenant$")],
    "linkcode": [("command", "linkcode"), ("callback", "^more:link$")],
    "deactivate": [("command", "deactivate"), ("callback", "^more:deact$")],
    "activate": [("command", "activate"), ("callback", "^more:act$")],
    "notify": [("command", "notify"), ("callback", "^more:notify$")],
    "broadcast": [("command", "broadcast"), ("callback", "^more:broadcast$")],
    "view": [("callback", "^more:view$")],
    "editpay": [("command", "editpayment"), ("callback", "^more:editpay$")],
    "editexp": [("command", "editexpense"), ("callback", "^more:editexp$")],
    "addfloors": [("command", "addfloors"), ("callback", "^more:addfloors$")],
    "editfloors": [("command", "editfloors"), ("callback", "^more:editfloors$")],
}


def make_flow_guard(key):
    """Builds the guard handler for one flow. Fires when THAT flow's own
    trigger arrives while a *different* flow is active."""
    async def _guard(update: Update, context: ContextTypes.DEFAULT_TYPE):
        if not is_admin(update.effective_user.id):
            return None
        current_key = context.user_data.get("flow_key")
        if current_key == key or current_key is None:
            # Re-tapping the flow you're already in (or nothing tracked
            # yet) — nothing to protect, ignore like before.
            return None
        context.user_data["pending_switch"] = {"key": key, "update": update}
        await send(
            update,
            f"You're still {FLOW_LABELS.get(current_key, 'in the middle of something')}. "
            f"Cancel that and start {FLOW_LABELS[key]} instead?",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("✅ Yes, switch", callback_data="switchflow:yes"),
                  InlineKeyboardButton("↩ No, keep going", callback_data="switchflow:no")]]
            ),
        )
        return None  # leave the in-progress conversation state untouched
    return _guard


async def switch_flow_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    pending = context.user_data.pop("pending_switch", None)
    if query.data == "switchflow:no" or not pending:
        await query.edit_message_text("Okay — continuing where you left off.")
        return None  # state stays exactly as it was
    key = pending["key"]
    original_update = pending["update"]
    await query.edit_message_text(f"Switching to {FLOW_LABELS[key]}...")
    starter = FLOW_STARTERS[key]
    return await starter(original_update, context)


def _build_guard_handlers():
    handlers = []
    for key, triggers in FLOW_TRIGGERS.items():
        guard_func = make_flow_guard(key)
        for kind, value in triggers:
            if kind == "command":
                handlers.append(CommandHandler(value, guard_func))
            elif kind == "regex":
                handlers.append(MessageHandler(filters.Regex(value), guard_func))
            elif kind == "callback":
                handlers.append(CallbackQueryHandler(guard_func, pattern=value))
    handlers.append(CallbackQueryHandler(switch_flow_confirm, pattern="^switchflow:"))
    return handlers


# ---------------------------------------------------------------------------
# The one big admin conversation (keeps all flows from colliding with each other)
# ---------------------------------------------------------------------------

admin_conv_entry_points = [
    CommandHandler("addshop", addshop_start),
    MessageHandler(filters.Regex("^➕ Add Shop$"), addshop_start),
    CommandHandler("pay", pay_start),
    MessageHandler(filters.Regex("^💰 Record Payment$"), pay_start),
    CommandHandler("expense", expense_start),
    MessageHandler(filters.Regex("^🧾 Add Expense$"), expense_start),
    CommandHandler("editrent", editrent_start),
    CallbackQueryHandler(editrent_start, pattern="^more:rent$"),
    CommandHandler("editshop", editshop_start),
    CallbackQueryHandler(editshop_start, pattern="^more:editshop$"),
    CommandHandler("edittenant", edittenant_start),
    CallbackQueryHandler(edittenant_start, pattern="^more:edittenant$"),
    CommandHandler("linkcode", linkcode_start),
    CallbackQueryHandler(linkcode_start, pattern="^more:link$"),
    CommandHandler("deactivate", deactivate_start),
    CallbackQueryHandler(deactivate_start, pattern="^more:deact$"),
    CommandHandler("activate", activate_start),
    CallbackQueryHandler(activate_start, pattern="^more:act$"),
    CommandHandler("notify", notify_start),
    CallbackQueryHandler(notify_start, pattern="^more:notify$"),
    CommandHandler("broadcast", broadcast_start),
    CallbackQueryHandler(broadcast_start, pattern="^more:broadcast$"),
    CallbackQueryHandler(view_start, pattern="^more:view$"),
    CommandHandler("addtenant", addtenant_start),
    CallbackQueryHandler(addtenant_start, pattern="^more:addtenant$"),
    CommandHandler("vacate", vacate_start),
    CallbackQueryHandler(vacate_start, pattern="^more:vacate$"),
    CommandHandler("editpayment", editpay_start),
    CallbackQueryHandler(editpay_start, pattern="^more:editpay$"),
    CommandHandler("editexpense", editexp_start),
    CallbackQueryHandler(editexp_start, pattern="^more:editexp$"),
    CommandHandler("addfloors", addfloors_start),
    CallbackQueryHandler(addfloors_start, pattern="^more:addfloors$"),
    CommandHandler("editfloors", editfloors_start),
    CallbackQueryHandler(editfloors_start, pattern="^more:editfloors$"),
]

admin_conv_states = {
    ADDSHOP_FLOOR: [CallbackQueryHandler(addshop_floor, pattern="^floor:")],
    ADDSHOP_NO: [MessageHandler(filters.TEXT & ~filters.COMMAND, addshop_no)],
    ADDSHOP_AREA: [MessageHandler(filters.TEXT & ~filters.COMMAND, addshop_area)],
    ADDSHOP_RENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, addshop_rent)],
    ADDSHOP_CONFIRM: [CallbackQueryHandler(addshop_confirm, pattern="^addshop:")],
    ADDTENANT_SELECT: [CallbackQueryHandler(addtenant_select, pattern="^addtenantshop:")],
    ADDTENANT_FLOOR: [CallbackQueryHandler(addtenant_floor, pattern="^addtenantfloor:")],
    ADDTENANT_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, addtenant_name)],
    ADDTENANT_PURPOSE: [MessageHandler(filters.TEXT & ~filters.COMMAND, addtenant_purpose)],
    ADDTENANT_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, addtenant_phone)],
    ADDTENANT_START: [MessageHandler(filters.TEXT & ~filters.COMMAND, addtenant_start_date)],
    ADDTENANT_LEASE_END: [MessageHandler(filters.TEXT & ~filters.COMMAND, addtenant_lease_end)],
    ADDTENANT_DOC: [
        MessageHandler(filters.PHOTO | filters.Document.ALL, addtenant_doc),
        CallbackQueryHandler(addtenant_doc_skip, pattern="^addtenantdoc:skip$"),
    ],
    ADDTENANT_DOC_MORE: [CallbackQueryHandler(addtenant_doc_more, pattern="^addtenantdoc:(more|done)$")],
    ADDTENANT_CONFIRM: [CallbackQueryHandler(addtenant_confirm, pattern="^addtenant:")],
    VACATE_SELECT: [CallbackQueryHandler(vacate_select, pattern="^vacateshop:")],
    VACATE_CONFIRM: [CallbackQueryHandler(vacate_confirm, pattern="^vacate:")],
    PAY_FLOOR: [CallbackQueryHandler(pay_floor, pattern="^payfloor:")],
    PAY_SELECT: [CallbackQueryHandler(pay_select, pattern="^payshop:")],
    PAY_MONTH: [CallbackQueryHandler(pay_month, pattern="^paymonth:")],
    PAY_RECEIPT: [
        MessageHandler(filters.PHOTO, pay_receipt_photo),
        CallbackQueryHandler(pay_receipt_skip, pattern="^payreceipt:skip$"),
    ],
    PAY_BANK: [CallbackQueryHandler(pay_bank, pattern="^paybank:")],
    PAY_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, pay_amount)],
    PAY_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, pay_date)],
    PAY_REF: [MessageHandler(filters.TEXT & ~filters.COMMAND, pay_ref)],
    PAY_CONFIRM: [CallbackQueryHandler(pay_confirm, pattern="^payconfirm:")],
    EXPENSE_MONTH: [CallbackQueryHandler(expense_month, pattern="^expmonth:")],
    EXPENSE_REASON: [CallbackQueryHandler(expense_reason, pattern="^expreason:")],
    EXPENSE_DESC: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_desc)],
    EXPENSE_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, expense_amount)],
    EXPENSE_RECEIPT: [
        MessageHandler(filters.PHOTO, expense_receipt_photo),
        CallbackQueryHandler(expense_receipt_skip, pattern="^expreceipt:skip$"),
    ],
    EXPENSE_CONFIRM: [CallbackQueryHandler(expense_confirm, pattern="^expconfirm:")],
    RENT_SELECT: [CallbackQueryHandler(rent_select, pattern="^rentshop:")],
    RENT_NEW: [MessageHandler(filters.TEXT & ~filters.COMMAND, rent_new)],
    RENT_EFFECTIVE: [CallbackQueryHandler(rent_effective_select, pattern="^rentmonth:")],
    RENT_CONFIRM: [CallbackQueryHandler(rent_confirm, pattern="^rentconfirm:")],
    EDITSHOP_SELECT: [CallbackQueryHandler(editshop_select, pattern="^editshop:")],
    EDITSHOP_FIELD: [CallbackQueryHandler(editshop_field, pattern="^editshopfield:")],
    EDITSHOP_VALUE: [
        MessageHandler(filters.TEXT & ~filters.COMMAND, editshop_value_text),
        CallbackQueryHandler(editshop_value_floor, pattern="^floor:"),
    ],
    EDITSHOP_CONFIRM: [CallbackQueryHandler(editshop_confirm, pattern="^editshopconfirm:")],
    EDITTENANT_SELECT: [CallbackQueryHandler(edittenant_select, pattern="^edittenantshop:")],
    EDITTENANT_FIELD: [CallbackQueryHandler(edittenant_field, pattern="^edittenantfield:")],
    EDITTENANT_VALUE: [MessageHandler(filters.TEXT & ~filters.COMMAND, edittenant_value)],
    EDITTENANT_CONFIRM: [CallbackQueryHandler(edittenant_confirm, pattern="^edittenantconfirm:")],
    VIEW_SELECT: [CallbackQueryHandler(view_select, pattern="^viewshop:")],
    LINK_SELECT: [CallbackQueryHandler(link_select, pattern="^linkshop:")],
    DEACT_SELECT: [CallbackQueryHandler(deact_select, pattern="^deactshop:")],
    DEACT_CONFIRM: [CallbackQueryHandler(deact_confirm, pattern="^deactconfirm:")],
    ACT_SELECT: [CallbackQueryHandler(act_select, pattern="^actshop:")],
    ACT_CONFIRM: [CallbackQueryHandler(act_confirm, pattern="^actconfirm:")],
    NOTIFY_SELECT: [CallbackQueryHandler(notify_select, pattern="^notifyshop:")],
    NOTIFY_MSG: [MessageHandler(filters.TEXT & ~filters.COMMAND, notify_message)],
    NOTIFY_CONFIRM: [CallbackQueryHandler(notify_confirm, pattern="^notifyconfirm:")],
    BROADCAST_MSG: [MessageHandler(filters.TEXT & ~filters.COMMAND, broadcast_message)],
    BROADCAST_CONFIRM: [CallbackQueryHandler(broadcast_confirm, pattern="^bc:")],
    EDITPAY_SELECT: [CallbackQueryHandler(editpay_select, pattern="^editpayshop:")],
    EDITPAY_PAYMENT: [CallbackQueryHandler(editpay_payment, pattern="^editpaypmt:")],
    EDITPAY_FIELD: [CallbackQueryHandler(editpay_field, pattern="^editpayfield:")],
    EDITPAY_VALUE: [
        MessageHandler(filters.TEXT & ~filters.COMMAND, editpay_value),
        MessageHandler(filters.PHOTO, editpay_receipt_photo),
        CallbackQueryHandler(editpay_bank_value, pattern="^editpaybank:"),
    ],
    EDITPAY_CONFIRM: [CallbackQueryHandler(editpay_confirm, pattern="^editpayconfirm:")],
    EDITEXP_MONTH: [CallbackQueryHandler(editexp_month, pattern="^editexpmonth:")],
    EDITEXP_PICK: [CallbackQueryHandler(editexp_pick, pattern="^editexppick:")],
    EDITEXP_FIELD: [CallbackQueryHandler(editexp_field, pattern="^editexpfield:")],
    EDITEXP_VALUE: [
        MessageHandler(filters.TEXT & ~filters.COMMAND, editexp_value),
        MessageHandler(filters.PHOTO, editexp_receipt_photo),
    ],
    EDITEXP_CONFIRM: [CallbackQueryHandler(editexp_confirm, pattern="^editexpconfirm:")],
    ADDFLOOR_LABEL: [MessageHandler(filters.TEXT & ~filters.COMMAND, addfloors_label)],
    ADDFLOOR_LUMPSUM: [CallbackQueryHandler(addfloors_lumpsum, pattern="^addfloorlump:")],
    ADDFLOOR_CONFIRM: [CallbackQueryHandler(addfloors_confirm, pattern="^addfloor:")],
    EDITFLOOR_SELECT: [CallbackQueryHandler(editfloors_select, pattern="^editfloor:")],
    EDITFLOOR_FIELD: [CallbackQueryHandler(editfloors_field, pattern="^editfloorfield:")],
    EDITFLOOR_VALUE: [
        MessageHandler(filters.TEXT & ~filters.COMMAND, editfloors_value_text),
        CallbackQueryHandler(editfloors_value_billing, pattern="^editfloorbilling:"),
    ],
    EDITFLOOR_CONFIRM: [CallbackQueryHandler(editfloors_confirm, pattern="^editfloorconfirm:")],
}

# Give every state a first look at each OTHER flow's trigger (and at the
# switch-flow confirm buttons) before its own normal handler(s) run. A
# guard only actually does anything if the trigger it's watching for
# doesn't belong to the flow currently in progress — see make_flow_guard.
_guard_handlers = _build_guard_handlers()
for _state_handlers in admin_conv_states.values():
    _state_handlers[0:0] = _guard_handlers

admin_conv = ConversationHandler(
    entry_points=admin_conv_entry_points,
    states=admin_conv_states,
    fallbacks=[CommandHandler("cancel", cancel)],
)


# ---------------------------------------------------------------------------
# Scheduled jobs
# ---------------------------------------------------------------------------

async def monthly_charge_job(context: ContextTypes.DEFAULT_TYPE):
    """Runs daily; only actually does something on the 1st of the Ethiopian
    month (rent periods are tracked in E.C., so the charge date follows suit —
    this is NOT the same day every Gregorian month)."""
    today = date.today()
    ey, em, ed = gregorian_to_ethiopian(today.year, today.month, today.day)
    if ed != 1:
        return
    applied = db.apply_monthly_charges_to_all()
    applied_expenses = db.apply_recurring_expenses_to_all(f"{ey:04d}-{em:02d}")
    logger.info(
        "Monthly charges applied to %d shop(s); permanent expenses applied: %s",
        len(applied), applied_expenses,
    )
    for shop in db.get_all_shops(active_only=True):
        if shop["shop_no"] in applied and shop["telegram_id"]:
            balance = db.get_balance(shop["shop_no"])
            try:
                await context.bot.send_message(
                    shop["telegram_id"],
                    f"This month's rent ({money(shop['monthly_rent'])}) has been charged. "
                    f"Your current balance is {money(balance)}.",
                )
            except Exception as e:
                logger.warning("Could not notify tenant %s: %s", shop["shop_no"], e)
    for admin_id in ADMIN_IDS:
        try:
            text = f"Monthly rent charged to {len(applied)} shop(s)."
            if applied_expenses:
                text += f"\nPermanent expenses applied: {', '.join(applied_expenses)}."
            await context.bot.send_message(admin_id, text)
        except Exception as e:
            logger.warning("Could not notify admin %s: %s", admin_id, e)


async def due_reminder_job(context: ContextTypes.DEFAULT_TYPE):
    """Weekly reminder to tenants who still owe money, plus a digest to admins."""
    shops = [s for s in db.get_all_balances(active_only=True) if s["balance"] > 0]
    for s in shops:
        if s["telegram_id"]:
            try:
                await context.bot.send_message(
                    s["telegram_id"],
                    f"Reminder: your shop ({s['shop_no']}) has an outstanding balance of {money(s['balance'])}.",
                )
            except Exception as e:
                logger.warning("Could not remind tenant %s: %s", s["shop_no"], e)
    if shops:
        lines = ["Weekly dues digest:"]
        for s in shops:
            lines.append(f"Shop {s['shop_no']} ({s['tenant_name']}): {money(s['balance'])}")
        for admin_id in ADMIN_IDS:
            try:
                await context.bot.send_message(admin_id, "\n".join(lines))
            except Exception as e:
                logger.warning("Could not send digest to admin %s: %s", admin_id, e)


# ---------------------------------------------------------------------------
# Native Telegram "/" command menu
# ---------------------------------------------------------------------------

ADMIN_COMMANDS = [
    BotCommand("start", "Show your menu"),
    BotCommand("menu", "Show the menu buttons"),
    BotCommand("addshop", "Add a new vacant shop"),
    BotCommand("addtenant", "Rent out a vacant shop"),
    BotCommand("vacate", "Mark a shop vacant"),
    BotCommand("pay", "Record a rent payment"),
    BotCommand("shops", "List all shops & balances"),
    BotCommand("dues", "Shops with outstanding balance"),
    BotCommand("latepayments", "Shops with unpaid months (🟢🟡🔴 by how overdue)"),
    BotCommand("expense", "Log a building expense"),
    BotCommand("expenses", "This month's expenses"),
    BotCommand("editexpense", "Edit a logged expense"),
    BotCommand("report", "Monthly or yearly collections vs expenses"),
    BotCommand("permanentexpenses", "Manage recurring expenses (electricity/water/salary)"),
    BotCommand("exportexcel", "Download the full-year rent table as Excel"),
    BotCommand("editrent", "Change a shop's rent"),
    BotCommand("editshop", "Edit a shop's number, floor, area or rent"),
    BotCommand("edittenant", "Edit a tenant's details"),
    BotCommand("linkcode", "Get a shop's tenant link code"),
    BotCommand("deactivate", "Mark a shop inactive"),
    BotCommand("activate", "Reactivate a shop"),
    BotCommand("chargenow", "Apply this month's rent now"),
    BotCommand("notify", "Message one tenant"),
    BotCommand("broadcast", "Message all tenants"),
    BotCommand("cancel", "Cancel the current action"),
    BotCommand("help", "Show help"),
]

TENANT_COMMANDS = [
    BotCommand("start", "Start"),
    BotCommand("register", "Link your account to your shop"),
    BotCommand("mybalance", "Your current balance"),
    BotCommand("myledger", "Your recent charges & payments"),
    BotCommand("myshop", "Your shop details"),
    BotCommand("myprofile", "View your profile"),
    BotCommand("registercode", "View your shop's link code again"),
    BotCommand("cancel", "Cancel the current action"),
    BotCommand("help", "Show help"),
]


async def _set_commands_with_retry(app: Application, commands, scope=None, attempts=3):
    """set_my_commands, retrying a couple of times on a slow/flaky connection
    instead of letting one timeout take the whole bot down before polling
    even starts (this used to be able to crash the process on startup —
    see the timeout traceback this was added for)."""
    for attempt in range(1, attempts + 1):
        try:
            kwargs = {"scope": scope} if scope is not None else {}
            await app.bot.set_my_commands(commands, **kwargs)
            return
        except (TimedOut, NetworkError) as e:
            logger.warning(
                "set_my_commands attempt %d/%d timed out: %s", attempt, attempts, e
            )
            if attempt < attempts:
                await asyncio.sleep(2)
    logger.warning(
        "Could not set command menu after %d attempts — bot will still start and "
        "work normally, the /-command menu just won't be populated until the next restart.",
        attempts,
    )


async def post_init(app: Application):
    await _set_commands_with_retry(app, TENANT_COMMANDS)  # default, for anyone not in the admin list
    for admin_id in ADMIN_IDS:
        await _set_commands_with_retry(app, ADMIN_COMMANDS, scope=BotCommandScopeChat(chat_id=admin_id))


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not BOT_TOKEN:
        raise SystemExit("Set TELEGRAM_BOT_TOKEN in your .env file.")
    if not ADMIN_IDS:
        logger.warning("No TELEGRAM_ADMIN_IDS configured — no one will have admin access.")

    db.init_db()

    app = (
        Application.builder()
        .bot(AutoCleanBot(
            token=BOT_TOKEN,
            request=HTTPXRequest(
                connect_timeout=20.0, read_timeout=20.0, write_timeout=20.0, pool_timeout=20.0,
            ),
        ))
        .post_init(post_init)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("menu", menu_cmd))
    app.add_handler(CommandHandler("help", help_cmd))
    app.add_handler(MessageHandler(filters.Regex("^❓ Help$"), help_cmd))

    app.add_handler(register_conv)
    app.add_handler(admin_conv)
    app.add_handler(profile_conv)
    app.add_handler(permexp_conv)

    # admin view-only commands / buttons (no guided input needed)
    app.add_handler(CommandHandler("shops", list_shops))
    app.add_handler(MessageHandler(filters.Regex("^🏬 Shops & Balances$"), list_shops))
    app.add_handler(CallbackQueryHandler(sb_floor_select, pattern="^sbfloor:"))
    app.add_handler(CallbackQueryHandler(list_shop_select, pattern="^listshop:"))
    app.add_handler(CallbackQueryHandler(view_document, pattern="^viewdoc:"))
    app.add_handler(CommandHandler("dues", dues))
    app.add_handler(MessageHandler(filters.Regex("^📉 Outstanding Dues$"), dues))
    app.add_handler(CommandHandler("latepayments", late_payments))
    app.add_handler(MessageHandler(filters.Regex("^⏰ Late Payments$"), late_payments))
    app.add_handler(CommandHandler("expenses", list_expenses))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("exportexcel", export_excel))
    app.add_handler(MessageHandler(filters.Regex("^📊 Monthly Report$"), report))
    app.add_handler(CallbackQueryHandler(report_period_select, pattern="^reportperiod:"))
    app.add_handler(CallbackQueryHandler(report_month_select, pattern="^reportmonth:"))
    app.add_handler(CallbackQueryHandler(report_year_select, pattern="^reportyear:"))
    app.add_handler(CallbackQueryHandler(report_export, pattern="^reportexport:"))
    app.add_handler(CommandHandler("chargenow", charge_now))
    app.add_handler(CallbackQueryHandler(charge_now_confirm, pattern="^chargenow:"))
    app.add_handler(CommandHandler("more", more_menu))
    app.add_handler(MessageHandler(filters.Regex("^⚙️ More$"), more_menu))
    app.add_handler(CallbackQueryHandler(more_router, pattern="^more:(exp|charge|close)$"))

    # tenant view-only commands / buttons
    app.add_handler(CommandHandler("mybalance", my_balance))
    app.add_handler(MessageHandler(filters.Regex("^💰 My Balance$"), my_balance))
    app.add_handler(CommandHandler("myledger", my_ledger))
    app.add_handler(MessageHandler(filters.Regex("^📜 My Ledger$"), my_ledger))
    app.add_handler(CommandHandler("myshop", my_shop))
    app.add_handler(CommandHandler("registercode", my_registercode))
    app.add_handler(MessageHandler(filters.Regex("^🏬 My Shop$"), my_shop))
    app.add_handler(CommandHandler("myprofile", my_profile))
    app.add_handler(MessageHandler(filters.Regex("^🪪 My Profile$"), my_profile))
    app.add_handler(CallbackQueryHandler(profile_close, pattern="^profile:close$"))

    # "One screen at a time" cleanup — see the AutoCleanBot/clean_incoming
    # comment above. group=100 makes sure this runs after every handler
    # registered above has already had its turn on this same update.
    app.add_handler(
        MessageHandler(filters.ChatType.PRIVATE & filters.ALL, clean_incoming), group=100
    )

    # scheduled jobs: daily check for the 1st of month, weekly dues reminder (Monday)
    app.job_queue.run_daily(monthly_charge_job, time=datetime.strptime("08:00", "%H:%M").time())
    app.job_queue.run_daily(
        due_reminder_job, time=datetime.strptime("09:00", "%H:%M").time(), days=(0,)
    )

    logger.info("Bot starting...")
    app.run_polling()


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
