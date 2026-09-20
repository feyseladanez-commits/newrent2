# Shop Rent Manager

Rent management for a 30-shop building, with two front doors that share the
same database (`rent.db`):

- **`bot.py`** — the original Telegram bot.
- **`webapp.py`** — a browser-based web app with full username/password
  logins for both roles. This is the newer, recommended way to use the
  system day-to-day; the bot is still useful for its scheduled Telegram
  notifications (monthly charge alerts, weekly dues reminders), so it's
  fine to run both at once.

Two roles either way:

- **Admin (you)** — add shops, record payments, track expenses, run reports.
- **Tenants** — log in and check their own balance and ledger.

## Web app setup (Windows, one-click)

1. **Install Python** if you don't have it: https://www.python.org/downloads/ — during install,
   check the box **"Add python.exe to PATH"**.
2. **Double-click `webapp.py`** (or **`Start Rent Web App.bat`** — either works).
   - First run: installs whatever's needed, creates a `.env` with a random
     session-signing key, then starts the server and opens your browser to
     `http://localhost:5000`.
   - You'll land on a one-time setup screen to create your own admin
     username and password.
3. Leave that window open while you want the web app running. Closing it stops it.
4. From the **Users** page (as admin) you can create more admin accounts.
   To create a tenant's login, open that shop's page → **Tenant Login** tab
   → set a username and password for them, and give it to them however you like.
5. From the **Floors** page (as admin) you can add new floors, rename or reorder
   existing ones, switch a floor between monthly rent and lump-sum billing, or
   delete a floor once it has no shops on it. Same underlying floors as the
   Telegram bot's `/addfloors` and `/editfloors` — either front door works.

The web app listens on all network interfaces (`0.0.0.0:5000` by default), so
if this PC is on your building's Wi-Fi, tenants can also reach it at
`http://<this-PC's-LAN-IP>:5000` from their own phones — no need to install
anything. Set `WEB_HOST` / `WEB_PORT` in `.env` to change the address or port.
Note this is a plain HTTP server meant for a trusted local/LAN network, not
the open internet — if you want it reachable from outside your building,
put it behind a reverse proxy (e.g. Caddy or nginx) with HTTPS, or deploy it
to a proper host.

Uploaded receipts and documents are stored under an `uploads/` folder next
to the script — back that up along with `rent.db`.

<details>
<summary>Manual setup (any OS, command line)</summary>

```bash
pip install -r requirements.txt
python webapp.py
```
</details>

## Telegram bot setup

## Setup (Windows, one-click)

1. **Install Python** if you don't have it: https://www.python.org/downloads/ — during install,
   check the box **"Add python.exe to PATH"** (the installer also associates `.py` files with
   Python by default, which is what makes double-clicking `bot.py` work).
2. **Create the bot**: message [@BotFather](https://t.me/BotFather) on Telegram, run `/newbot`,
   and copy the token it gives you.
3. **Find your admin ID**: message [@userinfobot](https://t.me/userinfobot) to get your numeric
   Telegram user ID.
4. **Double-click `bot.py`** (or `Start Rent Bot.bat` — either works, see below).
   - First time: it creates a `.env` file and opens it in Notepad. Paste in your bot token and
     admin ID, save, close Notepad. Then double-click again.
   - Second time: it installs whatever it needs and starts the bot.
5. Leave that window open while you want the bot running. Closing it stops the bot.

`bot.py` vs `Start Rent Bot.bat`: both now do the same first-run setup (create `.env`,
install dependencies) and both keep the window open if something goes wrong instead of it
flashing shut. `bot.py` installs packages into whatever Python is on your PATH; the `.bat`
file additionally creates an isolated `venv` folder for them, which is tidier if you have
other Python projects on the same machine. Either is fine for normal use.

A `rent.db` SQLite file is created automatically next to the script on first run — that's your
whole database, so back it up occasionally by copying that one file.

<details>
<summary>Manual setup (any OS, command line)</summary>

```bash
pip install -r requirements.txt
cp .env.example .env   # then edit .env with your token and admin ID
python bot.py
```

For it to run continuously on a server instead of your PC, use `screen`, `tmux`, `systemd`, or
a small VPS so it survives logging out.
</details>

## How it works day-to-day

Everything is menu-driven — you don't type long commands with arguments anymore.

- **Admin**: send `/start` and you'll get a button menu at the bottom of the chat
  (Add Shop, Record Payment, Shops & Balances, Dues, Add Expense, Report, More).
  Tap a button and the bot asks you one question at a time (shop number, then name,
  then rent, etc.) — answer in plain text or by tapping the inline buttons it shows.
  Anywhere mid-flow, send `/cancel` to back out.
- **Every command that changes something** ends with a review screen — Confirm/Save
  vs. Cancel — before anything is written or sent. Nothing is saved, applied, or
  delivered on the strength of your last answer alone; you always get one more tap
  to back out first.
- **Starting a second guided flow while one is already open** — e.g. tapping
  Record Payment while you're partway through Add Shop — doesn't silently drop
  what you were doing. The bot asks "You're still adding a shop. Cancel that and
  start recording a payment instead?" Tap **No, keep going** to pick up exactly
  where you left off, or **Yes, switch** to abandon the first one and begin the
  new one.
- **Tenant**: send `/start`, tap **Register**, and enter the code the admin gave you.
  After that you get **My Balance**, **My Ledger**, **My Shop** buttons.
- The native **/** command menu in Telegram (tap the icon next to the text box) also
  lists every command, tailored to whether you're the admin or a tenant.

## How rent tracking works

- Each shop has a `monthly_rent`. On the 1st of every month, the bot automatically adds one
  rent charge per active shop (you can also trigger this manually with `/chargenow`).
- Payments recorded with `/pay` are subtracted from the shop's running balance.
- A shop's **balance** = total charged − total paid. Positive means they owe money.

## Onboarding a tenant

1. Admin taps **Add Shop** (or sends `/addshop`) and answers the shop number, tenant
   name, phone, rent, and start date as the bot asks — the bot then shows a **link
   code**.
2. Give that code to the tenant (in person, SMS, WhatsApp — whatever you use).
3. The tenant opens the bot, taps **Register** (or sends `/register <code>`), and enters
   the code.
4. From then on, their Telegram account is tied to their shop and they get the tenant
   menu: **My Balance**, **My Ledger**, **My Shop**.

If a tenant loses their code, open **More → Link Code** and pick their shop to get it again.

## Admin commands (each opens a guided flow if it needs details)

| Command | Purpose |
|---|---|
| `/addshop` | Register a new shop |
| `/shops` | List all shops with balances |
| `/pay` | Record a rent payment |
| `/dues` | List shops with an outstanding balance |
| `/editrent` | Change a shop's monthly rent |
| `/editshop` | Edit a shop's number, floor, area, or rent |
| `/edittenant` | Edit a tenant's name, phone, purpose, dates, ID, or TIN |
| `/deactivate` / `/activate` | Stop/resume monthly charges for a shop |
| `/expense` | Log a building expense (repairs, utilities, etc.) |
| `/expenses` | This month's expenses |
| `/report` | Rent collected vs. expenses, net, total outstanding |
| `/chargenow` | Manually apply this month's rent to all shops |
| `/notify` | Message one tenant directly |
| `/broadcast` | Message every linked tenant |
| `/addfloors` | Add a new floor (name + monthly vs. lump-sum billing) |
| `/editfloors` | Rename a floor, change its billing type, reorder it, or delete it (if empty) |
| `/more` | Open the extra-actions menu (view shop, edit shop, edit tenant, link code, deactivate/activate, notify, broadcast, this month's expenses, charge now, add/edit floors) |
| `/cancel` | Stop whatever guided flow you're in the middle of |

## Tenant commands

| Command | Purpose |
|---|---|
| `/register <code>` | Link your Telegram account to your shop |
| `/mybalance` | Your current amount due |
| `/myledger` | Your recent charges and payments |
| `/myshop` | Your shop's details |

## Recording a payment

Tap **Record Payment**, pick the shop, then pick which month it's for (shown in
Amharic). Send a photo of the receipt — the bot reads the bank, amount, date and
reference number off it itself — and only asks you for whatever it couldn't
confidently read. You get one final screen to review and tap **Approve & Save**
before anything is stored; tap Cancel there to back out. (This is the same
confirm-before-you-commit pattern every other command uses now — see the note
above.)

Receipt reading needs the `tesseract-ocr` program installed on the machine
running the bot (separate from the `pytesseract`/`Pillow` Python packages,
which `requirements.txt` already installs):

- **Windows**: install from https://github.com/UB-Mannheim/tesseract/wiki and
  make sure it's on your PATH.
- **Linux**: `sudo apt install tesseract-ocr`

If it isn't installed, or a receipt can't be read, the bot just falls back to
asking for the bank, amount, date and reference number one at a time — nothing
breaks.

## Notes & things to adapt

- **This bot tracks money, it doesn't move it.** `/pay` just records that a payment happened
  (e.g. after you receive cash or a bank transfer) — it isn't a payment gateway. If you want
  tenants to pay *through* the bot, you'd add Telegram Payments or a local mobile-money API,
  which is a bigger integration.
- Scheduled jobs (monthly charge + weekly dues reminder) run in whatever timezone the server's
  clock uses — adjust `time=` in `bot.py`'s `main()` if you want a different local time.
- The database is a single SQLite file (`rent.db`). Back it up periodically (just copy the
  file) since there's no built-in backup.
