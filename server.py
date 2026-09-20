"""
Combined entrypoint for Shop Rent Manager: runs the Flask web app (via
waitress) in a background thread and the Telegram bot's polling loop in
the main thread — one process, one rent.db, so both stay in sync.

Why this exists: bot.py and webapp.py both import database.py, which
resolves rent.db relative to its own file location. That's correct, but
if bot.py and webapp.py are deployed as two SEPARATE processes/containers
(e.g. two separate AletCloud apps), each gets its own copy of the repo and
therefore its own unshared rent.db — so data entered through one never
shows up in the other. Running both inside this single process (and
therefore inside a single deployment) fixes that: there's only ever one
database.py, so DB_PATH, so rent.db.

bot.py and webapp.py are unchanged and still work standalone exactly as
the README describes (for local Windows one-click use). This file is only
for a single-process deployment (e.g. one AletCloud app instead of two).

Run:
    python server.py

Deploy on AletCloud as ONE app with this as the start command, instead of
two separate apps for bot.py and webapp.py.
"""
import logging
import os
import threading

import database as db
import bot
import webapp

logger = logging.getLogger(__name__)


def _run_web():
    host = os.getenv("WEB_HOST", "0.0.0.0")
    port = int(os.getenv("PORT") or os.getenv("WEB_PORT", "5000"))
    try:
        from waitress import serve
        logger.info("Web app starting on %s:%s (waitress)", host, port)
        serve(webapp.app, host=host, port=port)
    except ImportError:
        logger.warning("waitress not installed — falling back to Flask's dev server")
        webapp.app.run(host=host, port=port, debug=False)


def main():
    if not bot.BOT_TOKEN:
        raise SystemExit("Set TELEGRAM_BOT_TOKEN in your .env file.")

    # Shared init. bot.main() would also call this, but init_db() is
    # idempotent (CREATE TABLE IF NOT EXISTS), so doing it once up front
    # covers both.
    db.init_db()

    # ONE-TIME: import real data from rent_data_dump.sql on first deploy.
    # Safe to leave in permanently — it no-ops once the shops table has
    # any rows, and no-ops entirely once rent_data_dump.sql is deleted.
    # See import_initial_data.py's docstring for details.
    import import_initial_data
    import_initial_data.import_once()

    web_thread = threading.Thread(target=_run_web, daemon=True, name="webapp")
    web_thread.start()

    # bot.main() builds the Application, registers every handler and the
    # two scheduled jobs, and finishes with app.run_polling() — which
    # blocks and owns its own asyncio event loop, so it has to be the one
    # left running in the main thread. The waitress server above is
    # synchronous/thread-based, so the two don't conflict.
    logger.info("Bot starting (combined process)...")
    bot.main()


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (0, None):
            print(f"\n{e}" if str(e) else "")
        raise
    except KeyboardInterrupt:
        pass
