"""
ONE-TIME data import for the live deployment.

Why this exists: the deployed app starts with a fresh, empty rent.db
(database.py's init_db() only creates empty tables via CREATE TABLE IF
NOT EXISTS). This script loads the real data from rent_data_dump.sql
(exported from the developer's local rent.db) into that empty database,
exactly once.

Safety:
  - Runs only if rent.db has zero rows in the 'shops' table. Once data
    exists, it does nothing — so restarts, redeploys, and multiple
    container instances never re-import or duplicate rows.
  - Only ever INSERTs the dumped rows; never drops or alters tables.

After you've confirmed the live app shows your real data (shops, tenants,
payment history), you can delete this file, rent_data_dump.sql, and the
one line that calls this from server.py — it has no further purpose once
the live database is populated.
"""
import os
import sqlite3

import database as db

DUMP_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rent_data_dump.sql")


def import_once():
    if not os.path.exists(DUMP_PATH):
        return  # dump file removed after first successful import — nothing to do

    conn = sqlite3.connect(db.DB_PATH)
    cur = conn.cursor()

    # Guard: only import into a genuinely empty database.
    cur.execute("SELECT COUNT(*) FROM shops")
    if cur.fetchone()[0] > 0:
        conn.close()
        print("[import_initial_data] shops table already has data — skipping import.")
        return

    with open(DUMP_PATH, "r", encoding="utf-8") as f:
        sql_script = f.read()

    try:
        cur.executescript(sql_script)
        conn.commit()
        print("[import_initial_data] Imported initial data successfully.")
    except Exception as e:
        conn.rollback()
        print(f"[import_initial_data] Import FAILED, rolled back: {e}")
    finally:
        conn.close()


if __name__ == "__main__":
    db.init_db()
    import_once()
