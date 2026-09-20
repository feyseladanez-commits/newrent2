"""
Builds the classic spreadsheet-style rent report as an .xlsx workbook,
matching the original layout:

SHOP NO | phone num | name | contrat start | RENT | <12 Ethiopian months>

Used by both webapp.py (download from the browser) and bot.py (sent as a
Telegram document).
"""
from io import BytesIO
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

import database as db

COLUMNS = [
    "SHOP NO", "phone num", "name", "contrat start", "lease end", "RENT",
] + [db.ETHIOPIAN_MONTHS[m] for m in range(1, 13)]


def build_workbook(ec_year):
    rows = db.export_rows_for_year(ec_year)

    wb = Workbook()
    ws = wb.active
    ws.title = f"{ec_year} E.C."

    ws.append([f"CENTRAL BUILDING RENT FOR {ec_year} E.C."])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
    ws.cell(row=1, column=1).font = Font(bold=True, size=13)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center")

    ws.append(COLUMNS)
    for col in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=2, column=col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for r in rows:
        ws.append([
            r["shop_no"], r["phone"], r["tenant_name"], r["start_date_ec"],
            r["lease_end_date_ec"], r["monthly_rent"],
        ] + r["months"])

    widths = [10, 14, 20, 14, 14, 10] + [10] * 12
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=2, column=i).column_letter].width = w
    ws.freeze_panes = "A3"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Financial report export (collections vs expenses, by floor, by month) —
# separate from build_workbook() above, which is the full-year per-shop
# ledger table. This is the /report screen (bot.py and webapp.py) turned
# into a downloadable workbook.
# ---------------------------------------------------------------------------

_BOLD = Font(bold=True)
_TITLE = Font(bold=True, size=13)
_CENTER = Alignment(horizontal="center")
_HEADER_FILL = None  # kept simple / dependency-free — bold + borders only


def _header_row(ws, row, values, widths=None):
    ws.append(values)
    for col in range(1, len(values) + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = _BOLD
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(row=row, column=i).column_letter].width = w


def build_report_workbook(kind, key, exclude_floors=None, exclude_shop_ids=None):
    """kind: 'month' -> key is an Ethiopian 'YYYY-MM' period string.
             'year'  -> key is an Ethiopian year (int).
    exclude_floors: same meaning as everywhere else (e.g. db.LUMP_SUM_FLOORS).
    exclude_shop_ids: individual shops (by shops.id) to leave out of the report."""
    if kind == "year":
        ec_year = int(key)
        report = db.yearly_report(ec_year, exclude_floors=exclude_floors, exclude_shop_ids=exclude_shop_ids)
        label = f"{ec_year} E.C."
        start_iso, end_iso = db.ethiopian_year_bounds(ec_year)
        months = report["months"]
    else:
        ey, em = (int(x) for x in str(key).split("-"))
        next_ey, next_em = db.add_ethiopian_months(ey, em, 1)
        start_gy, start_gm, start_gd = db.ethiopian_to_gregorian(ey, em, 1)
        end_gy, end_gm, end_gd = db.ethiopian_to_gregorian(next_ey, next_em, 1)
        start_iso = date(start_gy, start_gm, start_gd).isoformat()
        end_iso = date(end_gy, end_gm, end_gd).isoformat()
        label = f"{db.ETHIOPIAN_MONTHS[em]} {ey}"
        report = db.monthly_report_range(
            start_iso, end_iso, label, exclude_floors=exclude_floors, exclude_shop_ids=exclude_shop_ids,
        )
        months = [report]

    scope_label = "Without Special & Bank Shops" if exclude_floors else "All Floors (incl. Special & Bank Shops)"
    if exclude_shop_ids:
        scope_label += f" — {len(exclude_shop_ids)} shop(s) left out"

    wb = Workbook()

    # --- Summary sheet ---
    ws = wb.active
    ws.title = "Summary"
    ws.append([f"RENT REPORT — {label}"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    ws.cell(row=1, column=1).font = _TITLE
    ws.append([f"Scope: {scope_label}"])
    ws.append([])
    summary_rows = [
        ("Rent expected", report.get("rent_expected")),
        ("Rent charged", report.get("rent_charged")),
        ("Rent collected", report["rent_collected"]),
        ("Expenses", report["expenses"]),
        ("Net (collected - expenses)", report["net"]),
    ]
    if kind == "year":
        summary_rows.append(("Total outstanding (all time)", report.get("total_outstanding_all_shops")))
    for name, value in summary_rows:
        if value is None:
            continue
        ws.append([name, value])
        ws.cell(row=ws.max_row, column=1).font = _BOLD
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 18

    # --- By Month sheet (yearly reports only show more than one row) ---
    ws2 = wb.create_sheet("By Month")
    _header_row(ws2, 1, ["Month", "Rent Collected", "Expenses", "Net"], widths=[16, 16, 16, 16])
    for m in months:
        ws2.append([m["period"], m["rent_collected"], m["expenses"], m["net"]])

    # --- By Floor sheet ---
    # Paid/Not Paid columns only make sense against a single period, so
    # they're only included for a monthly export (kind == "month"), where
    # `key` is already the Ethiopian 'YYYY-MM' period string.
    by_floor_period = key if kind == "month" else None
    by_floor = db.report_by_floor(
        start_iso, end_iso, period=by_floor_period, exclude_shop_ids=exclude_shop_ids,
    )
    ws3 = wb.create_sheet("By Floor")
    if by_floor_period:
        _header_row(
            ws3, 1, ["Floor", "Shops", "Rent Expected / mo", "Rent Collected", "Paid", "Not Paid"],
            widths=[20, 10, 20, 20, 10, 10],
        )
        for f in by_floor:
            ws3.append([
                f["label"], f["shop_count"], f["rent_expected"], f["rent_collected"],
                f.get("paid_count", 0), f.get("unpaid_count", 0),
            ])
    else:
        _header_row(
            ws3, 1, ["Floor", "Shops", "Rent Expected / mo", "Rent Collected"],
            widths=[20, 10, 20, 20],
        )
        for f in by_floor:
            ws3.append([f["label"], f["shop_count"], f["rent_expected"], f["rent_collected"]])

    # --- Expenses by category sheet ---
    ws4 = wb.create_sheet("Expenses")
    _header_row(ws4, 1, ["Category", "Total"], widths=[24, 16])
    for e in db.expenses_by_category(start_iso, end_iso):
        ws4.append([e["category"], e["total"]])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
