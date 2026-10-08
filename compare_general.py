#!/usr/bin/env python3
"""
=====================================================================
  ALERT TABLE COMPARISON TOOL
  Compares dbo.alerts_result against dbo.alerts_result_test, row by
  row and cell by cell, for every alert ID you choose, and writes an
  Excel report showing exactly which rows and columns differ.
=====================================================================

HOW TO USE  (follow the steps in order)
---------------------------------------
  STEP 0  One-time setup on your computer
          Install the Python packages (run this in a terminal / command prompt):
                 pip install pymssql pandas openpyxl
          That's all. No Microsoft ODBC driver is needed: pymssql brings
          its own SQL Server connector.

  STEPS 1-6  Scroll down to the section called "EDIT THESE SETTINGS".
             Every setting has a STEP number and an explanation.
             Only change the value to the RIGHT of the "=" sign.
             TIP: In DBeaver, right-click your connection > Edit Connection.
             The "Main" tab shows the Host, Port, Database and Username
             to copy into Steps 1 and 2.

  STEP 7  Run the script:
                 python compare_alert_tables.py
          It prints one line per alert ID while it works, and at the end
          tells you where the Excel report was saved.

WHAT THE EXCEL REPORT CONTAINS
------------------------------
  Summary             One line per alert ID: row counts in each table, how many
                      cells differ, and a MATCH / MISMATCH / NO DATA / ERROR status.
  Cell Differences    Every single cell that differs: alert ID, the row's key
                      (clientid, facilityid), the row number in each table,
                      the column name, and both values side by side.
  Missing Rows        Rows that exist in one table but not in the other.
  Column Differences  Columns that exist in only one of the two tables.
  Run Info            The settings used for this run (the password is never saved).

HOW ROWS ARE LINED UP
---------------------
  Rows are paired using the KEY_COLUMNS (clientid + facilityid by default),
  not just by position. That way, one extra or missing row does not make
  every row after it look different. "Row #" in the report is the row's
  position in that table's query result (sorted by your ORDER BY), so you
  can find it again by running your original SQL.
"""

# =====================================================================
#                 EDIT THESE SETTINGS  (Steps 1 - 6)
# =====================================================================

# ---------------------------------------------------------------------
# STEP 1: Where is the SQL Server?  (copy these from DBeaver's "Main" tab)
#   SERVER   = the Host: the host name or IP address of the SQL Server.
#              For a named instance, write it as  "myserver\\INSTANCENAME"
#   PORT     = the Port shown in DBeaver (SQL Server's default is 1433).
#   DATABASE = the database that contains dbo.alerts_result
# ---------------------------------------------------------------------
SERVER = "172.31.28.49"
PORT = 1433
DATABASE = "master"

# ---------------------------------------------------------------------
# STEP 2: Your SQL Server login.
#   Leave PASSWORD as "" (empty) and the script will ask you to type it
#   when it runs. That is safer than saving your password in this file.
# ---------------------------------------------------------------------
USERNAME = "akif_yeahia"
PASSWORD = "V7#qL9!mX2@pR8$k"

# ---------------------------------------------------------------------
# STEP 3: The values from your SQL script.
#   PROCESS_DATE can be written as "2026-10-07" or "20261007".
# ---------------------------------------------------------------------
VENDOR_ID = 4708
PROCESS_DATE = "2026-10-07"

# ---------------------------------------------------------------------
# STEP 4: Which alert IDs to compare.
#   Every alert ID from ALERT_ID_START to ALERT_ID_END (both included)
#   is compared, except the ones listed in EXCLUDED_ALERT_IDS.
# ---------------------------------------------------------------------
ALERT_ID_START = 1
ALERT_ID_END = 100
EXCLUDED_ALERT_IDS = [1, 9, 10, 12, 13, 18, 25, 26, 27, 30, 37, 42, 43, 46, 61, 65]

# ---------------------------------------------------------------------
# STEP 5: Tables and columns.
#   MAIN_TABLE / TEST_TABLE = the two tables being compared.
#   KEY_COLUMNS    = the columns that identify "the same row" in both
#                    tables. This matches the ORDER BY in your script.
#   IGNORE_COLUMNS = columns to leave out of the comparison, e.g. an
#                    identity/ID column or a "created on" timestamp that
#                    is always different between the tables.
#                    Example: IGNORE_COLUMNS = ["id", "insertdate"]
#                    Leave as [] to compare every column.
#   Column names are not case-sensitive here.
# ---------------------------------------------------------------------
MAIN_TABLE = "cmisupport.dbo.alerts_result"
TEST_TABLE = "cmisupport.dbo.alerts_result_test"
KEY_COLUMNS = ["clientid", "facilityid"]
IGNORE_COLUMNS = []

# ---------------------------------------------------------------------
# STEP 6: Where to save the Excel report.
#   ""  = the same folder this script is in.
#   Or give a folder, e.g.  r"C:\Reports"   (keep the r in front on Windows)
# ---------------------------------------------------------------------
REPORT_FOLDER = ""
OPEN_REPORT_WHEN_DONE = True
WRITE_EXCEL_REPORT = False

# ---------------------------------------------------------------------
# (Optional) The SQL that runs for EACH alert ID, once per table.
#   This is your original script. Python fills in the three "%s" with the
#   alert ID, vendor ID and process date (in that order), and replaces
#   {table} with MAIN_TABLE or TEST_TABLE.
#   You normally do NOT need to change this. If you do, keep the three
#   "%s" lines in the same order, and write any other % sign as %%.
# ---------------------------------------------------------------------
QUERY_TEMPLATE = """
SET NOCOUNT ON;
DECLARE @AlertId     int  = %s;
DECLARE @VendorId    int  = %s;
DECLARE @ProcessDate date = %s;
 
SELECT *
FROM {table}
WHERE venderid  = @VendorId
  AND alertid   = @AlertId
  AND alertdate >= @ProcessDate
  AND alertdate <  DATEADD(day, 1, @ProcessDate)
ORDER BY clientid, facilityid;
"""

# =====================================================================
#          NOTHING BELOW THIS LINE NEEDS TO BE CHANGED
# =====================================================================

import getpass
import json
import os
import re
import sys
import webbrowser
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pandas as pd

SUF_MAIN = "__main"
SUF_TEST = "__test"
EXCEL_MAX_ROWS = 1_000_000  # Excel's hard limit is 1,048,576 rows per sheet
HTML_MAX_ROWS = 5_000  # rows per alert ID kept in the HTML report (keeps the file fast)


# ---------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------
def parse_process_date(text):
    """Turn "2026-10-07" or "20261007" into a date."""
    for fmt in ("%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(str(text).strip(), fmt).date()
        except ValueError:
            pass
    sys.exit(
        f"PROCESS_DATE '{text}' is not a valid date. "
        f"Use YYYY-MM-DD, e.g. 2026-10-07 (see STEP 3)."
    )


def check_settings():
    """Stop early with a friendly message if the placeholders were not filled in."""
    problems = []
    if SERVER.startswith("your-"):
        problems.append("STEP 1: set SERVER to your SQL Server host name or IP.")
    if DATABASE.startswith("your-"):
        problems.append("STEP 1: set DATABASE to your database name.")
    if USERNAME.startswith("your-"):
        problems.append("STEP 2: set USERNAME to your SQL login.")
    if not KEY_COLUMNS:
        problems.append("STEP 5: KEY_COLUMNS must contain at least one column.")
    if problems:
        print("Please edit the settings at the top of this file first:\n")
        for p in problems:
            print("  - " + p)
        sys.exit(1)


def is_null(value):
    """True for SQL NULL (None) and for pandas' 'missing' markers (NaN/NaT)."""
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def values_equal(a, b):
    """Exact comparison. NULL equals NULL; NULL never equals a value."""
    if is_null(a) and is_null(b):
        return True
    if is_null(a) or is_null(b):
        return False
    try:
        return bool(a == b)
    except Exception:
        return False


_ILLEGAL_EXCEL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def show(value):
    """How a value is written in the report."""
    if is_null(value):
        return "NULL"
    if isinstance(value, (bytes, bytearray)):
        text = "0x" + bytes(value).hex().upper()
    else:
        text = str(value)
    # Excel cannot store some invisible control characters; show them as \xNN
    return _ILLEGAL_EXCEL_CHARS.sub(lambda m: f"\\x{ord(m.group()):02x}", text)


def json_value(value):
    """Value as stored in the HTML report: None for SQL NULL, otherwise text."""
    return None if is_null(value) else show(value)


def describe_difference(a, b):
    """A short hint about WHY two values differ (shown in the 'Note' column)."""
    if is_null(a):
        return f"NULL in {MAIN_TABLE}"
    if is_null(b):
        return f"NULL in {TEST_TABLE}"
    if isinstance(a, str) and isinstance(b, str):
        if a.strip() == b.strip():
            return "Only leading/trailing spaces differ"
        if a.lower() == b.lower():
            return "Only upper/lower case differs"
    if type(a) is not type(b) and str(a) == str(b):
        return (
            f"Same text, different data type ({type(a).__name__} vs {type(b).__name__})"
        )
    subtractable = (int, float, Decimal, date, datetime)
    if (
        isinstance(a, subtractable)
        and isinstance(b, subtractable)
        and not isinstance(a, bool)
        and not isinstance(b, bool)
    ):
        try:
            return f"test minus main = {b - a}"
        except TypeError:
            pass
    return ""


def make_row_key(record, keys):
    text = ", ".join(f"{k}={show(record[k])}" for k in keys)
    if record["_dup"] > 0:
        text += f"  (duplicate key, copy #{int(record['_dup']) + 1})"
    return text


# ---------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------
def connect():
    try:
        import pymssql
    except ImportError:
        sys.exit(
            "The 'pymssql' package is missing. Run:  pip install pymssql   (STEP 0)"
        )

    password = PASSWORD or getpass.getpass(f"Password for {USERNAME} on {SERVER}: ")

    print(f"Connecting to {SERVER}:{PORT} / {DATABASE} as {USERNAME} ...")
    try:
        return pymssql.connect(
            server=SERVER,
            port=PORT,
            user=USERNAME,
            password=password,
            database=DATABASE,
            login_timeout=30,
        )
    except pymssql.Error as err:
        message = str(err)
        print("\nCould not connect to SQL Server.\n")
        print("Error from the server:\n  " + message + "\n")
        print("Things to check (compare with DBeaver > Edit Connection > Main tab):")
        low = message.lower()
        if "login failed" in low:
            print(
                "  - STEP 2: the USERNAME or password is wrong, or the login has no access to DATABASE."
            )
        if (
            "unavailable or does not exist" in low
            or "timed out" in low
            or "connection refused" in low
        ):
            print(
                "  - STEP 1: check SERVER and PORT, and that you're on the right network/VPN."
            )
        if "cannot open database" in low:
            print("  - STEP 1: DATABASE name is wrong, or your login can't access it.")
        print("  - STEP 1: SERVER, PORT and DATABASE match what DBeaver uses.")
        sys.exit(1)


def run_query(conn, table, alert_id, process_date):
    """Run the query for one alert ID against one table. Returns a DataFrame."""
    cursor = conn.cursor()
    cursor.execute(
        QUERY_TEMPLATE.format(table=table), (alert_id, VENDOR_ID, process_date)
    )
    # Skip anything before the actual result set (e.g. 'rows affected' messages)
    while cursor.description is None:
        if not cursor.nextset():
            cursor.close()
            return pd.DataFrame()
    columns = [col[0].lower() for col in cursor.description]
    rows = [list(row) for row in cursor.fetchall()]
    cursor.close()
    # dtype=object keeps every value exactly as SQL Server returned it
    return pd.DataFrame(rows, columns=columns, dtype=object)


# ---------------------------------------------------------------------
# Comparison
# ---------------------------------------------------------------------
def blank_summary(alert_id):
    return {
        "Alert ID": alert_id,
        f"Rows in {MAIN_TABLE}": None,
        f"Rows in {TEST_TABLE}": None,
        "Rows paired by key": None,
        "Identical rows": None,
        "Rows with differences": None,
        "Different cells": None,
        f"Only in {MAIN_TABLE}": None,
        f"Only in {TEST_TABLE}": None,
        "Column differences": None,
        "Duplicate keys": None,
        "Status": "",
        "Error": "",
    }


def _prepare(df, keys, compare_cols):
    """Keep the needed columns, number the rows, and number duplicate keys
    so that rows sharing the same key are paired up in a consistent way."""
    df = df[keys + compare_cols].copy()
    df["_row"] = range(1, len(df) + 1)
    if df.empty:
        df["_dup"] = pd.Series(dtype="int64")
        return df
    if compare_cols:
        df["_sort"] = [
            " | ".join(show(v) for v in row)
            for row in df[compare_cols].itertuples(index=False)
        ]
    else:
        df["_sort"] = ""
    df = df.sort_values("_sort", kind="stable")
    df["_dup"] = df.groupby(keys, dropna=False, sort=False).cumcount().astype("int64")
    return df.drop(columns="_sort").sort_values("_row")


def compare_one_alert(alert_id, main_df, test_df):
    """Compare the two query results for one alert ID."""
    keys = [k.lower() for k in KEY_COLUMNS]
    ignore = {c.lower() for c in IGNORE_COLUMNS} - set(keys)

    cell_diffs, missing_rows, column_diffs = [], [], []

    # 1) Are the same columns in both tables?
    main_cols = [c for c in main_df.columns if c not in ignore]
    test_cols = [c for c in test_df.columns if c not in ignore]
    for c in main_cols:
        if c not in test_cols:
            column_diffs.append(
                {"Alert ID": alert_id, "Column": c, "Exists only in": MAIN_TABLE}
            )
    for c in test_cols:
        if c not in main_cols:
            column_diffs.append(
                {"Alert ID": alert_id, "Column": c, "Exists only in": TEST_TABLE}
            )

    for table_name, df in ((MAIN_TABLE, main_df), (TEST_TABLE, test_df)):
        missing_keys = [k for k in keys if k not in df.columns]
        if missing_keys:
            raise ValueError(
                f"Key column(s) {missing_keys} not found in {table_name}. "
                f"Check KEY_COLUMNS in STEP 5."
            )

    compare_cols = [c for c in main_cols if c in test_cols and c not in keys]
    has_dup_keys = bool(
        main_df.duplicated(subset=keys).any() or test_df.duplicated(subset=keys).any()
    )

    # 2) Line the rows up by key
    a = _prepare(main_df, keys, compare_cols)
    b = _prepare(test_df, keys, compare_cols)
    merged = a.merge(
        b,
        on=keys + ["_dup"],
        how="outer",
        suffixes=(SUF_MAIN, SUF_TEST),
        indicator=True,
    )
    merged = merged.sort_values(
        ["_row" + SUF_MAIN, "_row" + SUF_TEST], na_position="last"
    )

    # 3) Compare every cell of every paired row
    #    (the lists ending in _html keep full rows so the HTML report can
    #     show each difference in context)
    column_counts = {c: 0 for c in compare_cols}
    diff_rows_html, missing_rows_html = [], []
    identical_rows = rows_with_diffs = only_main = only_test = 0
    for rec in merged.to_dict("records"):
        row_key = make_row_key(rec, keys)
        key_values = [json_value(rec[k]) for k in keys]
        if rec["_merge"] == "both":
            differing, notes = [], {}
            for i, col in enumerate(compare_cols):
                va, vb = rec[col + SUF_MAIN], rec[col + SUF_TEST]
                if not values_equal(va, vb):
                    note = describe_difference(va, vb)
                    position = len(keys) + i
                    differing.append(position)
                    if note:
                        notes[position] = note
                    column_counts[col] += 1
                    cell_diffs.append(
                        {
                            "Alert ID": alert_id,
                            "Row key": row_key,
                            f"Row # in {MAIN_TABLE}": int(rec["_row" + SUF_MAIN]),
                            f"Row # in {TEST_TABLE}": int(rec["_row" + SUF_TEST]),
                            "Column": col,
                            f"Value in {MAIN_TABLE}": show(va),
                            f"Value in {TEST_TABLE}": show(vb),
                            "Note": note,
                        }
                    )
            if differing:
                rows_with_diffs += 1
                diff_rows_html.append(
                    {
                        "rm": int(rec["_row" + SUF_MAIN]),
                        "rt": int(rec["_row" + SUF_TEST]),
                        "dup": int(rec["_dup"]),
                        "m": key_values
                        + [json_value(rec[c + SUF_MAIN]) for c in compare_cols],
                        "t": key_values
                        + [json_value(rec[c + SUF_TEST]) for c in compare_cols],
                        "d": differing,
                        "n": notes,
                    }
                )
            else:
                identical_rows += 1
        else:
            in_main = rec["_merge"] == "left_only"
            suffix = SUF_MAIN if in_main else SUF_TEST
            if in_main:
                only_main += 1
            else:
                only_test += 1
            missing_rows.append(
                {
                    "Alert ID": alert_id,
                    "Found only in": MAIN_TABLE if in_main else TEST_TABLE,
                    "Missing from": TEST_TABLE if in_main else MAIN_TABLE,
                    "Row key": row_key,
                    "Row #": int(rec["_row" + suffix]),
                    "Row data": " | ".join(
                        f"{c}={show(rec[c + suffix])}" for c in compare_cols
                    ),
                }
            )
            missing_rows_html.append(
                {
                    "side": "main" if in_main else "test",
                    "r": int(rec["_row" + suffix]),
                    "dup": int(rec["_dup"]),
                    "v": key_values
                    + [json_value(rec[c + suffix]) for c in compare_cols],
                }
            )

    # 4) Decide the status
    if len(main_df) == 0 and len(test_df) == 0:
        status = "NO DATA"
    elif not cell_diffs and not missing_rows and not column_diffs:
        status = "MATCH"
    else:
        status = "MISMATCH"

    summary = blank_summary(alert_id)
    summary.update(
        {
            f"Rows in {MAIN_TABLE}": len(main_df),
            f"Rows in {TEST_TABLE}": len(test_df),
            "Rows paired by key": identical_rows + rows_with_diffs,
            "Identical rows": identical_rows,
            "Rows with differences": rows_with_diffs,
            "Different cells": len(cell_diffs),
            f"Only in {MAIN_TABLE}": only_main,
            f"Only in {TEST_TABLE}": only_test,
            "Column differences": len(column_diffs),
            "Duplicate keys": "Yes" if has_dup_keys else "No",
            "Status": status,
        }
    )

    detail = {
        "columns": keys + compare_cols,
        "keys": keys,
        "column_counts": sorted(
            ([c, n] for c, n in column_counts.items() if n), key=lambda x: -x[1]
        ),
        "only_columns": [
            [d["Column"], "main" if d["Exists only in"] == MAIN_TABLE else "test"]
            for d in column_diffs
        ],
        "diff_rows": diff_rows_html[:HTML_MAX_ROWS],
        "diff_rows_total": len(diff_rows_html),
        "missing": missing_rows_html[:HTML_MAX_ROWS],
        "missing_total": len(missing_rows_html),
    }
    return summary, cell_diffs, missing_rows, column_diffs, detail


# ---------------------------------------------------------------------
# Excel report
# ---------------------------------------------------------------------
def _format_sheet(ws, status_column=None):
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1F3864")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ws.auto_filter.ref = ws.dimensions

    for column_cells in ws.columns:
        longest = max(
            (len(str(c.value)) for c in column_cells[:500] if c.value is not None),
            default=8,
        )
        ws.column_dimensions[get_column_letter(column_cells[0].column)].width = min(
            max(10, longest + 2), 60
        )

    if status_column:
        colors = {
            "MATCH": "C6EFCE",
            "MISMATCH": "FFC7CE",
            "NO DATA": "EDEDED",
            "ERROR": "FFEB9C",
        }
        headers = [c.value for c in ws[1]]
        if status_column in headers:
            col_idx = headers.index(status_column) + 1
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                if cell.value in colors:
                    cell.fill = PatternFill("solid", fgColor=colors[cell.value])
                    cell.font = Font(bold=True)


def _write_sheet(writer, df, sheet_name, csv_base):
    if len(df) > EXCEL_MAX_ROWS:
        csv_path = f"{csv_base}_{sheet_name.replace(' ', '_')}.csv"
        df.to_csv(csv_path, index=False)
        print(
            f"  NOTE: '{sheet_name}' has {len(df):,} rows, too many for Excel. "
            f"Full list saved to {csv_path}; the Excel sheet shows the first {EXCEL_MAX_ROWS:,}."
        )
        df = df.head(EXCEL_MAX_ROWS)
    df.to_excel(writer, sheet_name=sheet_name, index=False)


def report_base_path(process_date, started):
    """Folder + file name (without extension) shared by the HTML and Excel reports."""
    folder = REPORT_FOLDER or os.path.dirname(os.path.abspath(__file__))
    os.makedirs(folder, exist_ok=True)
    return os.path.join(
        folder,
        f"alert_comparison_vendor{VENDOR_ID}_{process_date:%Y%m%d}"
        f"_{started:%Y%m%d_%H%M%S}",
    )


def write_excel_report(
    summaries,
    cell_diffs,
    missing_rows,
    column_diffs,
    process_date,
    alert_ids,
    started,
    base,
):
    path = base + ".xlsx"

    cell_cols = [
        "Alert ID",
        "Row key",
        f"Row # in {MAIN_TABLE}",
        f"Row # in {TEST_TABLE}",
        "Column",
        f"Value in {MAIN_TABLE}",
        f"Value in {TEST_TABLE}",
        "Note",
    ]
    missing_cols = [
        "Alert ID",
        "Found only in",
        "Missing from",
        "Row key",
        "Row #",
        "Row data",
    ]
    column_cols = ["Alert ID", "Column", "Exists only in"]

    statuses = [s["Status"] for s in summaries]
    run_info = pd.DataFrame(
        [
            ("Run started", started.strftime("%Y-%m-%d %H:%M:%S")),
            ("Server", SERVER),
            ("Database", DATABASE),
            ("User", USERNAME),
            ("Main table", MAIN_TABLE),
            ("Test table", TEST_TABLE),
            ("Vendor ID", VENDOR_ID),
            ("Process date", process_date.isoformat()),
            ("Alert IDs compared", ", ".join(map(str, alert_ids))),
            ("Alert IDs excluded", ", ".join(map(str, sorted(EXCLUDED_ALERT_IDS)))),
            ("Key columns (used to pair rows)", ", ".join(KEY_COLUMNS)),
            ("Ignored columns", ", ".join(IGNORE_COLUMNS) or "(none)"),
            ("Alerts: MATCH", statuses.count("MATCH")),
            ("Alerts: MISMATCH", statuses.count("MISMATCH")),
            ("Alerts: NO DATA (both tables empty)", statuses.count("NO DATA")),
            ("Alerts: ERROR", statuses.count("ERROR")),
            ("Total different cells", len(cell_diffs)),
            ("Total missing rows", len(missing_rows)),
            (
                "Row # meaning",
                "Position of the row in that table's query result, sorted by the ORDER BY.",
            ),
        ],
        columns=["Setting", "Value"],
    )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        _write_sheet(writer, pd.DataFrame(summaries), "Summary", base)
        _write_sheet(
            writer,
            pd.DataFrame(cell_diffs, columns=cell_cols),
            "Cell Differences",
            base,
        )
        _write_sheet(
            writer,
            pd.DataFrame(missing_rows, columns=missing_cols),
            "Missing Rows",
            base,
        )
        _write_sheet(
            writer,
            pd.DataFrame(column_diffs, columns=column_cols),
            "Column Differences",
            base,
        )
        _write_sheet(writer, run_info, "Run Info", base)

        _format_sheet(writer.sheets["Summary"], status_column="Status")
        for name in (
            "Cell Differences",
            "Missing Rows",
            "Column Differences",
            "Run Info",
        ):
            _format_sheet(writer.sheets[name])
    return path


# ---------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------
def write_html_report(summaries, details, process_date, started, base):
    alerts = []
    for s in summaries:
        alert_id = s["Alert ID"]
        alerts.append(
            {
                "id": alert_id,
                "status": s["Status"],
                "error": s["Error"],
                "rows_main": s[f"Rows in {MAIN_TABLE}"],
                "rows_test": s[f"Rows in {TEST_TABLE}"],
                "identical": s["Identical rows"],
                "rows_diff": s["Rows with differences"],
                "cells": s["Different cells"],
                "only_main": s[f"Only in {MAIN_TABLE}"],
                "only_test": s[f"Only in {TEST_TABLE}"],
                "col_diffs": s["Column differences"],
                "dup": s["Duplicate keys"] == "Yes",
                "detail": details.get(alert_id),
            }
        )
    data = {
        "main_table": MAIN_TABLE,
        "test_table": TEST_TABLE,
        "vendor": VENDOR_ID,
        "date": process_date.isoformat(),
        "run": started.strftime("%Y-%m-%d %H:%M"),
        "server": SERVER,
        "database": DATABASE,
        "keys": [k.lower() for k in KEY_COLUMNS],
        "ignored": [c.lower() for c in IGNORE_COLUMNS],
        "excluded": sorted(EXCLUDED_ALERT_IDS),
        "html_max_rows": HTML_MAX_ROWS,
        "alerts": alerts,
    }
    # "<" is escaped so no value from the database can break out of the <script> tag
    payload = json.dumps(data, default=str).replace("<", "\\u003c")
    path = base + ".html"
    with open(path, "w", encoding="utf-8") as f:
        f.write(HTML_TEMPLATE.replace("__REPORT_DATA__", payload))
    return path


HTML_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Alert comparison report</title>
<style>
:root{
  --paper:#F5F6F8; --surface:#FFFFFF; --ink:#1A2130; --muted:#586174; --faint:#8790A0;
  --rule:#E1E4EA; --rule-strong:#C7CCD6;
  --main:#2A57C0; --main-bg:#E8EEFB; --test:#7340C6; --test-bg:#F0E8FA;
  --mark:#FFE14F; --mark-ink:#1A2130; --mark-deep:#C99A00;
  --ok:#1C7A4E; --ok-bg:#E0F2E8; --bad:#B22E26; --bad-bg:#FBE4E2;
  --warn:#875C00; --warn-bg:#FBEFD0; --none:#5C6472; --none-bg:#EBEDF1;
  --sans:"Segoe UI Variable Text","Segoe UI",system-ui,-apple-system,"Helvetica Neue",Arial,sans-serif;
  --mono:ui-monospace,"Cascadia Mono","Cascadia Code",Consolas,"SF Mono",Menlo,monospace;
}
@media (prefers-color-scheme: dark){
  :root{
    --paper:#13161C; --surface:#1A1E26; --ink:#E5E8EE; --muted:#A1A9B7; --faint:#7A8291;
    --rule:#2A303B; --rule-strong:#3A414E;
    --main:#86A6F2; --main-bg:#1D2943; --test:#B790F2; --test-bg:#2A2240;
    --mark:#F0D046; --mark-ink:#15181E; --mark-deep:#F0D046;
    --ok:#5DC48F; --ok-bg:#15301F; --bad:#F0837A; --bad-bg:#3A1B19;
    --warn:#E6B64B; --warn-bg:#342912; --none:#A1A9B7; --none-bg:#252A33;
  }
}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font:15px/1.5 var(--sans);font-variant-numeric:tabular-nums}
a{color:var(--main)}
:focus-visible{outline:2px solid var(--main);outline-offset:2px}
select,button,input{font:inherit;color:inherit}
 
.bar{position:sticky;top:0;z-index:5;background:var(--surface);border-bottom:1px solid var(--rule)}
.bar-in{max-width:1320px;margin:0 auto;padding:10px 24px;display:flex;align-items:center;gap:12px 20px;flex-wrap:wrap}
.brand{margin-right:auto;line-height:1.3}
.brand strong{display:block;font-weight:600}
.brand span{color:var(--muted);font-size:13px}
.nav{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.navlink{padding:7px 12px;border-radius:6px;text-decoration:none;color:var(--ink);font-weight:500;border:1px solid transparent}
.navlink:hover{border-color:var(--rule-strong)}
.navlink[aria-current=page]{background:var(--main-bg);color:var(--main)}
.picker{display:flex;align-items:center;gap:6px}
.picker label{font-size:14px;color:var(--muted);white-space:nowrap}
select{padding:7px 10px;border:1px solid var(--rule-strong);border-radius:6px;background:var(--surface);min-width:250px;max-width:70vw}
.btn{padding:6px 11px;border:1px solid var(--rule-strong);border-radius:6px;background:var(--surface);cursor:pointer;line-height:1.4}
.btn:hover:not(:disabled){border-color:var(--ink)}
.btn:disabled{opacity:.35;cursor:default}
 
main{max-width:1320px;margin:0 auto;padding:36px 24px 72px}
.crumb{margin:0 0 10px;font-size:14px}
.titleline{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin:0 0 8px}
h1{font-size:32px;line-height:1.2;font-weight:650;letter-spacing:-.015em;margin:0;max-width:32ch}
.lede{color:var(--muted);margin:6px 0 26px;max-width:78ch;font-size:16px}
.lede b{color:var(--ink);font-weight:600}
h2{font-size:19px;font-weight:600;margin:44px 0 4px}
h3{font-size:16px;font-weight:600;margin:24px 0 8px}
.hint{color:var(--muted);font-size:14px;margin:0 0 14px;max-width:78ch}
 
.facts{display:flex;flex-wrap:wrap;margin:0;padding:0;border-top:1px solid var(--rule);border-bottom:1px solid var(--rule)}
.facts div{padding:12px 36px 12px 0}
.facts dt{font-size:13px;color:var(--muted)}
.facts dd{margin:0;font-size:24px;font-weight:600}
.facts dd.bad{color:var(--bad)}
 
.pill{display:inline-block;padding:2px 10px;border-radius:999px;font-size:13px;font-weight:600;white-space:nowrap;line-height:1.5}
.s-MATCH{background:var(--ok-bg);color:var(--ok)}
.s-MISMATCH{background:var(--bad-bg);color:var(--bad)}
.s-NO_DATA{background:var(--none-bg);color:var(--none)}
.s-ERROR{background:var(--warn-bg);color:var(--warn)}
.titleline .pill{font-size:15px;padding:3px 12px}
 
.tag{display:inline-block;padding:0 7px;border-radius:4px;font-size:12px;font-weight:600}
.tag-main{background:var(--main-bg);color:var(--main)}
.tag-test{background:var(--test-bg);color:var(--test)}
 
.panel{border:1px solid var(--rule);border-left-width:4px;border-radius:6px;padding:12px 16px;margin:16px 0;background:var(--surface);max-width:90ch}
.panel p{margin:0}
.panel p+p{margin-top:6px}
.panel.err{border-left-color:var(--bad)}
.panel.warn{border-left-color:var(--warn)}
.panel.ok{border-left-color:var(--ok)}
.panel code,.mono{font-family:var(--mono);font-size:13px}
 
.bars{display:grid;grid-template-columns:minmax(140px,max-content) minmax(120px,420px) max-content;gap:6px 14px;align-items:center}
.bars .col{all:unset;cursor:pointer;font-family:var(--mono);font-size:14px;padding:2px 6px;margin-left:-6px;border-radius:4px;overflow-wrap:anywhere}
.bars .col:hover{text-decoration:underline}
.bars .col:focus-visible{outline:2px solid var(--main)}
.bars .col[aria-pressed=true]{background:var(--mark);color:var(--mark-ink)}
.track{height:8px;background:var(--rule);border-radius:4px;overflow:hidden}
.fill{height:100%;background:var(--mark-deep);border-radius:4px}
.count{font-size:14px;color:var(--muted)}
 
.toolbar{display:flex;align-items:center;gap:10px 18px;flex-wrap:wrap;margin:0 0 12px;font-size:14px}
.toolbar label{display:flex;align-items:center;gap:6px;cursor:pointer}
.chips{display:flex;gap:6px;flex-wrap:wrap}
.chip{padding:4px 12px;border:1px solid var(--rule-strong);border-radius:999px;background:var(--surface);cursor:pointer;font-size:14px}
.chip[aria-pressed=true]{background:var(--ink);border-color:var(--ink);color:var(--paper)}
.filtered{background:var(--mark);color:var(--mark-ink);padding:2px 8px;border-radius:4px}
.linkbtn{all:unset;cursor:pointer;color:var(--main);text-decoration:underline}
.linkbtn:focus-visible{outline:2px solid var(--main)}
 
.tablewrap{overflow-x:auto;border:1px solid var(--rule);border-radius:8px;background:var(--surface)}
table{border-collapse:collapse;width:100%;font-size:14px}
th{text-align:left;font-weight:600;padding:9px 12px;border-bottom:1px solid var(--rule-strong);white-space:nowrap;vertical-align:bottom}
th.dcol{box-shadow:inset 0 -3px 0 var(--mark-deep)}
td{padding:7px 12px;border-bottom:1px solid var(--rule);vertical-align:top}
.num{text-align:right}
td.v{font-family:var(--mono);font-size:13px;white-space:nowrap}
td.v.long{white-space:normal;min-width:24ch;max-width:44ch;overflow-wrap:anywhere}
td:first-child{white-space:nowrap}
td.keyv{font-weight:600}
td.diff{background:var(--mark);color:var(--mark-ink)}
td.diff .null{color:var(--mark-ink);opacity:.7}
.null{color:var(--faint);font-style:italic;font-family:var(--sans)}
.ws{color:var(--bad);font-weight:700}
td.diff .ws{color:var(--bad)}
.note{display:block;font-family:var(--sans);font-size:12px;margin-top:3px;opacity:.85}
.dupnote{display:block;font-size:12px;color:var(--warn)}
tbody.pair tr:last-child td{border-bottom:2px solid var(--rule-strong)}
tbody.pair tr:first-child td{border-bottom:1px dashed var(--rule)}
td.side{white-space:nowrap;width:1%}
tr.click{cursor:pointer}
tr.click:hover td{background:var(--paper)}
.zero{color:var(--faint)}
.notes-cell{color:var(--muted);font-size:13px;max-width:44ch}
 
.more{margin:12px 0 0}
.empty{color:var(--muted);padding:18px 0}
.legend{font-size:14px;color:var(--muted);margin:0 0 6px}
.legend .tag{margin-right:4px}
.legend span+span{margin-left:16px}
.swatch{display:inline-block;width:14px;height:14px;border-radius:3px;background:var(--mark);vertical-align:-2px;margin-right:5px}
footer{max-width:1320px;margin:0 auto;padding:0 24px 40px;color:var(--faint);font-size:13px}
 
@media (max-width:720px){
  .bar-in,main,footer{padding-left:16px;padding-right:16px}
  h1{font-size:25px}
  select{min-width:0;width:100%}
  .picker{width:100%}
  .facts div{padding-right:22px}
  .bars{grid-template-columns:minmax(100px,max-content) 1fr max-content}
}
@media print{
  .bar{position:static}
  .nav{display:none}
  .tablewrap{overflow:visible}
}
</style>
</head>
<body>
<header class="bar">
  <div class="bar-in">
    <div class="brand"><strong>Alert table comparison</strong><span id="brand-sub"></span></div>
    <nav class="nav" aria-label="Report pages">
      <a class="navlink" id="nav-summary" href="#summary">Summary</a>
      <div class="picker">
        <label for="alert-select">Alert ID</label>
        <button class="btn" id="prev" type="button" aria-label="Previous alert ID">&#8249;</button>
        <select id="alert-select"></select>
        <button class="btn" id="next" type="button" aria-label="Next alert ID">&#8250;</button>
      </div>
    </nav>
  </div>
</header>
<main id="app"><noscript>This report needs JavaScript turned on in your browser.</noscript></main>
<footer id="foot"></footer>
<script type="application/json" id="report-data">__REPORT_DATA__</script>
<script>
(function () {
  "use strict";
  var D = JSON.parse(document.getElementById("report-data").textContent);
  var A = D.alerts;
  var byId = {};
  A.forEach(function (a) { byId[String(a.id)] = a; });
 
  var app = document.getElementById("app");
  var sel = document.getElementById("alert-select");
  var prevBtn = document.getElementById("prev");
  var nextBtn = document.getElementById("next");
  var navSummary = document.getElementById("nav-summary");
 
  var LABEL = { "MATCH": "Match", "MISMATCH": "Mismatch", "NO DATA": "No data", "ERROR": "Error" };
  var ORDER = ["MISMATCH", "ERROR", "MATCH", "NO DATA"];
  var PAGE = 100;
 
  // ---------- small helpers ----------
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function n(v) { return v === null || v === undefined ? "–" : Number(v).toLocaleString(); }
  function plural(v, one, many) { return n(v) + " " + (v === 1 ? one : (many || one + "s")); }
  function pill(s) { return '<span class="pill s-' + s.replace(" ", "_") + '">' + (LABEL[s] || s) + "</span>"; }
  function joinAnd(parts) {
    if (parts.length < 2) return parts.join("");
    return parts.slice(0, -1).join(", ") + " and " + parts[parts.length - 1];
  }
  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
  function val(v, showSpaces) {
    if (v === null) return '<span class="null">NULL</span>';
    if (v === "") return '<span class="null">(empty)</span>';
    var s = esc(v);
    if (showSpaces) {
      s = s.replace(/^ +| +$/g, function (m) {
        return '<span class="ws" title="' + m.length + ' space(s)">' + Array(m.length + 1).join("␣") + "</span>";
      });
    }
    return s;
  }
  function isLong(v) { return v !== null && String(v).length > 40; }
  function zero(v) { return v ? n(v) : '<span class="zero">0</span>'; }
  function missingCount(a) { return (a.only_main || 0) + (a.only_test || 0); }
 
  // ---------- header ----------
  document.getElementById("brand-sub").textContent =
    "Vendor " + D.vendor + ", process date " + D.date;
  document.getElementById("foot").innerHTML =
    "Created " + esc(D.run) + " from " + esc(D.server) + " / " + esc(D.database) +
    ". Rows are paired by " + esc(D.keys.join(" + ")) + "." +
    (D.ignored.length ? " Columns not compared: " + esc(D.ignored.join(", ")) + "." : "") +
    " Alert IDs excluded: " + esc(D.excluded.join(", ")) + ".";
 
  (function buildSelect() {
    var html = '<option value="">Choose an alert ID</option>';
    ORDER.forEach(function (st) {
      var list = A.filter(function (a) { return a.status === st; });
      if (!list.length) return;
      html += '<optgroup label="' + LABEL[st] + " (" + list.length + ')">';
      list.forEach(function (a) {
        var extra = "";
        if (st === "MISMATCH") {
          var bits = [];
          if (a.cells) bits.push(plural(a.cells, "cell"));
          if (missingCount(a)) bits.push(plural(missingCount(a), "missing row"));
          if (a.col_diffs) bits.push(plural(a.col_diffs, "column"));
          extra = bits.length ? " (" + bits.join(", ") + ")" : "";
        }
        html += '<option value="' + a.id + '">Alert ' + a.id + extra + "</option>";
      });
      html += "</optgroup>";
    });
    sel.innerHTML = html;
  })();
 
  sel.addEventListener("change", function () { if (sel.value) location.hash = "alert-" + sel.value; });
  prevBtn.addEventListener("click", function () { step(-1); });
  nextBtn.addEventListener("click", function () { step(1); });
  function currentIndex() {
    var m = location.hash.match(/^#alert-(\d+)$/);
    if (!m) return -1;
    for (var i = 0; i < A.length; i++) if (String(A[i].id) === m[1]) return i;
    return -1;
  }
  function step(d) {
    var i = currentIndex();
    var j = i === -1 ? (d > 0 ? 0 : A.length - 1) : i + d;
    if (j >= 0 && j < A.length) location.hash = "alert-" + A[j].id;
  }
  function setNav(alert) {
    var i = alert ? A.indexOf(alert) : -1;
    sel.value = alert ? String(alert.id) : "";
    prevBtn.disabled = alert ? i <= 0 : false;
    nextBtn.disabled = alert ? i >= A.length - 1 : false;
    if (alert) navSummary.removeAttribute("aria-current"); else navSummary.setAttribute("aria-current", "page");
  }
 
  // ---------- summary page ----------
  var sumState = { status: "ALL", col: null };
 
  function columnHotspots() {
    var map = {};
    A.forEach(function (a) {
      if (!a.detail) return;
      a.detail.column_counts.forEach(function (cc) {
        var m = map[cc[0]] || (map[cc[0]] = { col: cc[0], cells: 0, alerts: [] });
        m.cells += cc[1];
        m.alerts.push(a.id);
      });
    });
    return Object.keys(map).map(function (k) { return map[k]; })
      .sort(function (x, y) { return y.cells - x.cells; });
  }
 
  function renderSummary() {
    setNav(null);
    document.title = "Summary: alert comparison";
    var count = {};
    ORDER.forEach(function (s) { count[s] = 0; });
    var cells = 0, missing = 0;
    A.forEach(function (a) { count[a.status] = (count[a.status] || 0) + 1; cells += a.cells || 0; missing += missingCount(a); });
 
    var headline;
    if (!count.MISMATCH && !count.ERROR) {
      headline = count.MATCH ? "Every alert ID with data matches exactly." : "Neither table has rows for any of these alert IDs.";
    } else {
      headline = n(count.MISMATCH) + " of " + n(A.length) + " alert IDs don't match" +
        (count.ERROR ? ", and " + plural(count.ERROR, "couldn't be checked", "couldn't be checked") : "") + ".";
      if (count.ERROR && !count.MISMATCH) headline = plural(count.ERROR, "alert ID", "alert IDs") + " couldn't be checked.";
    }
 
    var h = "";
    h += '<div class="titleline"><h1>' + headline + "</h1></div>";
    h += '<p class="lede"><b>' + esc(D.main_table) + "</b> (Main) compared with <b>" + esc(D.test_table) +
      "</b> (Test) for vendor " + esc(D.vendor) + " on " + esc(D.date) + ", across " + plural(A.length, "alert ID") + ".</p>";
 
    h += '<dl class="facts">' +
      fact("Match", count.MATCH) + fact("Mismatch", count.MISMATCH, count.MISMATCH > 0) +
      fact("No data", count["NO DATA"]) + fact("Error", count.ERROR, count.ERROR > 0) +
      fact("Different cells", cells, cells > 0) + fact("Rows in one table only", missing, missing > 0) + "</dl>";
 
    var onlyCols = [];
    A.forEach(function (a) { if (a.detail) a.detail.only_columns.forEach(function (c) { if (onlyCols.indexOf(c[0] + "|" + c[1]) < 0) onlyCols.push(c[0] + "|" + c[1]); }); });
    if (onlyCols.length) {
      h += '<div class="panel warn"><p><b>The tables don’t have the same columns.</b> ' +
        onlyCols.map(function (s) { var p = s.split("|"); return '<code>' + esc(p[0]) + "</code> is only in " + (p[1] === "main" ? "Main" : "Test"); }).join("; ") +
        ". These columns can’t be compared.</p></div>";
    }
 
    var hot = columnHotspots();
    if (hot.length) {
      var max = hot[0].cells;
      h += "<h2>Columns that differ most</h2>";
      h += '<p class="hint">Total different cells per column, across all alert IDs. Select a column to list only the alert IDs where it differs.</p>';
      h += '<div class="bars">';
      hot.slice(0, 15).forEach(function (c) {
        h += '<button class="col" type="button" data-scol="' + esc(c.col) + '" aria-pressed="' + (sumState.col === c.col) + '">' + esc(c.col) + "</button>" +
          '<div class="track"><div class="fill" style="width:' + Math.max(2, Math.round(c.cells / max * 100)) + '%"></div></div>' +
          '<span class="count">' + plural(c.cells, "cell") + " in " + plural(c.alerts.length, "alert ID") + "</span>";
      });
      h += "</div>";
      if (hot.length > 15) h += '<p class="hint" style="margin-top:8px">' + (hot.length - 15) + " more columns have fewer differences; open an alert ID to see them.</p>";
    }
 
    h += "<h2>All alert IDs</h2>";
    h += '<p class="hint">Select a row to see exactly where that alert ID differs.</p>';
    h += '<div class="toolbar"><div class="chips" role="group" aria-label="Filter by status">' +
      chip("ALL", "All", A.length) + ORDER.map(function (s) { return count[s] ? chip(s, LABEL[s], count[s]) : ""; }).join("") +
      "</div>" + (sumState.col ? '<span class="filtered">Only alert IDs where <span class="mono">' + esc(sumState.col) + '</span> differs</span> <button class="linkbtn" type="button" data-clear-scol>Show all</button>' : "") + "</div>";
    h += '<div class="tablewrap"><table><thead><tr><th>Alert ID</th><th>Status</th><th class="num">Rows in Main</th><th class="num">Rows in Test</th><th class="num">Different cells</th><th class="num">Rows with differences</th><th class="num">Only in Main</th><th class="num">Only in Test</th><th>Notes</th></tr></thead><tbody>';
    var colAlerts = null;
    if (sumState.col) hot.forEach(function (c) { if (c.col === sumState.col) colAlerts = c.alerts; });
    var shown = A.filter(function (a) {
      return (sumState.status === "ALL" || a.status === sumState.status) && (!colAlerts || colAlerts.indexOf(a.id) >= 0);
    });
    shown.forEach(function (a) {
      var notes = [];
      if (a.error) notes.push(esc(a.error));
      if (a.col_diffs) notes.push(plural(a.col_diffs, "column") + " in one table only");
      if (a.dup) notes.push("Duplicate " + esc(D.keys.join(" + ")) + " values");
      h += '<tr class="click" data-alert="' + a.id + '"><td><a href="#alert-' + a.id + '">Alert ' + a.id + "</a></td><td>" + pill(a.status) +
        '</td><td class="num">' + n(a.rows_main) + '</td><td class="num">' + n(a.rows_test) +
        '</td><td class="num">' + zero(a.cells) + '</td><td class="num">' + zero(a.rows_diff) +
        '</td><td class="num">' + zero(a.only_main) + '</td><td class="num">' + zero(a.only_test) +
        '</td><td class="notes-cell">' + notes.join("<br>") + "</td></tr>";
    });
    if (!shown.length) h += '<tr><td colspan="9" class="empty">No alert IDs match this filter.</td></tr>';
    h += "</tbody></table></div>";
    app.innerHTML = h;
  }
  function fact(label, v, bad) { return "<div><dt>" + label + '</dt><dd class="' + (bad ? "bad" : "") + '">' + n(v || 0) + "</dd></div>"; }
  function chip(key, label, c) { return '<button class="chip" type="button" data-status="' + key + '" aria-pressed="' + (sumState.status === key) + '">' + label + " " + c + "</button>"; }
 
  // ---------- alert page ----------
  var al = { alert: null, col: null, all: false, limit: PAGE };
 
  function renderAlert(a) {
    if (al.alert !== a) al = { alert: a, col: null, all: false, limit: PAGE };
    setNav(a);
    document.title = "Alert " + a.id + ": " + (LABEL[a.status] || a.status);
    var d = a.detail;
    var h = '<p class="crumb"><a href="#summary">Summary of all alert IDs</a></p>';
    h += '<div class="titleline"><h1>Alert ' + a.id + "</h1>" + pill(a.status) + "</div>";
 
    if (a.status === "ERROR") {
      h += '<p class="lede">This alert ID couldn’t be compared because a query failed.</p>';
      h += '<div class="panel err"><p><b>Error from the database or script</b></p><p class="mono">' + esc(a.error) + "</p>" +
        "<p>Check that the column and table names in the script’s query match the database, then run the script again.</p></div>";
      app.innerHTML = h; return;
    }
 
    var cols = d.columns.length;
    var lede;
    if (a.status === "NO DATA") lede = "Neither table has rows for this alert ID on " + esc(D.date) + ", so there is nothing to compare.";
    else if (a.status === "MATCH") lede = "All " + plural(a.rows_main, "row") + " are identical in both tables, across all " + plural(cols, "compared column") + ".";
    else {
      var parts = [];
      if (a.cells) parts.push("<b>" + plural(a.cells, "cell") + "</b> " + (a.cells === 1 ? "differs" : "differ") + " across " + plural(a.rows_diff, "row"));
      if (a.only_main) parts.push("<b>" + plural(a.only_main, "row") + "</b> " + (a.only_main === 1 ? "is" : "are") + " in Main but missing from Test");
      if (a.only_test) parts.push("<b>" + plural(a.only_test, "row") + "</b> " + (a.only_test === 1 ? "is" : "are") + " in Test but missing from Main");
      if (a.col_diffs) parts.push("<b>" + plural(a.col_diffs, "column") + "</b> " + (a.col_diffs === 1 ? "exists" : "exist") + " in only one table");
      lede = cap(joinAnd(parts)) + ".";
    }
    h += '<p class="lede">' + lede + "</p>";
 
    h += '<dl class="facts">' + fact("Rows in Main", a.rows_main) + fact("Rows in Test", a.rows_test) +
      fact("Identical rows", a.identical) + fact("Different cells", a.cells, a.cells > 0) +
      fact("Only in Main", a.only_main, a.only_main > 0) + fact("Only in Test", a.only_test, a.only_test > 0) + "</dl>";
 
    if (a.status === "MATCH" || a.status === "NO DATA") { app.innerHTML = h; return; }
 
    h += '<p class="legend" style="margin-top:18px"><span><span class="tag tag-main">Main</span>' + esc(D.main_table) +
      '</span><span><span class="tag tag-test">Test</span>' + esc(D.test_table) +
      '</span><span><span class="swatch"></span>Cell differs</span></p>';
 
    if (d.only_columns.length) {
      h += '<div class="panel warn"><p><b>Columns that exist in only one table</b></p><p>' +
        d.only_columns.map(function (c) { return "<code>" + esc(c[0]) + "</code> is only in " + (c[1] === "main" ? "Main" : "Test"); }).join("; ") +
        ". These aren’t compared.</p></div>";
    }
    if (a.dup) {
      h += '<div class="panel warn"><p><b>Some rows share the same ' + esc(D.keys.join(" + ")) + ".</b> " +
        "Those rows were paired up as closely as possible, so a difference shown on a duplicate may just mean the copies were paired in a different order.</p></div>";
    }
 
    if (d.column_counts.length) {
      var max = d.column_counts[0][1];
      h += "<h2>Where the differences are</h2>";
      h += '<p class="hint">Number of rows where each column differs. Select a column to show only those rows below.</p><div class="bars">';
      d.column_counts.forEach(function (cc) {
        h += '<button class="col" type="button" data-col="' + esc(cc[0]) + '" aria-pressed="' + (al.col === cc[0]) + '">' + esc(cc[0]) + "</button>" +
          '<div class="track"><div class="fill" style="width:' + Math.max(2, Math.round(cc[1] / max * 100)) + '%"></div></div>' +
          '<span class="count">' + plural(cc[1], "row") + "</span>";
      });
      h += "</div>";
      h += '<h2 id="diff-rows">Rows with different values</h2>';
      h += '<p class="hint">Each pair shows the same row from both tables, one above the other. Highlighted cells are different. ' +
        "Row # is the row’s position in that table’s query result, so you can find it by running your SQL.</p>";
      h += '<div id="grid"></div>';
    }
 
    if (d.missing.length) {
      h += "<h2>Rows in only one table</h2>";
      h += '<p class="hint">No row with the same ' + esc(D.keys.join(" + ")) + " exists in the other table.</p>";
      h += missingTable(d, "main") + missingTable(d, "test");
      if (d.missing_total > d.missing.length) h += '<p class="hint">Showing the first ' + n(d.missing.length) + " of " + n(d.missing_total) + " rows. Set WRITE_EXCEL_REPORT = True in the script to get the full list.</p>";
    }
 
    app.innerHTML = h;
    if (d.column_counts.length) renderGrid();
  }
 
  function missingTable(d, side) {
    var rows = d.missing.filter(function (r) { return r.side === side; });
    if (!rows.length) return "";
    var nk = d.keys.length;
    var h = "<h3>" + (side === "main" ? '<span class="tag tag-main">Main</span> has ' : '<span class="tag tag-test">Test</span> has ') +
      plural(rows.length, "row") + " missing from " + (side === "main" ? "Test" : "Main") + "</h3>";
    h += '<div class="tablewrap"><table><thead><tr><th class="num">Row #</th>' +
      d.columns.map(function (c) { return "<th>" + esc(c) + "</th>"; }).join("") + "</tr></thead><tbody>";
    rows.forEach(function (r) {
      h += '<tr><td class="num">' + r.r + (r.dup ? '<span class="dupnote">duplicate key</span>' : "") + "</td>" +
        r.v.map(function (v, i) { return '<td class="v' + (i < nk ? " keyv" : "") + (isLong(v) ? " long" : "") + '">' + val(v, false) + "</td>"; }).join("") + "</tr>";
    });
    return h + "</tbody></table></div>";
  }
 
  function renderGrid() {
    var a = al.alert, d = a.detail, nk = d.keys.length;
    var rows = d.diff_rows;
    if (al.col !== null) {
      var ci = d.columns.indexOf(al.col);
      rows = rows.filter(function (r) { return r.d.indexOf(ci) >= 0; });
    }
    var show = [];
    if (al.all) { for (var i = 0; i < d.columns.length; i++) show.push(i); }
    else {
      var set = {};
      rows.forEach(function (r) { r.d.forEach(function (x) { set[x] = true; }); });
      for (var k = 0; k < nk; k++) show.push(k);
      Object.keys(set).map(Number).sort(function (x, y) { return x - y; }).forEach(function (x) { show.push(x); });
    }
    var anyDiff = {};
    rows.forEach(function (r) { r.d.forEach(function (x) { anyDiff[x] = true; }); });
 
    var h = '<div class="toolbar"><label><input type="checkbox" id="showall"' + (al.all ? " checked" : "") + "> Show every column (" + d.columns.length + ")</label>";
    if (al.col !== null) h += '<span class="filtered">Only rows where <span class="mono">' + esc(al.col) + '</span> differs</span> <button class="linkbtn" type="button" data-clear-col>Show all rows</button>';
    h += '<span class="count">' + plural(rows.length, "row") + "</span></div>";
 
    h += '<div class="tablewrap"><table><thead><tr><th>Table</th><th class="num">Row #</th>' +
      show.map(function (i) { return '<th class="' + (anyDiff[i] ? "dcol" : "") + '">' + esc(d.columns[i]) + "</th>"; }).join("") + "</tr></thead>";
    rows.slice(0, al.limit).forEach(function (r) {
      var dif = {};
      r.d.forEach(function (x) { dif[x] = true; });
      h += '<tbody class="pair">';
      h += '<tr><td class="side"><span class="tag tag-main">Main</span></td><td class="num">' + r.rm +
        (r.dup ? '<span class="dupnote">duplicate key, copy ' + (r.dup + 1) + "</span>" : "") + "</td>" +
        show.map(function (i) { return cell(r.m[i], dif[i], i < nk, null); }).join("") + "</tr>";
      h += '<tr><td class="side"><span class="tag tag-test">Test</span></td><td class="num">' + r.rt + "</td>" +
        show.map(function (i) { return cell(r.t[i], dif[i], i < nk, r.n[i]); }).join("") + "</tr>";
      h += "</tbody>";
    });
    h += "</table></div>";
    if (rows.length > al.limit) {
      h += '<p class="more"><button class="btn" type="button" id="more">Show ' + Math.min(PAGE, rows.length - al.limit) + " more rows</button> " +
        '<span class="count">Showing ' + n(al.limit) + " of " + n(rows.length) + "</span></p>";
    }
    if (d.diff_rows_total > d.diff_rows.length) {
      h += '<p class="hint">This report keeps the first ' + n(d.diff_rows.length) + " of " + n(d.diff_rows_total) +
        " rows with differences for this alert ID. Set WRITE_EXCEL_REPORT = True in the script to get every row.</p>";
    }
    document.getElementById("grid").innerHTML = h;
  }
 
  function cell(v, isDiff, isKey, note) {
    var cls = "v" + (isDiff ? " diff" : "") + (isKey ? " keyv" : "") + (isLong(v) ? " long" : "");
    return '<td class="' + cls + '">' + val(v, isDiff) + (note ? '<span class="note">' + esc(note) + "</span>" : "") + "</td>";
  }
 
  // ---------- events ----------
  app.addEventListener("click", function (e) {
    var t = e.target.closest("[data-col],[data-clear-col],[data-scol],[data-clear-scol],[data-status],#more,tr[data-alert]");
    if (!t) return;
    if (t.hasAttribute("data-col")) {
      var c = t.getAttribute("data-col");
      al.col = al.col === c ? null : c; al.limit = PAGE;
      app.querySelectorAll("[data-col]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.getAttribute("data-col") === al.col)); });
      renderGrid();
      var target = document.getElementById("diff-rows");
      if (target && al.col !== null) target.scrollIntoView({ block: "start" });
    } else if (t.hasAttribute("data-clear-col")) {
      al.col = null; al.limit = PAGE;
      app.querySelectorAll("[data-col]").forEach(function (b) { b.setAttribute("aria-pressed", "false"); });
      renderGrid();
    } else if (t.id === "more") {
      al.limit += PAGE; renderGrid();
    } else if (t.hasAttribute("data-scol")) {
      var sc = t.getAttribute("data-scol");
      sumState.col = sumState.col === sc ? null : sc; renderSummary();
    } else if (t.hasAttribute("data-clear-scol")) {
      sumState.col = null; renderSummary();
    } else if (t.hasAttribute("data-status")) {
      sumState.status = t.getAttribute("data-status"); renderSummary();
    } else if (t.hasAttribute("data-alert") && !e.target.closest("a")) {
      location.hash = "alert-" + t.getAttribute("data-alert");
    }
  });
  app.addEventListener("change", function (e) {
    if (e.target.id === "showall") { al.all = e.target.checked; renderGrid(); }
  });
 
  function route() {
    var m = location.hash.match(/^#alert-(\d+)$/);
    if (m && byId[m[1]]) renderAlert(byId[m[1]]); else renderSummary();
    window.scrollTo(0, 0);
  }
  window.addEventListener("hashchange", route);
  route();
})();
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------
# Main program
# ---------------------------------------------------------------------
def main():
    check_settings()
    started = datetime.now()
    process_date = parse_process_date(PROCESS_DATE)
    excluded = set(EXCLUDED_ALERT_IDS)
    alert_ids = [
        i for i in range(ALERT_ID_START, ALERT_ID_END + 1) if i not in excluded
    ]

    print(f"Comparing {MAIN_TABLE}  vs  {TEST_TABLE}")
    print(f"Vendor {VENDOR_ID}, date {process_date}, {len(alert_ids)} alert IDs\n")

    conn = connect()

    summaries, all_cells, all_missing, all_columns = [], [], [], []
    details = {}
    for n, alert_id in enumerate(alert_ids, start=1):
        try:
            main_df = run_query(conn, MAIN_TABLE, alert_id, process_date)
            test_df = run_query(conn, TEST_TABLE, alert_id, process_date)
            summary, cells, missing, columns, detail = compare_one_alert(
                alert_id, main_df, test_df
            )
        except Exception as err:
            summary, cells, missing, columns, detail = (
                blank_summary(alert_id),
                [],
                [],
                [],
                None,
            )
            summary["Status"] = "ERROR"
            summary["Error"] = str(err)

        summaries.append(summary)
        details[alert_id] = detail
        all_cells.extend(cells)
        all_missing.extend(missing)
        all_columns.extend(columns)

        line = f"[{n:>3}/{len(alert_ids)}] Alert {alert_id:>3}: {summary['Status']:<8}"
        if summary["Status"] == "ERROR":
            line += f"  {summary['Error']}"
        else:
            line += (
                f"  rows {summary[f'Rows in {MAIN_TABLE}']} vs {summary[f'Rows in {TEST_TABLE}']}"
                f", {summary['Different cells']} different cells"
                f", {summary[f'Only in {MAIN_TABLE}'] + summary[f'Only in {TEST_TABLE}']} missing rows"
            )
        print(line)

    conn.close()

    print("\nWriting report ...")
    base = report_base_path(process_date, started)
    html_path = write_html_report(summaries, details, process_date, started, base)
    excel_path = None
    if WRITE_EXCEL_REPORT:
        excel_path = write_excel_report(
            summaries,
            all_cells,
            all_missing,
            all_columns,
            process_date,
            alert_ids,
            started,
            base,
        )

    statuses = [s["Status"] for s in summaries]
    print("\n================ RESULT ================")
    print(f"  MATCH    : {statuses.count('MATCH')}")
    print(f"  MISMATCH : {statuses.count('MISMATCH')}")
    print(f"  NO DATA  : {statuses.count('NO DATA')}  (no rows in either table)")
    print(f"  ERROR    : {statuses.count('ERROR')}")
    print(f"\nHTML report saved to:\n  {html_path}")
    if excel_path:
        print(f"Excel report saved to:\n  {excel_path}")
    if OPEN_REPORT_WHEN_DONE:
        webbrowser.open(Path(html_path).resolve().as_uri())


if __name__ == "__main__":
    main()
