#!/usr/bin/env python3
"""
Dexcel — Database Schema Description Exporter

Exports an entire database schema into a text file like:

+--------------+
| TABLE: users |
+--------------+
+------------+--------------+------+-----+---------+----------------+
| Field      | Type         | Null | Key | Default | Extra          |
+------------+--------------+------+-----+---------+----------------+

Supports:
- MySQL / MariaDB
- PostgreSQL
- SQLite
- Microsoft SQL Server
- Oracle

Usage:
    dexcel
"""

import sys
import os
import getpass
import logging
import traceback
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(SCRIPT_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = os.path.join(LOG_DIR, f"dexcel_{_timestamp}.log")

logger = logging.getLogger("dexcel")
logger.setLevel(logging.DEBUG)

_file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
_file_handler.setLevel(logging.DEBUG)
_file_handler.setFormatter(
    logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
)
logger.addHandler(_file_handler)

_console_handler = logging.StreamHandler()
_console_handler.setLevel(logging.WARNING)
_console_handler.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(_console_handler)


DB_TYPES = {
    "1": {"name": "MySQL / MariaDB", "driver": "mysql"},
    "2": {"name": "PostgreSQL", "driver": "postgresql"},
    "3": {"name": "SQLite", "driver": "sqlite"},
    "4": {"name": "Microsoft SQL Server", "driver": "mssql"},
    "5": {"name": "Oracle", "driver": "oracle"},
}

DEFAULT_PORTS = {
    "mysql": 3306,
    "postgresql": 5432,
    "mssql": 1433,
    "oracle": 1521,
}

MAX_RETRIES = 3


def redact(value):
    return "*" * len(value) if value else ""


def ask(prompt, default=None, required=True):
    label = f"{prompt} [{default}]: " if default is not None else f"{prompt}: "

    while True:
        try:
            value = input(label).strip()
        except EOFError:
            logger.error("Input stream closed unexpectedly while prompting: %s", prompt)
            sys.exit(1)

        if not value and default is not None:
            return default

        if not value and not required:
            return ""

        if value:
            return value

        print("This value is required.")


def choose_db_type():
    print("\nSelect database type:")

    for key, info in DB_TYPES.items():
        print(f"  {key}. {info['name']}")

    while True:
        choice = input("Enter choice number: ").strip()

        if choice in DB_TYPES:
            logger.info("User selected database type: %s", DB_TYPES[choice]["name"])
            return DB_TYPES[choice]

        print("Invalid choice, try again.")


def connect_with_retry(connect_fn, max_retries=MAX_RETRIES):
    last_error = None

    for attempt in range(1, max_retries + 1):
        try:
            conn = connect_fn()
            logger.info("Connection succeeded on attempt %d", attempt)
            return conn
        except Exception as e:
            last_error = e
            logger.error("Connection attempt %d failed: %s", attempt, e)
            print(f"Connection failed: {e}")

            if attempt < max_retries:
                retry = ask("Try again? (y/n)", default="y")

                if retry.lower() not in ("y", "yes"):
                    break

    logger.critical("All connection attempts exhausted. Last error: %s", last_error)
    print("\nCould not connect after multiple attempts. Exiting.")
    print(f"Details were written to: {LOG_FILE}")
    sys.exit(1)


def build_connection(db_info):
    driver = db_info["driver"]

    if driver == "sqlite":
        path = ask("Path to SQLite .db/.sqlite file")

        if not os.path.isfile(path):
            logger.error("SQLite file not found: %s", path)
            print(f"File not found: {path}")
            sys.exit(1)

        import sqlite3

        def _connect():
            return sqlite3.connect(path)

        conn = connect_with_retry(_connect)
        db_name = os.path.splitext(os.path.basename(path))[0]

        return conn, db_name, driver

    default_port = DEFAULT_PORTS.get(driver, "")
    host = ask("Host", default="127.0.0.1")
    port = ask("Port", default=str(default_port) if default_port else None)
    user = ask("Username", default="root" if driver == "mysql" else None)
    password = getpass.getpass("Password (input hidden): ")
    database = ask("Database name")

    logger.info(
        "Connecting: driver=%s host=%s port=%s user=%s db=%s password=%s",
        driver,
        host,
        port,
        user,
        database,
        redact(password),
    )

    if driver == "mysql":
        import pymysql

        def _connect():
            return pymysql.connect(
                host=host,
                port=int(port),
                user=user,
                password=password,
                database=database,
                charset="utf8mb4",
                connect_timeout=10,
            )

    elif driver == "postgresql":
        import psycopg2

        def _connect():
            return psycopg2.connect(
                host=host,
                port=port,
                user=user,
                password=password,
                dbname=database,
                connect_timeout=10,
            )

    elif driver == "mssql":
        try:
            import pyodbc
        except ImportError as e:
            logger.critical("pyodbc import failed: %s", e)
            print(
                "\nThe SQL Server driver could not be loaded.\n"
                "Install Microsoft ODBC Driver 17 for SQL Server, then try again.\n"
            )
            sys.exit(1)

        def _connect():
            conn_str = (
                "DRIVER={ODBC Driver 17 for SQL Server};"
                f"SERVER={host},{port};"
                f"DATABASE={database};"
                f"UID={user};"
                f"PWD={password}"
            )

            return pyodbc.connect(conn_str, timeout=10)

    elif driver == "oracle":
        import oracledb

        def _connect():
            dsn = oracledb.makedsn(host, port, service_name=database)
            return oracledb.connect(user=user, password=password, dsn=dsn)

    else:
        logger.critical("Unsupported driver requested: %s", driver)
        sys.exit(f"Unsupported driver: {driver}")

    conn = connect_with_retry(_connect)

    return conn, database, driver


def list_tables(conn, driver):
    cursor = conn.cursor()

    if driver == "mysql":
        cursor.execute("SHOW FULL TABLES WHERE Table_type = 'BASE TABLE'")
        return [row[0] for row in cursor.fetchall()]

    if driver == "postgresql":
        cursor.execute(
            """
            SELECT table_name
            FROM information_schema.tables
            WHERE table_schema = 'public'
            AND table_type = 'BASE TABLE'
            ORDER BY table_name
            """
        )
        return [row[0] for row in cursor.fetchall()]

    if driver == "sqlite":
        cursor.execute(
            """
            SELECT name
            FROM sqlite_master
            WHERE type = 'table'
            AND name NOT LIKE 'sqlite_%'
            ORDER BY name
            """
        )
        return [row[0] for row in cursor.fetchall()]

    if driver == "mssql":
        cursor.execute(
            """
            SELECT TABLE_NAME
            FROM INFORMATION_SCHEMA.TABLES
            WHERE TABLE_TYPE = 'BASE TABLE'
            ORDER BY TABLE_NAME
            """
        )
        return [row[0] for row in cursor.fetchall()]

    if driver == "oracle":
        cursor.execute("SELECT table_name FROM user_tables ORDER BY table_name")
        return [row[0] for row in cursor.fetchall()]

    raise ValueError(f"Unsupported driver: {driver}")


def get_table_description(conn, driver, table):
    """
    Return rows in this common format:

    [
        {
            "Field": "...",
            "Type": "...",
            "Null": "...",
            "Key": "...",
            "Default": "...",
            "Extra": "...",
        }
    ]
    """
    cursor = conn.cursor()

    if driver == "mysql":
        cursor.execute(f"DESCRIBE `{table}`")
        rows = cursor.fetchall()

        return [
            {
                "Field": row[0],
                "Type": row[1],
                "Null": row[2],
                "Key": row[3],
                "Default": "NULL" if row[4] is None else row[4],
                "Extra": row[5],
            }
            for row in rows
        ]

    if driver == "postgresql":
        cursor.execute(
            """
            SELECT
                c.column_name,
                c.data_type,
                c.is_nullable,
                CASE
                    WHEN pk.column_name IS NOT NULL THEN 'PRI'
                    ELSE ''
                END AS column_key,
                c.column_default,
                CASE
                    WHEN c.column_default LIKE 'nextval%%' THEN 'auto_increment'
                    ELSE ''
                END AS extra
            FROM information_schema.columns c
            LEFT JOIN (
                SELECT ku.table_name, ku.column_name
                FROM information_schema.table_constraints tc
                JOIN information_schema.key_column_usage ku
                    ON tc.constraint_name = ku.constraint_name
                WHERE tc.constraint_type = 'PRIMARY KEY'
                  AND ku.table_schema = 'public'
            ) pk
              ON c.table_name = pk.table_name
             AND c.column_name = pk.column_name
            WHERE c.table_schema = 'public'
              AND c.table_name = %s
            ORDER BY c.ordinal_position
            """,
            (table,),
        )

        return [
            {
                "Field": row[0],
                "Type": row[1],
                "Null": "YES" if row[2] == "YES" else "NO",
                "Key": row[3] or "",
                "Default": "NULL" if row[4] is None else row[4],
                "Extra": row[5] or "",
            }
            for row in cursor.fetchall()
        ]

    if driver == "sqlite":
        cursor.execute(f'PRAGMA table_info("{table}")')
        rows = cursor.fetchall()

        return [
            {
                "Field": row[1],
                "Type": row[2],
                "Null": "NO" if row[3] else "YES",
                "Key": "PRI" if row[5] else "",
                "Default": "NULL" if row[4] is None else row[4],
                "Extra": "",
            }
            for row in rows
        ]

    if driver == "mssql":
        cursor.execute(
            """
            SELECT
                c.COLUMN_NAME,
                c.DATA_TYPE +
                    CASE
                        WHEN c.CHARACTER_MAXIMUM_LENGTH IS NOT NULL
                            THEN '(' + CAST(c.CHARACTER_MAXIMUM_LENGTH AS VARCHAR) + ')'
                        ELSE ''
                    END AS column_type,
                c.IS_NULLABLE,
                CASE
                    WHEN pk.COLUMN_NAME IS NOT NULL THEN 'PRI'
                    ELSE ''
                END AS column_key,
                c.COLUMN_DEFAULT,
                CASE
                    WHEN COLUMNPROPERTY(
                        OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME),
                        c.COLUMN_NAME,
                        'IsIdentity'
                    ) = 1 THEN 'auto_increment'
                    ELSE ''
                END AS extra
            FROM INFORMATION_SCHEMA.COLUMNS c
            LEFT JOIN (
                SELECT ku.TABLE_NAME, ku.COLUMN_NAME
                FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                    ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
                WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
            ) pk
              ON c.TABLE_NAME = pk.TABLE_NAME
             AND c.COLUMN_NAME = pk.COLUMN_NAME
            WHERE c.TABLE_NAME = ?
            ORDER BY c.ORDINAL_POSITION
            """,
            table,
        )

        return [
            {
                "Field": row[0],
                "Type": row[1],
                "Null": "YES" if row[2] == "YES" else "NO",
                "Key": row[3] or "",
                "Default": "NULL" if row[4] is None else row[4],
                "Extra": row[5] or "",
            }
            for row in cursor.fetchall()
        ]

    if driver == "oracle":
        cursor.execute(
            """
            SELECT
                c.column_name,
                c.data_type ||
                    CASE
                        WHEN c.data_type IN ('VARCHAR2', 'CHAR', 'NVARCHAR2', 'NCHAR')
                            THEN '(' || c.data_length || ')'
                        WHEN c.data_type = 'NUMBER' AND c.data_precision IS NOT NULL
                            THEN '(' || c.data_precision || ',' || c.data_scale || ')'
                        ELSE ''
                    END AS column_type,
                c.nullable,
                CASE
                    WHEN pk.column_name IS NOT NULL THEN 'PRI'
                    ELSE ''
                END AS column_key,
                c.data_default,
                ''
            FROM user_tab_columns c
            LEFT JOIN (
                SELECT cols.table_name, cols.column_name
                FROM user_constraints cons
                JOIN user_cons_columns cols
                    ON cons.constraint_name = cols.constraint_name
                WHERE cons.constraint_type = 'P'
            ) pk
              ON c.table_name = pk.table_name
             AND c.column_name = pk.column_name
            WHERE c.table_name = :table_name
            ORDER BY c.column_id
            """,
            table_name=table.upper(),
        )

        return [
            {
                "Field": row[0],
                "Type": row[1],
                "Null": "YES" if row[2] == "Y" else "NO",
                "Key": row[3] or "",
                "Default": "NULL" if row[4] is None else str(row[4]).strip(),
                "Extra": row[5] or "",
            }
            for row in cursor.fetchall()
        ]

    raise ValueError(f"Unsupported driver: {driver}")


def make_box_line(widths):
    return "+" + "+".join("-" * (width + 2) for width in widths) + "+"


def make_box_row(values, widths):
    cells = []

    for value, width in zip(values, widths):
        text = "" if value is None else str(value)
        cells.append(" " + text.ljust(width) + " ")

    return "|" + "|".join(cells) + "|"


def render_table_box(headers, rows):
    normalized_rows = []

    for row in rows:
        normalized_rows.append([row.get(header, "") for header in headers])

    widths = []

    for index, header in enumerate(headers):
        max_width = len(header)

        for row in normalized_rows:
            max_width = max(max_width, len(str(row[index])))

        widths.append(max_width)

    output = []
    line = make_box_line(widths)

    output.append(line)
    output.append(make_box_row(headers, widths))
    output.append(line)

    for row in normalized_rows:
        output.append(make_box_row(row, widths))

    output.append(line)

    return "\n".join(output)


def render_table_title(table_name):
    title = f"TABLE: {table_name}"
    width = len(title)

    return "\n".join(
        [
            "+" + "-" * (width + 2) + "+",
            "| " + title + " |",
            "+" + "-" * (width + 2) + "+",
        ]
    )


def export_schema_descriptions(conn, driver, tables, output_file):
    headers = ["Field", "Type", "Null", "Key", "Default", "Extra"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Table Descriptions"

    title_fill = PatternFill("solid", fgColor="1F4E78")
    title_font = Font(color="FFFFFF", bold=True, size=12)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)

    row_num = 1

    for table in tables:
        print(f"Describing table: {table}")
        logger.info("Describing table: %s", table)

        try:
            rows = get_table_description(conn, driver, table)
        except Exception as e:
            logger.error("Failed to describe table '%s': %s", table, e)

            ws.merge_cells(
                start_row=row_num,
                start_column=1,
                end_row=row_num,
                end_column=len(headers),
            )
            cell = ws.cell(row=row_num, column=1, value=f"TABLE: {table}")
            cell.fill = title_fill
            cell.font = title_font
            cell.alignment = left
            row_num += 1

            ws.cell(row=row_num, column=1, value=f"ERROR: Failed to describe table: {e}")
            row_num += 2
            continue

        # Table title row
        ws.merge_cells(
            start_row=row_num,
            start_column=1,
            end_row=row_num,
            end_column=len(headers),
        )
        title_cell = ws.cell(row=row_num, column=1, value=f"TABLE: {table}")
        title_cell.fill = title_fill
        title_cell.font = title_font
        title_cell.alignment = left
        title_cell.border = border

        for col in range(1, len(headers) + 1):
            ws.cell(row=row_num, column=col).border = border

        row_num += 1

        # Header row
        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=row_num, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center

        row_num += 1

        # Column rows
        for description in rows:
            for col_num, header in enumerate(headers, start=1):
                value = description.get(header, "")
                if value is None:
                    value = "NULL"

                cell = ws.cell(row=row_num, column=col_num, value=str(value))
                cell.border = border
                cell.alignment = left

            row_num += 1

        # Blank row between tables
        row_num += 1

    # Freeze top row? Not useful here because each table has its own header.
    # Instead, auto-size columns.
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)

        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = min(max_length + 2, 60)

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 25
    ws.column_dimensions["F"].width = 25

    wb.save(output_file)

    logger.info("Schema description Excel export finished: %s", output_file)


def main():
    print("=== Dexcel: Database Schema Description Exporter ===")
    logger.info("=== Dexcel schema export session started ===")

    db_info = choose_db_type()
    driver = db_info["driver"]

    conn, db_name, driver = build_connection(db_info)

    try:
        try:
            tables = list_tables(conn, driver)
        except Exception as e:
            print(f"\nCould not list tables: {e}")
            print(f"Details were written to: {LOG_FILE}")
            sys.exit(1)

        if not tables:
            print("No tables found in this database.")
            logger.warning("No tables found.")
            return

        print(f"\nFound {len(tables)} table(s): {', '.join(tables)}")

        output_file = ask(
            "Output Excel filename",
            default=f"{db_name}_table_descriptions.xlsx",
        )

        if not output_file.lower().endswith(".xlsx"):
            output_file += ".xlsx"

        if os.path.exists(output_file):
            overwrite = ask(
                f"'{output_file}' already exists. Overwrite? (y/n)",
                default="n",
            )

            if overwrite.lower() not in ("y", "yes"):
                output_file = ask("Enter a different filename")

                if not output_file.lower().endswith(".xlsx"):
                    output_file += ".xlsx"

        export_schema_descriptions(conn, driver, tables, output_file)

        print(f"\nDone. Exported schema descriptions for {len(tables)} table(s) to {output_file}")
        logger.info("=== Dexcel schema export session finished successfully ===")

    finally:
        try:
            conn.close()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.")
        logger.warning("Session cancelled by user.")
        sys.exit(1)
    except SystemExit:
        raise
    except Exception:
        logger.critical("Unhandled exception:\n%s", traceback.format_exc())
        print(f"\nAn unexpected error occurred. Details were written to: {LOG_FILE}")
        sys.exit(1)