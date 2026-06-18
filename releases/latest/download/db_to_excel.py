#!/usr/bin/env python3
"""
db_to_excel.py

Interactive CLI tool to export an entire database to a single Excel file
(one sheet per table), with auto-sized columns.

Supports: MySQL, PostgreSQL, SQLite, Microsoft SQL Server, Oracle.

Usage:
    dexcel

You will be prompted for the connection type and connection details.
Password input is hidden.

All activity is logged to logs/dexcel_<timestamp>.log for troubleshooting.
"""
import sys
import os
import re
import getpass
import logging
import traceback
from datetime import datetime

try:
    import pandas as pd
    from openpyxl.utils import get_column_letter
except ImportError as e:
    print(f"FATAL: required package not found ({e}).")
    print("Reinstall Dexcel and try again.")
    sys.exit(1)


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


def redact(value):
    return "*" * len(value) if value else ""


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

    try:
        if driver == "mysql":
            cursor.execute("SHOW TABLES")

        elif driver == "postgresql":
            cursor.execute(
                "SELECT table_name FROM information_schema.tables "
                "WHERE table_schema = 'public' "
                "AND table_type = 'BASE TABLE'"
            )

        elif driver == "sqlite":
            cursor.execute(
                "SELECT name FROM sqlite_master "
                "WHERE type = 'table' "
                "AND name NOT LIKE 'sqlite_%'"
            )

        elif driver == "mssql":
            cursor.execute(
                "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
                "WHERE TABLE_TYPE = 'BASE TABLE'"
            )

        elif driver == "oracle":
            cursor.execute("SELECT table_name FROM user_tables")

        else:
            raise ValueError(f"Unsupported driver: {driver}")

        tables = [row[0] for row in cursor.fetchall()]
        logger.info("Found %d table(s): %s", len(tables), tables)

        return tables

    except Exception as e:
        logger.critical("Failed to list tables: %s", e)
        raise


def quote_identifier(name, driver):
    if driver == "mysql":
        return f"`{name}`"

    if driver in ("postgresql", "sqlite", "oracle"):
        return f'"{name}"'

    if driver == "mssql":
        return f"[{name}]"

    return name


def sanitize_sheet_name(name, used_names):
    sheet_name = str(name)[:31]

    for bad_char in ["\\", "/", "*", "?", ":", "[", "]"]:
        sheet_name = sheet_name.replace(bad_char, "_")

    if not sheet_name:
        sheet_name = "sheet"

    original = sheet_name
    counter = 1

    while sheet_name in used_names:
        suffix = f"_{counter}"
        sheet_name = original[: 31 - len(suffix)] + suffix
        counter += 1

    used_names.add(sheet_name)

    return sheet_name

ILLEGAL_EXCEL_CHARS_RE = re.compile(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]")


def clean_excel_value(value):
    """
    Excel cannot store some control characters.
    Laravel/PHP serialized objects can contain null bytes, especially for
    protected/private properties. Remove those characters before writing.
    """
    if isinstance(value, str):
        return ILLEGAL_EXCEL_CHARS_RE.sub("", value)

    return value


def clean_dataframe_for_excel(df):
    """
    Clean only object/string columns before writing to Excel.
    """
    object_columns = df.select_dtypes(include=["object"]).columns

    for column in object_columns:
        df[column] = df[column].map(clean_excel_value)

    return df

def export_to_excel(conn, driver, tables, output_file):
    used_sheet_names = set()
    exported_count = 0
    skipped_tables = []

    base, ext = os.path.splitext(output_file)

    if not ext:
        ext = ".xlsx"

    tmp_output = f"{base}.tmp{ext}"

    try:
        with pd.ExcelWriter(tmp_output, engine="openpyxl") as writer:
            for table in tables:
                print(f"Exporting table: {table}")
                logger.info("Exporting table: %s", table)

                quoted = quote_identifier(table, driver)
                query = f"SELECT * FROM {quoted}"

                try:
                    df = pd.read_sql(query, conn)
                    df = clean_dataframe_for_excel(df)
                except Exception as e:
                    logger.error("Failed to read/clean table '%s': %s", table, e)
                    print(f"  Skipped '{table}' due to error: {e}")
                    skipped_tables.append(table)
                    continue

                sheet_name = sanitize_sheet_name(table, used_sheet_names)

                try:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                except Exception as e:
                    logger.error("Failed to write table '%s' to Excel: %s", table, e)
                    print(f"  Skipped '{table}' due to Excel write error: {e}")
                    skipped_tables.append(table)
                    continue
                worksheet = writer.sheets[sheet_name]

                for column_cells in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column_cells[0].column)

                    for cell in column_cells:
                        try:
                            if cell.value is not None:
                                max_length = max(max_length, len(str(cell.value)))
                        except Exception:
                            pass

                    worksheet.column_dimensions[column_letter].width = min(
                        max_length + 2,
                        50,
                    )

                exported_count += 1

        if exported_count == 0:
            logger.error("No tables were exported successfully.")
            os.remove(tmp_output)
            print("\nNo tables could be exported. No output file was created.")
            sys.exit(1)

        os.replace(tmp_output, output_file)

        logger.info(
            "Export finished. %d exported, %d skipped: %s",
            exported_count,
            len(skipped_tables),
            skipped_tables,
        )

    except Exception:
        logger.critical("Fatal error during export:\n%s", traceback.format_exc())

        if os.path.exists(tmp_output):
            os.remove(tmp_output)

        raise

    return exported_count, skipped_tables


def main():
    print("=== Dexcel: Database to Excel Exporter ===")
    logger.info("=== Dexcel session started ===")

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

        output_file = ask("Output Excel filename", default=f"{db_name}_export.xlsx")

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

        try:
            exported_count, skipped_tables = export_to_excel(
                conn,
                driver,
                tables,
                output_file,
            )
        except Exception as e:
            print(f"\nExport failed: {e}")
            print(f"Details were written to: {LOG_FILE}")
            sys.exit(1)

        print(f"\nDone. Exported {exported_count} table(s) to {output_file}")

        if skipped_tables:
            print(f"Skipped {len(skipped_tables)} table(s): {', '.join(skipped_tables)}")

        logger.info("=== Dexcel session finished successfully ===")

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