#!/usr/bin/env python3
"""
Dexcel — Database Schema Description Exporter
"""

import os
import sys
import platform
import subprocess
import logging
import traceback
from datetime import datetime
from typing import Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from textual.app import App, ComposeResult
from textual.containers import Horizontal, Vertical
from textual.screen import Screen
from textual.widgets import (
    Button, Input, Label,
    RadioSet, RadioButton, RichLog, Static,
)
from textual import work


# ── Logging ──────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(SCRIPT_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
LOG_FILE = os.path.join(LOG_DIR, f"dexcel_{_timestamp}.log")

logger = logging.getLogger("dexcel")
logger.setLevel(logging.DEBUG)
_fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
_fh.setLevel(logging.DEBUG)
_fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(_fh)
_ch = logging.StreamHandler()
_ch.setLevel(logging.WARNING)
_ch.setFormatter(logging.Formatter("%(message)s"))
logger.addHandler(_ch)


# ── Constants ────────────────────────────────────────────────────────

DB_TYPES_ORDERED = [
    ("mysql", "MySQL / MariaDB"),
    ("postgresql", "PostgreSQL"),
    ("sqlite", "SQLite"),
    ("mssql", "Microsoft SQL Server"),
    ("oracle", "Oracle"),
]

DEFAULT_PORTS = {
    "mysql": 3306,
    "postgresql": 5432,
    "mssql": 1433,
    "oracle": 1521,
}

HEADERS = ["Field", "Type", "Null", "Key", "Default", "Extra"]


# ── Helpers ──────────────────────────────────────────────────────────

def open_file_location(filepath: str) -> None:
    filepath = os.path.abspath(filepath)
    system = platform.system()
    try:
        if system == "Darwin":
            subprocess.Popen(["open", "-R", filepath])
        elif system == "Windows":
            subprocess.Popen(["explorer", "/select,", os.path.normpath(filepath)])
        else:
            subprocess.Popen(["xdg-open", os.path.dirname(filepath)])
    except Exception as e:
        logger.error("Failed to open file location: %s", e)


# ── Database Functions ──────────────────────────────────────────────

def build_connection(driver: str, **params):
    if driver == "sqlite":
        path = params["path"]
        if not os.path.isfile(path):
            raise FileNotFoundError(f"SQLite file not found: {path}")
        import sqlite3
        conn = sqlite3.connect(path)
        db_name = os.path.splitext(os.path.basename(path))[0]
        return conn, db_name

    host = params.get("host", "127.0.0.1")
    port = params.get("port", str(DEFAULT_PORTS.get(driver, "")))
    user = params.get("user", "")
    password = params.get("password", "")
    database = params.get("database", "")

    logger.info(
        "Connecting: driver=%s host=%s port=%s user=%s db=%s",
        driver, host, port, user, database,
    )

    if driver == "mysql":
        import pymysql
        conn = pymysql.connect(
            host=host, port=int(port), user=user, password=password,
            database=database, charset="utf8mb4", connect_timeout=10,
        )
    elif driver == "postgresql":
        import psycopg2
        conn = psycopg2.connect(
            host=host, port=port, user=user, password=password,
            dbname=database, connect_timeout=10,
        )
    elif driver == "mssql":
        try:
            import pyodbc
        except ImportError:
            raise ImportError(
                "SQL Server driver not available.\n"
                "Install Microsoft ODBC Driver 17 for SQL Server."
            )
        conn_str = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            f"SERVER={host},{port};DATABASE={database};"
            f"UID={user};PWD={password}"
        )
        conn = pyodbc.connect(conn_str, timeout=10)
    elif driver == "oracle":
        import oracledb
        dsn = oracledb.makedsn(host, port, service_name=database)
        conn = oracledb.connect(user=user, password=password, dsn=dsn)
    else:
        raise ValueError(f"Unsupported driver: {driver}")

    return conn, database


def list_tables(conn, driver: str):
    cursor = conn.cursor()
    if driver == "mysql":
        cursor.execute("SHOW FULL TABLES WHERE Table_type = 'BASE TABLE'")
        return [row[0] for row in cursor.fetchall()]
    if driver == "postgresql":
        cursor.execute(
            "SELECT table_name FROM information_schema.tables "
            "WHERE table_schema = 'public' AND table_type = 'BASE TABLE' "
            "ORDER BY table_name"
        )
        return [row[0] for row in cursor.fetchall()]
    if driver == "sqlite":
        cursor.execute(
            "SELECT name FROM sqlite_master "
            "WHERE type = 'table' AND name NOT LIKE 'sqlite_%' "
            "ORDER BY name"
        )
        return [row[0] for row in cursor.fetchall()]
    if driver == "mssql":
        cursor.execute(
            "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
            "WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME"
        )
        return [row[0] for row in cursor.fetchall()]
    if driver == "oracle":
        cursor.execute("SELECT table_name FROM user_tables ORDER BY table_name")
        return [row[0] for row in cursor.fetchall()]
    raise ValueError(f"Unsupported driver: {driver}")


def get_table_description(conn, driver: str, table: str):
    cursor = conn.cursor()

    if driver == "mysql":
        cursor.execute(f"DESCRIBE `{table}`")
        return [
            {"Field": r[0], "Type": r[1], "Null": r[2], "Key": r[3],
             "Default": "NULL" if r[4] is None else r[4], "Extra": r[5]}
            for r in cursor.fetchall()
        ]

    if driver == "postgresql":
        cursor.execute("""
            SELECT c.column_name, c.data_type, c.is_nullable,
                   CASE WHEN pk.column_name IS NOT NULL THEN 'PRI' ELSE '' END,
                   c.column_default,
                   CASE WHEN c.column_default LIKE 'nextval%%' THEN 'auto_increment' ELSE '' END
            FROM information_schema.columns c
            LEFT JOIN (
                SELECT ku.table_name, ku.column_name
                FROM information_schema.table_constraints tc
                JOIN information_schema.key_column_usage ku
                    ON tc.constraint_name = ku.constraint_name
                WHERE tc.constraint_type = 'PRIMARY KEY' AND ku.table_schema = 'public'
            ) pk ON c.table_name = pk.table_name AND c.column_name = pk.column_name
            WHERE c.table_schema = 'public' AND c.table_name = %s
            ORDER BY c.ordinal_position
        """, (table,))
        return [
            {"Field": r[0], "Type": r[1], "Null": "YES" if r[2] == "YES" else "NO",
             "Key": r[3] or "", "Default": "NULL" if r[4] is None else r[4],
             "Extra": r[5] or ""}
            for r in cursor.fetchall()
        ]

    if driver == "sqlite":
        cursor.execute(f'PRAGMA table_info("{table}")')
        return [
            {"Field": r[1], "Type": r[2], "Null": "NO" if r[3] else "YES",
             "Key": "PRI" if r[5] else "", "Default": "NULL" if r[4] is None else r[4],
             "Extra": ""}
            for r in cursor.fetchall()
        ]

    if driver == "mssql":
        cursor.execute("""
            SELECT c.COLUMN_NAME,
                   c.DATA_TYPE + CASE WHEN c.CHARACTER_MAXIMUM_LENGTH IS NOT NULL
                       THEN '(' + CAST(c.CHARACTER_MAXIMUM_LENGTH AS VARCHAR) + ')' ELSE '' END,
                   c.IS_NULLABLE,
                   CASE WHEN pk.COLUMN_NAME IS NOT NULL THEN 'PRI' ELSE '' END,
                   c.COLUMN_DEFAULT,
                   CASE WHEN COLUMNPROPERTY(OBJECT_ID(c.TABLE_SCHEMA+'.'+c.TABLE_NAME),
                       c.COLUMN_NAME, 'IsIdentity') = 1 THEN 'auto_increment' ELSE '' END
            FROM INFORMATION_SCHEMA.COLUMNS c
            LEFT JOIN (
                SELECT ku.TABLE_NAME, ku.COLUMN_NAME
                FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc
                JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE ku
                    ON tc.CONSTRAINT_NAME = ku.CONSTRAINT_NAME
                WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY'
            ) pk ON c.TABLE_NAME = pk.TABLE_NAME AND c.COLUMN_NAME = pk.COLUMN_NAME
            WHERE c.TABLE_NAME = ?
            ORDER BY c.ORDINAL_POSITION
        """, table)
        return [
            {"Field": r[0], "Type": r[1], "Null": "YES" if r[2] == "YES" else "NO",
             "Key": r[3] or "", "Default": "NULL" if r[4] is None else r[4],
             "Extra": r[5] or ""}
            for r in cursor.fetchall()
        ]

    if driver == "oracle":
        cursor.execute("""
            SELECT c.column_name,
                   c.data_type || CASE
                       WHEN c.data_type IN ('VARCHAR2','CHAR','NVARCHAR2','NCHAR')
                           THEN '(' || c.data_length || ')'
                       WHEN c.data_type = 'NUMBER' AND c.data_precision IS NOT NULL
                           THEN '(' || c.data_precision || ',' || c.data_scale || ')'
                       ELSE '' END,
                   c.nullable,
                   CASE WHEN pk.column_name IS NOT NULL THEN 'PRI' ELSE '' END,
                   c.data_default, ''
            FROM user_tab_columns c
            LEFT JOIN (
                SELECT cols.table_name, cols.column_name
                FROM user_constraints cons
                JOIN user_cons_columns cols
                    ON cons.constraint_name = cols.constraint_name
                WHERE cons.constraint_type = 'P'
            ) pk ON c.table_name = pk.table_name AND c.column_name = pk.column_name
            WHERE c.table_name = :table_name
            ORDER BY c.column_id
        """, table_name=table.upper())
        return [
            {"Field": r[0], "Type": r[1],
             "Null": "YES" if r[2] == "Y" else "NO",
             "Key": r[3] or "",
             "Default": "NULL" if r[4] is None else str(r[4]).strip(),
             "Extra": r[5] or ""}
            for r in cursor.fetchall()
        ]

    raise ValueError(f"Unsupported driver: {driver}")


def export_to_excel(conn, driver: str, tables, output_path: str,
                    on_progress=None) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Table Descriptions"

    title_fill = PatternFill("solid", fgColor="1F4E78")
    title_font = Font(color="FFFFFF", bold=True, size=12)
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_al = Alignment(horizontal="center", vertical="center")
    left_al = Alignment(horizontal="left", vertical="top", wrap_text=True)

    row_num = 1
    for table in tables:
        if on_progress:
            on_progress(f"Describing table: {table}")

        try:
            rows = get_table_description(conn, driver, table)
        except Exception as e:
            ws.merge_cells(
                start_row=row_num, start_column=1,
                end_row=row_num, end_column=len(HEADERS),
            )
            cell = ws.cell(row=row_num, column=1, value=f"TABLE: {table}")
            cell.fill = title_fill
            cell.font = title_font
            cell.alignment = left_al
            row_num += 1
            ws.cell(row=row_num, column=1, value=f"ERROR: {e}")
            row_num += 2
            continue

        ws.merge_cells(
            start_row=row_num, start_column=1,
            end_row=row_num, end_column=len(HEADERS),
        )
        tc = ws.cell(row=row_num, column=1, value=f"TABLE: {table}")
        tc.fill = title_fill
        tc.font = title_font
        tc.alignment = left_al
        tc.border = border
        for c in range(1, len(HEADERS) + 1):
            ws.cell(row=row_num, column=c).border = border
        row_num += 1

        for cn, h in enumerate(HEADERS, 1):
            cell = ws.cell(row=row_num, column=cn, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
            cell.alignment = center_al
        row_num += 1

        for row in rows:
            for cn, h in enumerate(HEADERS, 1):
                v = row.get(h, "") or "NULL"
                cell = ws.cell(row=row_num, column=cn, value=str(v))
                cell.border = border
                cell.alignment = left_al
            row_num += 1
        row_num += 1

    for cc in ws.columns:
        ml = max((len(str(c.value or "")) for c in cc), default=0)
        ws.column_dimensions[get_column_letter(cc[0].column)].width = min(ml + 2, 60)
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 25
    ws.column_dimensions["F"].width = 25

    wb.save(output_path)
    if on_progress:
        on_progress(f"Saved to: {output_path}")


# ── Screens ─────────────────────────────────────────────────────────

class MainScreen(Screen):
    def compose(self) -> ComposeResult:
        yield Static("Dexcel — Database Schema Exporter", id="titlebar")
        with Vertical():
            yield Label("Database Reader", id="screen-title")
            yield Label("Select a database type to describe.")
            with RadioSet(id="db-type"):
                for key, name in DB_TYPES_ORDERED:
                    yield RadioButton(f"{name}", id=key)
            with Horizontal():
                yield Button("Next", variant="primary", id="btn-next", disabled=True)

    def on_radio_set_changed(self, event: RadioSet.Changed) -> None:
        self.app.driver = event.pressed.id
        self.query_one("#btn-next", Button).disabled = False

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-next":
            d = self.app.driver
            if d == "sqlite":
                self.app.push_screen(SQLiteScreen())
            else:
                self.app.push_screen(NetScreen())


class SQLiteScreen(Screen):
    def compose(self) -> ComposeResult:
        yield Static("Dexcel — Database Schema Exporter", id="titlebar")
        with Vertical():
            yield Label("SQLite Database File", id="screen-title")
            yield Label("Path")
            yield Input(placeholder="Path to .db or .sqlite file", id="db-path")
            yield Label("", id="sqlite-error")
            with Horizontal():
                yield Button("Cancel", id="back")
                yield Button("Approve Describe", variant="primary", id="connect")

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "back":
            self.app.pop_screen()
        elif event.button.id == "connect":
            path = self.query_one("#db-path", Input).value.strip()
            if not path:
                self.query_one("#sqlite-error", Label).update("Enter a file path.")
                return
            if not os.path.isfile(path):
                self.query_one("#sqlite-error", Label).update(f"File not found: {path}")
                return
            self.app.db_params = {"path": path}
            self.app.push_screen(ExportScreen())


class NetScreen(Screen):
    def compose(self) -> ComposeResult:
        yield Static("Dexcel — Database Schema Exporter", id="titlebar")
        with Vertical():
            yield Label("Database Connection", id="screen-title")
            yield Label("Host")
            yield Input(placeholder="Host", id="host")
            yield Label("Port")
            yield Input(placeholder="Port", id="port")
            yield Label("Username")
            yield Input(placeholder="Username", id="username")
            yield Label("Password")
            yield Input(placeholder="Password", id="password", password=True)
            yield Label("Dataname")
            yield Input(placeholder="Dataname / Database name", id="database")
            yield Label("", id="net-error")
            with Horizontal():
                yield Button("Cancel", id="back")
                yield Button("Approve Describe", variant="primary", id="connect")

    def on_mount(self) -> None:
        driver = self.app.driver
        self.query_one("#host", Input).value = "127.0.0.1"
        if driver == "mysql":
            self.query_one("#username", Input).value = "root"
        port = DEFAULT_PORTS.get(driver)
        if port:
            self.query_one("#port", Input).value = str(port)

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "back":
            self.app.pop_screen()
        elif event.button.id == "connect":
            fields = {
                "host": self.query_one("#host", Input).value.strip(),
                "port": self.query_one("#port", Input).value.strip(),
                "user": self.query_one("#username", Input).value.strip(),
                "password": self.query_one("#password", Input).value,
                "database": self.query_one("#database", Input).value.strip(),
            }
            missing = [k for k, v in fields.items() if not v]
            if missing:
                self.query_one("#net-error", Label).update(
                    f"Missing required: {', '.join(missing)}"
                )
                return
            self.app.db_params = fields
            self.app.push_screen(ExportScreen())


class ExportScreen(Screen):
    def compose(self) -> ComposeResult:
        yield Static("Dexcel — Database Schema Exporter", id="titlebar")
        with Vertical():
            yield Label("Export Progress")
            yield RichLog(id="log", highlight=True, markup=True)
            yield Label("Saved Excel file")
            yield Button("Waiting for export...", id="file-link", disabled=True)
            with Horizontal():
                yield Button("Cancel", id="cancel")
                yield Button("Exit", variant="primary", id="exit", disabled=True)

    def on_mount(self) -> None:
        self.run_export()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "cancel":
            self.app.pop_screen()
        elif event.button.id == "exit":
            self.app.exit()
        elif event.button.id == "file-link":
            if self.app.output_file:
                open_file_location(self.app.output_file)

    @work(thread=True, exit_on_error=False)
    def run_export(self) -> None:
        def _log(msg: str) -> None:
            self.call_from_thread(self._append_log, msg)

        try:
            _log("Connecting to database...")
            conn, db_name = build_connection(self.app.driver, **self.app.db_params)
            self.app.conn = conn
            self.app.db_name = db_name
            _log("Connected successfully.")

            _log("Listing tables...")
            tables = list_tables(conn, self.app.driver)
            _log(f"Found {len(tables)} table(s).")

            if not tables:
                _log("No tables found in this database.")
                self.call_from_thread(self._finish, False)
                return

            output_file = f"{db_name}_table_descriptions.xlsx"
            output_path = os.path.abspath(output_file)
            self.app.output_file = output_path

            if os.path.exists(output_path):
                _log("File exists -- will overwrite")

            _log(f"Output: {output_path}")

            def on_progress(msg: str) -> None:
                self.call_from_thread(self._append_log, f"  {msg}")

            export_to_excel(conn, self.app.driver, tables, output_path, on_progress)

            _log("Export complete!")
            self.call_from_thread(self._finish, True, output_path)

        except Exception as e:
            _log(f"Error: {e}")
            logger.error("Export failed: %s", traceback.format_exc())
            self.call_from_thread(self._finish, False, error=str(e))
        finally:
            try:
                if hasattr(self.app, "conn") and self.app.conn:
                    self.app.conn.close()
            except Exception:
                pass

    def _append_log(self, message: str) -> None:
        self.query_one("#log", RichLog).write(message)

    def _finish(self, success: bool, file_path: Optional[str] = None,
                error: Optional[str] = None) -> None:
        self.query_one("#cancel", Button).disabled = True

        if success and file_path:
            self.query_one("#exit", Button).disabled = False
            file_link = self.query_one("#file-link", Button)
            file_link.disabled = False
            file_link.label = file_path
        elif error:
            self.query_one("#exit", Button).disabled = False
            file_link = self.query_one("#file-link", Button)
            file_link.label = "Export failed. Check the log above."
            file_link.disabled = True
        else:
            self.query_one("#exit", Button).disabled = False
            file_link = self.query_one("#file-link", Button)
            file_link.label = "Waiting for export..."
            file_link.disabled = True


# ── App ──────────────────────────────────────────────────────────────

class DexcelApp(App):
    """Dexcel — Database Schema Exporter"""

    CSS = """
    Screen {
        background: #191d28;
    }

    #titlebar {
        background: #7c5cff;
        color: #ffffff;
        text-style: bold;
        height: 1;
        content-align: center middle;
        width: 100%;
    }

    Vertical {
        align: center top;
        width: 76;
        height: auto;
        margin: 1 2;
        padding: 1 2;
        border: round #394056;
        background: #222836;
    }

    Input {
        margin: 0 0 1 0;
        width: 100%;
    }

    #screen-title {
        text-style: bold;
        color: #f2f4ff;
        margin: 0 0 1 0;
    }

    RichLog {
        height: 60%;
        min-height: 10;
        margin: 0 0 1 0;
        border: round #394056;
        background: #111521;
    }

    Horizontal {
        align: center middle;
        height: auto;
        margin: 1 0 0 0;
    }

    Horizontal Button {
        margin: 0 1;
        min-width: 14;
    }

    #file-link {
        width: 100%;
    }

    #net-error,
    #sqlite-error {
        color: #ff8a8a;
        margin: 0 0 1 0;
    }
    """

    def __init__(self):
        super().__init__()
        self.driver = None
        self.db_params = {}
        self.conn = None
        self.db_name = None
        self.output_file = None

    def on_mount(self) -> None:
        self.push_screen(MainScreen())


def main():
    app = DexcelApp()
    app.run()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except Exception:
        logger.critical("Unhandled exception:\n%s", traceback.format_exc())
        print(f"\nAn unexpected error occurred. Details: {LOG_FILE}")
        sys.exit(1)
