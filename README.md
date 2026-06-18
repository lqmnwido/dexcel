# Dexcel

Dexcel exports an entire database schema into one Excel file.

It generates a table description report for every table in your database, including:

- Field
- Type
- Null
- Key
- Default
- Extra

The output is an `.xlsx` file, similar to MySQL `DESCRIBE` / table description output.

Supported databases:

- MySQL / MariaDB
- PostgreSQL
- SQLite
- Microsoft SQL Server
- Oracle

---

## Install

### macOS / Linux

```bash
curl -fsSL https://dexcel.kohich.site/install.sh | bash
```

### Windows

```powershell
powershell -ExecutionPolicy Bypass -Command "irm https://dexcel.kohich.site/install.ps1 | iex"
```

---

## Run

After installation, open a new terminal or command prompt and run:

```bash
dexcel
```

---

## Example

```txt
=== Dexcel: Database Schema Description Exporter ===

Select database type:
  1. MySQL / MariaDB
  2. PostgreSQL
  3. SQLite
  4. Microsoft SQL Server
  5. Oracle

Enter choice number:
```

Then enter your database connection details.

Dexcel will export your database table descriptions into an Excel file, for example:

```txt
mypelantikan_table_descriptions.xlsx
```

---

## What the installer does

The installer:

1. Detects or installs Python 3.8+
2. Creates an isolated environment inside `.dexcel`
3. Downloads the latest Dexcel application files
4. Installs required Python packages
5. Creates a global `dexcel` command

After installation, you can run:

```bash
dexcel
```

from anywhere.

---

## Output

Dexcel creates one Excel file containing schema descriptions for all tables.

Each table section includes:

```txt
TABLE: users

Field               Type              Null  Key  Default  Extra
id                  bigint unsigned   NO    PRI  NULL     auto_increment
nama                varchar(255)      NO         NULL
emel                varchar(255)      NO    UNI  NULL
created_at          timestamp         YES        NULL
updated_at          timestamp         YES        NULL
```

---

## Logs

Dexcel writes logs into:

### macOS / Linux

```txt
~/.dexcel/logs
```

### Windows

```txt
C:\Users\<your-user>\.dexcel\logs
```

Passwords are never saved in plaintext.

---

## SQL Server note

For SQL Server, you may need to install Microsoft ODBC Driver 17 for SQL Server.

Other database types are not affected.
