"""Optional host-database compatibility check for check_compatibility (B3-14).

This is the *only* part of the reference server that reaches into a live
Microsoft Access file. It is deliberately isolated here so the core server
(`library.py`, `server.py`) stays importable with just `mcp`+`pyyaml`: the
`pyodbc` dependency is imported lazily inside `check_db`, never at module load.

Reading an `.accdb`/`.mdb` needs the Microsoft Access Database Engine (ACE)
ODBC driver, which is Windows-only. When either `pyodbc` or the driver is
absent, `check_db` returns an `available: False` result instead of raising, so
the server keeps working for everyone who never uses this tool.

Existence only: the check opens the database read-only and reads ODBC catalog
metadata (table and column names) — it never runs a query, never reads data,
and never writes anything.
"""

from pathlib import Path

_ACE_DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"
_ENABLE_HINT = (
    "Install pyodbc (pip install -r mcp-server/requirements-compat.txt) and the "
    "Microsoft Access Database Engine redistributable, on Windows."
)


def _unavailable(reason: str) -> dict:
    return {"available": False, "reason": reason, "how_to_enable": _ENABLE_HINT}


def _split_field(entry: str) -> tuple[str, str]:
    """Split a `Table.Field` requirement into (table, field); field '' if absent."""
    table, _, field = str(entry).partition(".")
    return table.strip(), field.strip()


def _read_catalog(conn) -> tuple[set, dict]:
    """Enumerate user tables and their columns from ODBC catalog metadata."""
    cur = conn.cursor()
    # Real user tables only (exclude system/ODBC catalog tables).
    tables = {row.table_name for row in cur.tables(tableType="TABLE")}
    columns = {}  # table (lower) -> set of column names (lower)
    for row in cur.columns():
        columns.setdefault(row.table_name.lower(), set()).add(row.column_name.lower())
    return tables, columns


def check_db(db_path: str, requires_tables, requires_fields) -> dict:
    """Report which required tables/fields exist in an Access database.

    Returns a dict with `available` plus, when available, `tables` and `fields`
    breakdowns (required / present / missing, and `table_absent` for fields
    whose parent table is itself missing). Never raises for a missing driver or
    unreadable file — those come back as `available: False` with a reason.
    """
    path = Path(db_path)
    if not path.is_file():
        return _unavailable(f"database file not found: {db_path}")

    try:
        import pyodbc
    except ImportError:
        return _unavailable("pyodbc is not installed")

    if not any(_ACE_DRIVER.lower() == d.lower() for d in pyodbc.drivers()):
        return _unavailable(f"ODBC driver not found: {_ACE_DRIVER}")

    # Read-only connection; catalog metadata only — no query touches data.
    conn_str = f"DRIVER={{{_ACE_DRIVER}}};DBQ={path};ReadOnly=1;"
    try:
        conn = pyodbc.connect(conn_str, readonly=True)
    except pyodbc.Error as exc:
        return _unavailable(f"could not open database: {exc}")

    try:
        try:
            db_tables, db_columns = _read_catalog(conn)
        except UnicodeDecodeError:
            # Known ACE quirk: the catalog's REMARKS column (field descriptions)
            # can carry bytes that break the default UTF-16 decode. Reconnect
            # decoding wide chars as latin-1 — table/column names come through
            # intact; only the remarks we never read are affected.
            conn.close()
            conn = pyodbc.connect(conn_str, readonly=True)
            conn.setdecoding(pyodbc.SQL_WCHAR, encoding="latin-1")
            db_tables, db_columns = _read_catalog(conn)
    finally:
        conn.close()
    lower_tables = {t.lower(): t for t in db_tables}

    req_tables = [str(t) for t in (requires_tables or [])]
    t_present = [t for t in req_tables if t.lower() in lower_tables]
    t_missing = [t for t in req_tables if t.lower() not in lower_tables]

    req_fields = [str(f) for f in (requires_fields or [])]
    f_present, f_missing, f_absent = [], [], []
    for entry in req_fields:
        table, field = _split_field(entry)
        if table.lower() not in lower_tables:
            f_absent.append(entry)
        elif field.lower() in db_columns.get(table.lower(), set()):
            f_present.append(entry)
        else:
            f_missing.append(entry)

    return {
        "available": True,
        "tables": {"required": req_tables, "present": t_present, "missing": t_missing},
        "fields": {"required": req_fields, "present": f_present,
                   "missing": f_missing, "table_absent": f_absent},
    }
