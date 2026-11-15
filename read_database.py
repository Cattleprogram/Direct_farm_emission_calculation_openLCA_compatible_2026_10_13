"""Read the farm processes and flows from the Access database for the transfer to openLCA.

The transfer needs the database WITH the calculated emissions, i.e. the dated copy that
run_all.py saves at the end ("Ergebnisse als neue Access-Datenbank speichern? j"), named
data/<date>_Alle_Daten_EmissionenBerechnet.accdb. The raw example database in config.py has
no emission rows yet.

Database used (first match):
    1. the path given with --db
    2. the newest data/*_EmissionenBerechnet*.accdb
"""

import sys
from pathlib import Path

import pyodbc

REPO_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO_DIR))
from config import DATA_DIR  # noqa: E402

# saved Access query: all flows per farm process (farm, process, flow, direction, value, unit, ...)
QUERY_NAME = "★_python_qry_Processes_Flows"


def resolve_db_path(db=None):
    """Path of the Access database with the calculated emissions (see module docstring)."""
    if db:
        path = Path(db).resolve()
        if not path.exists():
            raise FileNotFoundError(f"Datenbank nicht gefunden: {path}")
        return path
    candidates = sorted(DATA_DIR.glob("*_EmissionenBerechnet*.accdb"), key=lambda p: p.stat().st_mtime)
    if not candidates:
        raise FileNotFoundError(
            f"Keine Datenbank mit berechneten Emissionen in {DATA_DIR}. Zuerst 'python run_all.py' ausführen "
            "und am Ende mit 'j' speichern, oder den Pfad mit --db angeben."
        )
    return candidates[-1]


def query_db_data(db_path):
    """Return (column names, rows) of QUERY_NAME."""
    print(f"Access-Pfad: {db_path}")
    conn_str = f"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
    conn = pyodbc.connect(conn_str, timeout=5)
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM [{QUERY_NAME}]")
        column_names = [description[0] for description in cursor.description]
        rows = cursor.fetchall()
    finally:
        conn.close()
    print(f"{len(rows)} Zeilen aus {QUERY_NAME}")
    return column_names, rows


if __name__ == "__main__":
    column_names, rows = query_db_data(resolve_db_path(sys.argv[1] if len(sys.argv) > 1 else None))
    print(f"Spalten: {column_names}")
