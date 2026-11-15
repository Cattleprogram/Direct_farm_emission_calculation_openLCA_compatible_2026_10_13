"""
Central paths for all scripts. All paths are relative to the repository folder,
so the scripts run on any computer without changes.

To use your own data: change DB_FILENAME (or KTBL_PARAMETER_FILENAME) and put
the file into the data/ folder.
"""

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

# Access database with farm data (example dataset: farms 2 and 8)
DB_FILENAME = "2026_10_13_Alle_Daten_Example.accdb"
db_path = str(DATA_DIR / DB_FILENAME)

# KTBL parameter file (sheet "BEK-Parameterdatei"), used for the TAN share
# in 4_Emissions_Duengemanagement_N2O_NOx_NH3_Processlevel.py
KTBL_PARAMETER_FILENAME = "Parameterdatei_2024-12_Export.xlsx"
ktbl_parameter_path = DATA_DIR / KTBL_PARAMETER_FILENAME

# Output folder for CSV, logs, Excel exports and Word tables
OUTPUT_DIR = BASE_DIR / "results"
