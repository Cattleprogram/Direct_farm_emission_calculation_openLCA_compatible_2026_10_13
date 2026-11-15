"""
Runs all emission scripts one after the other and collects every calculated
emission in one long table (CSV).

The individual scripts are NOT changed. Instead:
  1. The database from config.py is copied to a temporary scratch copy.
  2. Each script runs via runpy, sees the scratch copy as db_path through a
     substituted `config` module, and every "(j/n)" prompt is answered with "j"
     automatically. So the scripts write as usual - only into the copy.
     This way later scripts see the results of earlier ones (calf ration,
     FlowPropertyDefaults, new P_EmissionMM_xy processes).
  3. After each emission script, its `insert_emissions_df` is taken from the
     script namespace (exactly what is written to tbl_4).
  4. Everything is enriched with farm/process/flow names and saved as CSV.
  5. Prompt whether the scratch copy should be saved as a new dated database next
     to the original database. The original file is never overwritten.

The output of each script goes to its own log file next to the CSV.
In addition, two Word tables are created for the example farms (EXAMPLE_FARMS):
an aggregated reproducibility table (paper) and a complete one (SI).

    python run_all.py                -> calculate everything, CSV + tables
    python run_all.py --tables-only  -> only rebuild the tables from the newest CSV
"""

# ZirKUH_IDs of the example farms for the Word tables (order = column order)
EXAMPLE_FARMS = [2, 8]

import builtins
import contextlib
import gc
import os
import runpy
import shutil
import sys
import tempfile
import time
import traceback
import types
import warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyodbc

# Without pooling, the ODBC driver does not keep the scratch DB open after conn.close();
# otherwise it cannot be copied at the end.
pyodbc.pooling = False

import config
from config import OUTPUT_DIR
from config import db_path as SOURCE_DB_PATH
from emission_tables import write_tables

SCRIPT_DIR = Path(__file__).resolve().parent

# Order = execution order. module=None: the script only prepares inputs
# (it runs, but is not added to the emission table).
SCRIPTS = [
    ("1_Update_P_Milchkuh_Kaelberration.py", None),
    ("2_Futtermittel_Default_KTBL.py", None),
    ("3_Emissions_Pflanzenbau_N2O_CO2_NH3_NO3_P2O5.py", "crop_N_CO2_P"),
    ("3_Emissions_Pflanzenbau_N2O_Wurzelrückstände.py", "crop_residues_N2O"),
    ("3_Emissions_Pflanzenbau_NMVOC_PM.py", "crop_NMVOC_PM"),
    ("4_Emissions_Duengemanagement_N2O_NOx_NH3_Processlevel.py", "manure_N"),
    ("4_Emissions_Duengemanagement_Weide_CH4_Processlevel.py", "manure_CH4"),
    ("4_Emission_Duengemanagement_PM_NMVOC_Processlevel.py", "manure_PM_NMVOC"),
    ("5_Emissions_Tierhaltung_CH4_Cattle.py", "enteric_cattle"),
    ("5_Emissions_Tierhaltung_CH4_Schwein_Schaf_Pferd.py", "enteric_other"),
]

CSV_COLUMNS = [
    "farm",
    "process",
    "process_instance",
    "BetriebProcessID",
    "description",
    "source",
    "flow",
    "category",
    "unit",
    "amount",
    "module",
]


# ============================================================================
# RUN ONE SCRIPT
# ============================================================================


def run_script(script_path, scratch_db, log_path):
    """Runs one script against the scratch DB and answers all prompts with "j".

    Returns (namespace of the script, list of answered prompts).
    """
    prompts = []

    def auto_yes(prompt=""):
        print(f"{prompt}j   [automatisch von run_all.py]")
        prompts.append(str(prompt).strip())
        return "j"

    # Take over all settings from config.py, only db_path points to the scratch copy
    fake_config = types.ModuleType("config")
    fake_config.__dict__.update(
        {k: v for k, v in vars(config).items() if not k.startswith("__")}
    )
    fake_config.db_path = str(scratch_db)

    original_input = builtins.input
    original_config = sys.modules.get("config")
    builtins.input = auto_yes
    sys.modules["config"] = fake_config
    try:
        with open(log_path, "w", encoding="utf-8") as log, contextlib.redirect_stdout(
            log
        ), contextlib.redirect_stderr(log):
            try:
                ns = runpy.run_path(str(script_path), run_name="__main__")
            except BaseException:
                traceback.print_exc()
                raise
    finally:
        builtins.input = original_input
        if original_config is not None:
            sys.modules["config"] = original_config
        else:
            sys.modules.pop("config", None)

    return ns, prompts


def close_connections(ns):
    """Closes all open pyodbc connections a script has left behind."""
    for value in list(ns.values()):
        if isinstance(value, pyodbc.Connection):
            try:
                value.close()
            except pyodbc.Error:
                pass


def extract_emissions(ns, module):
    """Takes insert_emissions_df from the script namespace and adds 'source'.

    source = internal emission type (e.g. NH3_Gas_House, N2O_kg), as far as the script
    keeps it in memory:
      - column 'Tierprozess' (manure management N2O/NOx/NH3), or
      - 'EmissionType' from emissions_long, if insert_emissions_df was derived from it
        row by row (same index, same values).
    """
    df = ns.get("insert_emissions_df")
    if not isinstance(df, pd.DataFrame):
        raise RuntimeError(
            f"[{module}] Kein insert_emissions_df im Skript gefunden - "
            "wurde die Variable umbenannt?"
        )
    df = df.copy()

    if "Tierprozess" in df.columns:
        df["source"] = df["Tierprozess"]
    else:
        df["source"] = pd.NA
        emissions_long = ns.get("emissions_long")
        if (
            isinstance(emissions_long, pd.DataFrame)
            and "EmissionType" in emissions_long.columns
            and df.index.isin(emissions_long.index).all()
        ):
            aligned = emissions_long.loc[df.index]
            if (aligned["Wert"].values == df["Wert"].values).all():
                df["source"] = aligned["EmissionType"].values

    if "TierkategorieEmissionFutter" not in df.columns:
        df["TierkategorieEmissionFutter"] = pd.NA

    df = df[
        ["BetriebProcessID", "FlowID", "Wert", "TierkategorieEmissionFutter", "source"]
    ].copy()
    df["module"] = module
    return df


# ============================================================================
# BUILD THE RESULTS TABLE
# ============================================================================


def load_lookup_tables(db):
    conn = pyodbc.connect(
        rf"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db};"
    )
    try:
        with warnings.catch_warnings():
            # pandas warns about pyodbc connections without SQLAlchemy - works anyway
            warnings.filterwarnings("ignore", message="pandas only supports SQLAlchemy")
            tbl3 = pd.read_sql(
                "SELECT BetriebProcessID, BetriebID, ProcessID FROM tbl_3_Betrieb_Process",
                conn,
            )
            processes = pd.read_sql(
                "SELECT ProcessID, Process FROM Katalog_Process", conn
            )
            flows = pd.read_sql(
                "SELECT FlowID, Flow, Category, Unit FROM Katalog_Flow", conn
            )
            farms = pd.read_sql(
                "SELECT BetriebID, ZirKUH_ID FROM Katalog_Betriebe", conn
            )
    finally:
        conn.close()
    return tbl3, processes, flows, farms


def build_results_table(raw, db):
    tbl3, processes, flows, farms = load_lookup_tables(db)

    # Instance number per farm + process (1, 2, ...), stable across runs,
    # unlike the BetriebProcessID itself.
    tbl3 = tbl3.sort_values("BetriebProcessID")
    tbl3["process_instance"] = tbl3.groupby(["BetriebID", "ProcessID"]).cumcount() + 1

    raw = raw.copy()
    raw["BetriebProcessID"] = pd.to_numeric(raw["BetriebProcessID"]).astype("Int64")
    raw["FlowID"] = pd.to_numeric(raw["FlowID"]).astype("Int64")
    tbl3["BetriebProcessID"] = tbl3["BetriebProcessID"].astype("Int64")
    flows["FlowID"] = flows["FlowID"].astype("Int64")

    result = (
        raw.merge(tbl3, on="BetriebProcessID", how="left")
        .merge(processes, on="ProcessID", how="left")
        .merge(flows, on="FlowID", how="left")
        .merge(farms, on="BetriebID", how="left")
    )

    result = result.rename(
        columns={
            "ZirKUH_ID": "farm",
            "Process": "process",
            "TierkategorieEmissionFutter": "description",
            "Flow": "flow",
            "Category": "category",
            "Unit": "unit",
            "Wert": "amount",
        }
    )

    result["process_instance"] = result["process_instance"].astype("Int64")

    module_order = {module: i for i, (_, module) in enumerate(SCRIPTS)}
    result["_module_order"] = result["module"].map(module_order)
    result = result.sort_values(
        [
            "farm",
            "_module_order",
            "process",
            "process_instance",
            "flow",
            "description",
            "source",
        ],
        na_position="last",
    )
    return result[CSV_COLUMNS].reset_index(drop=True)


def report_problems(result):
    problems = {
        "ohne Flow (FlowID nicht in Katalog_Flow)": result["flow"].isna(),
        "ohne Prozess (BetriebProcessID nicht in tbl_3)": result["process"].isna(),
        "ohne ZirKUH_ID": result["farm"].isna(),
        "amount leer": result["amount"].isna(),
    }
    for label, mask in problems.items():
        if mask.any():
            print(f"  ⚠️  {mask.sum()} Zeilen {label}")
            print(
                result.loc[mask, ["farm", "process", "flow", "amount", "module"]]
                .head(10)
                .to_string(index=False)
            )
    duplicate_key = [
        "farm",
        "process",
        "process_instance",
        "description",
        "source",
        "flow",
        "category",
    ]
    duplicates = result.duplicated(subset=duplicate_key, keep=False)
    if duplicates.any():
        print(
            f"  ⚠️  {duplicates.sum()} Zeilen mit doppeltem Schlüssel "
            f"({', '.join(duplicate_key)})"
        )


def print_summary(result):
    print("\n=== ZUSAMMENFASSUNG JE MODUL × FLOW (Summe über alle Betriebe) ===")
    summary = (
        result.groupby(["module", "flow", "unit"], dropna=False)["amount"]
        .agg(rows="count", total="sum")
        .reset_index()
    )
    module_order = {module: i for i, (_, module) in enumerate(SCRIPTS)}
    summary = summary.sort_values(
        ["module", "flow"], key=lambda s: s.map(module_order) if s.name == "module" else s
    )
    print(summary.round(2).to_string(index=False))


# ============================================================================
# SAVE THE SCRATCH DB AS A NEW DATED DATABASE
# ============================================================================


def next_free_db_path(source_db):
    folder = Path(source_db).parent
    base = f"{datetime.now():%Y_%m_%d}_Alle_Daten_EmissionenBerechnet"
    candidate = folder / f"{base}.accdb"
    counter = 2
    while candidate.exists():
        candidate = folder / f"{base}_{counter}.accdb"
        counter += 1
    return candidate


# ============================================================================
# WORD TABLES FOR THE EXAMPLE FARMS
# ============================================================================


def write_example_tables(result, stamp, source_name):
    module_order = [module for _, module in SCRIPTS if module is not None]
    paper_path, si_path = write_tables(
        result, EXAMPLE_FARMS, module_order, OUTPUT_DIR, stamp, source_name
    )
    print(f"\n📄 Tabellen für Betriebe {EXAMPLE_FARMS}:")
    print(f"   Paper (Modul × Flow): {paper_path}")
    print(f"   SI (vollständig):     {si_path}")


def tables_only():
    """Creates the Word tables from the newest emissions_*.csv without recalculating."""
    csv_files = sorted(OUTPUT_DIR.glob("emissions_*.csv"))
    if not csv_files:
        sys.exit(f"Keine emissions_*.csv in {OUTPUT_DIR} gefunden.")
    csv_path = csv_files[-1]
    print(f"Tabellen aus {csv_path.name}")
    result = pd.read_csv(csv_path, encoding="utf-8-sig")
    result["process_instance"] = result["process_instance"].astype("Int64")
    stamp = csv_path.stem.removeprefix("emissions_")
    write_example_tables(result, stamp, Path(SOURCE_DB_PATH).name)


# ============================================================================
# MAIN
# ============================================================================


def main():
    started = datetime.now()
    stamp = f"{started:%Y-%m-%d_%H%M}"
    log_dir = OUTPUT_DIR / f"logs_{stamp}"
    log_dir.mkdir(parents=True, exist_ok=True)

    source_db = Path(SOURCE_DB_PATH)
    scratch_dir = Path(tempfile.mkdtemp(prefix="run_all_"))
    scratch_db = scratch_dir / source_db.name
    shutil.copy2(source_db, scratch_db)

    print(f"Quelle (bleibt unverändert): {source_db}")
    print(f"Scratch-Kopie:               {scratch_db}")
    print(f"Logs:                        {log_dir}\n")

    # Run the scripts relative to the repository folder.
    os.chdir(SCRIPT_DIR)

    collected = []
    for i, (script_name, module) in enumerate(SCRIPTS, start=1):
        script_path = SCRIPT_DIR / script_name
        log_path = log_dir / f"{script_path.stem}.log"
        label = module or "input preparation"
        print(f"[{i}/{len(SCRIPTS)}] {script_name}  ({label}) ...", end=" ", flush=True)

        t0 = time.perf_counter()
        try:
            ns, prompts = run_script(script_path, scratch_db, log_path)
        except BaseException as exc:
            print("FEHLER")
            print(f"\n❌ {script_name} ist abgebrochen: {exc!r}")
            print(f"   Details: {log_path}")
            print(f"   Scratch-DB zum Nachsehen: {scratch_db}")
            print("   Es wurde nichts in die Originaldatenbank geschrieben.")
            sys.exit(1)

        close_connections(ns)
        n_rows = ""
        if module is not None:
            emissions = extract_emissions(ns, module)
            collected.append(emissions)
            n_rows = f", {len(emissions)} Emissionszeilen"
        del ns
        gc.collect()

        print(
            f"ok ({time.perf_counter() - t0:.0f} s, "
            f"{len(prompts)} Abfrage(n) mit 'j' beantwortet{n_rows})"
        )
        if not prompts:
            print("   ⚠️  Keine (j/n)-Abfrage - hat das Skript überhaupt geschrieben?")

    raw = pd.concat(collected, ignore_index=True)
    result = build_results_table(raw, scratch_db)

    csv_path = OUTPUT_DIR / f"emissions_{stamp}.csv"
    result.to_csv(csv_path, index=False, encoding="utf-8-sig")

    print(f"\n✅ {len(result)} Emissionszeilen aus {result['farm'].nunique()} Betrieben")
    print(f"   CSV: {csv_path}")
    report_problems(result)
    print_summary(result)
    write_example_tables(result, stamp, source_db.name)

    target_db = next_free_db_path(source_db)
    print(
        "\nDie Scratch-Kopie enthält jetzt alle Ergebnisse. Sie kann als neue Datenbank"
        f"\n   {target_db}\nabgelegt werden (die Originaldatei bleibt unverändert)."
    )
    confirm = input("Ergebnisse als neue Access-Datenbank speichern? (j/n): ")
    if confirm.strip().lower() == "j":
        shutil.copy2(scratch_db, target_db)
        print(f"\n✅ Gespeichert: {target_db}")
        print("   Zum Weiterarbeiten db_path in config.py auf diese Datei setzen.")
    else:
        print("\n❌ Nicht gespeichert. Die Originaldatenbank ist unverändert.")

    shutil.rmtree(scratch_dir, ignore_errors=True)
    print(f"\nLaufzeit gesamt: {(datetime.now() - started).seconds // 60} min")


if __name__ == "__main__":
    if "--tables-only" in sys.argv:
        tables_only()
    else:
        main()
