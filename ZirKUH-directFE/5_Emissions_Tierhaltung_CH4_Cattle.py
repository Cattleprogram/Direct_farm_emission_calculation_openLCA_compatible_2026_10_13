import math
import pickle
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import pyodbc
from config import OUTPUT_DIR, db_path

print(
    "Hello World! Let's calculate some emissions from the animal husbandry module! 🐄💨"
)
# ============================================================================
# DATABASE CONNECTION AND DATA LOADING
# ============================================================================


def connect_database(db_path):
    """Connects to the Access database."""
    conn = pyodbc.connect(
        rf"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
    )
    return conn


def load_all_tables(conn):
    """Loads all tables as pandas DataFrames into a dictionary."""
    cursor = conn.cursor()
    tables = cursor.tables()
    table_list = [
        row.table_name
        for row in tables
        if not row.table_name.startswith("MSys") and row.table_type == "TABLE"
    ]

    data = {}
    for table_name in table_list:
        try:
            data[table_name] = pd.read_sql_query(f"SELECT * FROM [{table_name}]", conn)
            print(f"✓ Tabelle '{table_name}' geladen: {len(data[table_name])} Zeilen")
        except Exception as e:
            print(f"✗ Fehler beim Laden von '{table_name}': {e}")

    return data


# Connect to the database and load data
conn = connect_database(db_path)
# print(f"\n📊 Lade alle Tabellen aus: {db_path}\n")
data = load_all_tables(conn)

# print(f"\n✅ Total {len(data)} Tabellen geladen\n")


# List available tables
# print("Verfügbare Tabellen für Berechnungen:")
# for table_name in data.keys():
#     print(
#         f"  - {table_name}: {len(data[table_name])} Zeilen, {len(data[table_name].columns)} Spalten"
#     )
# CATTLE: BUILD THE DATA BASIS FOR LATER CALCULATIONS
# ============================================================================


# Process -> relevant process parameters (Pp_name)
TARGET_PROCESSES = {
    "P_Milchkuh": ("Milchkuh / Mutterkuh", "Kälber", "Deckbulle"),
    "P_Aufzucht_Milchvieh": ("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
    "P_Rind_Mast": ("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
    "P_Mutterkuh": (
        "Milchkuh / Mutterkuh",
        "Kälber",
        "Deckbulle",
        "♂ Mast",
        "♀ Aufzucht / Masttiere (je nach MV oder MK)",
    ),
}


def to_bool_series(series):
    """Robustly converts Access True/False to bool (True can be e.g. -1, 1, 'true', 'ja')."""
    txt = series.astype(str).str.strip().str.lower()
    return (
        series.eq(True)
        | series.eq(1)
        | series.eq(-1)
        | txt.isin(["true", "wahr", "ja", "yes", "y", "1", "-1"])
    )


def is_target_pp(row):
    process_name = row["Process"]
    pp_name = row["Pp_name"]
    if process_name not in TARGET_PROCESSES:
        return False
    return pp_name in TARGET_PROCESSES[process_name]


# 1) Collect all relevant process parameter rows for TARGET_PROCESSES
target_pp_df = (
    data["tbl_6_Betrieb_ProcessParameterValue"]
    .merge(data["tbl_3_Betrieb_Process"], on="BetriebProcessID", how="left")
    .merge(
        data["Katalog_ProcessProcessparameter"][["ProcessPpID", "PpID"]],
        on="ProcessPpID",
        how="left",
    )
    .merge(data["Katalog_Processparameter"][["PpID", "Pp_name"]], on="PpID", how="left")
    .merge(
        data["Katalog_Process"][["ProcessID", "Process"]],
        on="ProcessID",
        how="left",
    )
)

target_pp_df = target_pp_df[target_pp_df.apply(is_target_pp, axis=1)].copy()

target_pp_df["Wert"] = (
    target_pp_df["Wert"].astype(str).str.replace(",", ".", regex=False).str.strip()
)
target_pp_df["Wert"] = pd.to_numeric(target_pp_df["Wert"], errors="coerce")

# print("\n=== RELEVANTE PROZESSPARAMETER (CATTLE) ===")
# print(
#     target_pp_df[
#         ["BetriebID", "BetriebProcessID", "ProcessID", "Process", "Pp_name", "Wert"]
#     ].head(20)
# )
target_pp_df.rename(columns={"Wert": "Pp_Wert"}, inplace=True)
# print(f"{len(target_pp_df)} relevante Parameterzeilen gefunden.")

# 2) Get flows from tbl_4 and filter on IsPotentialFeedstuff == True
flows_with_process_df = data["tbl_4_Betrieb_Process_FlowValue"].merge(
    target_pp_df[
        ["BetriebProcessID", "BetriebID", "ProcessID", "Process", "Pp_name", "Pp_Wert"]
    ].drop_duplicates(),
    on="BetriebProcessID",
    how="inner",
)

katalog_flow_df = data["Katalog_Flow"].copy()
if "IsPotentialFeedstuff" not in katalog_flow_df.columns:
    raise KeyError("Spalte 'IsPotentialFeedstuff' fehlt in Katalog_Flow.")

katalog_flow_df["_is_potential_feedstuff"] = to_bool_series(
    katalog_flow_df["IsPotentialFeedstuff"]
)

potential_feed_flows_df = flows_with_process_df.merge(
    katalog_flow_df[
        [
            c
            for c in [
                "FlowID",
                "Flow",
                "Category",
                "Subcategory",
                "IsPotentialFeedstuff",
                "_is_potential_feedstuff",
                "Kälberfutter",
            ]
            if c in katalog_flow_df.columns
        ]
    ],
    on="FlowID",
    how="left",
)

if (
    ~potential_feed_flows_df["_is_potential_feedstuff"]
    & (potential_feed_flows_df["TierkategorieEmissionFutter"] == "Kälber")
).any():
    raise ValueError(
        "Es gibt Zeilen mit TierkategorieEmissionFutter == 'Kälber', aber IsPotentialFeedstuff == False. "
    )

p_other_mask = (potential_feed_flows_df["Process"] != "P_Milchkuh") & (
    potential_feed_flows_df["_is_potential_feedstuff"]
)

p_milchkuh_milchkuh_mask = (
    (potential_feed_flows_df["Process"] == "P_Milchkuh")
    & (potential_feed_flows_df["_is_potential_feedstuff"])
    & (potential_feed_flows_df["TierkategorieEmissionFutter"] != "Kälber")
)

p_milchkuh_kälber_mask = (
    (potential_feed_flows_df["Process"] == "P_Milchkuh")
    & (potential_feed_flows_df["_is_potential_feedstuff"])
    & (potential_feed_flows_df["TierkategorieEmissionFutter"] == "Kälber")
)

filter_mask = p_other_mask | p_milchkuh_milchkuh_mask | p_milchkuh_kälber_mask

potential_feed_flows_df = potential_feed_flows_df[filter_mask].copy()

print(
    "\n=== RELEVANTE FLOWS JE TARGET PROCESS + Pp_name "
    "(Basis: IsPotentialFeedstuff; bei P_Milchkuh + Kälber zusätzlich Kälberfutter) ==="
)
print(
    potential_feed_flows_df[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "Pp_name",
            "FlowID",
            "Flow",
            "Wert",
        ]
    ].head(20)
)
print(f"{len(potential_feed_flows_df)} potenzielle Futtermittel-Flows gefunden.")

# Result DataFrame for the next step
target_potential_feed_flows_df = potential_feed_flows_df.copy()

print("\n=== FERTIG (STEP 1): target_potential_feed_flows_df ===")
print(f"{len(target_potential_feed_flows_df)} Zeilen")


# 3) Add feedstuff per flow (all columns from tbl_Futtermittel_Mapping)
futtermittel_mapping_df = data["tbl_Futtermittel_Mapping"].copy()
if "FlowID" not in futtermittel_mapping_df.columns:
    raise KeyError("Spalte 'FlowID' fehlt in tbl_Futtermittel_Mapping.")

# Expectation: every flow has exactly one feedstuff mapping
mapping_count_df = (
    futtermittel_mapping_df.groupby("FlowID", dropna=False)
    .size()
    .reset_index(name="mapping_count")
)

non_unique_mapping_df = mapping_count_df[mapping_count_df["mapping_count"] != 1]
if len(non_unique_mapping_df) > 0:
    print("\n⚠ Warnung: Für einige FlowIDs ist die Mapping-Anzahl ungleich 1.")
    print(non_unique_mapping_df.head(20))

target_potential_feed_flows_with_futtermittel_df = target_potential_feed_flows_df.merge(
    futtermittel_mapping_df,
    on="FlowID",
    how="left",
)

print("\n=== STEP 2: target_potential_feed_flows_with_futtermittel_df ===")
print(f"{len(target_potential_feed_flows_with_futtermittel_df)} Zeilen")

missing_futtermittel = (
    target_potential_feed_flows_with_futtermittel_df["FuttermittelID"].isna().sum()
    if "FuttermittelID" in target_potential_feed_flows_with_futtermittel_df.columns
    else None
)

if missing_futtermittel is not None:
    print(f"Fehlende Futtermittel-Zuordnungen: {missing_futtermittel}")

preview_cols = [
    "BetriebID",
    "BetriebProcessID",
    "Process",
    "Pp_name",
    "FlowID",
    "Flow",
    "FuttermittelID",
]
preview_cols = [
    c
    for c in preview_cols
    if c in target_potential_feed_flows_with_futtermittel_df.columns
]
print(target_potential_feed_flows_with_futtermittel_df[preview_cols].head(20))


# 4) Slim working view for creating new tbl_4 rows later
tbl4_input_cols = [
    "BetriebID",
    "BetriebProcessID",
    "BetriebProcessFlowValueID",
    "ProcessID",
    "Process",
    "Pp_name",
    "Pp_Wert",
    "FlowID",
    "Flow",
    "Wert",
    "FuttermittelID",
    "TierkategorieEmissionFutter",
]
tbl4_input_cols = [
    c
    for c in tbl4_input_cols
    if c in target_potential_feed_flows_with_futtermittel_df.columns
]

target_tbl4_input_df = target_potential_feed_flows_with_futtermittel_df[
    tbl4_input_cols
].copy()

target_tbl4_input_df.rename(columns={"Wert": "Flow_Wert"}, inplace=True)

print("\n=== STEP 3: target_tbl4_input_df (schlanke Sicht) ===")
print(f"{len(target_tbl4_input_df)} Zeilen")
print(target_tbl4_input_df.head(20))


# 5) Add all feedstuff info (GE, XL, XP, ...) from KTBL
ktbl_df = data["Katalog_Futtermittelgehalte_KTBL"].copy()
if "FuttermittelID" not in ktbl_df.columns:
    raise KeyError("Spalte 'FuttermittelID' fehlt in Katalog_Futtermittelgehalte_KTBL.")

target_tbl4_input_with_futtermittelinfo_df = target_tbl4_input_df.merge(
    ktbl_df,
    on="FuttermittelID",
    how="left",
)

print("\n=== STEP 4: target_tbl4_input_with_futtermittelinfo_df ===")
print(f"{len(target_tbl4_input_with_futtermittelinfo_df)} Zeilen")

missing_ktbl = (
    target_tbl4_input_with_futtermittelinfo_df["FuttermittelID"].notna().sum()
    - target_tbl4_input_with_futtermittelinfo_df.dropna(subset=["FuttermittelID"])
    .merge(
        ktbl_df[["FuttermittelID"]].drop_duplicates(), on="FuttermittelID", how="inner"
    )
    .shape[0]
)
print(f"Fehlende KTBL-Infos trotz FuttermittelID: {missing_ktbl}")

ktbl_preview_cols = [
    "BetriebID",
    "BetriebProcessID",
    "Process",
    "Pp_name",
    "FlowID",
    "Flow",
    "FuttermittelID",
    "GE",
    "XP",
    "XL",
]
ktbl_preview_cols = [
    cx
    for cx in target_tbl4_input_with_futtermittelinfo_df.columns
    for c in ktbl_preview_cols
    if c in cx
]
print(target_tbl4_input_with_futtermittelinfo_df[ktbl_preview_cols].head(20))


# 6) Match FlowProperties to KTBL columns and override with farm-specific values
flowproperty_df = data["Katalog_Flowproperty"].copy()
betrieb_flowproperty_df = data["tbl_5_Betrieb_FlowpropertyValue"].copy()

required_flowproperty_cols = {"FlowPropertyID", "FlowProperty"}
if not required_flowproperty_cols.issubset(flowproperty_df.columns):
    raise KeyError(
        "Katalog_Flowproperty muss die Spalten "
        f"{sorted(required_flowproperty_cols)} enthalten."
    )

if "FlowPropertyID" not in betrieb_flowproperty_df.columns:
    raise KeyError(
        "tbl_5_Betrieb_FlowpropertyValue muss mindestens die Spalte 'FlowPropertyID' enthalten."
    )

if "BetriebProcessFlowValueID" not in betrieb_flowproperty_df.columns:
    raise KeyError(
        "tbl_5_Betrieb_FlowpropertyValue muss die Spalte 'BetriebProcessFlowValueID' enthalten."
    )

if (
    "BetriebProcessFlowValueID"
    not in target_tbl4_input_with_futtermittelinfo_df.columns
):
    raise KeyError(
        "Im Ziel-DataFrame fehlt 'BetriebProcessFlowValueID'. "
        "Die ID muss aus tbl_4_Betrieb_Process_FlowValue übernommen werden, "
        "damit tbl_5-Overrides gematcht werden können."
    )

value_col_candidates = [
    "Wert",
    "Value",
    "FlowPropertyValue",
    "BetriebFlowPropertyValue",
]
value_col = next(
    (c for c in value_col_candidates if c in betrieb_flowproperty_df.columns),
    None,
)
if value_col is None:
    raise KeyError(
        "Keine Wertespalte in tbl_5_Betrieb_FlowpropertyValue gefunden. "
        f"Erwartet eine von: {value_col_candidates}"
    )

used_flowproperty_ids = (
    betrieb_flowproperty_df["FlowPropertyID"].dropna().drop_duplicates()
)
flowproperty_used_df = flowproperty_df[
    flowproperty_df["FlowPropertyID"].isin(used_flowproperty_ids)
].copy()

ktbl_columns = [c for c in ktbl_df.columns if c != "FuttermittelID"]


def _find_unique_ktbl_match(flowproperty_name):
    fp = str(flowproperty_name)
    matches = [c for c in ktbl_columns if fp in str(c)]
    if fp.startswith("XP"):
        matches = [c for c in matches if str(c).startswith("XP")]
    if len(matches) != 1:
        return matches
    return matches[0]


flowproperty_used_df["KTBLColumnMatch"] = flowproperty_used_df["FlowProperty"].apply(
    _find_unique_ktbl_match
)

invalid_match_df = flowproperty_used_df[
    flowproperty_used_df["KTBLColumnMatch"].apply(lambda x: isinstance(x, list))
].copy()

if len(invalid_match_df) > 0:
    details = invalid_match_df[["FlowPropertyID", "FlowProperty", "KTBLColumnMatch"]]
    print(
        "FlowProperty-Matching ist nicht eindeutig (muss genau 1 KTBL-Spalte sein):\n"
        f"{details.to_string(index=False)}"
    )

flowproperty_used_df = flowproperty_used_df[
    ~flowproperty_used_df["KTBLColumnMatch"].apply(lambda x: isinstance(x, list))
].copy()

flowproperty_match_df = flowproperty_used_df[
    ["FlowPropertyID", "FlowProperty", "KTBLColumnMatch"]
].copy()


betrieb_overrides_long_df = betrieb_flowproperty_df.merge(
    flowproperty_match_df,
    on="FlowPropertyID",
    how="inner",
)

merge_keys = ["BetriebProcessFlowValueID"]

dup_check_df = (
    betrieb_overrides_long_df.groupby(merge_keys + ["KTBLColumnMatch"], dropna=False)
    .size()
    .reset_index(name="n")
)
dup_rows = dup_check_df[dup_check_df["n"] > 1]
if len(dup_rows) > 0:
    raise ValueError(
        "Mehrfach vorhandene Betrieb-Overrides für denselben Key + FlowProperty gefunden:\n"
        f"{dup_rows.head(30).to_string(index=False)}"
    )

betrieb_overrides_wide_df = betrieb_overrides_long_df.pivot_table(
    index=merge_keys,
    columns="KTBLColumnMatch",
    values=value_col,
    aggfunc="first",
).reset_index()

override_cols = [c for c in betrieb_overrides_wide_df.columns if c not in merge_keys]
rename_map = {c: f"{c}__override" for c in override_cols}
betrieb_overrides_wide_df = betrieb_overrides_wide_df.rename(columns=rename_map)

target_tbl4_input_with_overrides_df = target_tbl4_input_with_futtermittelinfo_df.merge(
    betrieb_overrides_wide_df,
    on=merge_keys,
    how="left",
)

override_helper_cols = [
    f"{c}__override"
    for c in override_cols
    if f"{c}__override" in target_tbl4_input_with_overrides_df.columns
]

if len(override_helper_cols) > 0:
    has_tbl5_override_mask = (
        target_tbl4_input_with_overrides_df[override_helper_cols].notna().any(axis=1)
    )
else:
    has_tbl5_override_mask = pd.Series(
        False,
        index=target_tbl4_input_with_overrides_df.index,
    )

target_tbl4_input_with_overrides_df["FlowpropertyValueSource"] = (
    has_tbl5_override_mask.map(
        {
            True: "tbl_5_Betrieb_FlowpropertyValue",
            False: "Katalog_Futtermittelgehalte_KTBL",
        }
    )
)

for col in override_cols:
    override_col = f"{col}__override"
    if (
        col in target_tbl4_input_with_overrides_df.columns
        and override_col in target_tbl4_input_with_overrides_df.columns
    ):
        target_tbl4_input_with_overrides_df[col] = target_tbl4_input_with_overrides_df[
            override_col
        ].combine_first(target_tbl4_input_with_overrides_df[col])

target_tbl4_input_with_overrides_df = target_tbl4_input_with_overrides_df.drop(
    columns=[
        f"{c}__override"
        for c in override_cols
        if f"{c}__override" in target_tbl4_input_with_overrides_df.columns
    ]
)

target_tbl4_input_with_futtermittelinfo_df = target_tbl4_input_with_overrides_df.copy()

print("\n=== STEP 6: Betrieb-Overrides auf KTBL-Werte angewendet (Fallback-Logik) ===")
print(f"Merge-Keys für Overrides: {merge_keys}")
print(f"{len(target_tbl4_input_with_futtermittelinfo_df)} Zeilen")
print(
    target_tbl4_input_with_futtermittelinfo_df["FlowpropertyValueSource"].value_counts(
        dropna=False
    )
)
print(target_tbl4_input_with_futtermittelinfo_df[ktbl_preview_cols].head(20))

# 7)
#
# 7A) Add livestock unit (GVE) data for later weightings

gve_rind_df = data["tbl_2_Betrieb_Stammdaten"][
    ["BetriebID", "Stammdaten", "Value"]
].copy()

gve_rind_df = (
    gve_rind_df[gve_rind_df["Stammdaten"] == "GVE Rind"]
    .drop(columns=["Stammdaten"])
    .rename(columns={"Value": "GVE_Rind"})
)

# gve_rind_df["Value_num"] = pd.to_numeric(
#     gve_rind_df["Value"].str.replace(",", "."), errors="coerce"
# )
#
# if gve_rind_df["Value_num"].isna().any():
#     print(
#         "⚠ Warnung: Einige GVE Rind-Werte konnten nicht in numerische Werte konvertiert werden."
#     )
#     print(gve_rind_df[gve_rind_df["Value_num"].isna()][["BetriebID", "Value"]])

target_tbl4_input_with_futtermittelinfo_df = (
    target_tbl4_input_with_futtermittelinfo_df.merge(
        gve_rind_df,
        on="BetriebID",
        how="left",
    )
)

# 7A-2) Add farm type (suckler cow vs. dairy system), so that exports
# can be filtered by farm type.
betriebsform_df = data["tbl_2_Betrieb_Stammdaten"][
    ["BetriebID", "Stammdaten", "Value"]
].copy()
betriebsform_df = (
    betriebsform_df[betriebsform_df["Stammdaten"] == "Betriebsform"]
    .drop(columns=["Stammdaten"])
    .rename(columns={"Value": "Betriebsform"})
)

duplicate_betriebsform_df = betriebsform_df[
    betriebsform_df.duplicated(subset=["BetriebID"], keep=False)
]
if len(duplicate_betriebsform_df) > 0:
    print(
        "\n⚠ Warnung: Für einige BetriebIDs gibt es mehr als eine 'Betriebsform'-Zeile "
        "in tbl_2_Betrieb_Stammdaten:"
    )
    print(duplicate_betriebsform_df.to_string(index=False))

target_tbl4_input_with_futtermittelinfo_df = (
    target_tbl4_input_with_futtermittelinfo_df.merge(
        betriebsform_df,
        on="BetriebID",
        how="left",
    )
)

# 7A-3) Add ZirKUH_ID (human-readable farm ID for exports).
target_tbl4_input_with_futtermittelinfo_df = (
    target_tbl4_input_with_futtermittelinfo_df.merge(
        data["Katalog_Betriebe"][["BetriebID", "ZirKUH_ID"]],
        on="BetriebID",
        how="left",
    )
)

# 7B) Convert all columns relevant for the later calculation to numeric values
for column in [
    "Flow_Wert",
    "GVE_Rind",
    "Gross Energy (GE MJ/kg TM)",
    "XP (g/kg TM)",
    "XL (g/kg TM)",
    "XF (g/kg TM)",
    "NFE  (g/kg TM)",
]:
    # Convert any strings or other objects to float
    numeric_column = target_tbl4_input_with_futtermittelinfo_df[column].apply(
        lambda value: (
            float(value)
            if isinstance(value, (int, float))
            else float(value.replace(",", ".")) if isinstance(value, str) else np.nan
        )
    )
    if numeric_column.isna().any():
        incorrect_values = target_tbl4_input_with_futtermittelinfo_df[
            numeric_column.isna()
        ]
        print(
            f"Folgende Flows haben keine numerischen Werte in Spalte '{column}': {incorrect_values['Flow'].unique()}. Sie werden als 0 in die Berechnung einfließen."
        )
        # print(
        #     target_tbl4_input_with_futtermittelinfo_df[numeric_column.isna()][
        #         ["BetriebID", "Flow", column]
        #     ]
        # )

    target_tbl4_input_with_futtermittelinfo_df[column] = numeric_column.fillna(0)

# Manual exception: if a farm's GVE Rind value from the master data covers
# two processes together (instead of only one), the formulas further down
# (which divide by GVE_Rind per process) would be distorted. For such farms
# GVE_Rind can be overridden here as {BetriebID: GVE_Rind}; all other
# farms use the master data value. In the full project dataset this affected
# one farm; for the example dataset (farms 2 and 8) no
# exception is needed.
GVE_RIND_OVERRIDES = {}
betrieb_id_numeric = pd.to_numeric(
    target_tbl4_input_with_futtermittelinfo_df["BetriebID"], errors="coerce"
)
for betrieb_id, gve_rind_override in GVE_RIND_OVERRIDES.items():
    target_tbl4_input_with_futtermittelinfo_df.loc[
        betrieb_id_numeric == betrieb_id, "GVE_Rind"
    ] = gve_rind_override

FEEDSTUFF_EXPORT_DIR = OUTPUT_DIR / "Feedstuff"
FEEDSTUFF_EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# Export of target_tbl4_input_with_futtermittelinfo_df (export only, no
# effect on the further calculation): "Deckbulle" rows are pure duplicates of the
# corresponding "Milchkuh / Mutterkuh" row (same Flow_Werte, only tagged with a different
# Pp_name/Pp_Wert for the same BetriebProcessID) and are removed before the export.
tbl4_input_export_df = target_tbl4_input_with_futtermittelinfo_df[
    target_tbl4_input_with_futtermittelinfo_df["Pp_name"] != "Deckbulle"
].copy()

tbl4_input_export_path = (
    FEEDSTUFF_EXPORT_DIR
    / f"target_tbl4_input_with_futtermittelinfo_{date.today().isoformat()}.xlsx"
)
tbl4_input_export_df.to_excel(tbl4_input_export_path, index=False)
print(f"Excel-Export: {tbl4_input_export_path}")

# Export: feed amount actually eaten per ZirKUH_ID/BetriebID, per flow
# (deduplicated). target_tbl4_input_with_futtermittelinfo_df has, per
# BetriebProcessFlowValueID (= one physical tbl_4 flow), several rows, one per
# Pp_name category the flow is split into for the CH4 calculation.
# For the actual feed amount it must therefore first be deduplicated on BetriebProcessFlowValueID,
# before summing per ZirKUH_ID + flow (no grand total
# over all flows, but separately per flow name).
feedstuff_total_per_zirkuh_df = (
    target_tbl4_input_with_futtermittelinfo_df.drop_duplicates(
        subset=["BetriebProcessFlowValueID"]
    )
    .groupby(["ZirKUH_ID", "BetriebID", "FlowID", "Flow"], as_index=False)["Flow_Wert"]
    .sum()
    .rename(columns={"Flow_Wert": "Flow_Wert_total_kg"})
    .sort_values(["ZirKUH_ID", "Flow"])
)

print("\n=== EXPORT: Tatsächliche Futtermenge je ZirKUH_ID (dedupliziert) ===")
print(feedstuff_total_per_zirkuh_df.to_string(index=False))

feedstuff_total_per_zirkuh_export_path = (
    FEEDSTUFF_EXPORT_DIR / f"feedstuff_total_per_zirkuh_{date.today().isoformat()}.xlsx"
)
feedstuff_total_per_zirkuh_df.to_excel(
    feedstuff_total_per_zirkuh_export_path, index=False
)
print(f"Excel-Export: {feedstuff_total_per_zirkuh_export_path}")


# Export: feed inputs per Pp_name group and BetriebProcess (export only,
# no effect on the calculation). For every process/Pp_name group actually
# calculated further down, it reproduces exactly the same filter conditions that
# calculate_ch4_emissions_for_process_pp_group() uses internally (Process + Pp_name +,
# for P_Milchkuh, the TierkategorieEmissionFutter split), without calling
# or changing that function itself.
# Silenced: redundant with target_tbl4_input_with_futtermittelinfo export and
# feedstuff_total_per_zirkuh export above (kept for reference, not deleted).
# def build_feedstuff_group_export_view(df, process_name, grouped_pp_names, group_label):
#     group_df = df[
#         (df["Process"] == process_name) & (df["Pp_name"].isin(grouped_pp_names))
#     ].copy()
#     if process_name == "P_Milchkuh" and "Kälber" in grouped_pp_names:
#         group_df = group_df[group_df["TierkategorieEmissionFutter"] == "Kälber"]
#     if process_name == "P_Milchkuh" and "Kälber" not in grouped_pp_names:
#         group_df = group_df[group_df["TierkategorieEmissionFutter"] != "Kälber"]
#     group_df.insert(0, "PpGroup", group_label)
#     return group_df
#
#
# feedstuff_export_groups = [
#     (
#         "P_Aufzucht_Milchvieh",
#         ("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
#         "P_Aufzucht_Milchvieh: Mast + Aufzucht/Mast",
#     ),
#     (
#         "P_Rind_Mast",
#         ("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
#         "P_Rind_Mast: Mast + Aufzucht/Mast",
#     ),
#     (
#         "P_Mutterkuh",
#         ("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
#         "P_Mutterkuh: Mast + Aufzucht/Mast (GVE 0.6)",
#     ),
#     ("P_Mutterkuh", ("Kälber",), "P_Mutterkuh: Kälber (GVE 0.3)"),
#     (
#         "P_Mutterkuh",
#         ("Milchkuh / Mutterkuh", "Deckbulle"),
#         "P_Mutterkuh: Milchkuh/Mutterkuh + Deckbulle (Tier3)",
#     ),
#     (
#         "P_Milchkuh",
#         ("Milchkuh / Mutterkuh", "Deckbulle"),
#         "P_Milchkuh: Milchkuh/Mutterkuh + Deckbulle (Tier3)",
#     ),
#     ("P_Milchkuh", ("Kälber",), "P_Milchkuh: Kälber"),
# ]
#
# feedstuff_per_group_export_df = pd.concat(
#     [
#         build_feedstuff_group_export_view(
#             target_tbl4_input_with_futtermittelinfo_df,
#             process_name,
#             grouped_pp_names,
#             group_label,
#         )
#         for process_name, grouped_pp_names, group_label in feedstuff_export_groups
#     ],
#     ignore_index=True,
# )
#
# feedstuff_per_group_export_path = (
#     FEEDSTUFF_EXPORT_DIR
#     / f"feedstuff_per_pp_group_{date.today().isoformat()}.xlsx"
# )
# feedstuff_per_group_export_df.to_excel(feedstuff_per_group_export_path, index=False)
# print(f"Excel-Export: {feedstuff_per_group_export_path}")


# Export: calf milk intake. Physically not a tbl_4 flow, so it does not appear in any of
# the exports above, but it enters the calculation inside calculate_ch4_emissions_for_process_pp_group()
# (P_Milchkuh + Kälber, Tier 2). Loads/calculates the same values as
# there (see milk_df further down), purely for the export, without touching the calculation.
milk_intake_export_df = pickle.load(
    open(Path(__file__).parent / "temp_calf_milk_intake.pkl", "rb")
)
milk_intake_export_df = milk_intake_export_df.rename(columns={"Wert": "Flow_Wert"})
milk_intake_export_df = milk_intake_export_df[
    milk_intake_export_df["Process"].isin(["P_Milchkuh", "P_Mutterkuh"])
].copy()
milk_intake_export_df["Futtermittel"] = "Vollmilch 4,2 Fett 3,5 Eiweiß"
milk_intake_export_df["Flow"] = "Vollmilch 4,2 Fett 3,5 Eiweiß"
milk_intake_export_df["Pp_name"] = "Kälber"

milchkuh_kalber_reference_df = target_tbl4_input_with_futtermittelinfo_df[
    (target_tbl4_input_with_futtermittelinfo_df["Process"] == "P_Milchkuh")
    & (target_tbl4_input_with_futtermittelinfo_df["Pp_name"] == "Kälber")
    & (
        target_tbl4_input_with_futtermittelinfo_df["TierkategorieEmissionFutter"]
        == "Kälber"
    )
]
milk_pp_werts = milchkuh_kalber_reference_df.groupby("BetriebProcessID")[
    "Pp_Wert"
].first()
milk_gve_rinds = milchkuh_kalber_reference_df.groupby("BetriebProcessID")[
    "GVE_Rind"
].first()
milk_intake_export_df["Pp_Wert"] = milk_intake_export_df["BetriebProcessID"].map(
    milk_pp_werts
)
milk_intake_export_df["GVE_Rind"] = milk_intake_export_df["BetriebProcessID"].map(
    milk_gve_rinds
)

milk_intake_export_path = (
    FEEDSTUFF_EXPORT_DIR / f"kaelber_milchintake_{date.today().isoformat()}.xlsx"
)
milk_intake_export_df.to_excel(milk_intake_export_path, index=False)
print(f"Excel-Export: {milk_intake_export_path}")

# 8) Grouped CH4 calculation for P_Aufzucht_Milchvieh
# Tier 2: CH4 conversion factors (share of gross energy) and energy content of CH4 (MJ/kg):
# Arbeitsgruppe BEK (2026): Methodenbeschreibung Sachbilanz Tierhaltung der BEK-Fachgruppen.
X_CH4_GE_NON_CALVES = 0.063
X_CH4_GE_CALVES = 0.041
N_CH4 = 55.65


def get_unique_column(candidates, available_columns, label):
    matches = [
        c
        for c in available_columns
        if any(k.lower() in str(c).lower() for k in candidates)
    ]
    if len(matches) == 0:
        raise KeyError(
            f"Keine passende Spalte für '{label}' gefunden. "
            f"Gesucht nach Teilstrings: {candidates}"
        )
    if len(matches) > 1:
        raise ValueError(
            f"Mehrdeutige Spalten für '{label}': {matches}. "
            "Bitte Mapping präzisieren."
        )
    return matches[0]


def calculate_ch4_emissions_for_process_pp_group(
    df, process_name, grouped_pp_names, consider_gve, gve_weight, is_tier_2
):
    # A process that does not occur in the dataset (e.g. no beef fattening farm)
    # has nothing to calculate: return an empty result with the usual columns.
    if process_name not in df["Process"].unique():
        print(f"Prozess '{process_name}' nicht im Datensatz - keine Berechnung.")
        return pd.DataFrame(
            columns=[
                "BetriebID",
                "ProcessID",
                "Process",
                "BetriebProcessID",
                "Notizen",
                "TierkategorieEmissionFutter",
                "Wert",
                "Flow",
                "FlowID",
                "Direction",
            ]
        )

    # Make sure all grouped_pp_names exist in the DataFrame
    for pp_name in grouped_pp_names:
        assert (
            pp_name in df["Pp_name"].unique()
        ), f"Pp_name '{pp_name}' nicht in DataFrame gefunden."

    # Filter all rows for the given process_name and grouped_pp_names
    filtered_pp_df = df[
        (df["Process"] == process_name) & (df["Pp_name"].isin(grouped_pp_names))
    ].copy()

    # Marks feed rows as "not a milk flow". Set to True on the milk rows further
    # down, so that the milk can be excluded specifically from the GVE weighting
    # (see comment at "Wert_x_GE_weighted").
    filtered_pp_df["is_milk"] = False

    if process_name == "P_Milchkuh" and "Kälber" in grouped_pp_names:
        filtered_pp_df = filtered_pp_df[
            filtered_pp_df["TierkategorieEmissionFutter"] == "Kälber"
        ]
    if process_name == "P_Milchkuh" and "Kälber" not in grouped_pp_names:
        # TODO: make sure that Milchkuh, Milchkuh/Mutterkuh and Deckbulle do not use the calf flows
        filtered_pp_df = filtered_pp_df[
            filtered_pp_df["TierkategorieEmissionFutter"] != "Kälber"
        ]

    if is_tier_2:
        # Add milk rows
        # TODO: check whether this applies to both processes
        if (
            process_name in ["P_Milchkuh", "P_Mutterkuh"]
        ) and "Kälber" in grouped_pp_names:
            milk_df = pickle.load(
                open(Path(__file__).parent / "temp_calf_milk_intake.pkl", "rb")
            )
            milk_df.rename(columns={"Wert": "Flow_Wert"}, inplace=True)
            milk_df = milk_df[milk_df["Process"] == process_name]
            # Fix 06.10.2026 convert fresh milk to DM
            MILK_DENSITY_KG_PER_L = 1.03  # whole milk
            milk_ktbl = ktbl_df[ktbl_df["FuttermittelID"] == 98].iloc[
                0
            ]  # whole milk 4.2 fat 3.5 protein
            milk_df["Milch_L"] = milk_df["Flow_Wert"]
            milk_df["Flow_Wert"] = (
                milk_df["Milch_L"]
                * MILK_DENSITY_KG_PER_L
                * milk_ktbl["TM (g/kg)"]
                / 1000
            )  # → kg DM, like all other feeds
            milk_df["Gross Energy (GE MJ/kg TM)"] = milk_ktbl[
                "Gross Energy (GE MJ/kg TM)"
            ]  # instead of hard-coded 26.4
            milk_df["Futtermittel"] = "Vollmilch 4,2 Fett 3,5 Eiweiß"
            milk_df["Flow"] = "Vollmilch 4,2 Fett 3,5 Eiweiß"
            milk_df["Pp_name"] = "Kälber"
            milk_df["is_milk"] = True
            milk_df["ProcessID"] = filtered_pp_df["ProcessID"].iloc[0]
            pp_werts = filtered_pp_df.groupby("BetriebProcessID")["Pp_Wert"].first()
            milk_df["Pp_Wert"] = milk_df["BetriebProcessID"].map(pp_werts)
            gve_rinds = filtered_pp_df.groupby("BetriebProcessID")["GVE_Rind"].first()
            milk_df["GVE_Rind"] = milk_df["BetriebProcessID"].map(gve_rinds)
            if milk_df["Pp_Wert"].isna().any() or milk_df["GVE_Rind"].isna().any():
                incorrect_pp_rows = milk_df[
                    milk_df["Pp_Wert"].isna() | milk_df["GVE_Rind"].isna()
                ]
                print(
                    f"n⚠ Warnung: Für einige BetriebProcessID konnte kein Pp_Wert oder GVE_Rind für Kälber gefunden werden: {incorrect_pp_rows['BetriebProcessID'].tolist()}."
                )
            filtered_pp_df = pd.concat([filtered_pp_df, milk_df], ignore_index=True)

        filtered_pp_df["Wert_x_GE"] = (
            filtered_pp_df["Flow_Wert"] * filtered_pp_df["Gross Energy (GE MJ/kg TM)"]
        )
        if consider_gve:
            # GE_Flow * Wert_Flow / GVE_Rind / 0.6 * (number fattening + number rearing / fattening)
            # =
            # GE_Flow * Wert_Flow / GVE_Rind / 0.6 * number fattening +
            # GE_Flow * Wert_Flow / GVE_Rind / 0.6 * number rearing / fattening
            # -> calculation per row, then sum per BetriebProcessID
            if gve_weight is None:
                raise ValueError(
                    "consider_gve is True, but gve_weight is None. Please provide a gve_weight."
                )
            filtered_pp_df["Wert_x_GE_weighted"] = (
                filtered_pp_df["Wert_x_GE"]
                / filtered_pp_df["GVE_Rind"]
                * filtered_pp_df["Pp_Wert"]
                * gve_weight
            )

            # Do NOT distribute milk (is_milk) via GVE_Rind/Pp_Wert/gve_weight: this
            # apportionment is only correct for solid feed, which is recorded as one common
            # herd ration in tbl_4 and must be split proportionally between fattening/rearing/calves/etc.
            # The calf milk, in contrast, is already a purely
            # calf-specific amount (see Kälbermilch_L in
            # 1_Update_P_Milchkuh_Kaelberration.py) - additionally dividing it by GVE_Rind
            # and multiplying by Pp_Wert*gve_weight would wrongly treat it
            # like a herd share and dilute it artificially. Milk therefore enters
            # the sum unweighted (Wert_x_GE only).
            filtered_pp_df.loc[filtered_pp_df["is_milk"], "Wert_x_GE_weighted"] = (
                filtered_pp_df.loc[filtered_pp_df["is_milk"], "Wert_x_GE"]
            )

        else:
            # Check that Wert_x_GE is unique for each (BetriebProcessID, FlowID) pair
            # When not considering GVE we expect that the Wert_x_GE is the same for each (BetriebProcessID, FlowID) pair, otherwise we have a conflict
            conflicts = filtered_pp_df.groupby(["BetriebProcessID", "FlowID"])[
                "Wert_x_GE"
            ].nunique()

            conflicts = conflicts[conflicts > 1]

            if not conflicts.empty:
                raise ValueError(
                    f"Found {len(conflicts)} (BetriebProcessID, FlowID) pairs with different Wert_x_GE values."
                )

            filtered_pp_df["Wert_x_GE_unweighted"] = filtered_pp_df["Wert_x_GE"]

            # Drop duplicates of (BetriebProcessID, FlowID) to avoid double counting in the sum

            if not (
                filtered_pp_df.groupby(["BetriebProcessID", "FlowID"]).agg(
                    Flows=("Pp_name", set)
                )["Flows"]
                == set(grouped_pp_names)
            ).all():
                print(
                    f"Warning: For Process {process_name}, some (BetriebProcessID, FlowID) pairs, not all grouped_pp_names are present. "
                    "Please check the input data."
                )

            filtered_pp_df["Pp_name"] = " | ".join(sorted(grouped_pp_names))

            # TODO: Implement check that maximum of 1 duplicate is present for each ("BetriebProcessID", "FlowID") combination
            filtered_pp_df.drop_duplicates(
                subset=["BetriebProcessID", "FlowID"], inplace=True
            )

    else:
        # TODO: do not implement weighting for Milchkuh Milchkuh/Mutterkuh

        # Tier 3 formula (coefficients a-e): Arbeitsgruppe BEK (2026): Methodenbeschreibung
        # Sachbilanz Tierhaltung der BEK-Fachgruppen. KTBL Handbuch, Darmstadt.
        # For Tier 3, we sum the different parameters of the Futtermittel with factors
        # Then we multiply by the Flow Wert
        # Then Divide by GVE_Rind, multiply by gve_weight and pp_Wert (which is the number of animals in that category)
        a = 0.079
        b = 0.01
        c = 0.026
        d = -0.212  # added -
        e = 365 * 0.063

        # BUGFIX (e-term): e is a flat per-animal-per-year constant, not a
        # feed-nutrient term. It must be counted exactly once per animal in
        # this category. Previously it was added inside the per-feed-row
        # polynomial below and then multiplied by that feed's Flow_Wert, so
        # it got scaled by feed mass and repeated once per distinct feed
        # type consumed. We compute it separately here from Pp_Wert
        # (headcount), once per (BetriebProcessID, Pp_name), so it survives
        # the later per-feed summation unaffected by feed mass or feed count.
        e_per_pp_df = filtered_pp_df.drop_duplicates(
            subset=["BetriebProcessID", "Pp_name"]
        )[["BetriebProcessID", "Pp_Wert"]]
        e_contribution_df = e_per_pp_df.groupby("BetriebProcessID", as_index=False)[
            "Pp_Wert"
        ].sum()
        e_contribution_df["Wert_x_Tier3_E"] = (
            e * e_contribution_df["Pp_Wert"] * (gve_weight if consider_gve else 1)
        )
        e_contribution_df = e_contribution_df[["BetriebProcessID", "Wert_x_Tier3_E"]]

        filtered_pp_df["Wert_x_Tier3_Formula"] = (
            a
            * (
                filtered_pp_df["XF (g/kg TM)"] / 1000
            )  # divide by 1000 to make sure it is kg in the formula
            + b * (filtered_pp_df["NFE  (g/kg TM)"] / 1000)
            + c * (filtered_pp_df["XP (g/kg TM)"] / 1000)
            + d * (filtered_pp_df["XL (g/kg TM)"] / 1000)
        ) * filtered_pp_df["Flow_Wert"]
        if consider_gve:
            if gve_weight is None:
                raise ValueError(
                    "consider_gve is True, but gve_weight is None. Please provide a gve_weight."
                )
            filtered_pp_df["Wert_x_Tier3_Formula_weighted"] = (
                filtered_pp_df["Wert_x_Tier3_Formula"]
                / filtered_pp_df["GVE_Rind"]
                * filtered_pp_df["Pp_Wert"]
            ) * gve_weight
        else:
            # BUGFIX (double counting): Wert_x_Tier3_Formula is identical for
            # every Pp_name sharing the same (BetriebProcessID, FlowID) - the
            # earlier merge with tbl_6 cross-joins each feed row with every
            # tracked Pp_name (e.g. "Milchkuh / Mutterkuh" AND "Deckbulle"),
            # even when the second category has 0 animals. Without collapsing
            # this back to one row per (BetriebProcessID, FlowID) before the
            # groupby-sum below, each feed item's contribution gets added
            # once per Pp_name present (2x here), silently doubling the
            # result. Mirrors the equivalent dedup already done for the
            # Tier 2 unweighted branch above.
            conflicts = filtered_pp_df.groupby(["BetriebProcessID", "FlowID"])[
                "Wert_x_Tier3_Formula"
            ].nunique()

            conflicts = conflicts[conflicts > 1]

            if not conflicts.empty:
                raise ValueError(
                    f"Found {len(conflicts)} (BetriebProcessID, FlowID) pairs with different Wert_x_Tier3_Formula values."
                )

            filtered_pp_df["Wert_x_Tier3_Formula_unweighted"] = filtered_pp_df[
                "Wert_x_Tier3_Formula"
            ]

            filtered_pp_df["Pp_name"] = " | ".join(sorted(grouped_pp_names))

            filtered_pp_df.drop_duplicates(
                subset=["BetriebProcessID", "FlowID"], inplace=True
            )

    # Group per BetriebProcessID and sum the values, aggregate other columns with first() since they should be unique per BetriebProcessID
    grouped_df = filtered_pp_df.groupby("BetriebProcessID", as_index=False)

    # TODO: Update for Milckuh with tier3 unweighted
    if is_tier_2 and consider_gve:
        sum_agg = ("Wert_x_GE_weighted", "sum")
    elif is_tier_2 and not consider_gve:
        sum_agg = ("Wert_x_GE_unweighted", "sum")
    elif not is_tier_2 and consider_gve:
        sum_agg = ("Wert_x_Tier3_Formula_weighted", "sum")
    elif not is_tier_2 and not consider_gve:
        sum_agg = ("Wert_x_Tier3_Formula_unweighted", "sum")
    else:
        raise ValueError(
            "Invalid combination of is_tier_2 and consider_gve. Please check the input parameters."
        )

    additional_cols = ["BetriebID", "ProcessID", "Process"]
    for col in additional_cols:
        assert (
            grouped_df[col].nunique()[col] == 1
        ).all(), f"Expected unique {col} per BetriebProcessID, but found multiple."

    # Aggregate the filtered_pp_df
    grouped_sum_df = filtered_pp_df.groupby("BetriebProcessID", as_index=False).agg(
        **{col: (col, "first") for col in additional_cols},
        **{
            "Flows": (
                "Flow",
                lambda s: " | ".join(sorted({str(x) for x in s.dropna()})),
            )
        },
        **{sum_agg[0]: (sum_agg[0], sum_agg[1])},
    )

    grouped_pp_df = (
        filtered_pp_df.groupby("BetriebProcessID")["Pp_name"]
        .apply(lambda s: " | ".join(sorted({str(x) for x in s.dropna()})))
        .reset_index(name="TierkategorieEmissionFutter")
    )
    # "Kälber" is also used as a feedstuff tag (see IsPotentialFeedstuff check above);
    # tagging the emission row the same way makes it indistinguishable from a feedstuff
    # row on the next run. Use a distinct label for the emission side.
    grouped_pp_df["TierkategorieEmissionFutter"] = grouped_pp_df[
        "TierkategorieEmissionFutter"
    ].replace({"Kälber": "Emission_Kälber"})

    calculation_result_df = grouped_sum_df.merge(
        grouped_pp_df, on="BetriebProcessID", how="left"
    )

    if is_tier_2:
        factor_to_take = (
            X_CH4_GE_CALVES if "Kälber" in grouped_pp_names else X_CH4_GE_NON_CALVES
        )
        column_to_use = "Wert_x_GE_weighted" if consider_gve else "Wert_x_GE_unweighted"
        calculation_result_df["EFT_Wert"] = (
            calculation_result_df[column_to_use] * factor_to_take / N_CH4
        )
    else:
        column_to_use = (
            "Wert_x_Tier3_Formula_weighted"
            if consider_gve
            else "Wert_x_Tier3_Formula_unweighted"
        )
        calculation_result_df["EFT_Wert"] = calculation_result_df[column_to_use]
        # BUGFIX (e-term): add back the per-animal constant computed above
        # (once per BetriebProcessID, based on headcount) instead of the
        # removed per-feed-row "+ e" term.
        calculation_result_df = calculation_result_df.merge(
            e_contribution_df, on="BetriebProcessID", how="left"
        )
        calculation_result_df["EFT_Wert"] += calculation_result_df[
            "Wert_x_Tier3_E"
        ].fillna(0)

    methane_flow_df = data["Katalog_Flow"][
        data["Katalog_Flow"]["Flow"] == "Methane, non-fossil"
    ].copy()
    if len(methane_flow_df) == 0:
        raise ValueError(
            "Flow 'Methane, non-fossil' wurde in Katalog_Flow nicht gefunden."
        )
    if len(methane_flow_df["FlowID"].drop_duplicates()) > 1:
        raise ValueError(
            "Flow 'Methane, non-fossil' ist in Katalog_Flow nicht eindeutig. "
            "Bitte präzisieren (z.B. über Category)."
        )

    methane_flow_id = methane_flow_df["FlowID"].drop_duplicates().iloc[0]

    tbl4_calculation_results_df = calculation_result_df[
        [
            "BetriebID",
            "ProcessID",
            "Process",
            "BetriebProcessID",
            "Flows",
            "TierkategorieEmissionFutter",
            "EFT_Wert",
        ]
    ].copy()
    tbl4_calculation_results_df["Flow"] = "Methane, non-fossil"
    tbl4_calculation_results_df["FlowID"] = methane_flow_id
    tbl4_calculation_results_df["Direction"] = "Output"
    tbl4_calculation_results_df.rename(
        columns={"EFT_Wert": "Wert", "Flows": "Notizen"}, inplace=True
    )

    return tbl4_calculation_results_df


ch4_emissions_aufzucht_milchvieh_df = calculate_ch4_emissions_for_process_pp_group(
    target_tbl4_input_with_futtermittelinfo_df,
    process_name="P_Aufzucht_Milchvieh",
    grouped_pp_names=("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
    consider_gve=False,  # If False: summary without weighting
    gve_weight=None,
    is_tier_2=True,
)


ch4_emissions_rind_mast_df = calculate_ch4_emissions_for_process_pp_group(
    target_tbl4_input_with_futtermittelinfo_df,
    process_name="P_Rind_Mast",
    grouped_pp_names=("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
    consider_gve=False,  # If False: summary without weighting
    gve_weight=None,
    is_tier_2=True,
)

ch4_emissions_mutterkuh_mast_aufzucht_mast_df = (
    calculate_ch4_emissions_for_process_pp_group(
        target_tbl4_input_with_futtermittelinfo_df,
        process_name="P_Mutterkuh",
        grouped_pp_names=("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
        consider_gve=True,
        gve_weight=0.6,
        is_tier_2=True,
    )
)

ch4_emissions_mutterkuh_kalber_df = calculate_ch4_emissions_for_process_pp_group(
    target_tbl4_input_with_futtermittelinfo_df,
    process_name="P_Mutterkuh",
    grouped_pp_names=("Kälber",),
    consider_gve=True,
    gve_weight=0.3,
    is_tier_2=True,
)


ch4_emissions_mutterkuh_milchkuh_mutterkuh_deckbulle_df = (
    calculate_ch4_emissions_for_process_pp_group(
        target_tbl4_input_with_futtermittelinfo_df,
        process_name="P_Mutterkuh",
        grouped_pp_names=("Milchkuh / Mutterkuh", "Deckbulle"),
        consider_gve=True,
        gve_weight=1,
        is_tier_2=False,  # Tier 3 calculation, therefore False
    )
)

ch4_emissions_milchkuh_milchkuh_mutterkuh_deckbulle_df = (
    calculate_ch4_emissions_for_process_pp_group(
        target_tbl4_input_with_futtermittelinfo_df,
        process_name="P_Milchkuh",
        grouped_pp_names=("Milchkuh / Mutterkuh", "Deckbulle"),
        consider_gve=False,
        gve_weight=1,
        is_tier_2=False,  # Tier 3 calculation, therefore False
    )
)

ch4_emissions_milchkuh_kalber_df = calculate_ch4_emissions_for_process_pp_group(
    target_tbl4_input_with_futtermittelinfo_df,
    process_name="P_Milchkuh",
    grouped_pp_names=("Kälber",),
    consider_gve=False,
    gve_weight=None,
    is_tier_2=True,
)


# Combine all calculation results in one DataFrame
final_tbl4_calculation_results_df = pd.concat(
    [
        ch4_emissions_aufzucht_milchvieh_df,
        ch4_emissions_rind_mast_df,
        ch4_emissions_mutterkuh_mast_aufzucht_mast_df,
        ch4_emissions_mutterkuh_kalber_df,
        ch4_emissions_mutterkuh_milchkuh_mutterkuh_deckbulle_df,
        ch4_emissions_milchkuh_milchkuh_mutterkuh_deckbulle_df,
        ch4_emissions_milchkuh_kalber_df,
    ],
    ignore_index=True,
).sort_values(["BetriebID", "ProcessID", "TierkategorieEmissionFutter"])


# TODOS:
# - treat process Rind Mast the same as P_Aufzucht_Milchvieh
# - P_Mutterkuh: Mast + Aufzucht/Mast the same as P_Aufzucht Milchvieh, but with the 0.6 formula
#    - get GVE Rind (total number of cows on the farm. tbl_2 master data. StammdatenID: catalogue master data "GVE Rind").
#    - tbl_6 refers to Katalog_ProcessProcessparameter, which refers to Katalog_Processparameter -> Pp_name Mast and Aufzucht/Mast


# - P_Mutterkuh: Kälber works the same way, but with 0.3 instead of 0.6 and of course only with the calf parameter


# - P_Mutterkuh: Milchkuh/Mutterkuh with the Tier 3 formula (all data already available), but must be weighted (divided by GVE Rind) times (Milchkuh/Mutterkuh parameter + Deckbulle)

# - P_Milchkuh: needs no weighting at all, because the feed split is done via the "Kälberfutter" entry

# ===============================================================================
# Manual check calculations for the example farms 8 (BetriebProcessID 31) and
# 2 (BetriebProcessID 494), to validate the calculation. The input values
# are taken from the example dataset; for other data remove or adapt
# this block.
# ==============================================================================

# P_Mutterkuh: Mast + Aufzucht/Mast

## Process P_Mutterkuh ID 2 for farm ID 8 -> BetriebProcessID 31
## Only the PPs Mast + Aufzucht/Mast here (0.6 formula) (from Katalog_ProcessParameter): PPIDs: 58, 60
## From Katalog_ProcessProcessParameter, ProcessPPIDs for Process ID 2: 58: 77, 60: 79
## From tbl_6_Betrieb_ProcessParameterValue, PP values: 77: 30, 79: 16
## Flow IDs (input) for BPID (from tbl_4_Betrieb_Process_FlowValue): 12, 14, 20, 242,
## Of these, only those with IsPotentialFeedstuff = True (from Katalog_Flow): 12, 14, 20, 242
## Their flow values (Wert) from tbl_4_Betrieb_Process_FlowValue: 12: 72196, 14: 21926, 20: 245536, 242: 81096
## For the flow IDs, the BetriebProcessFlowValueIDs from tbl_4_Betrieb_Process_FlowValue: 12: 34, 14: 35, 20: 36, 242: 37
## FuttermittelID for flow IDs from tbl_Futtermittel_Mapping: 12: 29, 14: 38, 20: 2, 242: 35
## GE FlowPropertyID: 28 (to look up farm-specific overrides). In tbl_5_Betrieb_FlowpropertyValue: no entries for BetriebProcessFlowValueIDs 34, 35, 36, 37. So no overrides.
## GE values from Katalog_Futtermittelgehalte KTBL: 29: 18.69, 38: 18.46, 2: 18.59, 35: 18.44
## GVE_Rind (Stammdaten ID 22) for BetriebProcessID 31 from tbl_2_Betrieb_Stammdaten: 79.5

## Calculation

betrieb_process_id_31_manual = (
    (18.69 * 72196 + 18.46 * 21926 + 18.59 * 245536 + 18.44 * 81096)
    / 79.5
    * (30 + 16)
    * 0.6
    * X_CH4_GE_NON_CALVES
    / N_CH4
)
print(
    "ID 31 calculated: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Mast")
        )
    ]["Wert"].iloc[0],
)
print("ID 31 manual: ", betrieb_process_id_31_manual)
print(
    "ID 31 difference: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Mast")
        )
    ]["Wert"].iloc[0]
    - betrieb_process_id_31_manual,
)

assert math.isclose(
    betrieb_process_id_31_manual,
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Mast")
        )
    ].iloc[0]["Wert"],
    rel_tol=1e-9,
)

# P_Mutterkuh_Kälber: Kälber
# Farm ID 8 Process ID 2 -> BetriebProcessID 31
# Only the PPs Kälber here (from Katalog_ProcessParameter): PPIDs: 57
# From Katalog_ProcessProcessParameter, ProcessPPIDs for Process ID 2: 57: 76
# From tbl_6_Betrieb_ProcessParameterValue, PP values: 76: 19
# Flow IDs (input) for BPID (from tbl_4_Betrieb_Process_FlowValue): 12, 14, 20, 242
# Of these, only those with IsPotentialFeedstuff = True (from Katalog_Flow): 12, 14, 20, 242
# Their flow values (Wert) from tbl_4_Betrieb_Process_FlowValue: 12: 72196, 14: 21926, 20: 245536, 242: 81096
# For the flow IDs, the BetriebProcessFlowValueIDs from tbl_4_Betrieb_Process_FlowValue: 12: 34, 14: 35, 20: 36, 242: 37
# FuttermittelID for flow IDs from tbl_Futtermittel_Mapping: 12: 29, 14: 38, 20: 2, 242: 35
# GE FlowPropertyID: 28 (to look up farm-specific overrides). In tbl_5_Betrieb_FlowpropertyValue: no entries for BetriebProcessFlowValueIDs 34, 35, 36, 37. So no overrides.
# GE values from Katalog_Futtermittelgehalte KTBL: 29: 18.69, 38: 18.46, 2: 18.59, 35: 18.44
# GVE_Rind (Stammdaten ID 22) for BetriebProcessID 31 from tbl_2_Betrieb_Stammdaten: 79.5

betrieb_process_id_31_calf_manual = (
    (
        (18.69 * 72196 + 18.46 * 21926 + 18.59 * 245536 + 18.44 * 81096)
        / 79.5
        * 19
        * 0.3
        + (
            92058.12 * 1.03 * 135 / 1000 * 26.4
        )  # Kälbermilch_L (from temp_calf_milk_intake.pkl, BetriebProcessID=31, Process=P_Mutterkuh)
        # FIX 2026-10-06: L fresh milk * 1.03 kg/L * 135 g DM/kg -> kg DM, then * GE 26.4 MJ/kg DM
    )
    * X_CH4_GE_CALVES
    / N_CH4
)


print(
    "ID 31 Kälber calculated: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df["TierkategorieEmissionFutter"].apply(
                lambda row: "Kälber" in row
            )
        )
    ].iloc[0]["Wert"],
)
print("ID 31 Kälber manual: ", betrieb_process_id_31_calf_manual)
print(
    "Difference: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Kälber")
        )
    ].iloc[0]["Wert"]
    - betrieb_process_id_31_calf_manual,
)

assert math.isclose(
    betrieb_process_id_31_calf_manual,
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Kälber")
        )
    ].iloc[0]["Wert"],
    rel_tol=1e-9,
)

# P_Milchkuh_Milchkuh/Mutterkuh + Deckbulle: Milchkuh/Mutterkuh + Deckbulle (Tier 3)
# Farm ID 2 Process ID 1 -> BetriebProcessID 494
# Tier 3 formula as in calculate_ch4_emissions_for_process_pp_group (a,b,c,d,e). For
# P_Milchkuh consider_gve=False -> no GVE weighting, no X_CH4/N_CH4 factor at the
# end (unlike Tier 2): EFT = sum(Wert_x_Tier3_Formula per flow) + e * Pp_Wert_total * 1
# Flow IDs (input, potential feedstuff, not tagged Kälber) from tbl_4_Betrieb_Process_FlowValue:
# 23, 232, 20, 12, 109, 276, 182
# Flow values (as of 2026-10-19, corrected cow ration of farm 2, checked directly in the DB):
# 23: 175791, 232: 134176, 20: 81030, 12: 248924, 109: 28361, 276: 101288, 182: 64824
# BetriebProcessFlowValueIDs: 23: 1841, 232: 1842, 20: 1843, 12: 1844, 109: 1845, 276: 1846, 182: 1847
# FuttermittelID: 23: 22, 232: 96, 20: 2, 12: 29, 109: 83, 276: 87, 182: 57
# tbl_5_Betrieb_FlowpropertyValue overrides (IsBetriebswert=True) for XP/XF/XL/NFE, checked directly in the DB:
#   1841 (FlowID 23): XP=127, XF=278, XL=33, NFE=434 (all four overridden)
#   1842 (FlowID 232): XP=71, XF=161, XL=37, NFE=687 (all four overridden)
#   1844 (FlowID 12): XP=143, XF=276, XL=28 overridden, NFE NOT overridden -> KTBL 422
#   1843, 1845, 1846, 1847 (FlowID 20, 109, 276, 182): no overrides -> pure KTBL values
# KTBL values (XF/NFE/XP/XL in g/kg DM): 20: 205/472/195/39, 109: 29/794/138/20, 276: 85/648/205/43, 182: 89/557/298/16
# Pp_Wert (tbl_6): Milchkuh / Mutterkuh: 110, Deckbulle: 0

a_tier3 = 0.079
b_tier3 = 0.01
c_tier3 = 0.026
d_tier3 = -0.212
e_tier3 = 365 * 0.063


def tier3_term(xf, nfe, xp, xl, wert):
    return (
        a_tier3 * (xf / 1000)
        + b_tier3 * (nfe / 1000)
        + c_tier3 * (xp / 1000)
        + d_tier3 * (xl / 1000)
    ) * wert


betrieb_process_id_494_manual = (
    tier3_term(278, 434, 127, 33, 175791)  # FlowID 23, Silo_DG, tbl_5-Override
    + tier3_term(161, 687, 71, 37, 134176)  # FlowID 232, Silomais, tbl_5-Override
    + tier3_term(205, 472, 195, 39, 81030)  # FlowID 20, pasture intake, KTBL
    + tier3_term(
        276, 422, 143, 28, 248924
    )  # FlowID 12, Silo_LKG, tbl_5-Override (NFE: KTBL)
    + tier3_term(29, 794, 138, 20, 28361)  # FlowID 109, WW, KTBL
    + tier3_term(
        85, 648, 205, 43, 101288
    )  # FlowID 276, Milchleistungsfutter 20_4, KTBL
    + tier3_term(89, 557, 298, 16, 64824)  # FlowID 182, Ackerbohne, KTBL
    + e_tier3
    * (110 + 0)
    * 1  # e term: Pp_Wert Milchkuh/Mutterkuh(110)+Deckbulle(0), gve_weight factor=1 (consider_gve=False)
)

print(
    "ID 494 Milchkuh/Deckbulle calculated: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 494)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Deckbulle")
        )
    ].iloc[0]["Wert"],
)
print("ID 494 Milchkuh/Deckbulle manual: ", betrieb_process_id_494_manual)
print(
    "Difference: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 494)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Deckbulle")
        )
    ].iloc[0]["Wert"]
    - betrieb_process_id_494_manual,
)

assert math.isclose(
    betrieb_process_id_494_manual,
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 494)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Deckbulle")
        )
    ].iloc[0]["Wert"],
    rel_tol=1e-9,
)

# P_Mutterkuh_Milchkuh/Mutterkuh + Deckbulle: Milchkuh/Mutterkuh + Deckbulle (Tier 3, GVE-weighted)
# Farm ID 8 Process ID 2 -> BetriebProcessID 31
# consider_gve=True, gve_weight=1 for P_Mutterkuh -> the feed sum is divided once by GVE_Rind
# and multiplied by (Pp_Wert Milchkuh/Mutterkuh + Pp_Wert Deckbulle) (same derivation as for the
# Mast+Aufzucht check above: the 2 Pp_names are duplicated per flow and each copy is summed weighted
# with its own Pp_Wert - equivalent to a single division/multiplication of the sum).
# e term also multiplied by gve_weight (here 1), because consider_gve=True.
# Flow IDs (input, potential feedstuff, not tagged Kälber): 242, 12, 20, 14
# Flow values: 242: 81096, 12: 72196, 20: 245536, 14: 21926
# BetriebProcessFlowValueIDs: 242: 34, 12: 35, 20: 36, 14: 37
# FuttermittelID: 242: 35, 12: 29, 20: 2, 14: 38
# tbl_5_Betrieb_FlowpropertyValue: no overrides for BetriebProcessFlowValueIDs 34-37 -> pure KTBL values
# KTBL values (XF/NFE/XP/XL in g/kg DM): 242: 260/507/140/25, 12: 300/422/154/40, 20: 205/472/195/39, 14: 300/419/180/20
# Pp_Wert (tbl_6): Milchkuh / Mutterkuh: 43, Deckbulle: 2
# GVE_Rind (Stammdaten ID 22) for BetriebID 8: 79.5

betrieb_process_id_31_milchkuh_deckbulle_manual = (
    tier3_term(260, 507, 140, 25, 81096)  # FlowID 242, Heu_DG
    + tier3_term(300, 422, 154, 40, 72196)  # FlowID 12, Silo_LKG
    + tier3_term(205, 472, 195, 39, 245536)  # FlowID 20, pasture intake
    + tier3_term(300, 419, 180, 20, 21926)  # FlowID 14, Heu_LKG
) / 79.5 * (  # GVE_Rind BetriebID 8
    43 + 2
) * 1 + e_tier3 * (  # Pp_Wert Milchkuh/Mutterkuh(43) + Deckbulle(2)  # gve_weight
    43 + 2
) * 1  # e term, also weighted with gve_weight (consider_gve=True)

print(
    "ID 31 Milchkuh/Deckbulle calculated: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Deckbulle")
        )
    ].iloc[0]["Wert"],
)
print(
    "ID 31 Milchkuh/Deckbulle manual: ", betrieb_process_id_31_milchkuh_deckbulle_manual
)
print(
    "Difference: ",
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Deckbulle")
        )
    ].iloc[0]["Wert"]
    - betrieb_process_id_31_milchkuh_deckbulle_manual,
)

assert math.isclose(
    betrieb_process_id_31_milchkuh_deckbulle_manual,
    final_tbl4_calculation_results_df[
        (final_tbl4_calculation_results_df["BetriebProcessID"] == 31)
        & (
            final_tbl4_calculation_results_df[
                "TierkategorieEmissionFutter"
            ].str.contains("Deckbulle")
        )
    ].iloc[0]["Wert"],
    rel_tol=1e-9,
)

print("\n=== STEP 7: final_tbl4_calculation_results_df ===")
print(f"{len(final_tbl4_calculation_results_df)} Zeilen")
print(final_tbl4_calculation_results_df)

excel_export_path = (
    OUTPUT_DIR / f"final_tbl4_calculation_results_{date.today().isoformat()}.xlsx"
)
final_tbl4_calculation_results_df.to_excel(excel_export_path, index=False)
print(f"Excel-Export: {excel_export_path}")


# ============================================================================
# CREATE FORMAT FOR tbl_4_Betrieb_Process_FlowValue
# ============================================================================

insert_emissions_df = final_tbl4_calculation_results_df[
    ["BetriebProcessID", "FlowID", "Wert", "TierkategorieEmissionFutter"]
].copy()

insert_emissions_df["Direction"] = "Output"

print(insert_emissions_df.head())
if len(insert_emissions_df) > 0:
    print(insert_emissions_df.sample(min(10, len(insert_emissions_df))))
else:
    print("No emissions rows to sample from.")
print(f"{len(insert_emissions_df)} Emissionszeilen vorbereitet")


# ============================================================================
# ROBUST UPSERT: UPDATE EXISTING EMISSIONS, ADD NEW ONES
# ============================================================================

cursor = conn.cursor()

# Load existing emission rows from Access to avoid writing duplicates
existing_df = pd.read_sql(
    """
    SELECT BetriebProcessID, FlowID, Direction, BetriebProcessFlowValueID, TierkategorieEmissionFutter
    FROM tbl_4_Betrieb_Process_FlowValue
    WHERE Direction = 'Output'
    """,
    conn,
)

# Link with the newly calculated emissions
merge_df = insert_emissions_df.merge(
    existing_df,
    on=["BetriebProcessID", "FlowID", "Direction", "TierkategorieEmissionFutter"],
    how="left",
)

# Existing rows -> UPDATE
update_rows = merge_df[merge_df["BetriebProcessFlowValueID"].notna()]

# Rows that do not exist yet -> INSERT
insert_rows = merge_df[merge_df["BetriebProcessFlowValueID"].isna()]

# Summary of the planned changes
print(f"{len(update_rows)} vorhandene Emissionen werden aktualisiert.")
print(f"{len(insert_rows)} neue Emissionen werden eingefügt.")

# Show new emissions
if len(insert_rows) > 0:
    print("\n=== NEUE EMISSIONEN DIE HINZUGEFÜGT WERDEN ===")
    # Enrich with flow names and process names for readability
    new_emissions_display = (
        insert_rows.merge(
            data["Katalog_Flow"][["FlowID", "Flow", "Category"]],
            on="FlowID",
            how="left",
        )
        .merge(
            data["tbl_3_Betrieb_Process"][["BetriebProcessID", "ProcessID"]],
            on="BetriebProcessID",
            how="left",
        )
        .merge(
            data["Katalog_Process"][["ProcessID", "Process"]],
            on="ProcessID",
            how="left",
        )
    )
    print(
        new_emissions_display[
            [
                "BetriebProcessID",
                "Process",
                "Flow",
                "Category",
                "Wert",
                "TierkategorieEmissionFutter",
            ]
        ].to_string(index=False)
    )
    print("=" * 50)

confirm = input("Änderungen in die Access-Datenbank schreiben? (j/n): ")

if confirm.lower() == "j":

    # Update existing values
    for _, row in update_rows.iterrows():
        cursor.execute(
            """
            UPDATE tbl_4_Betrieb_Process_FlowValue
            SET Wert = ?
            WHERE BetriebProcessFlowValueID = ?
            """,
            row["Wert"],
            int(row["BetriebProcessFlowValueID"]),
        )

    # Insert new values
    if len(insert_rows) > 0:
        rows_to_insert = list(
            insert_rows[
                [
                    "BetriebProcessID",
                    "FlowID",
                    "Wert",
                    "Direction",
                    "TierkategorieEmissionFutter",
                ]
            ].itertuples(index=False, name=None)
        )

        cursor.executemany(
            """
            INSERT INTO tbl_4_Betrieb_Process_FlowValue
            (BetriebProcessID, FlowID, Wert, Direction, TierkategorieEmissionFutter)
            VALUES (?, ?, ?, ?,?)
            """,
            rows_to_insert,
        )
    else:
        print("Keine neuen Emissionen zum Einfügen.")

    conn.commit()


conn.close()
