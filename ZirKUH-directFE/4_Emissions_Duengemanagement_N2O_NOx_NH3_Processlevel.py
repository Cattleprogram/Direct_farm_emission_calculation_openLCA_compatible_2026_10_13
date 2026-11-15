import pyodbc
import pandas as pd
from config import db_path, ktbl_parameter_path

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
print(f"\n📊 Lade alle Tabellen aus: {db_path}\n")
data = load_all_tables(conn)

print(f"\n✅ Total {len(data)} Tabellen geladen\n")


# List available tables
print("Verfügbare Tabellen für Berechnungen:")
for table_name in data.keys():
    print(
        f"  - {table_name}: {len(data[table_name])} Zeilen, {len(data[table_name].columns)} Spalten"
    )
# CATTLE: BUILD THE DATA BASIS FOR LATER CALCULATIONS
# ============================================================================


# Process -> relevant process parameters (Pp_name)
TARGET_PROCESSES = {
    "P_Milchkuh": ("Milchkuh / Mutterkuh", "Kälber"),
    "P_Aufzucht_Milchvieh": ("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
    "P_Mutterkuh": (
        "Milchkuh / Mutterkuh",
        "Kälber",
        "♂ Mast",
        "♀ Aufzucht / Masttiere (je nach MV oder MK)",
    ),
    "P_Rind_Mast": ("♂ Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"),
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

print("\n=== RELEVANTE PROZESSPARAMETER (CATTLE) ===")
print(
    target_pp_df[
        ["BetriebID", "BetriebProcessID", "ProcessID", "Process", "Pp_name", "Wert"]
    ].head(20)
)
print(f"{len(target_pp_df)} relevante Parameterzeilen gefunden.")

# 2) Get flows from tbl_4 and filter on IsPotentialFeedstuff == True
# ----------------------------------------------------------------------------
# FIX 2026-10-06 (fan-out removed):
# Previously "Pp_name" was also merged here. Since target_pp_df has one row
# per BetriebProcessID AND animal category (Pp_name), every
# tbl_4 feed row was multiplied per animal category (P_Milchkuh 2x,
# P_Aufzucht_Milchvieh 2x, P_Mutterkuh 4x) and feed_N_kg (-> m_feedlo) further
# down was counted several times accordingly. This fan-out is only needed in
# 5_Emissions_Tierhaltung_CH4_Cattle.py (weighting per animal category);
# here only the feed amount per process instance is needed.
# The merge now only restricts the data to the cattle processes in
# TARGET_PROCESSES (one row per BetriebProcessID, no Pp_name).
# Old:
#     target_pp_df[
#         ["BetriebProcessID", "BetriebID", "ProcessID", "Process", "Pp_name"]
#     ].drop_duplicates(),
# ----------------------------------------------------------------------------
flows_with_process_df = data["tbl_4_Betrieb_Process_FlowValue"].merge(
    target_pp_df[
        ["BetriebProcessID", "BetriebID", "ProcessID", "Process"]
    ].drop_duplicates(subset=["BetriebProcessID"]),
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

# FIX 2026-10-06 (fan-out removed): explicitly mark the calf ration for P_Milchkuh.
# The calf feed rows created by 1_Update_P_Milchkuh_Kaelberration.py
# (TierkategorieEmissionFutter == "Kälber") deliberately stay in the feed N, because mexcr
# for P_Milchkuh includes the calves (5.5 kg N/calf). For P_Mutterkuh,
# P_Aufzucht_Milchvieh and P_Rind_Mast the herd ration is already complete.
potential_feed_flows_df["Futterration"] = "Herde / Kuhration"
potential_feed_flows_df.loc[p_milchkuh_kälber_mask[filter_mask], "Futterration"] = (
    "Kälberration (P_Milchkuh)"
)

print(
    "\n=== RELEVANTE FLOWS JE TARGET PROCESS "
    "(Basis: IsPotentialFeedstuff; bei P_Milchkuh inkl. Kälberration) ==="
)
print(
    potential_feed_flows_df[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "Futterration",  # FIX 2026-10-06: instead of "Pp_name" (no longer exists)
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

# Defensive deduplication: avoid 1-to-many mappings causing duplicated rows
# during merges. Keep unique (FlowID, FuttermittelID) pairs and, if multiple
# entries exist for the same FlowID, keep the first one.
if {"FlowID", "FuttermittelID"}.issubset(futtermittel_mapping_df.columns):
    futtermittel_mapping_df = futtermittel_mapping_df.drop_duplicates(
        subset=["FlowID", "FuttermittelID"]
    )
    # If still multiple mappings per FlowID, reduce to a single mapping to
    # prevent multiplication of rows downstream. This is defensive; since
    # you fixed the DB, this will usually be a no-op.
    if futtermittel_mapping_df.groupby("FlowID").size().gt(1).any():
        print(
            "⚠️ Doppelte Mapping-Einträge für FlowID gefunden — reduziere auf erstes Mapping pro FlowID"
        )
        futtermittel_mapping_df = futtermittel_mapping_df.drop_duplicates(
            subset=["FlowID"], keep="first"
        )

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
    "Futterration",  # FIX 2026-10-06: carry the calf ration flag along
    "FlowID",
    "Flow",
    "Wert",
    "FuttermittelID",
]
tbl4_input_cols = [
    c
    for c in tbl4_input_cols
    if c in target_potential_feed_flows_with_futtermittel_df.columns
]

target_tbl4_input_df = target_potential_feed_flows_with_futtermittel_df[
    tbl4_input_cols
].copy()

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

print("\n=== STEP 5: Eindeutiges FlowProperty -> KTBL-Spalten-Matching ===")
print(flowproperty_match_df.head(30))

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
# ============================================================================
# Prepare the calculations. First the total nitrogen excretions (mexcr)
# and TAN excretions (mexcr, TAN) are taken from DLG e. V. (2014): Bilanzierung der Nährstoffausscheidungen landwirtschaftlicher Nutztiere. Arbeiten der DLG 199,
# Frankfurt am Main, Deutsche Landwirtschafts-Gesellschaft, 2nd ed., to use them in the later calculations for
# N2O, NOx and NH3.
# The calculations are based on the entries in the animal modules, but the emissions
# are then all written to the process 1_Jahr_Duengemanagement!
# ============================================================================
print(
    "=== Vorbereitung der Berechnungen: Gesamtstickstoffausscheidungen (mexcr) und TAN Ausscheidungen (TAN) ==="
)
# Total nitrogen excretions (mexcr)
# =============================================================================
# P_Milchkuh:
# 5.5 kg N per year and calf + 117 kg N per year and cow (Table 1.6, p. 37)
# ECM = FPCM * ((0.38*4 + 0.21 * 3.3 + 1.05) / 3.28)

# P_Mutterkuh:
# 114.5 kg N per year and cow with calf + 54 kg N per year and rearing or fattening animal (Table 1.4, p. 33)

# P_Aufzucht_Milchvieh:
# 54 kg N per year and rearing or fattening animal (Table 1.6, p. 37)

# P_Rind_Mast:
# 54 kg N per year and rearing or fattening animal (Table 1.6, p. 37)

# P_Mastschwein:
# 4.48 kg N per year and fattening pig (Table 2.13, p. 72)

# P_Schafhaltung_Mast:
# 20.1 kg N per year and fattening sheep (Table 5.1, p. 95)

# P_Legehennen:
# 0.764 kg N per year and laying hen (Table 3.3, p. 78) (number of laying hens)

# P_Bruderhahnmast:
# 0.311 kg N per year and brother rooster (Table 3.7, p. 82) (number of fattening chickens)

# P_Pferdehaltung:
# 53.6 kg N per year and horse (Table 4.2, p. 92) (horse)

# ===========================================================================
# TAN excretions (mexcr, TAN)
# ===========================================================================
# ============================================================================
# STICKSTOFFAUSSCHEIDUNGEN (MEXCR) - TOTAL NITROGEN EXCRETION CALCULATION
# ============================================================================
# Sources:
# DLG e. V. (2014): Bilanzierung der Nährstoffausscheidungen landwirtschaftlicher Nutztiere
# Arbeiten der DLG 199, Frankfurt am Main, Deutsche Landwirtschafts-Gesellschaft, 2. Aufl.
# The ECM-based derivation of the N excretion was not applied, because the
# underlying milk yield data partly contain only the marketed milk
# amount. With dam- or nurse-cow rearing of calves, the milk consumed by the
# calves is not recorded, so the actual milk yield
# and thus the N excretion are systematically underestimated. This led
# in single cases to calculated TAN shares > 1. To ensure a
# consistent and biologically plausible parametrisation, three different
# MEXCR factors are therefore used for dairy cows, depending on the annual yield (kg milk A+B per cow and year):
# - < 6000 kg milk/year: 103 kg N/cow/year
# - 6000-8000 kg milk/year: 117 kg N/cow/year
# - > 8000 kg milk/year: 133.9 kg N/cow/year
# Define mexcr values (kg N per year) for each process and animal category
# If correct milk recording (MLP) yield data and urea contents are available for all farms,
# the BEK method based on ECM should definitely be applied next time
MEXCR_FACTORS = {
    "P_Milchkuh": {
        "Milchkuh / Mutterkuh": {
            "formula": "direct",
            "value": 117,
            "unit": "kg N/Kuh/Jahr",
        },
        "Kälber": {"formula": "direct", "value": 5.5, "unit": "kg N/Kalb/Jahr"},
        # FIX 2026-10-06 (breeding bull): breeding bulls were previously ignored completely,
        # because without an entry here the tbl_6 parameter "Deckbulle" is filtered out.
        # 53 kg N/animal/year according to DLG (2014), Table 1.8a.
        "Deckbulle": {"formula": "direct", "value": 53, "unit": "kg N/Tier/Jahr"},
    },
    "P_Mutterkuh": {
        "Milchkuh / Mutterkuh": {
            "formula": "direct",
            "value": 114.5,
            "unit": "kg N/Kuh+Kalb/Jahr",
        },
        "Kälber": {"formula": "included_in_mutterkuh", "value": 0, "unit": "kg N"},
        # FIX 2026-10-06 (breeding bull): DLG (2014), Table 1.8a
        "Deckbulle": {"formula": "direct", "value": 53, "unit": "kg N/Tier/Jahr"},
        "♂ Mast": {"formula": "direct", "value": 54, "unit": "kg N/Tier/Jahr"},
        "♀ Aufzucht / Masttiere (je nach MV oder MK)": {
            "formula": "direct",
            "value": 54,
            "unit": "kg N/Tier/Jahr",
        },
    },
    "P_Aufzucht_Milchvieh": {
        "♂ Mast": {"formula": "direct", "value": 54, "unit": "kg N/Tier/Jahr"},
        "♀ Aufzucht / Masttiere (je nach MV oder MK)": {
            "formula": "direct",
            "value": 54,
            "unit": "kg N/Tier/Jahr",
        },
    },
    "P_Rind_Mast": {
        "♂ Mast": {"formula": "direct", "value": 54, "unit": "kg N/Tier/Jahr"},
        "♀ Aufzucht / Masttiere (je nach MV oder MK)": {
            "formula": "direct",
            "value": 54,
            "unit": "kg N/Tier/Jahr",
        },
        # FIX 2026-10-06 (breeding bull): DLG (2014), Table 1.8a
        "Deckbulle": {"formula": "direct", "value": 53, "unit": "kg N/Tier/Jahr"},
    },
    "P_Mastschwein": {
        "Mastschwein": {"formula": "direct", "value": 4.48, "unit": "kg N/Tier/Jahr"},
    },
    "P_Schafhaltung_Mast": {
        "Mutterschaf": {
            "formula": "direct",
            "value": 20.1,
            "unit": "kg N/Tier/Jahr",
        },
        "Bock": {
            "formula": "direct",
            "value": 20.1,
            "unit": "kg N/Tier/Jahr",
        },
        "Lämmer": {
            "formula": "direct",
            "value": 20.1,
            "unit": "kg N/Tier/Jahr",
        },
        "Mastschaf": {
            "formula": "direct",
            "value": 20.1,
            "unit": "kg N/Tier/Jahr",
        },
    },
    "P_Legehennen": {
        "Anzahl Legehennen": {
            "formula": "direct",
            "value": 0.764,
            "unit": "kg N/Tier/Jahr",
        },
    },
    "P_Bruderhahnmast": {
        "Anzahl Masthähnchen": {
            "formula": "direct",
            "value": 0.311,
            "unit": "kg N/Tier/Jahr",
        },
    },
    "P_Pferdehaltung": {
        "Pferd": {"formula": "direct", "value": 53.6, "unit": "kg N/Tier/Jahr"},
    },
}


def calculate_ecm_from_fpcm(fpcm):
    """
    ECM = FPCM * ((0,38*4 + 0,21 * 3,3 + 1,05) / 3,28)
    ECM = FPCM * ((1.52 + 0.693 + 1.05) / 3.28)
    ECM = FPCM * (3.263 / 3.28)
    ECM = FPCM * 0.9954
    """
    if pd.isna(fpcm) or fpcm == 0:
        return 0
    return fpcm * ((0.38 * 4 + 0.21 * 3.3 + 1.05) / 3.28)


def extract_numeric_value(val):
    """Robustly converts values to numeric, replacing commas with points."""
    if pd.isna(val):
        return None
    val_str = str(val).strip().replace(",", ".")
    try:
        return float(val_str)
    except (ValueError, TypeError):
        return None


MEXCR_MILCHKUH_PRODUCTION_PARAMETER_NAMES = [
    "Ø Jahresleistung Milch kg A+B/ Kuh/ Jahr",
]


def get_mexcr_factor_for_milchkuh(production_value):
    """Selects the MEXCR factor for dairy cows based on the annual milk yield."""
    production_value = extract_numeric_value(production_value)
    if pd.isna(production_value):
        return 117.0
    if production_value < 6000:
        return 103.0
    if production_value <= 8000:
        return 117.0
    return 133.9


def calculate_mexcr_for_process(process_name, data_dict):
    """
    Calculates the total nitrogen excretion (mexcr) for one process
    based on animal numbers from the process parameters.

    Returns:
        DataFrame with BetriebID, BetriebProcessID, Process, Tierkategorie,
        Anzahl, mexcr_factor (kg N/animal), and Gesamt_mexcr (kg N total)
    """
    if process_name not in MEXCR_FACTORS:
        print(f"⚠ Prozess '{process_name}' nicht in MEXCR_FACTORS definiert.")
        return pd.DataFrame()

    # Get all process parameter values for this process
    tbl6 = data_dict["tbl_6_Betrieb_ProcessParameterValue"].copy()
    tbl3 = data_dict["tbl_3_Betrieb_Process"].copy()
    pp_cat = data_dict["Katalog_ProcessProcessparameter"].copy()
    pp_def = data_dict["Katalog_Processparameter"].copy()
    proc_cat = data_dict["Katalog_Process"].copy()

    # Merge to get full process parameter information
    merged = (
        tbl6.merge(tbl3, on="BetriebProcessID", how="left")
        .merge(pp_cat[["ProcessPpID", "PpID"]], on="ProcessPpID", how="left")
        .merge(pp_def[["PpID", "Pp_name"]], on="PpID", how="left")
        .merge(proc_cat[["ProcessID", "Process"]], on="ProcessID", how="left")
    )

    # Filter for this process
    process_df = merged[merged["Process"] == process_name].copy()

    if len(process_df) == 0:
        print(f"⚠ Keine Daten für Prozess '{process_name}' gefunden.")
        return pd.DataFrame()

    # Convert values to numeric
    process_df["Wert_num"] = process_df["Wert"].apply(extract_numeric_value)

    # Extract production values for P_Milchkuh cluster logic
    production_df = process_df[
        process_df["Pp_name"].isin(MEXCR_MILCHKUH_PRODUCTION_PARAMETER_NAMES)
    ].copy()
    production_by_process = (
        production_df.dropna(subset=["Wert_num"])
        .groupby("BetriebProcessID", dropna=False)["Wert_num"]
        .first()
        .to_dict()
    )

    if process_name == "P_Milchkuh":
        all_milchkuh_ids = set(process_df["BetriebProcessID"].dropna().unique())
        missing_ids = sorted(all_milchkuh_ids - set(production_by_process.keys()))
        if missing_ids:
            missing_df = process_df[process_df["BetriebProcessID"].isin(missing_ids)][
                ["BetriebProcessID", "Pp_name", "Wert"]
            ].drop_duplicates()
            print(
                "\n⚠ Fehlende Milchleistungswerte für P_Milchkuh in folgenden BetriebProcessIDs:"
            )
            print(missing_ids)
            print(
                "Bitte prüfen Sie in tbl_6_Betrieb_ProcessParameterValue, welche Pp_name-Werte für diese Prozess-IDs vorhanden sind.\n"
                + "Empfohlen: Fügen Sie den Wert für 'Ø Jahresleistung Milch kg A+B/ Kuh/ Jahr' hinzu."
            )
            print(
                missing_df.sort_values(["BetriebProcessID", "Pp_name"]).to_string(
                    index=False
                )
            )

    # Filter for animal category process parameters that we need
    target_pp_names = list(MEXCR_FACTORS[process_name].keys())
    process_df = process_df[process_df["Pp_name"].isin(target_pp_names)].copy()

    if len(process_df) == 0:
        print(
            f"⚠ Keine relevanten Tierkategorie-Parameter für '{process_name}' gefunden."
        )
        return pd.DataFrame()

    print(f"\n=== MEXCR Berechnung für {process_name} ===")
    print(
        process_df[
            ["BetriebID", "BetriebProcessID", "Pp_name", "Wert", "Wert_num"]
        ].head(10)
    )

    # Initialize result list
    results = []

    # All processes use direct kg N per animal/year logic now
    for _, row in process_df.iterrows():
        pp_name = row["Pp_name"]
        animal_count = row["Wert_num"]
        betrieb_id = row["BetriebID"]
        betrieb_process_id = row["BetriebProcessID"]

        if pd.isna(animal_count) or animal_count == 0:
            continue

        mexcr_info = MEXCR_FACTORS[process_name].get(pp_name)
        if not mexcr_info:
            continue

        # Skip if this is included in another parameter
        if mexcr_info.get("formula") == "included_in_mutterkuh":
            continue

        production_value = None
        if process_name == "P_Milchkuh" and pp_name == "Milchkuh / Mutterkuh":
            production_value = production_by_process.get(betrieb_process_id)
            mexcr_per_animal = get_mexcr_factor_for_milchkuh(production_value)
        else:
            mexcr_per_animal = mexcr_info.get("value", 0)

        total_mexcr = animal_count * mexcr_per_animal

        results.append(
            {
                "BetriebID": betrieb_id,
                "BetriebProcessID": betrieb_process_id,
                "Process": process_name,
                "Tierkategorie": pp_name,
                "Anzahl": animal_count,
                "mexcr_per_Tier (kg N/Jahr)": mexcr_per_animal,
                "Gesamt_mexcr (kg N/Jahr)": total_mexcr,
                "Milchleistung_Jahr (kg A+B/Kuh/Jahr)": production_value,
                "mexcr_factor_source": (
                    "milk_production_cluster"
                    if process_name == "P_Milchkuh"
                    and pp_name == "Milchkuh / Mutterkuh"
                    else "static"
                ),
            }
        )

    result_df = pd.DataFrame(results)
    return result_df


# Calculate mexcr for all defined processes
all_processes = list(MEXCR_FACTORS.keys())
all_mexcr_results = []

for process in all_processes:
    result = calculate_mexcr_for_process(process, data)
    if len(result) > 0:
        all_mexcr_results.append(result)

if all_mexcr_results:
    mexcr_detailed_df = pd.concat(all_mexcr_results, ignore_index=True)
else:
    mexcr_detailed_df = pd.DataFrame()

print("\n\n" + "=" * 80)
print("GESAMTERGEBNIS: STICKSTOFFAUSSCHEIDUNGEN (MEXCR) - DETAILLIERT")
print("=" * 80)
print(mexcr_detailed_df.head(10).to_string(index=False))

# Summary by Process
# First aggregate by Process and BetriebProcessID (sum all Pp_name rows per process instance)
# Then aggregate by Process only (to get total per process)
if len(mexcr_detailed_df) > 0:
    # Step 1: Group by Process and BetriebProcessID to sum all Pp_names within each process instance
    process_instance_df = (
        mexcr_detailed_df.groupby(["Process", "BetriebProcessID"], as_index=False)
        .agg(
            {
                "Gesamt_mexcr (kg N/Jahr)": "sum",
                "BetriebID": "first",  # Get BetriebID for reference
            }
        )
        .rename(
            columns={"Gesamt_mexcr (kg N/Jahr)": "Gesamt_mexcr_pro_BetriebProcessID"}
        )
    )

    print("\n" + "=" * 80)
    print("CHECK: TOTAL KG N/Jahr PER BETRIEBPROCESSID")
    print("=" * 80)
    print(
        process_instance_df[
            [
                "BetriebID",
                "BetriebProcessID",
                "Process",
                "Gesamt_mexcr_pro_BetriebProcessID",
            ]
        ]
        .sort_values(["Process", "BetriebProcessID"])
        .to_string(index=False)
    )

    # ============================================================================
    # LOAD GRAZING DAYS AND SPLIT MEXCR INTO PASTURE / BARN
    # ============================================================================

    tbl6 = data["tbl_6_Betrieb_ProcessParameterValue"].copy()
    process_pp = data["Katalog_ProcessProcessparameter"].copy()
    processparam = data["Katalog_Processparameter"].copy()

    tbl6 = tbl6.merge(
        process_pp[
            [
                "ProcessPpID",
                "PpID",
                "ProcessID",
            ]
        ],
        on="ProcessPpID",
        how="left",
    )

    tbl6 = tbl6.merge(
        processparam[
            [
                "PpID",
                "Pp_name",
            ]
        ],
        on="PpID",
        how="left",
    )

    weidetage_df = tbl6[tbl6["Pp_name"] == "Weidetage"].copy()
    print(f"\n🌱 Weidetage gefunden: {len(weidetage_df)}")

    weidetage_subset = weidetage_df[
        [
            "BetriebProcessID",
            "Wert",
        ]
    ].rename(columns={"Wert": "Weidetage"})

    allocation_summary = process_instance_df.merge(
        weidetage_subset,
        on="BetriebProcessID",
        how="left",
    )

    allocation_summary["Weidetage"] = pd.to_numeric(
        allocation_summary["Weidetage"],
        errors="coerce",
    ).fillna(0)

    allocation_summary["Grazing_fraction"] = (
        allocation_summary["Weidetage"] / 365
    ).clip(0, 1)

    allocation_summary["Stall_fraction"] = (
        1 - allocation_summary["Grazing_fraction"]
    ).clip(0, 1)

    # Straw as bedding
    # Bedding straw only, not feed straw
    straw_flow_ids = (
        data["Katalog_Flow"]
        .loc[data["Katalog_Flow"]["Flow"].astype(str).str.strip().eq("Stroh"), "FlowID"]
        .tolist()
    )

    straw_flows_df = data["tbl_4_Betrieb_Process_FlowValue"].copy()
    if "Direction" in straw_flows_df.columns:
        straw_flows_df = straw_flows_df[
            straw_flows_df["Direction"].astype(str).str.lower().eq("input")
        ]
    straw_flows_df = straw_flows_df[straw_flows_df["FlowID"].isin(straw_flow_ids)]

    straw_flows_df["Wert"] = pd.to_numeric(
        straw_flows_df["Wert"].astype(str).str.replace(",", "."),
        errors="coerce",
    ).fillna(0)

    straw_by_process = (
        straw_flows_df.groupby("BetriebProcessID", as_index=False)["Wert"]
        .sum()
        .rename(columns={"Wert": "Stroh_kg"})
    )

    allocation_summary = allocation_summary.merge(
        straw_by_process,
        on="BetriebProcessID",
        how="left",
    )
    allocation_summary["Stroh_kg"] = allocation_summary["Stroh_kg"].fillna(0)

    potential_feed_df = target_tbl4_input_with_futtermittelinfo_df.copy()

    # Check feed straw duplicates
    print("\n=== CHECK DUPLIKATE FUTTERSTROH VOR N-BERECHNUNG ===")

    check = potential_feed_df[
        potential_feed_df["Flow"]
        .astype(str)
        .str.contains("Futterstroh", case=False, na=False)
    ]

    print(
        check[
            [
                "BetriebProcessID",
                "FlowID",
                "Flow",
                "Wert",
                "FuttermittelID",
                "Futtermittel",
                "XP (g/kg TM)",
            ]
        ]
        .sort_values("BetriebProcessID")
        .to_string(index=False)
    )

    def _find_best_column(df, candidates):
        for candidate in candidates:
            if candidate in df.columns:
                return candidate
        normalized_columns = {str(col).lower(): col for col in df.columns}
        for candidate in candidates:
            normalized_candidate = candidate.lower()
            for normalized_col, original_col in normalized_columns.items():
                if (
                    normalized_candidate in normalized_col
                    or normalized_col in normalized_candidate
                ):
                    return original_col
        return None

    xp_column = _find_best_column(
        potential_feed_df,
        [
            "XP",
            "XP (g/kg TM)",
            "XP (g/kg DM)",
            "XP (%)",
            "Rohprotein",
        ],
    )

    potential_feed_df["Wert_numeric"] = pd.to_numeric(
        potential_feed_df["Wert"].astype(str).str.replace(",", "."),
        errors="coerce",
    ).fillna(0)

    if xp_column is None:
        print(
            "⚠ Kein XP-Spaltenname in target_tbl4_input_with_futtermittelinfo_df gefunden. "
            "feed_n_kg wird vorläufig mit 0 berechnet."
        )
        potential_feed_df["XP_numeric"] = 0
    else:
        potential_feed_df["XP_numeric"] = pd.to_numeric(
            potential_feed_df[xp_column].astype(str).str.replace(",", "."),
            errors="coerce",
        ).fillna(0)

    potential_feed_df["feed_n_kg"] = (
        potential_feed_df["Wert_numeric"]
        * potential_feed_df["XP_numeric"]
        / 1000.0
        / 6.25
    )
    # FIX 2026-10-06 (fan-out removed): safeguard against double counting.
    # Every tbl_4 feed row (BetriebProcessFlowValueID) may enter
    # feed_N_kg exactly once. Aborts if the fan-out is reintroduced.
    dup_feed_rows = potential_feed_df["BetriebProcessFlowValueID"].duplicated(keep=False)
    if dup_feed_rows.any():
        raise ValueError(
            "Futter-N: tbl_4-Zeilen mehrfach vorhanden (Fan-out?):\n"
            + potential_feed_df.loc[
                dup_feed_rows, ["BetriebProcessID", "BetriebProcessFlowValueID", "Flow"]
            ].to_string(index=False)
        )

    # FIX 2026-10-06: print feed N per process, split into cow/herd ration and
    # calf ration (P_Milchkuh only), as a check.
    print("\n=== Futter-N je BetriebProcessID (kg N), getrennt nach Futterration ===")
    print(
        potential_feed_df.pivot_table(
            index=["BetriebProcessID", "Process"],
            columns="Futterration",
            values="feed_n_kg",
            aggfunc="sum",
            fill_value=0,
        )
        .round(1)
        .to_string()
    )

    feed_n_loss_by_process = (
        potential_feed_df.groupby("BetriebProcessID", as_index=False)["feed_n_kg"]
        .sum()
        .rename(columns={"feed_n_kg": "feed_N_kg"})
    )

    allocation_summary = allocation_summary.merge(
        feed_n_loss_by_process,
        on="BetriebProcessID",
        how="left",
    )
    allocation_summary["feed_N_kg"] = allocation_summary["feed_N_kg"].fillna(0)

    # Calculate the straw N content from KTBL per BetriebProcessID. The flow "Stroh" is always bedding
    # and feed straw is always IsPotentialFeedstuff
    straw_ktbl = ktbl_df[
        ktbl_df["Futtermittel"].astype(str).str.contains("Stroh", case=False, na=False)
        | ktbl_df["Kategorie"].astype(str).str.contains("Stroh", case=False, na=False)
    ]
    if len(straw_ktbl) == 0:
        raise KeyError(
            "Kein Stroh-Eintrag in Katalog_Futtermittelgehalte_KTBL gefunden."
        )
    straw_xp_g_per_kg = pd.to_numeric(
        straw_ktbl["XP (g/kg TM)"].dropna().astype(float).iloc[0]
    )
    straw_n_fraction = straw_xp_g_per_kg / 1000 / 6.25

    allocation_summary["mbedding_kg_N"] = (
        allocation_summary["Stroh_kg"] * straw_n_fraction
    )

    allocation_summary["mexcr_graz"] = (
        allocation_summary["Gesamt_mexcr_pro_BetriebProcessID"]
        * allocation_summary["Grazing_fraction"]
    )
    allocation_summary["mexcr_stall"] = (
        allocation_summary["Gesamt_mexcr_pro_BetriebProcessID"]
        * allocation_summary["Stall_fraction"]
    )
    allocation_summary["Xhouse"] = 1
    # Always 1 here, because all animals of a process are treated the same and totals are used, not per-animal values.

    print("\n=== MEXCR WEIDE/STALL AUFTEILUNG ===")
    print(
        allocation_summary[
            [
                "BetriebID",
                "BetriebProcessID",
                "Process",
                "Weidetage",
                "Grazing_fraction",
                "Stall_fraction",
                "Xhouse",
                "Stroh_kg",
                "mbedding_kg_N",
                "mexcr_graz",
                "mexcr_stall",
            ]
        ]
        .sort_values(["Process", "BetriebProcessID"])
        .head(20)
        .to_string(index=False)
    )

    print("\n=== MEXCR WEIDE/STALL/BEDDING SUMMEN PRO PROCESS ===")
    print(
        allocation_summary.groupby("Process", as_index=False)[
            [
                "Gesamt_mexcr_pro_BetriebProcessID",
                "mexcr_graz",
                "mexcr_stall",
                "Stroh_kg",
                "mbedding_kg_N",
            ]
        ]
        .sum()
        .to_string(index=False)
    )
# ============================================================================
# Check whether straw is treated as bedding and whether feed straw appears correctly in the feed N calculation
# =============================================================================
print("\n=== Einstreu-Stroh Flow Check ===")

print(
    data["Katalog_Flow"][data["Katalog_Flow"]["FlowID"].isin(straw_flow_ids)][
        ["FlowID", "Flow"]
    ]
)

# ============================================================================
# CHECK: FEED STRAW FLOW
# ============================================================================

feed_straw_flow_ids = (
    data["Katalog_Flow"]
    .loc[
        data["Katalog_Flow"]["Flow"].astype(str).str.strip().eq("Futterstroh"), "FlowID"
    ]
    .tolist()
)

print("\n=== Futterstroh Flow Check ===")

print(
    data["Katalog_Flow"][data["Katalog_Flow"]["FlowID"].isin(feed_straw_flow_ids)][
        [
            "FlowID",
            "Flow",
            "IsPotentialFeedstuff",
        ]
    ].to_string(index=False)
)
# ============================================================================
# CHECK: FEED STRAW IN FEED N CALCULATION
# ============================================================================

feed_straw_check = potential_feed_df[
    potential_feed_df["Flow"]
    .astype(str)
    .str.contains("Futterstroh", case=False, na=False)
]

print("\n=== Futterstroh N-Berechnung Check ===")

if len(feed_straw_check) == 0:

    print("Kein Futterstroh als Futtermittel gefunden.")

else:

    print(
        feed_straw_check[
            [
                "BetriebProcessID",
                "Flow",
                "Wert_numeric",
                "XP_numeric",
                "feed_n_kg",
            ]
        ]
        .sort_values("BetriebProcessID")
        .to_string(index=False)
    )
# =============================================================================
# Note: Gesamt_mexcr_pro_BetriebProcessID is mexcr and thus the first variable
# for every BetriebProcessID that is used in the later calculations for N2O, NOx and NH3.
# =============================================================================
# TAN content (TAN)
# =============================================================================
print("2.Inputvariable berechnen - TaN-Gehalt (XTAN, EXKR)")
# Sources: KTBL BEK 2026
# for all animal categories except dairy cattle, the TAN share of the total nitrogen excretion
# is taken from the parameter "TAN share in the excretion" of the KTBL parameter file 2024.
# For dairy cattle the TAN share of the total nitrogen excretion
# is calculated with the following equation: mexcr,TAN = -6.89 + 10.1 ∙ x MU + 50.34, where XMU = milk urea in mg dl-1 and
# mexcr, TAN in g per animal and day.
# Read the KTBL parameter file:
print("\n=== Einlesen der TAN-Anteil Parameter aus KTBL ===")
parameterdatei_KTBL_df = pd.read_excel(
    ktbl_parameter_path,
    sheet_name="BEK-Parameterdatei",
    usecols=[
        "KATEGORIE",
        "UNTERKATEGORIE",
        "EMISSIONSQUELLE",
        "PARAMETER (formatiert)",
        "BESCHREIBUNG (formatiert)",
        "BEZEICHNUNG (formatiert)",
        "WERT (formatiert)",
        "EINHEIT (formatiert)",
    ],
)
print(
    "Parameterdatei KTBL - Vorschau der relevanten Spalten:\n",
    parameterdatei_KTBL_df[
        [
            "EMISSIONSQUELLE",
            "PARAMETER (formatiert)",
            "WERT (formatiert)",
            "EINHEIT (formatiert)",
        ]
    ].head(10),
)


# Filter for relevant parameters (TAN share in the excretion)
def get_ktbl_tan_values():
    return parameterdatei_KTBL_df[
        (parameterdatei_KTBL_df["EMISSIONSQUELLE"] == "Tier")
        & (
            (parameterdatei_KTBL_df["PARAMETER (formatiert)"] == "TAN-Gehalt")
            | (parameterdatei_KTBL_df["PARAMETER (formatiert)"] == "UAN-Gehalt")
        )
        & (
            parameterdatei_KTBL_df["BESCHREIBUNG (formatiert)"].str.contains(
                "Ausscheidung", na=False
            )
        )
    ]


tan_ausscheidung_parameter_df = get_ktbl_tan_values()

tan_lookup = dict(
    zip(
        tan_ausscheidung_parameter_df["UNTERKATEGORIE"],
        tan_ausscheidung_parameter_df["WERT (formatiert)"],
    )
)
print("\nTAN-Anteil Lookup-Tabelle:\n", tan_lookup)


def normalize_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip().lower()
    replacements = {
        "ö": "oe",
        "ä": "ae",
        "ü": "ue",
        "ß": "ss",
        "♀": "f",
        "♂": "m",
        "/": " ",
        "+": " ",
        ",": "",
        "-": " ",
        ".": "",
        "(": "",
        ")": "",
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return (
        "".join(ch for ch in text if ch.isalnum() or ch == " ")
        .replace(" ", " ")
        .strip()
    )


# Sometimes the animal category alone is not enough to determine the correct TAN factor (e.g. dairy cow vs. suckler cow). In such cases a combination of process and animal category is used to determine the lookup key.
# Here a mapping table is defined that maps specific combinations of process and animal category to the corresponding TAN lookup keys. If such a combination is found in the data, the assigned key is used for the lookup. Otherwise the normalised animal category is used as fallback.
# Note: all beef fattening categories were assigned to young cattle husbandry, because farm-specific values for daily weight gain are missing.
process_tan_key_map = {
    ("P_Milchkuh", "Milchkuh / Mutterkuh"): "Milchkuhhaltung",
    ("P_Milchkuh", "Kälber"): "Kälberhaltung",
    ("P_Mutterkuh", "Milchkuh / Mutterkuh"): "Mutterkuhhaltung",
    ("P_Mutterkuh", "Kälber"): "Kälberhaltung",
    ("P_Mutterkuh", "♂ Mast"): "Jungrinderhaltung",
    ("P_Mutterkuh", "♀ Aufzucht / Masttiere (je nach MV oder MK)"): "Jungrinderhaltung",
    ("P_Aufzucht_Milchvieh", "♂ Mast"): "Jungrinderhaltung",
    (
        "P_Aufzucht_Milchvieh",
        "♀ Aufzucht / Masttiere (je nach MV oder MK)",
    ): "Jungrinderhaltung",
    ("P_Rind_Mast", "♂ Mast"): "Jungrinderhaltung",
    ("P_Rind_Mast", "♀ Aufzucht / Masttiere (je nach MV oder MK)"): "Jungrinderhaltung",
    # FIX 2026-10-06 (breeding bull): KTBL parameter file, subcategory breeding bull husbandry (TAN 0.6)
    ("P_Milchkuh", "Deckbulle"): "Deckbullenhaltung",
    ("P_Mutterkuh", "Deckbulle"): "Deckbullenhaltung",
    ("P_Rind_Mast", "Deckbulle"): "Deckbullenhaltung",
    ("P_Mastschwein", "Mastschwein"): "Schweinemast",
    ("P_Schafhaltung_Mast", "Mutterschaf + Bock + Lämmer + Mastschaf"): "Schafhaltung",
    ("P_Pferdehaltung", "Pferd"): "Pferdehaltung",
    ("P_Legehennen", "Anzahl Legehennen"): "Legehennenhaltung",
    ("P_Bruderhahnmast", "Anzahl Masthähnchen"): "Hühnermast",
}

custom_tierkategorie_map = {
    "milchkuh / mutterkuh": "milchkuhhaltung",
    "kälber": "kaelberhaltung",
    "kaelber": "kaelberhaltung",
    "mastschein": "schweinemast",
    "♀ aufzucht  masttiere je nach mv oder mk": "jungrinderhaltung",
    "mutterschaf": "schafhaltung",
    "bock": "schafhaltung",
    "lammer": "schafhaltung",
    "mastschaf": "schafhaltung",
    "pferd": "pferdehaltung",
}

normalized_lookup = {normalize_text(k): v for k, v in tan_lookup.items()}


def get_tan_factor(process, tierkategorie):
    process_key = (process, tierkategorie)
    if process_key in process_tan_key_map:
        lookup_key = normalize_text(process_tan_key_map[process_key])
    else:
        normalized_key = normalize_text(tierkategorie)
        mapped_key = custom_tierkategorie_map.get(normalized_key)
        lookup_key = mapped_key if mapped_key else normalized_key

    if lookup_key in normalized_lookup:
        raw = normalized_lookup[lookup_key]
    else:
        fuzzy_matches = [
            value
            for lookup_key2, value in normalized_lookup.items()
            if lookup_key in lookup_key2 or lookup_key2 in lookup_key
        ]
        if fuzzy_matches:
            raw = fuzzy_matches[0]
        else:
            print(
                f"⚠ Warnung: Kein TAN-Faktor für Prozess='{process}', Tierkategorie='{tierkategorie}' gefunden."
            )
            return None

    try:
        return float(str(raw).replace(",", "."))
    except ValueError:
        print(
            f"⚠ Warnung: TAN-Faktor für Prozess='{process}', Tierkategorie='{tierkategorie}' konnte nicht in float konvertiert werden. "
            f"Rohwert: {raw}"
        )
        return None


# Use the already calculated mexcr_detailed_df as the basis for the TAN calculation.
result_df = mexcr_detailed_df.copy()

# Merge the aggregated mexcr values per BetriebProcessID
if len(mexcr_detailed_df) > 0:
    result_df = result_df.merge(
        process_instance_df[
            ["Process", "BetriebProcessID", "Gesamt_mexcr_pro_BetriebProcessID"]
        ],
        on=["Process", "BetriebProcessID"],
        how="left",
    )

result_df["TAN_Faktor"] = result_df.apply(
    lambda row: get_tan_factor(row["Process"], row["Tierkategorie"]), axis=1
)

# Fattening pig always 0.55, because a farm-specific daily weight gain of 700 g/day on farm 36 was taken as average value.
result_df.loc[result_df["Process"] == "P_Mastschwein", "TAN_Faktor"] = 0.55

print("\n--- MEXCR, TAN_Faktor head 20 ---")
print(
    result_df[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "Tierkategorie",
            "TAN_Faktor",
        ]
    ]
    .head(20)
    .to_string(index=False)
)

# Summary with TAN_Faktor only (no absolute numbers)
summary_df = (
    result_df.groupby(["Process", "Tierkategorie"], as_index=False)["TAN_Faktor"]
    .mean()
    .sort_values(["Process", "Tierkategorie"])
)
print("\n--- Zusammenfassung: Process / Tierkategorie / TAN_Faktor ---")
print(summary_df.head(20).to_string(index=False))


# Calculation of farm-specific values for P_Milchvieh according to BEK Eq. 15
# XMU in mg/dl-1 = 19.57
# mexcr,TAN in kg per animal and day = (-6.89 + 10.1 ∙ x MU + 50.34)/1000
# TAN_Faktor = mexcr,TAN / mexcr
print(
    mexcr_detailed_df.loc[
        mexcr_detailed_df["Process"] == "P_Milchkuh",
        ["Anzahl", "mexcr_per_Tier (kg N/Jahr)"],
    ].head(10)
)


def calculate_tan_factor_milchvieh(row):
    if (
        row["Process"] == "P_Milchkuh"
        and row["Tierkategorie"] == "Milchkuh / Mutterkuh"
    ):
        x_mu = 19.57
        mexcr_tan = (-6.89 + 10.1 * x_mu + 50.34) / 1000
        mexcr_per_day = row["Gesamt_mexcr_pro_BetriebProcessID"] / 365
        mexcr_per_tier_per_day = (
            mexcr_per_day / row["Anzahl"] if row["Anzahl"] > 0 else 0
        )
        tan_faktor = (
            mexcr_tan / mexcr_per_tier_per_day if mexcr_per_tier_per_day > 0 else 0
        )
        return tan_faktor
    return row["TAN_Faktor"]


result_df["TAN_Faktor"] = result_df.apply(calculate_tan_factor_milchvieh, axis=1)
print(
    "\n--- Aktualisierte TAN_Faktor mit betriebsindividuellen Werten für Milchvieh ---"
)
print(
    result_df[
        (result_df["Process"] == "P_Milchkuh")
        & (result_df["Tierkategorie"] == "Milchkuh / Mutterkuh")
    ][
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "Tierkategorie",
            "TAN_Faktor",
        ]
    ]
    .head(20)
    .to_string(index=False)
)
# ============================================================
# Reduce the TAN factor to one row per BetriebProcessID for the later mapping
# Priority: Milchkuh / Mutterkuh before Kälber
# ============================================================

result_df["tan_priority"] = 1

result_df.loc[
    (result_df["Process"] == "P_Milchkuh")
    & (result_df["Tierkategorie"] == "Milchkuh / Mutterkuh"),
    "tan_priority",
] = 0

# FIX 2026-10-06 (breeding bull): never choose the breeding bull TAN (0.6) as the TAN factor of the whole
# process as long as other animal categories are present.
result_df.loc[result_df["Tierkategorie"] == "Deckbulle", "tan_priority"] = 2


result_df = (
    result_df.sort_values(["BetriebProcessID", "tan_priority"])
    .drop_duplicates(subset=["BetriebProcessID"], keep="first")
    .drop(columns="tan_priority")
)
# Check
print("\n Check ob nur ein TAN Faktor für jede BetriebsProcessID")
print(
    result_df[result_df["Process"] == "P_Milchkuh"][
        ["BetriebProcessID", "Tierkategorie", "TAN_Faktor"]
    ].to_string(index=False)
)
# Tests why TAN factors can be > 1
check = result_df[
    (result_df["Process"] == "P_Milchkuh")
    & (result_df["Tierkategorie"] == "Milchkuh / Mutterkuh")
].copy()

check["mexcr_per_day"] = (
    check["Gesamt_mexcr_pro_BetriebProcessID"] / check["Anzahl"] / 365
)

print(
    check[
        [
            "BetriebProcessID",
            "Anzahl",
            "mexcr_per_Tier (kg N/Jahr)",
            "Gesamt_mexcr_pro_BetriebProcessID",
            "mexcr_per_day",
            "TAN_Faktor",
        ]
    ].sort_values("TAN_Faktor", ascending=False)
)

print("\n=== Milchkuh Debug ===")

milch_debug = result_df[
    (result_df["Process"] == "P_Milchkuh")
    & (result_df["Tierkategorie"] == "Milchkuh / Mutterkuh")
].copy()
print(
    milch_debug[
        [
            "BetriebProcessID",
            "Anzahl",
            "mexcr_per_Tier (kg N/Jahr)",
            "Gesamt_mexcr_pro_BetriebProcessID",
        ]
    ]
)

# =============================================================================
# Now mhouse (kg N per barn and year), Xhouse (1 - allocation_summary["Grazing_fraction"]),
# mbedding (kg N from bedding per year), m_feedlo (kg N from feed losses per year)
# and mexcr (kg N per animal and year) are integrated into the input table for the emission calculation.
# mTANhouse (kg TAN per barn and year), to calculate N2O, NO and NH3 emissions from
# manure management and grazing later.
# =============================================================================
# Since feed losses are not recorded explicitly in the Access database, where exact ration calculations
# or total feed intakes were used, a general 5% of the N amount from the potential feed inputs is assumed here.
# m_feedlo = 0.05 * N from potential feedstuff XP
allocation_summary["m_feedlo"] = 0.05 * allocation_summary["feed_N_kg"]
# mhouse = mexcr_stall + mbedding + m_feedlo
allocation_summary["mhouse"] = (
    allocation_summary["mexcr_stall"]
    + allocation_summary[
        "mbedding_kg_N"
    ]  # m_bedding contains bedding and feed straw (which applies to one farm.
    # The two straw types could not be distinguished in the Access table,
    # so they were assigned to bedding.)
    + allocation_summary["m_feedlo"]
)
# mTANhouse = mexcr_stall * TAN_Faktor
allocation_summary = allocation_summary.merge(
    result_df[["BetriebProcessID", "TAN_Faktor"]],
    on="BetriebProcessID",
    how="left",
)
allocation_summary["mTANhouse"] = (
    allocation_summary["mexcr_stall"]
    * allocation_summary["TAN_Faktor"]
    * allocation_summary["Xhouse"]
)
print("\n=== Finaler Check: mhouse, mTANhouse, m_feedlo ===")
print(
    allocation_summary[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "Gesamt_mexcr_pro_BetriebProcessID",
            "Grazing_fraction",
            "Xhouse",
            "mbedding_kg_N",
            "m_feedlo",
            "mhouse",
            "TAN_Faktor",
            "mTANhouse",
        ]
    ].sort_values("mTANhouse", ascending=False)
)
# mtanbedding = mbedding * TAN_Faktor (0,5)
allocation_summary["mtanbedding"] = allocation_summary["mbedding_kg_N"] * 0.5
# mtanfeedlo = m_feedlo * TAN_FAktor (0,5)
allocation_summary["mtanfeedlo"] = allocation_summary["m_feedlo"] * 0.5
# # For bedding and feed loss N a flat mineralisation share
# of 50 % is assumed. This share is added to the TAN pool in manure management.
# zugerechnet.
# =============================================================================
# EF N2O and NH3 from the KTBL parameter file
# =============================================================================
# Extract emission factors from the KTBL parameter file 2024 for the different emission sources (e.g. N2O, NH3) and animal categories (e.g. dairy cow husbandry, pig fattening) and store them in a lookup dictionary.
# EFN2O house
# EF N2Odeposition
# EF N20 grazing
# EF NH3 house
# EF NH3 MM: the following storage / housing systems exist in the system (red DB)
#                Biogas; slurry general; slurry natural crust; slurry with foil cover;
#                slurry under slatted floor; solid manure; deep litter
# EF NH3 grazing
# =============================================================================
# Define the EF from the KTBL parameter file by animal species and emission source
# =============================================================================

# All cattle husbandry processes are grouped, because their EF do not differ in the KTBL parameter file.

EF_PROCESS_GROUPS = {
    "P_Milchkuh": "Rind",
    "P_Mutterkuh": "Rind",
    "P_Aufzucht_Milchvieh": "Rind",
    "P_Rind_Mast": "Rind",
    "P_Mastschwein": "Schwein",
    "P_Schafhaltung_Mast": "Schaf",
    "P_Legehennen": "Geflügel",
    "P_Bruderhahnmast": "Geflügel",
}
# EF N2O MM covers housing and storage
EF_N2O_MM = {
    "Rind": {
        "Gülle mit Folienabdeckung": 0,
        "Gülle nat. Schwimmschicht": 0.005,
        "Gülle allg.": 0.005,
        "Gülle unter Spaltenboden": 0.002,
        "Tiefstreu": 0.01,
        "Biogas": 0.005,  # assuming an open digestate store with natural floating cover.
        # simplified assumption: the farms only bring the slurry to the biogas plant
        # and therefore the emissions of the pre-storage were left out, as they count as 0 anyway (KTBL)
    },
    "Schwein": {
        "Tiefstreu": 0.01,
    },
    "Geflügel": {
        "Tiefstreu": 0.001,
    },
    "Schaf": {
        "Tiefstreu": 0.005,
    },
}

# EF N2O deposition
EF_N2O_DEPOSITION = {
    "Rind": 0.01,
    "Schwein": 0.01,
    "Geflügel": 0.01,
    "Schaf": 0.01,
}

# EF N2O grazing
EF_N2O_GRAZING = {
    "Rind": 0.01,
    "Schwein": 0.02264,
    "Geflügel": 0.001,
    "Schaf": 0.01264,
}

# EF NH3 house
EF_NH3_HOUSE = {
    "Rind": {
        "Laufstall güllebasiert": 0.197,
        "Tiefstreu": 0.197,
    },
    "Schwein": {
        "Tiefstreu": 0.4,
    },
    "Geflügel": {
        "Tiefstreu": 0.09,  # organic, without manure belt ventilation
    },
    "Schaf": {
        "Tiefstreu": 0.22,
    },
}

# EF NH3 MM
EF_NH3_MM = {
    "Rind": {
        "Gülle mit Folienabdeckung": 0.023,
        "Gülle nat. Schwimmschicht": 0.045,
        "Gülle unter Spaltenboden": 0.045,
        "Gülle allg.": 0.045,
        "Tiefstreu": 0.6,  # corresponds to the manure heap
        "Biogas": 0.045,  # assuming an open digestate store with natural floating cover.
    },
    "Schwein": {
        "Tiefstreu": 0.6,
    },
    "Geflügel": {
        "Tiefstreu": 0.08,  # organic, without manure belt ventilation
    },
    "Schaf": {
        "Tiefstreu": 0.32,
    },
}

# EF NH3 grazing
EF_NH3_GRAZING = {
    "Rind": 0.14,
    "Schwein": 0.31,
    "Geflügel": 0.35,
    "Schaf": 0.09,
}

# Get the housing system and MM system from the red DB (ensure tbl6 available)
tbl6_pp = data.get("tbl_6_Betrieb_ProcessParameterValue")
process_pp = data.get("Katalog_ProcessProcessparameter")
processparam = data.get("Katalog_Processparameter")

if tbl6_pp is None or process_pp is None or processparam is None:
    raise KeyError(
        "Required tables for extracting process parameters (tbl6/process_pp/processparam) not found in data."
    )

tbl6_pp = tbl6_pp.copy()
process_pp = process_pp.copy()
processparam = processparam.copy()

process_parameter_values = tbl6_pp.merge(
    process_pp[["ProcessPpID", "PpID"]], on="ProcessPpID", how="left"
).merge(processparam[["PpID", "Pp_name"]], on="PpID", how="left")

# Extract Stallsystem for each BetriebProcessID
housing_system_df = process_parameter_values[
    process_parameter_values["Pp_name"] == "Stallsystem"
][["BetriebProcessID", "Wert"]].rename(columns={"Wert": "Housing_System"})
# Add to allocation_summary
allocation_summary = allocation_summary.merge(
    housing_system_df, on="BetriebProcessID", how="left"
)
# Add EF animal group
allocation_summary["EF_Tiergruppe"] = allocation_summary["Process"].map(
    EF_PROCESS_GROUPS
)
print("\n=== Check: Housing System hinzugefügt ===")
print(
    allocation_summary[
        ["BetriebID", "BetriebProcessID", "Process", "Housing_System"]
    ].head(20)
)


# ============================================================
# Derive the NH3 housing system
# ============================================================


def derive_nh3_house_system(row):

    # Cattle: all slurry-based systems
    # use the same NH3 housing EF
    if row["EF_Tiergruppe"] == "Rind":

        if row["Housing_System"] == "Tiefstreu":
            return "Tiefstreu"

        return "Laufstall güllebasiert"

    # take other animal species directly
    return row["Housing_System"]


allocation_summary["NH3_Housing_System"] = allocation_summary.apply(
    derive_nh3_house_system,
    axis=1,
)


# ============================================================
# Generic EF lookup function
# ============================================================


def get_EF(
    ef_dict,
    tiergruppe,
    system=None,
):

    value = ef_dict.get(tiergruppe)

    if value is None:
        print(f"⚠ Keine EF-Gruppe: {tiergruppe}")
        return None

    # EF without system distinction
    if isinstance(value, (float, int)):
        return value

    ef = value.get(system)

    if ef is None:
        print(f"⚠ Kein EF für {tiergruppe} + {system}")

    return ef


# ============================================================
# EF mapping to BetriebProcessID
# ============================================================


allocation_summary["EF_N2O_MM"] = allocation_summary.apply(
    lambda row: get_EF(
        EF_N2O_MM,
        row["EF_Tiergruppe"],
        row["Housing_System"],
    ),
    axis=1,
)


allocation_summary["EF_NH3_MM"] = allocation_summary.apply(
    lambda row: get_EF(
        EF_NH3_MM,
        row["EF_Tiergruppe"],
        row["Housing_System"],
    ),
    axis=1,
)


allocation_summary["EF_NH3_house"] = allocation_summary.apply(
    lambda row: get_EF(
        EF_NH3_HOUSE,
        row["EF_Tiergruppe"],
        row["NH3_Housing_System"],
    ),
    axis=1,
)


allocation_summary["EF_N2O_deposition"] = allocation_summary["EF_Tiergruppe"].apply(
    lambda x: get_EF(EF_N2O_DEPOSITION, x)
)


allocation_summary["EF_N2O_grazing"] = allocation_summary["EF_Tiergruppe"].apply(
    lambda x: get_EF(EF_N2O_GRAZING, x)
)


allocation_summary["EF_NH3_grazing"] = allocation_summary["EF_Tiergruppe"].apply(
    lambda x: get_EF(EF_NH3_GRAZING, x)
)


# ============================================================
# Check output
# ============================================================

print(
    allocation_summary[
        [
            "BetriebProcessID",
            "Process",
            "EF_Tiergruppe",
            "Housing_System",
            "NH3_Housing_System",
            "EF_N2O_MM",
            "EF_NH3_house",
            "EF_NH3_MM",
            "EF_N2O_grazing",
            "EF_NH3_grazing",
        ]
    ]
    .drop_duplicates()
    .to_string(index=False)
)

# ============================================================
# Check EF mapping completeness
# ============================================================

required_ef_columns = [
    "EF_N2O_MM",
    "EF_N2O_deposition",
    "EF_N2O_grazing",
    "EF_NH3_house",
    "EF_NH3_MM",
    "EF_NH3_grazing",
]


missing_ef_check = allocation_summary[
    allocation_summary[required_ef_columns].isna().any(axis=1)
]


if missing_ef_check.empty:

    print("\n✅ Alle Tierhaltungsprozesse haben vollständige Emissionsfaktoren.")


else:

    print("\n⚠ WARNUNG: Fehlende Emissionsfaktoren gefunden:")

    print(
        missing_ef_check[
            [
                "BetriebID",
                "BetriebProcessID",
                "Process",
                "EF_Tiergruppe",
                "Housing_System",
                "NH3_Housing_System",
            ]
            + required_ef_columns
        ].to_string(index=False)
    )
    print("\n=== EF Mapping Kombinationen ===")

print(
    allocation_summary[
        [
            "Process",
            "EF_Tiergruppe",
            "Housing_System",
            "NH3_Housing_System",
        ]
        + required_ef_columns
    ]
    .drop_duplicates()
    .sort_values(
        [
            "EF_Tiergruppe",
            "Housing_System",
        ]
    )
    .to_string(index=False)
)

# # NH3 N house = mTANhouse * EF_NH3_House Eq. 39
allocation_summary["NH3_N_House"] = (
    allocation_summary["mTANhouse"] * allocation_summary["EF_NH3_house"]
)
# NH3 house gas = NH3_N_house * (17/14) (Eq. 38)
allocation_summary["NH3_Gas_House"] = allocation_summary["NH3_N_House"] * (17 / 14)

# NH3 N storage = mTanstorage * EF_NH3_MM
# mTANstorage, different depending on the housing system
# Liquid manure systems are calculated with equation 28:
# mTANstorage liquid = (mTANhouse - NH3_N_House) + (mexcr, house - mTAN_house)
# * 0.1 + (mTAN_bedding + mTAN_feedlo))
# Solid manure is calculated with equation 34: since frac leach is 0 for deep litter, mTAN_storage_leach drops out and:
# mTANstorage solid = (mTAN_house - NH3_N_House - 0 ) * (1 - 0.4) + (mTAN_bedding + mTAN_feedlo)

allocation_summary["mTANstorageliquid"] = (
    (allocation_summary["mTANhouse"] - allocation_summary["NH3_N_House"])
    + (allocation_summary["mexcr_stall"] - allocation_summary["mTANhouse"]) * 0.1
    + (allocation_summary["mtanbedding"] + allocation_summary["mtanfeedlo"])
)

allocation_summary["mTANstoragesolid"] = (
    allocation_summary["mTANhouse"] - allocation_summary["NH3_N_House"]
) * (1 - 0.4) + (allocation_summary["mtanbedding"] + allocation_summary["mtanfeedlo"])

allocation_summary["mTANstorage"] = allocation_summary.apply(
    lambda row: (
        row["mTANstorageliquid"]
        if row["NH3_Housing_System"] == "Laufstall güllebasiert"
        else row["mTANstoragesolid"]
    ),
    axis=1,
)
# NH3 N MM = mTANstorage * EF_NH3_MM
allocation_summary["NH3_N_MM"] = (
    allocation_summary["mTANstorage"] * allocation_summary["EF_NH3_MM"]
)
# NH3 MM gas = NH3_N_MM * (17/14)
allocation_summary["NH3_Gas_MM"] = allocation_summary["NH3_N_MM"] * (17 / 14)
# NH3 N Grazing = mexcr_graz * TAN_Faktor * EF_NH3_Grazing
allocation_summary["NH3_N_Grazing"] = (
    allocation_summary["mexcr_graz"]
    * allocation_summary["TAN_Faktor"]
    * allocation_summary["EF_NH3_grazing"]
)

# NH3 Grazing gas = NH3_N_Grazing * (17/14)
allocation_summary["NH3_Gas_Grazing"] = allocation_summary["NH3_N_Grazing"] * (17 / 14)
# === N2O calculation ===
print("\n=== N2O Berechnung ===")
# N2O MM = mhouse * EF_N2O_MM Eq. 35
# Calculation of mhouse
# mhouse = mexcr_stall * Xhouse + mbedding + m_feedlo
allocation_summary["m_house"] = (
    allocation_summary["mexcr_stall"] * allocation_summary["Xhouse"]
    + allocation_summary["mbedding_kg_N"]
    + allocation_summary["m_feedlo"]
)

allocation_summary["N2O_N_MM"] = (
    allocation_summary["m_house"] * allocation_summary["EF_N2O_MM"]
)
# N2O MM Gas = N2O_MM * (44/28)
allocation_summary["N2O_MM_Gas"] = allocation_summary["N2O_N_MM"] * (44 / 28)
# ============================================================================
# N2O deposition ========
print("\n=== N2O Deposition Berechnung ===")
# ============================================================================
# N2O deposition = (NO_N_MM + NH3_N_MM)* EF_N2O_deposition Eq. 36
# NO_N_MM = 0.1 * N2O_N_MM  Eq. 37
allocation_summary["NO_deposition"] = 0.1 * allocation_summary["N2O_N_MM"]
# NO as gas = NO_N_MM * (30/14)
allocation_summary["NO_Gas_deposition"] = allocation_summary["NO_deposition"] * (
    30 / 14
)
# N2O N deposition = (NO_N_MM + NH3_N_MM)* EF_N2O_deposition Eq. 36
allocation_summary["N2O_N_deposition"] = (
    allocation_summary["NO_deposition"] + allocation_summary["NH3_N_MM"]
) * allocation_summary["EF_N2O_deposition"]

# N2O deposition Gas = N2O_N_deposition * (44/28)
allocation_summary["N2O_Gas_deposition"] = allocation_summary["N2O_N_deposition"] * (
    44 / 28
)
# N2O grazing = mexcr_graz * EF_N2O_grazing
allocation_summary["N2O_N_Grazing"] = (
    allocation_summary["mexcr_graz"] * allocation_summary["EF_N2O_grazing"]
)
# N2O Grazing Gas = N2O_N_Grazing * (44/28)
allocation_summary["N2O_Gas_Grazing"] = allocation_summary["N2O_N_Grazing"] * (44 / 28)

print("\n=== Finaler Check: N2O und NH3 Berechnung ===")
print(
    allocation_summary[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "m_house",
            "EF_N2O_MM",
            "N2O_N_MM",
            "N2O_MM_Gas",
            "NO_deposition",
            "NO_Gas_deposition",
            "N2O_N_deposition",
            "N2O_Gas_deposition",
            "N2O_N_Grazing",
            "N2O_Gas_Grazing",
            "NH3_N_House",
            "NH3_Gas_House",
            "NH3_N_MM",
            "NH3_Gas_MM",
            "NH3_N_Grazing",
            "NH3_Gas_Grazing",
        ]
    ].sort_values("N2O_N_MM", ascending=False)
)

# =============================================================================
# Finally, the calculated values are integrated into the input table for the emission calculation. This ensures that the values are assigned correctly to the respective BetriebProcessIDs, for a precise calculation of the emissions from manure management and grazing.
# =============================================================================
emission_mapping = {
    "NH3_Gas_House": {
        "Flow": "Ammonia",
        "Quelle": "House",
    },
    "NH3_Gas_MM": {
        "Flow": "Ammonia",
        "Quelle": "MM",
    },
    "NH3_Gas_Grazing": {
        "Flow": "Ammonia",
        "Quelle": "Grazing",
    },
    "N2O_Gas_deposition": {
        "Flow": "Dinitrogen monoxide",
        "Quelle": "Deposition",
    },
    "N2O_MM_Gas": {
        "Flow": "Dinitrogen monoxide",
        "Quelle": "MM",
    },
    "N2O_Gas_Grazing": {
        "Flow": "Dinitrogen monoxide",
        "Quelle": "Grazing",
    },
    "NO_Gas_deposition": {
        "Flow": "Nitrogen monoxide",
        "Quelle": "Deposition",
    },
}

# ============================================================================
# Build long emissions DataFrame from calculated N2O/NH3 results
# ============================================================================
emission_columns = list(emission_mapping.keys())

n_emissions_long = allocation_summary[
    ["BetriebID", "BetriebProcessID", "Process"] + emission_columns
].copy()

n_emissions_long = n_emissions_long.melt(
    id_vars=["BetriebID", "BetriebProcessID", "Process"],
    value_vars=emission_columns,
    var_name="EmissionType",
    value_name="Wert",
)

n_emissions_long = n_emissions_long[
    n_emissions_long["Wert"].notna() & (n_emissions_long["Wert"] != 0)
].copy()

n_emissions_long["Flow"] = n_emissions_long["EmissionType"].map(
    lambda x: emission_mapping[x]["Flow"]
)

n_emissions_long["TierkategorieEmissionFutter"] = n_emissions_long.apply(
    lambda row: f"{row['Process']}, {emission_mapping[row['EmissionType']]['Quelle']}",
    axis=1,
)

n_emissions_long["Tierprozess"] = n_emissions_long["EmissionType"]

# Map the calculated emissions to FlowIDs from Katalog_Flow
katalog_flow_df = data["Katalog_Flow"][["FlowID", "Flow"]].copy()
katalog_flow_df["Flow_lower"] = katalog_flow_df["Flow"].str.lower()

n_emissions_long["Flow_lower"] = n_emissions_long["Flow"].str.lower()

n_emissions_long = n_emissions_long.merge(
    katalog_flow_df[["FlowID", "Flow_lower"]],
    on="Flow_lower",
    how="left",
)

missing_flow_mapping = n_emissions_long[n_emissions_long["FlowID"].isna()]
if len(missing_flow_mapping) > 0:
    print(
        "\n⚠ Warnung: Für einige berechnete Emissionen konnte kein FlowID-Mapping gefunden werden:"
    )
    print(
        missing_flow_mapping[["EmissionType", "Flow", "Wert"]]
        .drop_duplicates()
        .to_string(index=False)
    )

n_emissions_long = n_emissions_long.drop(columns=["Flow_lower", "Flow"])

# ============================================================================
# TARGET PROCESSES PER LIVESTOCK PROCESS: P_EmissionMM_xy
# ============================================================================
# Instead of putting all N2O/NH3/NO emissions into one single manure management process
# per farm, every livestock instance (BetriebProcessID) gets its own
# P_EmissionMM_xy process. Same approach as in
# 4_Emission_Duengemanagement_PM_NMVOC_Processlevel.py, so that both scripts write to
# the same target process instances. The animal tag at the end of the
# source process name is the key, e.g. P_Mastschwein -> P_EmissionMM_Mastschwein.
PROCESS_TO_EMISSIONMM = {
    "P_Milchkuh": "P_EmissionMM_Milchkuh",
    "P_Mutterkuh": "P_EmissionMM_Mutterkuh",
    "P_Rind_Mast": "P_EmissionMM_Rind_Mast",
    "P_Schafhaltung_Mast": "P_EmissionMM_Schafhaltung_Mast",
    "P_Mastschwein": "P_EmissionMM_Mastschwein",
    "P_Legehennen": "P_EmissionMM_Legehennen",
    "P_Bruderhahnmast": "P_EmissionMM_Bruderhahnmast",
    "P_Aufzucht_Milchvieh": "P_EmissionMM_Aufzucht_Milchvieh",
    "P_Pferdehaltung": "P_EmissionMM_Pferdehaltung",
}

katalog_process_df = data["Katalog_Process"][["ProcessID", "Process"]]
process_name_to_id = dict(
    zip(katalog_process_df["Process"], katalog_process_df["ProcessID"])
)

missing_target_processes = [
    target
    for target in PROCESS_TO_EMISSIONMM.values()
    if target not in process_name_to_id
]
if missing_target_processes:
    raise KeyError(
        "Folgende P_EmissionMM_xy Prozesse fehlen in Katalog_Process und müssen dort "
        f"zuerst angelegt werden, bevor dieses Skript laufen kann: {missing_target_processes}"
    )

target_process_id_by_source = {
    source: process_name_to_id[target]
    for source, target in PROCESS_TO_EMISSIONMM.items()
}

# ============================================================================
# DETERMINE TARGET BETRIEBPROCESS PER LIVESTOCK INSTANCE (1:1) AND CREATE NEW ONES
# ============================================================================
# Same pairing as in the PM/NMVOC script: existing P_EmissionMM_xy
# instances are assigned deterministically (sorted by BetriebProcessID) to the
# source instances before new ones are created. This way both
# scripts end up on the same target BetriebProcessIDs, regardless of their run order.

betrieb_process_df = data["tbl_3_Betrieb_Process"].copy()

source_instances = betrieb_process_df.merge(
    katalog_process_df, on="ProcessID", how="left"
)
source_instances = source_instances[
    source_instances["Process"].isin(PROCESS_TO_EMISSIONMM.keys())
][["BetriebID", "BetriebProcessID", "Process"]].copy()
source_instances["TargetProcessID"] = source_instances["Process"].map(
    target_process_id_by_source
)

next_bp_id = int(betrieb_process_df["BetriebProcessID"].max()) + 1

new_target_rows = []
source_to_target = {}

for (betrieb_id, target_process_id), group in source_instances.groupby(
    ["BetriebID", "TargetProcessID"]
):
    sources_sorted = group.sort_values("BetriebProcessID")["BetriebProcessID"].tolist()
    existing_targets_sorted = sorted(
        betrieb_process_df.loc[
            (betrieb_process_df["BetriebID"] == betrieb_id)
            & (betrieb_process_df["ProcessID"] == target_process_id),
            "BetriebProcessID",
        ].tolist()
    )

    for i, source_bp_id in enumerate(sources_sorted):
        if i < len(existing_targets_sorted):
            target_bp_id = existing_targets_sorted[i]
        else:
            target_bp_id = next_bp_id
            next_bp_id += 1
            new_target_rows.append(
                {
                    "BetriebProcessID": target_bp_id,
                    "BetriebID": betrieb_id,
                    "ProcessID": target_process_id,
                }
            )
        source_to_target[source_bp_id] = target_bp_id

new_target_processes_df = pd.DataFrame(
    new_target_rows, columns=["BetriebProcessID", "BetriebID", "ProcessID"]
)

if len(new_target_processes_df) > 0:
    # Update tbl_3 in memory so that later merges (display/debug) already know the new
    # processes, even before they are written to Access.
    data["tbl_3_Betrieb_Process"] = pd.concat(
        [data["tbl_3_Betrieb_Process"], new_target_processes_df], ignore_index=True
    )

print("\n=== NEUE P_EmissionMM_xy PROZESSINSTANZEN (tbl_3_Betrieb_Process) ===")
if len(new_target_processes_df) > 0:
    preview_new_targets = new_target_processes_df.merge(
        katalog_process_df, on="ProcessID", how="left"
    )
    print(
        preview_new_targets[["BetriebID", "BetriebProcessID", "Process"]].to_string(
            index=False
        )
    )
else:
    print("Keine neuen Zielprozess-Instanzen erforderlich, alle bereits vorhanden.")
print(f"{len(new_target_processes_df)} neue P_EmissionMM_xy Prozesse werden angelegt.")

# Re-map emissions from the source to the target BetriebProcessID
n_emissions_tbl4 = n_emissions_long.copy()
n_emissions_tbl4["Duenge_BetriebProcessID"] = n_emissions_tbl4["BetriebProcessID"].map(
    source_to_target
)

unmapped_n_emissions = n_emissions_tbl4[
    n_emissions_tbl4["Duenge_BetriebProcessID"].isna()
]
if len(unmapped_n_emissions) > 0:
    print(
        "⚠️ WARNING: Für einige Tierhaltungsinstanzen konnte kein Zielprozess ermittelt werden!"
    )
    print(
        unmapped_n_emissions[["BetriebID", "BetriebProcessID", "Process"]]
        .drop_duplicates()
        .to_string(index=False)
    )

n_emissions_tbl4 = n_emissions_tbl4.dropna(subset=["Duenge_BetriebProcessID"]).copy()
n_emissions_tbl4["BetriebProcessID"] = n_emissions_tbl4[
    "Duenge_BetriebProcessID"
].astype(int)

print("\n=== Check P_EmissionMM_xy-Zuordnung ===")
print(
    new_target_processes_df.head(20).to_string(index=False)
    if len(new_target_processes_df) > 0
    else "Keine neuen Instanzen."
)

print(f"""
Betriebe mit Tieremissionen:
{allocation_summary['BetriebID'].nunique()}

Tierhaltungsinstanzen mit P_EmissionMM_xy Zielprozess:
{len(source_to_target)}
""")

# ============================================================================
# VALIDATION: sum per farm unchanged compared with the old farm-level script
# ============================================================================
# The old script (4_Emissions_Duengemanagement_N2O_NOx_NH3.py) wrote all
# emissions per farm into one single P_1_Jahr_Duengemanagement process.
# Here the same emission values are only distributed over several
# P_EmissionMM_xy instances per livestock process — so the sum per
# farm must stay exactly the same. Any differences point to
# unassigned (unmapped) livestock instances.
old_style_betrieb_sum = (
    n_emissions_long.groupby("BetriebID")["Wert"]
    .sum()
    .rename("Wert_alt (Betriebsebene)")
)
new_style_betrieb_sum = (
    n_emissions_tbl4.groupby("BetriebID")["Wert"]
    .sum()
    .rename("Wert_neu (P_EmissionMM_xy je Instanz)")
)

betrieb_sum_comparison = pd.concat(
    [old_style_betrieb_sum, new_style_betrieb_sum], axis=1
).fillna(0)
betrieb_sum_comparison["Differenz"] = (
    betrieb_sum_comparison["Wert_alt (Betriebsebene)"]
    - betrieb_sum_comparison["Wert_neu (P_EmissionMM_xy je Instanz)"]
)

print("\n" + "=" * 70)
print(
    "VALIDIERUNG: Betriebssumme alt (1 Duengemanagement-Prozess) "
    "vs. neu (P_EmissionMM_xy je Instanz)"
)
print("=" * 70)
abweichungen = betrieb_sum_comparison[betrieb_sum_comparison["Differenz"].abs() > 1e-6]
if len(abweichungen) == 0:
    print("✅ Alle Betriebssummen stimmen exakt überein.")
else:
    print(
        f"⚠️ WARNUNG: {len(abweichungen)} Betriebe mit abweichender Summe "
        "(vermutlich nicht zugeordnete Instanzen):"
    )
    print(abweichungen.to_string())
print(betrieb_sum_comparison.round(2).to_string())
# ============================================================
# Prepare N emissions for tbl_4 upload
# ============================================================

insert_emissions_df = n_emissions_tbl4.copy()


insert_emissions_df["Direction"] = "Output"


insert_emissions_df = insert_emissions_df[
    [
        "BetriebProcessID",
        "FlowID",
        "Wert",
        "Direction",
        "TierkategorieEmissionFutter",
        "Tierprozess",
    ]
]


# Remove zero emissions
insert_emissions_df = insert_emissions_df[
    insert_emissions_df["Wert"].notna() & (insert_emissions_df["Wert"] != 0)
]

# ============================================================================
# NO3 LEACHING FROM PASTURE EXCRETA (mexcr_graz) -> PASTURE PROCESSES
# ============================================================================
# Above, mexcr_graz is only used for NH3 and N2O grazing. Nitrate leaching from
# pasture excreta is calculated here with the same PEF factor as for applied
# fertiliser in script 3 (1.33 kg NO3 / kg N) and assigned to the pasture
# processes (P_DG_Weide, P_DG_Maehweide) instead of P_EmissionMM_xy:
# 1. NO3 per farm = sum of mexcr_graz over all animal processes * 1.33
# 2. Split over the pasture processes by their share of the grazed feed
#    (Weidefutteraufnahme output per ha * area)
# 3. Divide by the area, since the DG processes are modelled per ha
# The P input from pasture excreta is split the same way: P2O5 = mexcr_graz *
# P2O5/N ratio of the excreta per animal process, then summed per farm. Like the
# P2O5 row from script 3, it is booked as emission to soil/agricultural with factor 1.
EF_NO3_GRAZING = 1.33  # kg NO3 / kg N, PEF 2021 (same as EF["NO3"] in script 3)
# Source: Achilles, W., Anter, J., Belau, T., & Blankenburg, J. (2018). Faustzahlen für
# die Landwirtschaft (15th ed.). KTBL. Based on:
# dairy cow = forage crop farm with grazing, 8000 kg ECM (p. 487); suckler cow = suckler
# cow with calf, 700 kg LW (p. 487); rearing / fattening = young cattle (p. 486);
# sheep = ewe (p. 491)
P2O5_PER_N_GRAZING = {  # kg P2O5 / kg N excreted on pasture
    "P_Milchkuh": 0.359,
    "P_Mutterkuh": 0.295,
    "P_Aufzucht_Milchvieh": 0.323,
    "P_Rind_Mast": 0.323,
    "P_Schafhaltung_Mast": 0.284,
}
PASTURE_PROCESSES = ["P_DG_Weide", "P_DG_Maehweide"]
TAG_NO3_GRAZING = "Weide Exkremente"

graz_processes = allocation_summary[allocation_summary["mexcr_graz"] > 0]
missing_p2o5_ratio = set(graz_processes["Process"]) - set(P2O5_PER_N_GRAZING)
if missing_p2o5_ratio:
    raise KeyError(
        f"Grazing animal processes without P2O5/N ratio: {sorted(missing_p2o5_ratio)}"
    )

no3_graz_farm = (
    allocation_summary.assign(
        NO3_Grazing_kg=allocation_summary["mexcr_graz"] * EF_NO3_GRAZING,
        P2O5_Grazing_kg=allocation_summary["mexcr_graz"]
        * allocation_summary["Process"].map(P2O5_PER_N_GRAZING).fillna(0),
    )
    .groupby("BetriebID", as_index=False)[
        ["mexcr_graz", "NO3_Grazing_kg", "P2O5_Grazing_kg"]
    ]
    .sum()
)
no3_graz_farm = no3_graz_farm[no3_graz_farm["NO3_Grazing_kg"] > 0]

# Grazed feed output and area per pasture process
pasture_tbl4 = (
    data["tbl_4_Betrieb_Process_FlowValue"]
    .merge(data["Katalog_Flow"][["FlowID", "Flow"]], on="FlowID", how="left")
    .merge(
        data["tbl_3_Betrieb_Process"][["BetriebProcessID", "BetriebID", "ProcessID"]],
        on="BetriebProcessID",
        how="left",
    )
    .merge(katalog_process_df, on="ProcessID", how="left")
)
pasture_tbl4 = pasture_tbl4[
    (pasture_tbl4["Flow"] == "Weidefutteraufnahme")
    & (pasture_tbl4["Direction"] == "Output")
    & pasture_tbl4["Process"].isin(PASTURE_PROCESSES)
].copy()
pasture_tbl4["Weidefutter_kg_pro_ha"] = pd.to_numeric(
    pasture_tbl4["Wert"].astype(str).str.replace(",", "."), errors="coerce"
).fillna(0)
pasture_tbl4 = (
    pasture_tbl4.groupby(["BetriebID", "BetriebProcessID", "Process"], as_index=False)[
        "Weidefutter_kg_pro_ha"
    ]
    .sum()
)

flaeche_df = process_parameter_values[
    process_parameter_values["Pp_name"] == "Fläche"
][["BetriebProcessID", "Wert"]].copy()
flaeche_df["Flaeche_ha"] = pd.to_numeric(
    flaeche_df["Wert"].astype(str).str.replace(",", "."), errors="coerce"
)
# tbl_6 contains empty duplicate "Fläche" rows
flaeche_df = flaeche_df.dropna(subset=["Flaeche_ha"]).drop_duplicates("BetriebProcessID")
pasture_tbl4 = pasture_tbl4.merge(
    flaeche_df[["BetriebProcessID", "Flaeche_ha"]], on="BetriebProcessID", how="left"
)
pasture_tbl4["Weidefutter_kg_gesamt"] = (
    pasture_tbl4["Weidefutter_kg_pro_ha"] * pasture_tbl4["Flaeche_ha"]
)
pasture_tbl4 = pasture_tbl4[pasture_tbl4["Flaeche_ha"] > 0]
# Split by grazed feed; if a farm has no grazed feed entered at all
# (e.g. few grazing days but pasture output 0), split by area instead
pasture_tbl4["Basis"] = pasture_tbl4["Weidefutter_kg_gesamt"].where(
    pasture_tbl4.groupby("BetriebID")["Weidefutter_kg_gesamt"].transform("sum") > 0,
    pasture_tbl4["Flaeche_ha"],
)
pasture_tbl4["Anteil"] = pasture_tbl4["Basis"] / pasture_tbl4.groupby("BetriebID")[
    "Basis"
].transform("sum")
pasture_tbl4 = pasture_tbl4[pasture_tbl4["Anteil"] > 0]

no3_pasture = pasture_tbl4.merge(no3_graz_farm, on="BetriebID", how="inner")
no3_pasture["NO3_kg_gesamt"] = no3_pasture["NO3_Grazing_kg"] * no3_pasture["Anteil"]
no3_pasture["Wert"] = no3_pasture["NO3_kg_gesamt"] / no3_pasture["Flaeche_ha"]
no3_pasture["P2O5_kg_gesamt"] = no3_pasture["P2O5_Grazing_kg"] * no3_pasture["Anteil"]
no3_pasture["P2O5_kg_pro_ha"] = no3_pasture["P2O5_kg_gesamt"] / no3_pasture["Flaeche_ha"]

# No pasture N may get lost: farms with grazing animals but no pasture process
farms_without_pasture = set(no3_graz_farm["BetriebID"]) - set(no3_pasture["BetriebID"])
if farms_without_pasture:
    raise ValueError(
        "Farms with mexcr_graz > 0 but no pasture process with grazed feed "
        f"and area: {sorted(farms_without_pasture)}"
    )
no3_check = (
    no3_pasture.groupby("BetriebID")["NO3_kg_gesamt"].sum()
    - no3_graz_farm.set_index("BetriebID")["NO3_Grazing_kg"]
)
if (no3_check.abs() > 1e-6).any():
    raise ValueError(f"NO3 grazing: split loses NO3:\n{no3_check}")
p2o5_check = (
    no3_pasture.groupby("BetriebID")["P2O5_kg_gesamt"].sum()
    - no3_graz_farm.set_index("BetriebID")["P2O5_Grazing_kg"]
)
if (p2o5_check.abs() > 1e-6).any():
    raise ValueError(f"P2O5 grazing: split loses P2O5:\n{p2o5_check}")

print("\n=== NO3 AND P2O5 FROM PASTURE EXCRETA PER FARM AND PASTURE PROCESS ===")
no3_display = no3_pasture.merge(
    data["Katalog_Betriebe"][["BetriebID", "ZirKUH_ID"]], on="BetriebID", how="left"
)
print(
    no3_display[
        [
            "ZirKUH_ID",
            "BetriebProcessID",
            "Process",
            "mexcr_graz",
            "NO3_Grazing_kg",
            "Weidefutter_kg_pro_ha",
            "Flaeche_ha",
            "Anteil",
            "NO3_kg_gesamt",
            "Wert",
            "P2O5_Grazing_kg",
            "P2O5_kg_gesamt",
            "P2O5_kg_pro_ha",
        ]
    ]
    .rename(columns={"Wert": "NO3_kg_pro_ha"})
    .sort_values(["ZirKUH_ID", "Process"])
    .round(2)
    .to_string(index=False)
)

nitrate_flow_id = data["Katalog_Flow"].loc[
    (data["Katalog_Flow"]["Flow"] == "Nitrate")
    & data["Katalog_Flow"]["Category"].str.contains("Emission to water", na=False),
    "FlowID",
]
if len(nitrate_flow_id) != 1:
    raise KeyError(f"Nitrate (Emission to water) not unique in Katalog_Flow: {list(nitrate_flow_id)}")

no3_pasture_rows = no3_pasture[["BetriebProcessID", "Wert"]].assign(
    FlowID=int(nitrate_flow_id.iloc[0]),
    Direction="Output",
    TierkategorieEmissionFutter=TAG_NO3_GRAZING,
    Tierprozess="NO3_Grazing",
)
p2o5_flow_id = data["Katalog_Flow"].loc[
    (data["Katalog_Flow"]["Flow"] == "Phosphorus pentoxide")
    & data["Katalog_Flow"]["Category"].str.contains("Emission to soil/agricultural", na=False),
    "FlowID",
]
if len(p2o5_flow_id) != 1:
    raise KeyError(
        "Phosphorus pentoxide (Emission to soil/agricultural) not unique in "
        f"Katalog_Flow: {list(p2o5_flow_id)}"
    )

p2o5_pasture_rows = (
    no3_pasture[["BetriebProcessID", "P2O5_kg_pro_ha"]]
    .rename(columns={"P2O5_kg_pro_ha": "Wert"})
    .assign(
        FlowID=int(p2o5_flow_id.iloc[0]),
        Direction="Output",
        TierkategorieEmissionFutter=TAG_NO3_GRAZING,
        Tierprozess="P2O5_Grazing",
    )
)
insert_emissions_df = pd.concat(
    [
        insert_emissions_df,
        no3_pasture_rows[insert_emissions_df.columns],
        p2o5_pasture_rows[insert_emissions_df.columns],
    ],
    ignore_index=True,
)

# Round Wert to 2 decimal places
insert_emissions_df["Wert"] = insert_emissions_df["Wert"].round(2)
# ============================================================================
# Debug
# ============================================================================
check = allocation_summary[allocation_summary["Process"] == "P_Milchkuh"]
print("\ncheck P_Milchkuh===============")
print(
    check[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "NH3_Gas_House",
            "NH3_Gas_MM",
            "N2O_MM_Gas",
        ]
    ].to_string(index=False)
)
check = allocation_summary[allocation_summary["Process"] == "P_Milchkuh"]

print(
    check[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "TAN_Faktor",
            "mexcr_stall",
            "mTANhouse",
            "NH3_Gas_House",
            "NH3_Gas_MM",
            "NH3_Gas_Grazing",
            "N2O_MM_Gas",
            "N2O_Gas_Grazing",
            "N2O_Gas_deposition",
            "NO_Gas_deposition",
        ]
    ].to_string(index=False)
)
print("\n===========================================\n")
# Debug=====================================================================
print("\n Debuglines \n")
print(
    allocation_summary[allocation_summary["Process"] == "P_Aufzucht_Milchvieh"][
        [
            "Process",
            "Housing_System",
            "mexcr_stall",
            "mTANhouse",
            "NH3_N_House",
            "mTANstorageliquid",
            "mTANstoragesolid",
            "mTANstorage",
            "EF_NH3_MM",
            "NH3_Gas_MM",
        ]
    ].to_string(index=False)
)
print("\nEnde Debug Line=======================================\n")
# ============================================================================
# ROBUST UPSERT: UPDATE EXISTING EMISSIONS, ADD NEW ONES
# ============================================================================

cursor = conn.cursor()

# Load existing emission rows from Access to avoid writing duplicates
existing_columns = data["tbl_4_Betrieb_Process_FlowValue"].columns
existing_select = [
    "BetriebProcessFlowValueID",
    "BetriebProcessID",
    "FlowID",
    "Direction",
]
merge_keys = ["BetriebProcessID", "FlowID", "Direction"]

if "TierkategorieEmissionFutter" in existing_columns:
    existing_select.append("TierkategorieEmissionFutter")
    merge_keys.append("TierkategorieEmissionFutter")

existing_sql = f"SELECT {', '.join(existing_select)} FROM tbl_4_Betrieb_Process_FlowValue WHERE Direction = 'Output'"
existing_df = pd.read_sql(existing_sql, conn)

# Link with the newly calculated emissions
merge_df = insert_emissions_df.merge(
    existing_df,
    on=merge_keys,
    how="left",
)

# Existing rows -> UPDATE
update_rows = merge_df[merge_df["BetriebProcessFlowValueID"].notna()].copy()

# Rows that do not exist yet -> INSERT
insert_rows = merge_df[merge_df["BetriebProcessFlowValueID"].isna()].copy()

# Summary of the planned changes
print("\n" + "=" * 60)
print("ZUSAMMENFASSUNG DER GEPLANTEN ÄNDERUNGEN")
print("=" * 60)
print(
    f"{len(new_target_processes_df)} neue P_EmissionMM_xy Prozessinstanzen "
    "werden in tbl_3_Betrieb_Process angelegt."
)
print(f"{len(update_rows)} vorhandene Emissionen werden aktualisiert.")
print(f"{len(insert_rows)} neue Emissionen werden eingefügt.")

if len(new_target_processes_df) > 0:
    print("\n=== NEUE P_EmissionMM_xy PROZESSINSTANZEN ===")
    preview_new_targets = new_target_processes_df.merge(
        katalog_process_df, on="ProcessID", how="left"
    )
    print(
        preview_new_targets[["BetriebID", "BetriebProcessID", "Process"]].to_string(
            index=False
        )
    )

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
    display_cols = ["BetriebProcessID", "Process", "Flow", "Category", "Wert"]
    if "TierkategorieEmissionFutter" in insert_rows.columns:
        display_cols.insert(4, "TierkategorieEmissionFutter")
    print(new_emissions_display[display_cols].to_string(index=False))
    print("=" * 50)

confirm = input("Änderungen in die Access-Datenbank schreiben? (j/n): ")

if confirm.lower() == "j":
    # Create new P_EmissionMM_xy process instances
    if len(new_target_processes_df) > 0:
        for _, row in new_target_processes_df.iterrows():
            cursor.execute(
                "INSERT INTO tbl_3_Betrieb_Process (BetriebProcessID, BetriebID, ProcessID) VALUES (?, ?, ?)",
                (
                    int(row["BetriebProcessID"]),
                    int(row["BetriebID"]),
                    int(row["ProcessID"]),
                ),
            )

    # Update existing values
    if "TierkategorieEmissionFutter" in merge_keys:
        update_sql = """
            UPDATE tbl_4_Betrieb_Process_FlowValue
            SET Wert = ?, TierkategorieEmissionFutter = ?
            WHERE BetriebProcessFlowValueID = ?
            """
    else:
        update_sql = """
            UPDATE tbl_4_Betrieb_Process_FlowValue
            SET Wert = ?
            WHERE BetriebProcessFlowValueID = ?
            """

    for _, row in update_rows.iterrows():
        if "TierkategorieEmissionFutter" in merge_keys:
            cursor.execute(
                update_sql,
                row["Wert"],
                row["TierkategorieEmissionFutter"],
                int(row["BetriebProcessFlowValueID"]),
            )
        else:
            cursor.execute(
                update_sql,
                row["Wert"],
                int(row["BetriebProcessFlowValueID"]),
            )

    # Insert new values
    insert_cols = ["BetriebProcessID", "FlowID", "Wert", "Direction"]
    insert_sql_cols = "(BetriebProcessID, FlowID, Wert, Direction)"
    insert_placeholders = "VALUES (?, ?, ?, ?)"
    if "TierkategorieEmissionFutter" in merge_keys:
        insert_cols.append("TierkategorieEmissionFutter")
        insert_sql_cols = (
            "(BetriebProcessID, FlowID, Wert, Direction, TierkategorieEmissionFutter)"
        )
        insert_placeholders = "VALUES (?, ?, ?, ?, ?)"

    if len(insert_rows) > 0:
        rows_to_insert = list(
            insert_rows[insert_cols].itertuples(index=False, name=None)
        )
        cursor.executemany(
            f"""
            INSERT INTO tbl_4_Betrieb_Process_FlowValue
            {insert_sql_cols}
            {insert_placeholders}
            """,
            rows_to_insert,
        )
    else:
        print("Keine neuen Emissionen zum Einfügen.")

    conn.commit()
    print("\n✅ Änderungen wurden in die Access-Datenbank geschrieben.")
else:
    print("\n❌ Änderungen wurden verworfen.")
