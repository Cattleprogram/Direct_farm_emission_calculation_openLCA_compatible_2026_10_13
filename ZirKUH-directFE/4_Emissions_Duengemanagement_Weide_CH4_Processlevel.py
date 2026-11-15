print(
    "Hello world!, Let us calculate the CH4 emissions for pasture and manure management using VS and DOM 🐄💨"
)

import pyodbc
import pandas as pd
import numpy as np
from config import db_path


def require_columns(df, df_name, columns):
    missing = [col for col in columns if col not in df.columns]
    if missing:
        raise KeyError(
            f"Missing required columns in {df_name}: {missing}. "
            f"Available columns: {list(df.columns)}"
        )


def to_numeric_de(series):
    """
    Converts values in German number format (e.g. "205,5" or "1.234,5")
    to float. Unreadable, empty and "None"/"nan" values become NaN.
    """
    s = series.astype(str).str.replace(r"\s+", "", regex=True)

    # Remove thousands separator only if both "." and "," occur
    has_both = s.str.contains(".", regex=False) & s.str.contains(",", regex=False)
    s = s.where(~has_both, s.str.replace(".", "", regex=False))

    s = s.str.replace(",", ".", regex=False)
    s = s.replace(["", "None", "nan", "NaN", "<NA>"], np.nan)

    return pd.to_numeric(s, errors="coerce")


# ============================================================================
# DATABASE CONNECTION AND DATA LOADING
# ============================================================================


def connect_database(db_path):
    """
    Connects to the Access database.
    """
    conn = pyodbc.connect(
        rf"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
    )
    return conn


def load_all_tables(conn):
    """
    Loads all Access tables as pandas DataFrames
    into a dictionary.
    """
    cursor = conn.cursor()
    tables = cursor.tables()

    # Load real tables only (no Access system tables)
    table_list = [
        row.table_name
        for row in tables
        if not row.table_name.startswith("MSys") and row.table_type == "TABLE"
    ]

    data = {}

    for table_name in table_list:

        try:
            data[table_name] = pd.read_sql_query(
                f"SELECT * FROM [{table_name}]",
                conn,
            )

            print(
                f"✓ Tabelle '{table_name}' geladen: " f"{len(data[table_name])} Zeilen"
            )

        except Exception as e:

            print(f"✗ Fehler beim Laden von '{table_name}': {e}")

    return data


# ============================================================================
# LOAD DATABASE
# ============================================================================

conn = connect_database(db_path)

print(f"\n📊 Lade alle Tabellen aus: {db_path}\n")

data = load_all_tables(conn)

print(f"\n✅ Total {len(data)} Tabellen geladen\n")


# ============================================================================
# OVERVIEW OF LOADED TABLES
# ============================================================================

print("Verfügbare Tabellen für Berechnungen:")

for table_name in data.keys():

    print(
        f"  - {table_name}: "
        f"{len(data[table_name])} Zeilen, "
        f"{len(data[table_name].columns)} Spalten"
    )


# ============================================================================
# IDENTIFY FEED INPUTS
# ============================================================================

# Table with all process flows of the farms
flow_df = data["tbl_4_Betrieb_Process_FlowValue"]

# Load flow catalogue
katalog_flow_df = data["Katalog_Flow"]

# Select only flows
# that are marked as potential feedstuff in the catalogue
feed_flows = katalog_flow_df[katalog_flow_df["IsPotentialFeedstuff"] == True][
    ["FlowID", "Flow"]
]
# Table with the farm processes
betrieb_process_df = data["tbl_3_Betrieb_Process"]
betrieb_process_subset = betrieb_process_df[
    [
        "BetriebProcessID",
        "ProcessID",
    ]
]

# Link feed flows with the farm flows
feed_inputs = flow_df.merge(
    feed_flows,
    on="FlowID",
    how="inner",
)

# Consider inputs only
# (remove outputs such as milk, meat etc.)
feed_inputs = feed_inputs[feed_inputs["Direction"] == "Input"]

print(f"\n🐄 Futtermittelinputs gefunden: {len(feed_inputs)}")

# ============================================================================
# ATTACH PROCESSID TO FEED INPUTS
# ============================================================================

betrieb_process_df = data["tbl_3_Betrieb_Process"]

betrieb_process_subset = betrieb_process_df[
    ["BetriebProcessID", "ProcessID", "BetriebID"]
]

feed_inputs = feed_inputs.merge(
    betrieb_process_subset,
    on="BetriebProcessID",
    how="left",
)

print("\n🔗 ProcessID ergänzt")

# ============================================================================
# LOAD FARM-SPECIFIC FLOWPROPERTIES
# ============================================================================

# Table with farm-specific FlowProperties
tbl5 = data["tbl_5_Betrieb_FlowpropertyValue"]

# Catalogue of FlowProperties
flowprop = data["Katalog_Flowproperty"]

# For the VS calculation only XF and XA are relevant for now
xa_xf_props = flowprop[flowprop["FlowProperty"].isin(["XA", "XF"])][
    ["FlowPropertyID", "FlowProperty"]
]

# Keep only XF/XA properties from tbl_5
tbl5 = tbl5.merge(
    xa_xf_props,
    on="FlowPropertyID",
    how="inner",
)

print(f"\n📋 Betriebsindividuelle XF/XA Werte gefunden: {len(tbl5)}")


# ============================================================================
# LOAD KTBL DEFAULT VALUES
# ============================================================================

# Mapping between flows and KTBL feedstuffs
mapping_df = data["tbl_Futtermittel_Mapping"]

# KTBL default feedstuff values
ktbl_df = data["Katalog_Futtermittelgehalte_KTBL"]

# Link feed inputs with the KTBL mapping
feed_inputs = feed_inputs.merge(
    mapping_df,
    on="FlowID",
    how="left",
)

print("\n🔗 Futtermittel mit KTBL-Mapping verknüpft")


# ============================================================================
# ADD XF/XA VALUES
# ============================================================================

# Goal:
# XF and XA are to be added for every feedstuff.
#
# Priority:
# 1. Farm-specific values from tbl_5
# 2. If empty there -> KTBL default values


# ============================================================================
# PREPARE FARM-SPECIFIC XF/XA VALUES
# ============================================================================

# Select XF/XA properties
xa_xf_props = flowprop[flowprop["FlowProperty"].isin(["XA", "XF"])][
    ["FlowPropertyID", "FlowProperty"]
]


# Turn rows into columns
#
# before:
# BetriebProcessFlowValueID | FlowProperty | Wert
#
# after:
# BetriebProcessFlowValueID | XF | XA

# Keep raw values for diagnostics, then decimal comma -> float
tbl5["Wert_raw"] = tbl5["Wert"]
tbl5["Wert"] = to_numeric_de(tbl5["Wert"])

tbl5_pivot = tbl5.pivot_table(
    index="BetriebProcessFlowValueID",
    columns="FlowProperty",
    values="Wert",
    aggfunc="first",
).reset_index()

# Clean up pivot column names
tbl5_pivot.columns.name = None

# Pivot the raw values as well (XF_raw, XA_raw)
tbl5_raw_pivot = tbl5.pivot_table(
    index="BetriebProcessFlowValueID",
    columns="FlowProperty",
    values="Wert_raw",
    aggfunc="first",
)
tbl5_raw_pivot.columns = [f"{col}_raw" for col in tbl5_raw_pivot.columns]
tbl5_pivot = tbl5_pivot.merge(
    tbl5_raw_pivot.reset_index(),
    on="BetriebProcessFlowValueID",
    how="left",
)

print("\n📋 Pivotierte XF/XA Betriebswerte:")
print(tbl5_pivot.head())


# ============================================================================
# ATTACH FARM VALUES TO FEED INPUTS
# ============================================================================

feed_inputs = feed_inputs.merge(
    tbl5_pivot,
    on="BetriebProcessFlowValueID",
    how="left",
)

print("\n🔗 Futtermittelinputs mit Betriebswerten verknüpft")


# ============================================================================
# ADD KTBL DEFAULT VALUES
# ============================================================================

# Select relevant KTBL columns
ktbl_subset = ktbl_df[
    [
        "FuttermittelID",
        "XF (g/kg TM)",
        "XA (g/kg TM)",
    ]
]

# Attach KTBL values
feed_inputs = feed_inputs.merge(
    ktbl_subset,
    on="FuttermittelID",
    how="left",
)

print("\n📚 KTBL XF/XA Werte ergänzt")


# ============================================================================
# PREFER FARM VALUES, fallback to KTBL
# ============================================================================

# If farm-specific values exist:
# -> use them
#
# otherwise:
# -> use KTBL default values

# Keep raw values for diagnostics and the old/new comparison
feed_inputs["Wert_raw"] = feed_inputs["Wert"]
feed_inputs["KTBL_XF_raw"] = feed_inputs["XF (g/kg TM)"]
feed_inputs["KTBL_XA_raw"] = feed_inputs["XA (g/kg TM)"]

# Convert to numbers first (decimal comma!), then combine_first.
# This way the KTBL default applies when a farm value is unreadable.
for col in ["XF", "XA", "XF (g/kg TM)", "XA (g/kg TM)", "Wert"]:
    feed_inputs[col] = to_numeric_de(feed_inputs[col])

feed_inputs["XF_final"] = feed_inputs["XF"].combine_first(feed_inputs["XF (g/kg TM)"])

feed_inputs["XA_final"] = feed_inputs["XA"].combine_first(feed_inputs["XA (g/kg TM)"])

print("\n✅ Finale XF/XA Werte erzeugt")


# ============================================================================
# DIAGNOSTICS: FEEDSTUFFS WITHOUT A VALID XF/XA VALUE
# ============================================================================

missing_xf_xa = feed_inputs["XF_final"].isna() | feed_inputs["XA_final"].isna()

print(f"\n🔍 Futtermittelzeilen ohne gültigen XF oder XA Wert: {missing_xf_xa.sum()}")

if missing_xf_xa.any():
    print(
        feed_inputs.loc[
            missing_xf_xa,
            [
                "BetriebID",
                "BetriebProcessID",
                "Flow",
                "Wert",
                "Wert_raw",
                "XF_raw",
                "KTBL_XF_raw",
                "XA_raw",
                "KTBL_XA_raw",
            ],
        ].to_string(index=False)
    )


# ============================================================================
# CALCULATE WEIGHTED MEAN XF/XA CONTENTS
# ============================================================================

# Calculate weighted XF/XA amounts
feed_inputs["XF_weighted"] = feed_inputs["Wert"] * feed_inputs["XF_final"]

feed_inputs["XA_weighted"] = feed_inputs["Wert"] * feed_inputs["XA_final"]

# DM of feedstuffs with a valid XF or XA value only
feed_inputs["Wert_XF"] = feed_inputs["Wert"].where(feed_inputs["XF_final"].notna())

feed_inputs["Wert_XA"] = feed_inputs["Wert"].where(feed_inputs["XA_final"].notna())

# Reproduce the old calculation (only for the old/new comparison):
# combine_first on raw values, comma replaced in Wert only
wert_alt = pd.to_numeric(
    feed_inputs["Wert_raw"].astype(str).str.replace(",", ".", regex=False),
    errors="coerce",
)
xf_alt = pd.to_numeric(
    feed_inputs["XF_raw"].combine_first(feed_inputs["KTBL_XF_raw"]), errors="coerce"
)
xa_alt = pd.to_numeric(
    feed_inputs["XA_raw"].combine_first(feed_inputs["KTBL_XA_raw"]), errors="coerce"
)
feed_inputs["Wert_alt"] = wert_alt
feed_inputs["XF_weighted_alt"] = wert_alt * xf_alt
feed_inputs["XA_weighted_alt"] = wert_alt * xa_alt

# Sum per BetriebProcessID
summary = (
    feed_inputs.groupby(["BetriebID", "BetriebProcessID", "ProcessID"])
    .agg(
        # total DM eaten per year
        DM=("Wert", "sum"),
        # DM with a valid XF/XA content
        DM_XF=("Wert_XF", "sum"),
        DM_XA=("Wert_XA", "sum"),
        # weighted XF/XA sums
        XF_sum=("XF_weighted", "sum"),
        XA_sum=("XA_weighted", "sum"),
        # old calculation
        DM_alt=("Wert_alt", "sum"),
        XF_sum_alt=("XF_weighted_alt", "sum"),
        XA_sum_alt=("XA_weighted_alt", "sum"),
    )
    .reset_index()
)

# Calculate mean XF/XA contents (only over feedstuffs with a valid content)
summary["XF_mean"] = summary["XF_sum"] / summary["DM_XF"].replace(0, np.nan)

summary["XA_mean"] = summary["XA_sum"] / summary["DM_XA"].replace(0, np.nan)

print("\n🧮 Gewichtete XF/XA-Gehalte berechnet")

print(summary.head())

# Warning if part of the DM has no valid content
for prop in ["XF", "XA"]:
    incomplete = summary[summary[f"DM_{prop}"] < summary["DM"]]
    for _, row in incomplete.iterrows():
        missing_share = (1 - row[f"DM_{prop}"] / row["DM"]) * 100
        print(
            f"⚠️ WARNUNG: BetriebProcessID {row['BetriebProcessID']} "
            f"(BetriebID {row['BetriebID']}): {missing_share:.1f} % der TM "
            f"ohne gültigen {prop}-Wert"
        )

# ============================================================================
# COMPARISON OF XF/XA MEANS OLD VS. NEW
# ============================================================================

xf_xa_comparison = pd.DataFrame(
    {
        "BetriebID": summary["BetriebID"],
        "BetriebProcessID": summary["BetriebProcessID"],
        "XF_mean_alt": summary["XF_sum_alt"] / summary["DM_alt"],
        "XF_mean_neu": summary["XF_mean"],
        "XA_mean_alt": summary["XA_sum_alt"] / summary["DM_alt"],
        "XA_mean_neu": summary["XA_mean"],
    }
)
xf_xa_comparison["XF_diff"] = (
    xf_xa_comparison["XF_mean_neu"] - xf_xa_comparison["XF_mean_alt"]
)
xf_xa_comparison["XA_diff"] = (
    xf_xa_comparison["XA_mean_neu"] - xf_xa_comparison["XA_mean_alt"]
)

print("\n🔍 XF/XA Mittelwerte alt vs. neu (g/kg TM):\n")

print(xf_xa_comparison.round(2).to_string(index=False))

# ============================================================================
# Short debug window for farm 8 (BetriebProcessID 31 = P_Mutterkuh)
# ============================================================================

betrieb_id = 31

betrieb_debug = feed_inputs[feed_inputs["BetriebProcessID"] == betrieb_id].copy()

betrieb_debug["XF_source"] = np.where(
    betrieb_debug["XF"].notna(), "Betriebswert", "KTBL"
)
total_dm = betrieb_debug["Wert"].sum()

betrieb_debug["DM_share_percent"] = betrieb_debug["Wert"] / total_dm * 100

betrieb_debug["XF_contribution"] = (
    betrieb_debug["XF_weighted"] / betrieb_debug["XF_weighted"].sum() * 100
)
debug_cols = [
    "Flow",
    "Wert",
    "DM_share_percent",
    "XF_final",
    "XF_source",
    "XF_weighted",
    "XF_contribution",
]

betrieb_debug = betrieb_debug[debug_cols]

betrieb_debug = betrieb_debug.sort_values("DM_share_percent", ascending=False)

pd.set_option("display.max_rows", 100)

print("\n🔍 Diagnose Betrieb 8 (BetriebProcessID 31):\n")

print(betrieb_debug)

# ============================================================================
# CALCULATE DOM
# ============================================================================

# ============================================================================
# ATTACH PROCESS INFORMATION TO SUMMARY
# ============================================================================

# Load process catalogue
process_df = data["Katalog_Process"]

# Select relevant columns only
process_subset = process_df[
    [
        "ProcessID",
        "Process",
        "ProcessCategoryID",
    ]
]

# Attach process info to feed_inputs
feed_inputs = feed_inputs.merge(
    process_subset,
    on="ProcessID",
    how="left",
)

# Attach process info to summary
summary = summary.reset_index().merge(
    process_subset,
    on="ProcessID",
    how="left",
)

print("\n🔗 Prozessinformationen ergänzt")


# ============================================================================
# DEFINE DOM TYPES
# ============================================================================

# All processes
# that are treated with the cattle equation

cattle_processes = [
    "P_Milchkuh",
    "P_Mutterkuh",
    "P_Rind_Mast",
    "P_Aufzucht_Milchvieh",
    "P_Pferdehaltung",
    "P_Schafhaltung_Mast",
]

# All processes
# that are treated with the pig equation

pig_processes = [
    "P_Mastschwein",
    "P_Legehennen",
    "P_Bruderhahnmast",
]

# Assign DOM type
summary["DOM_type"] = np.where(
    summary["Process"].isin(cattle_processes),
    "DOMRind",
    np.where(
        summary["Process"].isin(pig_processes),
        "DOMSchwein",
        "UNDEFINED",
    ),
)

print("\n🐄🐖 DOM-Typen zugewiesen")


# ============================================================================
# CALCULATE DOM
# ============================================================================
# ============================================================================
# ORGANIC MATTER DIGESTIBILITY (X_DOM)
# ============================================================================
# Source: Kirchgeßner (1997), as cited in Arbeitsgruppe BEK (2026, Eqs. 7-8).
#
# The original regressions are formulated in percent:
#     X_DOM [%] = 87.6 - 0.81 * XF [% DM]   (cattle, sheep, horses)
#     X_DOM [%] = 92.2 - 1.68 * XF [% DM]   (pigs, poultry)
#
# BEK states the units of X_DOM and x_CF as kg kg-1 but keeps the percent
# coefficients. Taken literally, this gives X_DOM of about 87, not a fraction.
#
# Here both variables are expressed in kg kg-1. Substituting
# X_DOM[%] = 100 * X_DOM[kg/kg] and XF[%] = 100 * XF[kg/kg] and dividing
# by 100 scales only the intercept; the slope stays the same because the
# unit factor cancels:
#     X_DOM [kg/kg] = 0.876 - 0.81 * XF [kg/kg DM]
#     X_DOM [kg/kg] = 0.922 - 1.68 * XF [kg/kg DM]
#
# Correction note: an earlier version used (87.6 - 0.81 * XF_kgkg) / 100.
# That divided the slope by 100 twice, so X_DOM stayed at ~0.874 regardless
# of diet fibre content. VS and manure CH4 were underestimated by a factor
# of ~2.3-2.5 for the case farms.
# Example: XF = 205.5 g/kg DM -> X_DOM = 0.876 - 0.81 * 0.2055 = 0.710.
# ============================================================================

# Convert XF from g/kg DM -> kg/kg
summary["XF_kgkg"] = summary["XF_mean"] / 1000

# Create empty DOM column
summary["X_DOM"] = np.nan


# ============================================================================
# CATTLE EQUATION
# ============================================================================

summary.loc[
    summary["DOM_type"] == "DOMRind",
    "X_DOM",
] = (87.6 / 100) - 0.81 * summary.loc[
    summary["DOM_type"] == "DOMRind",
    "XF_kgkg",
]


# ============================================================================
# PIG EQUATION
# ============================================================================

summary.loc[
    summary["DOM_type"] == "DOMSchwein",
    "X_DOM",
] = (92.2 / 100) - 1.68 * summary.loc[
    summary["DOM_type"] == "DOMSchwein",
    "XF_kgkg",
]


print("\n🧮 X_DOM berechnet")


# ============================================================================
# DOM DEBUG TABLE
# ============================================================================

dom_debug = summary[
    [
        "BetriebProcessID",
        "Process",
        "DOM_type",
        "DM",
        "XF_mean",
        "XA_mean",
        "X_DOM",
    ]
].copy()

# Plausibility check
dom_debug["DOM_flag"] = np.where(
    (dom_debug["X_DOM"] < 0.5) | (dom_debug["X_DOM"] > 0.95),
    "CHECK",
    "OK",
)

# Sorting
dom_debug = dom_debug.sort_values(["DOM_type", "X_DOM"])

print("\n🔍 DOM Debug Tabelle:\n")

print(dom_debug.head(50))

# =================================================================
# VS CALCULATION
# ============================================================================

# ============================================================================
# CALCULATE VS
# ============================================================================

# Convert XA from g/kg DM -> kg/kg
summary["XA_kgkg"] = summary["XA_mean"] / 1000

# ============================================================================
# CALCULATE TOTAL VS
# ============================================================================

# Formula according to BEK:
#
# VS = DMfeed,int * (1 - X_DOM) * (1 - x_ash)
#
# Result:
# volatile solids (kg VS / year)

summary["VS_total"] = summary["DM"] * (1 - summary["X_DOM"]) * (1 - summary["XA_kgkg"])

print("\n💨 VS_total berechnet")


# ============================================================================
# VS DEBUG TABLE
# ============================================================================

vs_debug = summary[
    [
        "BetriebProcessID",
        "Process",
        "ProcessCategoryID",
        "DOM_type",
        "DM",
        "XF_mean",
        "XA_mean",
        "X_DOM",
        "XA_kgkg",
        "VS_total",
    ]
].copy()


# ============================================================================
# PLAUSIBILITY CHECK
# ============================================================================

vs_debug["VS_flag"] = np.where(
    (vs_debug["VS_total"] < 0) | (vs_debug["VS_total"] > vs_debug["DM"]),
    "CHECK",
    "OK",
)

# ============================================================================
# WARNING FOR PLAUSIBILITY ERRORS
# ============================================================================

n_checks = (vs_debug["VS_flag"] == "CHECK").sum()

if n_checks > 0:

    print(f"\n⚠️ WARNUNG: {n_checks} Prozesse mit auffälligen VS-Werten gefunden!\n")

    print(
        vs_debug.loc[
            vs_debug["VS_flag"] == "CHECK",
            [
                "Betrieb\nProcessID",
                "Tierprozess",
                "DM\nkg/a",
                "DOM\nkg/kg",
                "XA\nkg/kg",
                "VS\nkg/a",
                "Flag",
            ],
        ].to_string(index=False)
    )

else:

    print("\n✅ Alle VS-Werte plausibel\n")

# ============================================================================
# LOAD GRAZING DAYS
# ============================================================================

# Table with farm-specific process parameters
tbl6 = data["tbl_6_Betrieb_ProcessParameterValue"]

# Mapping:
# which ProcessPpID belongs to which parameter
process_pp = data["Katalog_ProcessProcessparameter"]

# Catalogue of parameter names
processparam = data["Katalog_Processparameter"]


# ============================================================================
# ATTACH PROCESSPARAMETERS TO tbl6
# ============================================================================

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

# ============================================================================
# ATTACH PARAMETER NAMES
# ============================================================================

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


# ============================================================================
# SELECT GRAZING DAYS ONLY
# ============================================================================

weidetage_df = tbl6[tbl6["Pp_name"] == "Weidetage"].copy()

print(f"\n🌱 Weidetage gefunden: {len(weidetage_df)}")

# ============================================================================
# ATTACH GRAZING DAYS TO SUMMARY
# ============================================================================

weidetage_subset = weidetage_df[
    [
        "BetriebProcessID",
        "Wert",
    ]
].rename(columns={"Wert": "Weidetage"})

summary = summary.merge(
    weidetage_subset,
    on="BetriebProcessID",
    how="left",
)

# Missing values = no grazing
summary["Weidetage"] = summary["Weidetage"].fillna(0)

# make numeric (mind the decimal comma)
summary["Weidetage"] = to_numeric_de(summary["Weidetage"]).fillna(0)

print("\n🌱 Weidetage an summary ergänzt")

# ============================================================================
# SPLIT VS INTO PASTURE AND HOUSING
# ============================================================================

# Share of grazing
summary["Grazing_fraction"] = summary["Weidetage"] / 365

# Safety limit
summary["Grazing_fraction"] = summary["Grazing_fraction"].clip(0, 1)

# VS on pasture
summary["VS_graz"] = summary["VS_total"] * summary["Grazing_fraction"]

# VS in the barn
summary["VS_stall"] = summary["VS_total"] * (1 - summary["Grazing_fraction"])

print("\n🐄 VS_graz und VS_stall berechnet")

# ============================================================================
# VS PASTURE / BARN DEBUG
# ============================================================================

vs_split_debug = summary[
    [
        "BetriebProcessID",
        "Process",
        "DOM_type",
        "Weidetage",
        "VS_total",
        "VS_graz",
        "VS_stall",
    ]
].copy()

# ============================================================================
# At this point the total nitrogen excretion
# and the TAN excretion could also be calculated, to calculate the N2O emissions later.
# ============================================================================


# =================================================================
# Now comes the formula part for the CH4 emissions, based on the VS values.
# ============================================================================

# ============================================================================
# CALCULATE CH4 FROM MANURE MANAGEMENT
# ============================================================================

# Equation from Arbeitsgruppe BEK (2026):
#
# CH4_MM,i = VS_i * B0 * MCF_i * rho_CH4
#
# where:
#
# CH4_MM,i   = methane emissions from manure management [kg CH4]
# VS_i       = volatile solids [kg VS]
# B0         = maximum methane producing capacity [m3 CH4 / kg VS]
# MCF_i      = methane conversion factor of the housing system [-]
# rho_CH4    = density of methane [kg CH4 / m3 CH4]
#
# rho_CH4 = 0.67 kg/m3
# ============================================================================
# ============================================================================
# DEFINE B0 VALUES
# ============================================================================

summary["Animal_type"] = np.where(
    summary["Process"].isin(
        [
            "P_Milchkuh",
            "P_Mutterkuh",
            "P_Rind_Mast",
            "P_Aufzucht_Milchvieh",
        ]
    ),
    "Rind",
    np.where(
        summary["Process"].isin(
            [
                "P_Mastschwein",
            ]
        ),
        "Schwein",
        np.where(
            summary["Process"].isin(
                [
                    "P_Legehennen",
                    "P_Bruderhahnmast",
                ]
            ),
            "Huhn",
            np.where(
                summary["Process"].isin(
                    [
                        "P_Pferdehaltung",
                    ]
                ),
                "Pferd",
                np.where(
                    summary["Process"].isin(
                        [
                            "P_Schafhaltung_Mast",
                        ]
                    ),
                    "Schaf",
                    "UNDEFINED",
                ),
            ),
        ),
    ),
)

# Assign B0 values
summary["B0"] = np.nan

summary.loc[summary["Animal_type"] == "Rind", "B0"] = 0.23

summary.loc[summary["Animal_type"] == "Schwein", "B0"] = 0.30

summary.loc[summary["Animal_type"] == "Huhn", "B0"] = 0.39

summary.loc[summary["Animal_type"] == "Pferd", "B0"] = 0.30

summary.loc[summary["Animal_type"] == "Schaf", "B0"] = 0.19
# Load housing system
# ============================================================================
# LOAD HOUSING SYSTEM
# ============================================================================

stallsystem_df = tbl6[tbl6["PpID"] == 79].copy()

print(f"\n🏠 Stallsysteme gefunden: {len(stallsystem_df)}")
# Attach housing system to summary
stallsystem_subset = stallsystem_df[
    [
        "BetriebProcessID",
        "Wert",
    ]
].rename(columns={"Wert": "Stallsystem"})

summary = summary.merge(
    stallsystem_subset,
    on="BetriebProcessID",
    how="left",
)
# Assign MCF values

# ============================================================================
# DEFINE MCF VALUES
# ============================================================================

summary["MCF"] = np.nan

# Pasture
summary["MCF_graz"] = 0.01

# Housing management
summary.loc[
    summary["Stallsystem"].str.contains(
        "Festmist",
        case=False,
        na=False,
    ),
    "MCF",
] = 0.02

summary.loc[
    summary["Stallsystem"].str.contains(
        "Tiefstreu",
        case=False,
        na=False,
    ),
    "MCF",
] = 0.17

summary.loc[
    summary["Stallsystem"].str.contains(
        "Gülle",
        case=False,
        na=False,
    ),
    "MCF",
] = 0.17

summary.loc[
    summary["Stallsystem"].str.contains(
        "Schwimmschicht",
        case=False,
        na=False,
    ),
    "MCF",
] = 0.10


# ============================================================================
# BIOGAS
# ============================================================================

# Due to missing plant-specific information on methane slip
#  and methane use, a simplified effective methane conversion factor
# was used for biogas plants.
# For plants in medium capacity classes, as is standard on Hessian
# farms (often 75 kW), a medium default value of 0.05
# was set, based on typical agricultural on-farm biogas plants
#  in Germany

MCF_BIOGAS = 0.05

summary.loc[
    summary["Stallsystem"].str.contains(
        "Biogas",
        case=False,
        na=False,
    ),
    "MCF",
] = MCF_BIOGAS


print("\n💨 MCF Werte definiert")

# ============================================================================
# CALCULATE CH4
# ============================================================================

rho_CH4 = 0.67

# CH4 on pasture
summary["CH4_graz"] = summary["VS_graz"] * summary["B0"] * summary["MCF_graz"] * rho_CH4

# CH4 from housing management (manure management)
summary["CH4_barn"] = summary["VS_stall"] * summary["B0"] * summary["MCF"] * rho_CH4

# ============================================================================
# ROUND CH4 TO 2 DECIMAL PLACES
# ============================================================================

summary["CH4_graz"] = summary["CH4_graz"].round(2)
summary["CH4_barn"] = summary["CH4_barn"].round(2)

print("\n💨 CH4-Emissionen berechnet")

# Debug table
ch4_debug = summary[
    ["BetriebProcessID", "Process", "Animal_type", "CH4_graz", "CH4_barn"]
].copy()

print("\n🔍 CH4 DEBUG TABELLE:\n")

print(ch4_debug.to_string(index=False))

# ============================================================================
# TARGET PROCESSES PER LIVESTOCK PROCESS: P_EmissionMM_xy
# ============================================================================
# Instead of putting all CH4 emissions (pasture + barn) into one single
# manure management process per farm, every livestock instance
# (BetriebProcessID) gets its own P_EmissionMM_xy process. Same
# approach as in 4_Emission_Duengemanagement_PM_NMVOC_Processlevel.py and
# 4_Emissions_Duengemanagement_N2O_NOx_NH3_Processlevel.py, so that all three
# scripts write to the same target process instances. The animal tag at the end
# of the source process name is the key, e.g. P_Mastschwein ->
# P_EmissionMM_Mastschwein.
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
# Same pairing as in the other two manure management process-level
# scripts: existing P_EmissionMM_xy instances are assigned deterministically
# (sorted by BetriebProcessID) to the source instances before new ones
# are created. This way all three scripts end up on the same target
# BetriebProcessIDs, regardless of their run order.

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

# Map every livestock instance (BetriebProcessID) directly to its own
# P_EmissionMM_xy target process (instead of one collective process per farm).
summary["DuengeProcessID"] = summary["BetriebProcessID"].map(source_to_target)

print("\n💩 P_EmissionMM_xy Zielprozess je Tierhaltungsinstanz an summary angehängt")

# ============================================================================
# VALIDATION: sum per farm unchanged compared with the old farm-level script
# ============================================================================
# The old script (4_Emissions_Duengemanagement_Weide_CH4.py) wrote all
# CH4 emissions (pasture + barn) per farm into one single
# P_1_Jahr_Duengemanagement process. Here the same values are only
# distributed over several P_EmissionMM_xy instances per livestock process —
# so the sum per farm must stay exactly the same. Differences point to
# unassigned (unmapped) livestock instances.
summary["CH4_total"] = summary["CH4_barn"].fillna(0) + summary["CH4_graz"].fillna(0)

old_style_betrieb_sum = (
    summary.groupby("BetriebID")["CH4_total"].sum().rename("CH4_alt (Betriebsebene)")
)
new_style_betrieb_sum = (
    summary.dropna(subset=["DuengeProcessID"])
    .groupby("BetriebID")["CH4_total"]
    .sum()
    .rename("CH4_neu (P_EmissionMM_xy je Instanz)")
)

betrieb_sum_comparison = pd.concat(
    [old_style_betrieb_sum, new_style_betrieb_sum], axis=1
).fillna(0)
betrieb_sum_comparison["Differenz"] = (
    betrieb_sum_comparison["CH4_alt (Betriebsebene)"]
    - betrieb_sum_comparison["CH4_neu (P_EmissionMM_xy je Instanz)"]
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
        "(vermutlich Instanzen ohne einen der 9 definierten Tierhaltungsprozesse):"
    )
    print(abweichungen.to_string())
print(betrieb_sum_comparison.round(2).to_string())

# ============================================================================
# DEBUG
# ============================================================================

missing_duenge = summary["DuengeProcessID"].isna().sum()

print(f"\n⚠️ Instanzen ohne P_EmissionMM_xy Zielprozess: {missing_duenge}")

if missing_duenge > 0:
    print(
        summary[summary["DuengeProcessID"].isna()][
            [
                "BetriebID",
                "Process",
            ]
        ]
        .drop_duplicates()
        .to_string(index=False)
    )

# Goal: assign all CH4 emissions (pasture + barn) to the manure management process
# without splitting further pasture types. CH4_graz and CH4_barn are used
# and both are written to the process P_1_Jahr_Duengemanagement.

# ============================================================================
# FIND EXISTING CH4 FLOW
# ============================================================================

ch4_flows = data["Katalog_Flow"][
    data["Katalog_Flow"]["Flow"].str.contains(
        "CH4|Methan",
        case=False,
        na=False,
    )
]

print("\n🔍 Verfügbare CH4-Flows:\n")

print(ch4_flows[["FlowID", "Flow", "Category"]].to_string(index=False))

# ============================================================================
# SELECT THE CORRECT FLOW HERE
# ============================================================================

# TODO:
# select the correct CH4 flow

FLOW_CH4 = ch4_flows.iloc[0]["FlowID"]

print(f"\n💨 Verwendeter CH4 FlowID: {FLOW_CH4}")


# ============================================================================
# ============================================================================
# BUILD INSERT DATAFRAME
# ============================================================================

# CH4 from housing management (manure management) -> P_1_Jahr_Duengemanagement
flows_mm = pd.DataFrame(
    {
        "BetriebProcessID": summary["DuengeProcessID"],
        "FlowID": FLOW_CH4,
        "Wert": summary["CH4_barn"],
        "Direction": "Output",
        "TierkategorieEmissionFutter": "Stall",
    }
)

# CH4 from grazing -> P_1_Jahr_Duengemanagement
flows_graz = pd.DataFrame(
    {
        "BetriebProcessID": summary["DuengeProcessID"],
        "FlowID": FLOW_CH4,
        "Wert": summary["CH4_graz"],
        "Direction": "Output",
        "TierkategorieEmissionFutter": "Weide",
    }
)

# ============================================================================
# MERGE ALL FLOWS
# ============================================================================

insert_emissions_df = pd.concat(
    [flows_mm, flows_graz],
    ignore_index=True,
)

# Additional diagnostics: show how many rows have NaN
print("\n🔍 PRE-DROPNA DIAGNOSE:")
print(f"   Gesamtzeilen in Flows: {len(insert_emissions_df)}")
print(
    f"   Zeilen mit NaN BetriebProcessID: {insert_emissions_df['BetriebProcessID'].isna().sum()}"
)
print(f"   Zeilen mit NaN Wert: {insert_emissions_df['Wert'].isna().sum()}")

# CH4 in NaN rows: calculate sum of CH4 in rows lacking BetriebProcessID
ch4_in_nan = insert_emissions_df.loc[
    insert_emissions_df["BetriebProcessID"].isna(), "Wert"
].sum(min_count=1)
print(f"   CH4 in NaN Zeilen: {ch4_in_nan:.2f}")

print("\n🔍 FEHLENDE ZIELPROZESSE (per Betrieb):\n")

missing_summary = summary[summary["DuengeProcessID"].isna()].copy()
if not missing_summary.empty:
    missing_summary["missing_CH4_kg"] = missing_summary[["CH4_barn", "CH4_graz"]].sum(
        axis=1
    )
    per_betrieb = (
        missing_summary.groupby("BetriebID")["missing_CH4_kg"]
        .sum()
        .reset_index()
        .sort_values("missing_CH4_kg", ascending=False)
    )

    print(per_betrieb.to_string(index=False))

    total_missing = per_betrieb["missing_CH4_kg"].sum()
    print(f"\n🔎 Total missing CH4 (kg): {total_missing:.2f}")
else:
    print("✅ Keine Betriebe mit fehlendem Düngemanagementprozess gefunden.")


# ============================================================================
# REMOVE EMPTY ROWS
# ============================================================================

insert_emissions_df = insert_emissions_df.dropna(subset=["BetriebProcessID", "Wert"])


# ============================================================================
# AGGREGATE IDENTICAL TARGET PROCESSES
# ============================================================================

insert_emissions_df = insert_emissions_df.groupby(
    [
        "BetriebProcessID",
        "FlowID",
        "Direction",
        "TierkategorieEmissionFutter",
    ],
    as_index=False,
)["Wert"].sum()

print("\n💨 CH4-Flows aggregiert")


# ============================================================================
# DEBUG OF THE FINAL FLOWS
# ============================================================================

print("\n🔍 FINALE CH4 FLOWS:\n")

print(insert_emissions_df.head(50).to_string(index=False))


# ============================================================================
# DETAILED BALANCE CHECK
# ============================================================================

mm_assigned = flows_mm[flows_mm["BetriebProcessID"].notna()]["Wert"].sum()

graz_assigned = flows_graz[flows_graz["BetriebProcessID"].notna()]["Wert"].sum()

mm_missing = summary["CH4_barn"].sum() - mm_assigned

graz_missing = summary["CH4_graz"].sum() - graz_assigned

print("\n🔍 Detaillierte Zuordnung:")
print(
    f"   CH4_barn:   {summary['CH4_barn'].sum():.2f} kg "
    f"(zugeordnet: {mm_assigned:.2f}, fehlend: {mm_missing:.2f})"
)
print(
    f"   CH4_graz:   {summary['CH4_graz'].sum():.2f} kg "
    f"(zugeordnet: {graz_assigned:.2f}, fehlend: {graz_missing:.2f})"
)

print("\n🔍 Wahrscheinliche Ursachen:")
if mm_missing > 0:
    print(f"   - {mm_missing:.2f} kg CH4_barn: P_EmissionMM_xy Zielprozess fehlend")
if graz_missing > 0:
    print(f"   - {graz_missing:.2f} kg CH4_graz: P_EmissionMM_xy Zielprozess fehlend")
if mm_missing <= 0 and graz_missing <= 0:
    print(
        "   - Alle CH4-Emissionen sind pro Tierhaltungsinstanz auf P_EmissionMM_xy Zielprozesse zugeordnet."
    )

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
