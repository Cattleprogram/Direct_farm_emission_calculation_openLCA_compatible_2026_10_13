print(
    "Hello, World!, Lets update P_Milchkuh to the feed ratio of the calves and assume that calves consume the same than the lactating cows but receive only the amount * 0,24 + the calculated milk"
)
# Update P_Milchkuh with the feed ration of the calves and assume that calves consume the same
# as the lactating cows but receive only the amount * 0.24 (roughly according to the feed intake in the suckler cow system,
# where 14.5 * 0.3 gives 4.35 kg DM. In the dairy rations
# the daily feed intake varies by farm and can be between 18 and 24 kg DM)
# + the calculated milk
# The original feed flows of dairy farming were interpreted as a per-cow ration and, to account for suckling calves,
# supplemented proportionally with the calf factor explained above. New flows are created
# that contain the term "Kälberfutter" (calf feed) in the notes.
# This applies only to the feed flows of the process P_Milchkuh. For P_Mutterkuh the calves were already included in the feed calculation
#  Milk intake:
# if Eimer_Vollmilch == 45.5 l * drinking period in weeks * total calvings
# if Automat_MAT == 56 L * total calvings (the rest of the milk comes from milk replacer (MAT), which is already in the DB)
# if muttergebunden (dam-rearing) == 1000 l * total calvings
# if ammengebunden (nurse cow) == 1000 l * total calvings
# For suckler cow farms == total calvings * weaning age * 7 days * 7.29 L/day

from pathlib import Path

import pyodbc
import pandas as pd
import numpy as np
from config import db_path

print(
    "Hello World! Let's map the FlowID with the Flowproperties and get all the necessary ingredients for the used feedstuffs!"
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
# IDENTIFY DAIRY COW + SUCKLER COW
# ============================================================================
process_df = data["Katalog_Process"]
betrieb_process_df = data["tbl_3_Betrieb_Process"]

tier_process_ids = process_df[
    process_df["Process"].isin(
        [
            "P_Milchkuh",
            "P_Mutterkuh",
        ]
    )
][["ProcessID", "Process"]]


tier_betriebe = betrieb_process_df.merge(tier_process_ids, on="ProcessID", how="inner")

print(tier_betriebe[["BetriebProcessID", "Process"]].drop_duplicates())
# ============================================================================
# LOAD FARM PROCESSES P_MILCHKUH
# ============================================================================

milchkuh_betriebe = betrieb_process_df[
    betrieb_process_df["ProcessID"].isin(tier_process_ids["ProcessID"])
].copy()

print(f"{len(milchkuh_betriebe)} Milchkuh-Prozesse gefunden")


# ============================================================================
# LOAD FLOW VALUES
# ============================================================================

flowvalue_df = data["tbl_4_Betrieb_Process_FlowValue"]

milchkuh_flows = flowvalue_df.merge(
    tier_betriebe,
    on="BetriebProcessID",
    how="inner",
)

print(f"{len(milchkuh_flows)} Flowwerte gefunden")


# ============================================================================
# FEED FLOWS ONLY
# ============================================================================

flow_df = data["Katalog_Flow"]

feed_flows = flow_df[flow_df["IsPotentialFeedstuff"] == True][
    [
        "FlowID",
        "Flow",
    ]
]

milchkuh_feedflows = milchkuh_flows.merge(
    feed_flows,
    on="FlowID",
    how="inner",
)

print(f"{len(milchkuh_feedflows)} Futtermittelflows gefunden")

# ============================================================================
# FIND FLOWS ALREADY FULLY ASSIGNED TO CALVES (e.g. MAT Milchaustauscher)
# ============================================================================
# Katalog_Flow has no TierkategorieEmissionFutter column of its own (it only exists per
# row in tbl_4). Some flows (e.g. "MAT Milchaustauscher") are, according to their
# description, pure calf feed by nature and must therefore NOT additionally
# be duplicated via the calf feed factor - the whole flow already belongs
# to the calves. This is detected via the text "Kälberfutter" in Katalog_Flow.Description.
calf_only_flow_ids = flow_df.loc[
    flow_df["Description"]
    .astype(str)
    .str.contains("Kälberfutter", case=False, na=False),
    "FlowID",
].tolist()

print(
    f"\n{len(calf_only_flow_ids)} Flow(s) laut Description bereits vollständig Kälberfutter "
    f"(werden von der Kälberfutter-Duplizierung ausgenommen): "
    f"{flow_df[flow_df['FlowID'].isin(calf_only_flow_ids)]['Flow'].tolist()}"
)

# Warning in case such flows in tbl_4 have not yet been tagged manually with
# TierkategorieEmissionFutter = 'Kälber' (as required by the description)
# - otherwise 5_Emissions_Tierhaltung_CH4_Cattle.py later assigns them
# wrongly to the dairy cow/suckler cow group instead of the calf group.
if "TierkategorieEmissionFutter" in milchkuh_feedflows.columns and calf_only_flow_ids:
    untagged_calf_only_df = milchkuh_feedflows[
        milchkuh_feedflows["FlowID"].isin(calf_only_flow_ids)
        & (milchkuh_feedflows["TierkategorieEmissionFutter"] != "Kälber")
    ]
    if len(untagged_calf_only_df) > 0:
        print(
            "⚠ Warnung: Folgende BetriebProcessID/FlowID-Zeilen sind laut Description "
            "reines Kälberfutter, aber in tbl_4 noch nicht mit "
            "TierkategorieEmissionFutter = 'Kälber' getaggt. Bitte manuell in Access "
            "korrigieren:"
        )
        print(
            untagged_calf_only_df[["BetriebProcessID", "FlowID", "Flow"]].to_string(
                index=False
            )
        )


# ============================================================================
# LOAD PARAMETER VALUES
# ============================================================================

param_values_df = data["tbl_6_Betrieb_ProcessParameterValue"]

process_processparam_df = data["Katalog_ProcessProcessparameter"]

process_param_df = data["Katalog_Processparameter"]


# ============================================================================
# 1. ProcessPpID → PpID
# ============================================================================

param_values_df = param_values_df.merge(
    process_processparam_df[
        [
            "ProcessPpID",
            "PpID",
        ]
    ],
    on="ProcessPpID",
    how="left",
)


# ============================================================================
# 2. PpID → Pp_name
# ============================================================================

param_values_df = param_values_df.merge(
    process_param_df[
        [
            "PpID",
            "Pp_name",
        ]
    ],
    on="PpID",
    how="left",
)


# ============================================================================
# Check
# ============================================================================

print("\nVerfügbare Parameter:")

print(param_values_df["Pp_name"].drop_duplicates().sort_values().tolist())
# Robust decimal separator

# ============================================================================
# ROBUST NUMBER CONVERSION (Access comma/point)
# ============================================================================


def convert_numeric(series):
    return pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )


# ============================================================================
# LOAD ANIMAL NUMBERS PER FARM
# ============================================================================

tierparameter = [
    "Milchkuh / Mutterkuh",
    "Deckbulle",
    "Kälber",
    "Kalbungen gesamt",
    "Aufzuchtverfahren_Kälber",
    "Absetzalter",
]

tier_df = param_values_df[param_values_df["Pp_name"].isin(tierparameter)].copy()

tier_df = tier_df.pivot_table(
    index="BetriebProcessID",
    columns="Pp_name",
    values="Wert",
    aggfunc="first",
).reset_index()

tier_df = tier_df.fillna(0)

print("\nTierzahlen:")
print(tier_df.head())


# ============================================================================
# LINK FEED FLOWS WITH ANIMAL NUMBERS
# ============================================================================

milchkuh_feedflows = milchkuh_feedflows.merge(
    tier_df,
    on="BetriebProcessID",
    how="left",
)

# Make sure the animal numbers are numeric, otherwise "+" concatenates strings.
for col in [
    "Milchkuh / Mutterkuh",
    "Deckbulle",
    "Kälber",
    "Kalbungen gesamt",
    "Absetzalter",
]:
    milchkuh_feedflows[col] = convert_numeric(milchkuh_feedflows[col]).fillna(0)

# Make sure the flow value is numeric, otherwise the division fails.
milchkuh_feedflows["Wert"] = convert_numeric(milchkuh_feedflows["Wert"]).fillna(0)


# ============================================================================
# CALCULATE BASE ANIMAL NUMBER
# ============================================================================

milchkuh_feedflows["BasisTiere"] = (
    milchkuh_feedflows["Milchkuh / Mutterkuh"] + milchkuh_feedflows["Deckbulle"]
)

# Avoid division by 0
milchkuh_feedflows = milchkuh_feedflows[milchkuh_feedflows["BasisTiere"] > 0]


# ============================================================================
# FEED PER ANIMAL
# ============================================================================

milchkuh_feedflows["FutterProTier"] = (
    milchkuh_feedflows["Wert"] / milchkuh_feedflows["BasisTiere"]
)

# ============================================================================
# CALCULATE CALF FEED
# ============================================================================

KALB_FAKTOR = 0.24  # based on the livestock unit (GVE) value of 0.3 but reduced somewhat,
# because otherwise the feed intake would be overestimated. GVE is not a suitable measure to quantify feed intake.

milchkuh_feedflows["Kälberfutter"] = 0

# FIX 2026-10-06 (repeated run): use only the original cow ration as the base.
# From the second run on, the previously written calf feed rows
# (TierkategorieEmissionFutter == "Kälber") are also in milchkuh_feedflows. Without
# this exclusion the calf factor was applied to them again ("calf ration
# of the calf ration", e.g. farm 2 Silo_DG: 4481 -> 284 kg), and drop_duplicates
# further down kept this wrong value depending on the row order, which would then
# have overwritten the correct calf ration in Access via UPDATE.
# Old:
#     mask_milchkuh = (milchkuh_feedflows["Process"] == "P_Milchkuh") & (
#         ~milchkuh_feedflows["FlowID"].isin(calf_only_flow_ids)
#     )
mask_milchkuh = (
    (milchkuh_feedflows["Process"] == "P_Milchkuh")
    & (~milchkuh_feedflows["FlowID"].isin(calf_only_flow_ids))
    & (milchkuh_feedflows["TierkategorieEmissionFutter"] != "Kälber")
)

milchkuh_feedflows.loc[mask_milchkuh, "Kälberfutter"] = (
    milchkuh_feedflows.loc[mask_milchkuh, "FutterProTier"]
    * KALB_FAKTOR
    * milchkuh_feedflows.loc[mask_milchkuh, "Kälber"]
).round(0)
# ===========================================================================
# Calculate calf milk intake
# if Eimer_Vollmilch and P_Milchkuh == 45.5 l/week * milk-fed calves * 52.14 weeks # because restrictive feeding assumes 6.5 l milk per day
# if Automat_MAT and P_Milchkuh == 8 l/week * milk-fed calves * 52.14 weeks (the rest of the milk comes from milk replacer (MAT), which is already in the DB)
#  This is a strong simplification, as it assumes a constant amount of milk, while in reality the MAT mix contains more milk at the beginning and less at the end
# if muttergebunden (dam-rearing) and P_Milchkuh == 51 l/week * milk-fed calves * 52.14 weeks (in ad lib systems 7.29 L per day is assumed)
# if ammengebunden (nurse cow) and P_Milchkuh == 51 l/week * milk-fed calves * 52.14 weeks (in ad lib systems 7.29 L per day is assumed)
# if muttergebunden (dam-rearing) and P_Mutterkuh == calvings * weaning age * 7 days/week * 7.29 (in ad lib systems 7.29 L per day is assumed)
# Suckler cow (MK) farms do not need the HIT logic, because all calves are fed milk on the farm and are sold as weaners at the earliest.
# Milk-fed calves = "Kälber" HIT * (weaning age in weeks / HIT period of the category calf (26 weeks))
#
# # ============================================================================
# LOGIC OF THE CALF MILK CALCULATION: P_Milchkuh vs. P_Mutterkuh
# ============================================================================
# Both processes estimate the annual calf milk amount, but via different
# reference values - on purpose, because the way calves leave the farm differs between
# the systems:
#
# P_Milchkuh: tränkekälber (milk-fed calves) = Kälber (HIT) * (weaning age / 26 weeks)
#             Milk = tränkekälber * weekly rate * 52.14 weeks
#   Uses the HIT stock "Kälber" (age and sex statistics, already a
#   one-year average, not a reference-date value) instead of the total calvings, because
#   dairy calves are often sold off the farm within a few weeks.
#   The HIT stock therefore realistically reflects how many calves actually drink
#   milk on THIS farm - a calculation based on total calvings would
#   systematically overestimate in this system.
#
# P_Mutterkuh: Milk = total calvings * weaning age * 7 days * 7.29 L/day
#   Uses total calvings directly, because suckler calves usually stay with the
#   dam on the farm until weaning - no early sale that would require
#   a HIT correction.
#
# Known limitation: on suckler cow farms the weaning age can exceed 26 weeks (6
# months), so milk is also counted for animals that by definition
# would already be "Aufzucht" (rearing) instead of "Kälber" (calves). Since rearing animals get no
# milk in this model, capping at 26 weeks would lose this milk amount entirely
# from the calculation instead of reallocating it correctly. For the farm
# total this is harmless (misallocation between calves/rearing, no
# over- or under-counting); for an evaluation by animal category (e.g. "how much
# CH4 do calves cause") it leads to a bias in favour of the calf group.

# ===========================================================================
# ===========================================================================
# CALCULATE CALF MILK INTAKE
# ===========================================================================


def calculate_calf_milk(row):

    process = str(row["Process"]).strip()
    system = str(row["Aufzuchtverfahren_Kälber"]).lower().strip()

    kalbungen = row["Kalbungen gesamt"]

    # average number of milk-drinking calves
    # HIT calf stock (0-6 months) corrected to the milk feeding phase
    tränkekälber = row["Kälber"] * (row["Absetzalter"] / 26)

    # ------------------------------------------------------
    # Dairy
    # ------------------------------------------------------
    if process == "P_Milchkuh":

        if "eimer" in system:
            return 45.5 * tränkekälber * 52.14

        elif "automat" in system:
            return 8 * tränkekälber * 52.14

        elif "mutter" in system or "ammen" in system:
            return 51 * tränkekälber * 52.14

    # ------------------------------------------------------
    # Suckler cow husbandry
    # ------------------------------------------------------
    elif process == "P_Mutterkuh":

        return (
            kalbungen
            * row["Absetzalter"]
            * 7  # Days per week
            * 7.29  # L per day assumed for ad lib.
        )


# ============================================================================
# CREATE INTERNAL MILK FLOW FOR THE EMISSION CALCULATION
# ============================================================================

milchkuh_feedflows["Kälbermilch_L"] = milchkuh_feedflows.apply(
    calculate_calf_milk, axis=1
)

MILK_FLOW_ID = 523

calf_milk_intake_df = (
    milchkuh_feedflows[
        [
            "BetriebID",
            "BetriebProcessID",
            "Process",
            "Kälbermilch_L",
        ]
    ]
    .drop_duplicates(subset=["BetriebProcessID"])
    .copy()
)

calf_milk_intake_df["FlowID"] = MILK_FLOW_ID
calf_milk_intake_df["Direction"] = "Input"

calf_milk_intake_df = calf_milk_intake_df.rename(columns={"Kälbermilch_L": "Wert"})

calf_milk_intake_df["TierkategorieEmissionFutter"] = "Kälber"

print("\nBerechnete Kälberfutter-Flows:")
print(
    milchkuh_feedflows[
        [
            "BetriebProcessID",
            "Flow",
            "Wert",
            "FutterProTier",
            "Kälber",
            "Kälberfutter",
            "Absetzalter",
        ]
    ].head(20)
)

# ============================================================================
# DEBUG: CHECK CALF MILK
# ============================================================================

calf_milk_check = (
    milchkuh_feedflows[
        [
            "BetriebProcessID",
            "Process",
            "Aufzuchtverfahren_Kälber",
            "Kalbungen gesamt",
            "Absetzalter",
            "Kälbermilch_L",
        ]
    ]
    .drop_duplicates(subset=["BetriebProcessID"])
    .copy()
)

print("\n========== Check Kälbermilch je Betrieb ==========\n")

print(
    calf_milk_check.sort_values(["Process", "Kälbermilch_L"], ascending=[True, False])
    .round(1)
    .to_string(index=False)
)

print(
    "\nGesamte Kälbermilch:",
    round(calf_milk_check["Kälbermilch_L"].sum(), 0),
    "Liter/Jahr",
)

print("\n==================================================\n")
# ============================================================================
# CREATE NEW FLOW ROWS
# ============================================================================
new_flows_df = milchkuh_feedflows[(milchkuh_feedflows["Kälberfutter"] > 0)][
    [
        "BetriebProcessID",
        "FlowID",
        "Direction",
        "Kälberfutter",
    ]
].copy()

new_flows_df = new_flows_df.rename(
    columns={
        "Kälberfutter": "Wert",
    }
)

new_flows_df["Wert"] = pd.to_numeric(
    new_flows_df["Wert"],
    errors="coerce",
).fillna(0)

new_flows_df["TierkategorieEmissionFutter"] = "Kälber"

print(f"\n{len(new_flows_df)} neue Kälberfutter-Flows erzeugt")

# If new_Flows_df is loaded here without the milk_flows and there is no pd.concat,
# they exist only internally and are not sent to the Access DB. This makes sense,
# because the calves are considered part of the dairy cow and suckler cow processes, so the milk is an internal flow.
new_flows_df = pd.concat([new_flows_df], ignore_index=True)

# But calf_milk_intake_df must be exported!
# ============================================================================
# SAVE INTERNAL MILK FLOW FOR THE EMISSION CALCULATION
# ============================================================================

calf_milk_intake_df.to_pickle(Path(__file__).parent / "temp_calf_milk_intake.pkl")

print(f"{len(calf_milk_intake_df)} interne Kälbermilch-Flows gespeichert")

# ============================================================================
# LOAD EXISTING CALF FEED FLOWS
# ============================================================================

existing_df = pd.read_sql(
    """
    SELECT
        BetriebProcessFlowValueID,
        BetriebProcessID,
        FlowID,
        TierkategorieEmissionFutter
    FROM tbl_4_Betrieb_Process_FlowValue
    WHERE TierkategorieEmissionFutter = 'Kälber'
    """,
    conn,
)

# ============================================================================
# PREPARE UPSERT
# ============================================================================
print(
    "Doppelte neue Flows: ",
    new_flows_df.duplicated(
        subset=[
            "BetriebProcessID",
            "FlowID",
            "TierkategorieEmissionFutter",
        ]
    ).sum(),
)

new_flows_df = new_flows_df.drop_duplicates(
    subset=[
        "BetriebProcessID",
        "FlowID",
        "TierkategorieEmissionFutter",
    ]
)

merge_df = new_flows_df.merge(
    existing_df,
    on=[
        "BetriebProcessID",
        "FlowID",
        "TierkategorieEmissionFutter",
    ],
    how="left",
)

print(merge_df.columns.tolist())

update_rows = merge_df[merge_df["BetriebProcessFlowValueID"].notna()]

insert_rows = merge_df[merge_df["BetriebProcessFlowValueID"].isna()]

print(f"\n{len(update_rows)} bestehende Kälberflows werden aktualisiert")

print(f"{len(insert_rows)} neue Kälberflows werden eingefügt")


# ============================================================================
# CONFIRM WRITING
# ============================================================================

confirm = input("\nKälberfutter-Flows in Access schreiben? (j/n): ")

if confirm.lower() == "j":

    cursor = conn.cursor()

    # UPDATE
    update_sql = """
        UPDATE tbl_4_Betrieb_Process_FlowValue
        SET Wert = ?
        WHERE BetriebProcessFlowValueID = ?
    """
    update_data = [
        (
            0 if pd.isna(row["Wert"]) else float(row["Wert"]),
            int(row["BetriebProcessFlowValueID"]),
        )
        for _, row in update_rows.iterrows()
    ]

    if update_data:
        cursor.executemany(update_sql, update_data)

    # INSERT
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
            (
                BetriebProcessID,
                FlowID,
                Wert,
                Direction,
                TierkategorieEmissionFutter
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            rows_to_insert,
        )

    conn.commit()

    print("\n✅ Kälberfutter-Flows erfolgreich geschrieben")

else:

    print("\nÄnderungen verworfen")
