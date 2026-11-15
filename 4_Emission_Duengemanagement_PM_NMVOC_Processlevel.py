import pyodbc
import pandas as pd
import numpy as np
from config import db_path

print("Hello World! Let's calculate some emissions from manure management! 🐄💨")
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

###Load the tables and queries needed for the calculations

# 1) Load the prepared Access query
duenger_df = pd.read_sql_query("SELECT * FROM [☀_N_P_Duengemittel_qry]", conn)
process_flows_df = pd.read_sql_query(
    "SELECT * FROM [★_python_qry_Processes_Flows]", conn
)
flows_def = pd.read_sql_query("SELECT * FROM [Katalog_Flow]", conn)

print(f"✓ Duenger: {len(duenger_df)} Zeilen")
print(f"✓ Prozess-Flows: {len(process_flows_df)} Zeilen")
print(f"✓ Flow-Katalog: {len(flows_def)} Zeilen")
# ============================================================================
# Emission factors
# ============================================================================

# Emission factors for the process category manure management
# according to the EMEP/EEA air pollutant emission inventory guidebook 2023
EF_NMVOC_Duem = {
    "dairy_cattle": {
        "with_silo": 17.937,  # kg NMVOC / n animals
        "without_silo": 8.047,  # kg NMVOC / n animals
    },
    "non_dairy_cattle": {
        "with_silo": 8.902,  # kg NMVOC / n animals
        "without_silo": 3.602,  # kg NMVOC / n animals
    },
    "swine": {"without_silo": 0.551},  # e.g. no silage case present
    "sheep": {
        "with_silo": 0.279,  # kg NMVOC / n animals
        "without_silo": 0.169,  # kg NMVOC / n animals
    },
    "horse": {
        "with_silo": 7.781,  # kg NMVOC / n animals
        "without_silo": 4.275,  # kg NMVOC / n animals
    },
}
EF_PM10_Duem = {
    "dairy_cattle": 0.63,  # kg PM10 / n animals
    "non_dairy_cattle": 0.27,  # kg PM10 / n animals
    "swine": 0.14,  # kg PM10 / n animals
    "sheep": 0.06,  # kg PM10 / n animals
    "horse": 0.06,  # kg PM10 / n animals
}
EF_PM2_5_Duem = {
    "dairy_cattle": 0.41,  # kg PM2.5 / n animals
    "non_dairy_cattle": 0.18,  # kg PM2.5 / n animals
    "swine": 0.006,  # kg PM2.5 / n animals
    "sheep": 0.02,  # kg PM2.5 / n animals
    "horse": 0.02,  # kg PM2.5 / n animals
}


# ============================================================================
# CALCULATIONS
# ============================================================================


# Calculation of the NMVOC emissions from manure management per process
# First identify the animal category per process, then the feeding regime, then assign the matching emission factor
def map_pp_to_animal(pp_name, process=None):
    pp_name = str(pp_name)

    if "Milchkuh" in pp_name:
        # FIX 2026-10-06 (suckler cow): the parameter "Milchkuh / Mutterkuh" is shared by
        # P_Milchkuh and P_Mutterkuh. Suckler cows are not
        # milked and count as "non-dairy cattle" according to EMEP/EEA (2023).
        # Old: return "dairy_cattle"  (also for suckler cows)
        if process == "P_Mutterkuh":
            return "non_dairy_cattle"
        return "dairy_cattle"
    elif "Mastschwein" in pp_name:
        return "swine"
    elif "Ferkel" in pp_name:
        return "swine"
    elif "Mutterschaf" in pp_name or "Mastschaf" in pp_name:
        return "sheep"
    elif "Pferd" in pp_name:
        return "horse"
    elif "Aufzucht" in pp_name or "Mast" in pp_name or "Deckbulle" in pp_name:
        return "non_dairy_cattle"
    else:
        return None


# Map the animal category per process in tier_df
# Assemble tier_df from the available tables
tier_df = (
    data["tbl_6_Betrieb_ProcessParameterValue"]
    .merge(data["tbl_3_Betrieb_Process"], on="BetriebProcessID", how="left")
    .merge(
        data["Katalog_ProcessProcessparameter"][["ProcessPpID", "PpID"]],
        on="ProcessPpID",
        how="left",
    )
    .merge(data["Katalog_Processparameter"][["PpID", "Pp_name"]], on="PpID", how="left")
    .merge(
        data["Katalog_Process"][["ProcessID", "Process"]], on="ProcessID", how="left"
    )
)
print(tier_df[["Process", "Pp_name", "Wert"]].head(20))

# Convert Wert to numeric values (comma as decimal separator)
tier_df["Wert"] = (
    tier_df["Wert"].astype(str).str.replace(",", ".", regex=False).str.strip()
)
tier_df["Wert"] = pd.to_numeric(tier_df["Wert"], errors="coerce")
# FIX 2026-10-06 (suckler cow): pass the process as well, so that suckler cows are
# recognised as non_dairy_cattle.
# Old: tier_df["animal_type"] = tier_df["Pp_name"].apply(map_pp_to_animal)
tier_df["animal_type"] = tier_df.apply(
    lambda r: map_pp_to_animal(r["Pp_name"], r["Process"]), axis=1
)

# Debug: Check tier_df content
print("\n=== TIER_DF DEBUG INFORMATION ===")
print(f"Total rows in tier_df: {len(tier_df)}")

# Filter for ProcessID 14
tier_df_14 = tier_df[tier_df["ProcessID"] == 14]
print(f"Rows in tier_df for ProcessID 14: {len(tier_df_14)}")

if len(tier_df_14) > 0:
    print(f"\nUnique process names in ProcessID 14: {tier_df_14['Pp_name'].unique()}")
    print(
        f"Detected animal types: {tier_df_14['animal_type'].value_counts().to_dict()}"
    )
    print(f"Non-null Wert values: {tier_df_14['Wert'].notna().sum()}")
    print(f"Non-zero Wert values: {(tier_df_14['Wert'] > 0).sum()}")
    print(f"Sample Wert values: {tier_df_14['Wert'].describe()}")
else:
    print("No rows found for ProcessID 14!")

print("=" * 60 + "\n")

flows_df = data["tbl_4_Betrieb_Process_FlowValue"].merge(
    data["Katalog_Flow"][["FlowID", "Flow"]], on="FlowID", how="left"
)
flows_per_process = flows_df.groupby("BetriebProcessID")["Flow"].apply(list).to_dict()


# Function that calculates the NMVOC, PM10 and PM2.5 emissions per row in tier_df
def calc_emissions(row):
    process_id = row["ProcessID"]
    betrieb_id = row["BetriebProcessID"]
    animal = row["animal_type"]
    n = row["Wert"]

    if not animal or pd.isna(n):
        return pd.Series([0, 0, 0])

    flows = flows_per_process.get(betrieb_id, [])
    has_silo = any("silo_" in f.lower() for f in flows)

    # NMVOC
    ef_nmvoc = EF_NMVOC_Duem.get(animal, {}).get(
        "with_silo" if has_silo else "without_silo",
        EF_NMVOC_Duem.get(animal, {}).get("without_silo", 0),
    )

    # PM
    ef_pm10 = EF_PM10_Duem.get(animal, 0)
    ef_pm25 = EF_PM2_5_Duem.get(animal, 0)

    return pd.Series([n * ef_nmvoc, n * ef_pm10, n * ef_pm25])


# Calculate NMVOC, PM10 and PM2.5 emissions with the calc_emissions function
tier_df[["NMVOC_kg", "PM10_kg", "PM2_5_kg"]] = tier_df.apply(calc_emissions, axis=1)

# Calculate PM10 minus PM2.5, i.e. coarse PM (PM2.5-10)
tier_df["PM2_5_10_kg"] = tier_df["PM10_kg"] - tier_df["PM2_5_kg"]
tier_df["PM2_5_10_kg"] = tier_df["PM2_5_10_kg"].clip(lower=0)


# ============================================================================
# TARGET PROCESSES PER LIVESTOCK PROCESS: P_EmissionMM_xy
# ============================================================================
# Instead of summing all emissions at farm level into one single manure management
# process, every livestock instance (BetriebProcessID) gets
# its own P_EmissionMM_xy process. The animal tag at the end of the source process
# name is the key, e.g. P_Mastschwein -> P_EmissionMM_Mastschwein.
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
# AGGREGATE EMISSIONS PER LIVESTOCK BETRIEBPROCESSID (instead of per farm!)
# ============================================================================

relevant_tier_df = tier_df[tier_df["Process"].isin(PROCESS_TO_EMISSIONMM.keys())].copy()

Duem_emissions = (
    relevant_tier_df.groupby(["BetriebID", "BetriebProcessID", "Process"])[
        ["NMVOC_kg", "PM10_kg", "PM2_5_kg", "PM2_5_10_kg"]
    ]
    .sum()
    .reset_index()
)

print("\n=== AGGREGATED EMISSIONS PER TIERHALTUNGS-BETRIEBPROCESSID ===")
print(Duem_emissions.head())

# ============================================================================
# DETERMINE TARGET BETRIEBPROCESS PER LIVESTOCK INSTANCE (1:1) AND CREATE NEW ONES
# ============================================================================
# Every livestock process instance gets its own P_EmissionMM_xy instance.
# If instances already exist for a farm + target process (e.g. from an
# earlier run), they are assigned deterministically (sorted by BetriebProcessID)
# to the source instances before new ones are created. This keeps the
# run idempotent, even if a farm has several instances of the same
# livestock process.

betrieb_process_df = data["tbl_3_Betrieb_Process"].copy()

# All existing livestock instances of the relevant processes, regardless
# of calculated NMVOC/PM emissions: the P_EmissionMM_xy instance must also
# exist when no emission is calculated here (yet), because other
# emission types are written to the same process in later scripts.
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
Duem_emissions["BetriebProcessID"] = Duem_emissions["BetriebProcessID"].map(
    source_to_target
)

unmapped = Duem_emissions[Duem_emissions["BetriebProcessID"].isna()]
if len(unmapped) > 0:
    print(
        "⚠️ WARNING: Für einige Tierhaltungsinstanzen konnte kein Zielprozess ermittelt werden!"
    )
    print(unmapped)

Duem_emissions = Duem_emissions.dropna(subset=["BetriebProcessID"])
Duem_emissions["BetriebProcessID"] = Duem_emissions["BetriebProcessID"].astype(int)

Duem_emissions = (
    Duem_emissions.groupby("BetriebProcessID")[
        ["NMVOC_kg", "PM10_kg", "PM2_5_kg", "PM2_5_10_kg"]
    ]
    .sum()
    .reset_index()
)

print("\n=== FINAL EMISSIONS ASSIGNED TO P_EmissionMM_xy PROZESSEN ===")
print(Duem_emissions.head())
print(Duem_emissions.sample(min(10, len(Duem_emissions))))

emissions_df = Duem_emissions

print(emissions_df.head())
print(emissions_df.tail(20))
if len(emissions_df) > 0:
    print(emissions_df.sample(min(10, len(emissions_df))))
else:
    print("No emissions data to sample from.")


# ============================================================================
# CONVERT TO LONG FORMAT
# ============================================================================

emissions_long = pd.melt(
    emissions_df,
    id_vars=["BetriebProcessID"],
    value_vars=[
        "NMVOC_kg",
        "PM2_5_10_kg",
        "PM2_5_kg",
    ],
    var_name="EmissionType",
    value_name="Wert",
)

# Keep only emissions > 0
emissions_long = emissions_long[emissions_long["Wert"] > 0]

print(emissions_long.head())
print(emissions_long.tail(20))
if len(emissions_long) > 0:
    print(emissions_long.sample(min(10, len(emissions_long))))
else:
    print("No positive emissions to sample from.")

# ============================================================================
# Map calculated emissions to the emission flows (red DB and openLCA)
# ============================================================================

# Mapping between the calculated emissions and the flow names in Katalog_Flow
flow_mapping = {
    "NMVOC_kg": {
        "Flow": "NMVOC, non-methane volatile organic compounds",
        "Category": "Elementary flows/Emission to air/low population density, long-term",
    },
    "PM2_5_10_kg": {
        "Flow": "Particulate Matter, > 2.5 um and < 10um",
        "Category": "Elementary flows/Emission to air/low population density",
    },
    "PM2_5_kg": {
        "Flow": "Particulate Matter, < 2.5 um",
        "Category": "Elementary flows/Emission to air/low population density",
    },
}

# Add the emission names
emissions_long["FlowName"] = emissions_long["EmissionType"].map(
    lambda x: flow_mapping[x]["Flow"]
)
emissions_long["Category"] = emissions_long["EmissionType"].map(
    lambda x: flow_mapping[x]["Category"]
)

# Get the matching FlowIDs from Katalog_Flow
katalog_flow_df = data["Katalog_Flow"][["FlowID", "Flow", "Category"]]

# Add both names in lower case to Katalog_Flow so the merge works
katalog_flow_df["Flow_lower"] = katalog_flow_df["Flow"].str.lower()
katalog_flow_df["Category_lower"] = katalog_flow_df["Category"].str.lower()

emissions_long["Flow_lower"] = emissions_long["FlowName"].str.lower()
emissions_long["Category_lower"] = emissions_long["Category"].str.lower()

# Merge to get the FlowIDs
emissions_long = emissions_long.merge(
    katalog_flow_df[["FlowID", "Flow_lower", "Category_lower"]],
    on=["Flow_lower", "Category_lower"],
    how="left",
)

# Check whether all flows were found
missing = emissions_long[emissions_long["FlowID"].isna()]

mapped = emissions_long[emissions_long["FlowID"].notna()]

print("Erfolgreich gemappte Emissionen:")
print(
    mapped[["EmissionType", "FlowName", "Category", "FlowID"]]
    .drop_duplicates()
    .sort_values("EmissionType")
)

if len(missing) > 0:
    print("\nFolgende Emissionen konnten nicht gemappt werden:")
    print(missing[["EmissionType", "FlowName", "Category"]].drop_duplicates())
else:
    print("Alle Emissionen wurden erfolgreich gemappt!")

# ============================================================================
# CREATE FORMAT FOR tbl_4_Betrieb_Process_FlowValue
# ============================================================================

insert_emissions_df = emissions_long[["BetriebProcessID", "FlowID", "Wert"]].copy()

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
    SELECT BetriebProcessID, FlowID, Direction, BetriebProcessFlowValueID
    FROM tbl_4_Betrieb_Process_FlowValue
    WHERE Direction = 'Output'
    """,
    conn,
)

# Link with the newly calculated emissions
merge_df = insert_emissions_df.merge(
    existing_df,
    on=["BetriebProcessID", "FlowID", "Direction"],
    how="left",
)

# Existing rows -> UPDATE
update_rows = merge_df[merge_df["BetriebProcessFlowValueID"].notna()]

# Rows that do not exist yet -> INSERT
insert_rows = merge_df[merge_df["BetriebProcessFlowValueID"].isna()]

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
    print(
        new_emissions_display[
            ["BetriebProcessID", "Process", "Flow", "Category", "Wert"]
        ].to_string(index=False)
    )
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
            insert_rows[["BetriebProcessID", "FlowID", "Wert", "Direction"]].itertuples(
                index=False, name=None
            )
        )

        cursor.executemany(
            """
            INSERT INTO tbl_4_Betrieb_Process_FlowValue
            (BetriebProcessID, FlowID, Wert, Direction)
            VALUES (?, ?, ?, ?)
            """,
            rows_to_insert,
        )
    else:
        print("Keine neuen Emissionen zum Einfügen.")

    conn.commit()
    print("\n✅ Änderungen wurden in die Access-Datenbank geschrieben.")
else:
    print("\n❌ Änderungen wurden verworfen.")
