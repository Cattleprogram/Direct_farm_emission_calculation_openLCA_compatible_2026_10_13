import pyodbc
import pandas as pd
import numpy as np
from config import db_path

print(
    "Hello World! Let's calculate CH4 emissions from the animal husbandry module for pig, horse and sheep ! 🐄💨"
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
# CH4 EMISSIONS FROM LIVESTOCK (Tier 1)
# ============================================================================

# Tier 1 emission factors (kg CH4 per head and year): IPCC (2006), Vol. 4, Ch. 10, Table 10.10
EF_CH4 = {"swine": 1.5, "sheep": 8.0, "horse": 18.0}

# Define the relevant processes
TARGET_PROCESSES = ["P_Schafhaltung_Mast", "P_Mastschwein", "P_Pferdehaltung"]

# Prepare animal data
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

# Relevant processes only
tier_df = tier_df[tier_df["Process"].isin(TARGET_PROCESSES)].copy()


# Function that maps the process parameters to animal species
def map_pp_to_animal(pp_name):
    pp_name = str(pp_name)

    if "Milchkuh" in pp_name:
        return "dairy_cattle"
    elif "Mastschwein" in pp_name:
        return "swine"
    elif "Ferkel" in pp_name:
        return "swine"
    elif "Mutterschaf" in pp_name or "Mastschaf" in pp_name:
        return "sheep"
    elif "Pferd" in pp_name:
        return "horse"
    elif "Aufzucht" in pp_name or "♂ Mast" in pp_name or "Deckbulle" in pp_name:
        return "non_dairy_cattle"
    else:
        return None


# Clean numeric values
tier_df["Wert"] = (
    tier_df["Wert"].astype(str).str.replace(",", ".", regex=False).str.strip()
)
tier_df["Wert"] = pd.to_numeric(tier_df["Wert"], errors="coerce")

tier_df["animal_type"] = tier_df["Pp_name"].apply(map_pp_to_animal)

# Keep only the relevant animals
tier_df = tier_df[tier_df["animal_type"].notna()]

print("\n=== TIERDATEN CH4 ===")
print(tier_df[["Process", "Pp_name", "animal_type", "Wert"]].head())

# ============================================================================
# CALCULATE EMISSIONS (directly per process)
# ============================================================================

tier_df["CH4_kg"] = tier_df.apply(
    lambda row: row["Wert"] * EF_CH4.get(row["animal_type"], 0), axis=1
)

# Sum up per process
emissions_df = tier_df.groupby("BetriebProcessID")["CH4_kg"].sum().reset_index()

print("\n=== CH4 EMISSIONEN PRO PROZESS ===")
print(emissions_df.head())

# ============================================================================
# CONVERT TO LONG FORMAT
# ============================================================================

emissions_long = pd.melt(
    emissions_df,
    id_vars=["BetriebProcessID"],
    value_vars=[
        "CH4_kg",
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
    "CH4_kg": {
        "Flow": "Methane, non-fossil",
        "Category": "Elementary flows/Emission to air/low population density",
    }
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
# Spaces around "/" are normalised, because Access stores them inconsistently
# (e.g. "Elementary flows / Emission to air" instead of "Elementary flows/Emission to air")
katalog_flow_df["Flow_lower"] = katalog_flow_df["Flow"].str.lower()
katalog_flow_df["Category_lower"] = (
    katalog_flow_df["Category"].str.lower().str.replace(r"\s*/\s*", "/", regex=True)
)

emissions_long["Flow_lower"] = emissions_long["FlowName"].str.lower()
emissions_long["Category_lower"] = (
    emissions_long["Category"].str.lower().str.replace(r"\s*/\s*", "/", regex=True)
)

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

insert_emissions_df = mapped[["BetriebProcessID", "FlowID", "Wert"]].copy()

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
            ["BetriebProcessID", "Process", "Flow", "Category", "Wert"]
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
