import pyodbc
import pandas as pd
import numpy as np
from config import db_path

print("Hello World! Let's calculate PM, NMVOC and N2O emissions from Pflanzenbau!")
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

# ============================================================================
# Emission factors
# ============================================================================

# ============================================================================
# CROP PRODUCTION EMISSIONS (Tier 1, 1 ha per process) from EMEP 2023
# ============================================================================

EF_PFLANZENBAU = {
    "NMVOC": 0.86,  # kg / ha
    "PM10": 1.56,  # kg / ha
    "PM2_5": 0.06,  # kg / ha
}

# Load processes + names
process_df = data["tbl_3_Betrieb_Process"].merge(
    data["Katalog_Process"][["ProcessID", "Process"]], on="ProcessID", how="left"
)

# Filter crop production processes only
pflanzenbau_df = process_df[
    process_df["Process"].str.startswith(("P_Acker", "P_DG"), na=False)
].copy()

print("\n=== PFLANZENBAU PROZESSE ===")
print(pflanzenbau_df[["BetriebProcessID", "Process"]].head())

# ============================================================================
# CALCULATE EMISSIONS (1 ha per process)
# ============================================================================

pflanzenbau_df["ha"] = 1  # <-- KEY SIMPLIFIER

pflanzenbau_df["NMVOC_kg"] = pflanzenbau_df["ha"] * EF_PFLANZENBAU["NMVOC"]
pflanzenbau_df["PM10_kg"] = pflanzenbau_df["ha"] * EF_PFLANZENBAU["PM10"]
pflanzenbau_df["PM2_5_kg"] = pflanzenbau_df["ha"] * EF_PFLANZENBAU["PM2_5"]

pflanzenbau_df["PM2_5_10_kg"] = (
    pflanzenbau_df["PM10_kg"] - pflanzenbau_df["PM2_5_kg"]
).clip(lower=0)

print("\n=== BERECHNETE EMISSIONEN (1 ha BASIS) ===")
print(pflanzenbau_df.head())

# ============================================================================
# FORMAT FOR FURTHER PROCESSING
# ============================================================================

emissions_df = pflanzenbau_df[
    ["BetriebProcessID", "NMVOC_kg", "PM2_5_kg", "PM2_5_10_kg"]
].copy()


flows_df = data["tbl_4_Betrieb_Process_FlowValue"].merge(
    data["Katalog_Flow"][["FlowID", "Flow"]], on="FlowID", how="left"
)
flows_per_process = flows_df.groupby("BetriebProcessID")["Flow"].apply(list).to_dict()

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
insert_emissions_df["TierkategorieEmissionFutter"] = "Pflanzenbau"

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
    # FIX 2026-10-06: summary indented into the if block. Before, it was
    # outside and aborted with a NameError when there were no new rows (every
    # repeated run) - before the j/n prompt, so updates were never possible.
    summary = new_emissions_display.groupby("Flow").size().reset_index(name="n_rows")

    print("\n=== ANZAHL EMISSIONS-EINTRÄGE ===")
    print(summary.to_string(index=False))
    print("=" * 50)

print(f"{len(update_rows)} vorhandene Emissionen werden aktualisiert.")
print(f"{len(insert_rows)} neue Emissionen werden eingefügt.")

confirm = input("Änderungen in die Access-Datenbank schreiben? (j/n): ")

if confirm.lower() == "j":

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
