print(
    "Hello World, this is the Emissionsberechnung den Pflanzenbau. Emissionen aus Duengemittelausbringung, Kalkung und Dieselverbrauch 🚜💨"
)
import pyodbc
import pandas as pd
import numpy as np
from config import db_path

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

print(duenger_df.head())
print(process_flows_df.head())
print(flows_def.head())
# ============================================================================
# Emission factors
# ============================================================================

# Emission factors for the emissions from fertiliser application
#  according to the EU PEF guidelines 2021
EF = {
    "N2O_direct": 0.022,  # kg N2O-N / kg N applied
    "NO3": 1.33,  # kg NO3-N / kg N applied
    "P2O5": 1,  # kg P2O5 / kg P2O5 applied. Previously the factor was 0.02, but that is
    # wrong, because the soil compartment is chosen and leaching and run-off then remain part of the impact assessment method.
    "CO2_Diesel": 2.606,  # kg CO2 / Liter Diesel IPCC: EF 19610 (3140 g/kg diesel)
    # 3140/100 *0,83 [kg/L]
    "CO2_Kalk": 0.44,  # kg CO2 / kg limestone
}
NH3_FACTORS = {
    "Mineralduenger": 0.12,
    "Wirtschaftsduenger": 0.24,
    "Duengemittel": 0.12,  # Assumption: organic fertilisers have the same NH3 factor as synthetic fertilisers
}
# ============================================================================
# ASSIGN EMISSION FACTORS PER FERTILISER ROW
# ============================================================================


def get_nh3_factor(subcategory):
    subcategory = str(subcategory)

    for key, value in NH3_FACTORS.items():
        if key in subcategory:
            return value

    return 0


# Apply factors to every row
duenger_df["NH3_factor"] = duenger_df["Subcategory"].apply(get_nh3_factor)


# ============================================================================
# CALCULATIONS
# ============================================================================

# Calculate emissions per individual fertiliser row
duenger_df["NH3_kg"] = duenger_df["kgN"] * duenger_df["NH3_factor"]
duenger_df["N2O_kg"] = duenger_df["kgN"] * EF["N2O_direct"]
duenger_df["NO3_kg"] = duenger_df["kgN"] * EF["NO3"]
duenger_df["P2O5_kg"] = duenger_df["kgP2O5"] * EF["P2O5"]

# CALCULATE EMISSIONS FROM DIESEL CONSUMPTION

# Convert Wert to numeric values (comma as decimal separator)
flow_values_df = data["tbl_4_Betrieb_Process_FlowValue"].copy()
flow_values_df["Wert"] = (
    flow_values_df["Wert"].astype(str).str.replace(",", ".", regex=False).str.strip()
)
flow_values_df["Wert"] = pd.to_numeric(flow_values_df["Wert"], errors="coerce")

# Join with Katalog_Flow to get the flow names
flow_values_df = flow_values_df.merge(
    flows_def[["FlowID", "Flow"]], on="FlowID", how="left"
)

# CO2 emissions from diesel consumption: this calculation is switched off for now,
# because the scope 1 emissions from diesel combustion are already included in the ecoinvent flow (diesel, burned in agricultural machinery)
# diesel_df = (
#    flow_values_df.loc[
#        flow_values_df["Flow"] == "diesel, burned in agricultural machinery"
#    ]
#    .assign(CO2_Diesel_kg=lambda x: x["Wert"] * EF["CO2_Diesel"])
#    .groupby("BetriebProcessID", as_index=False)["CO2_Diesel_kg"]
#    .sum()
# )

# Calculate emissions from liming
kalk_df = (
    flow_values_df.loc[flow_values_df["Flow"] == "limestone, milled, loose"]
    .assign(CO2_Kalk_kg=lambda x: x["Wert"] * EF["CO2_Kalk"])
    .groupby("BetriebProcessID", as_index=False)["CO2_Kalk_kg"]
    .sum()
)
# ============================================================================
# SUM UP AT PROCESS LEVEL
# ============================================================================
emissions_df = (
    duenger_df.groupby("BetriebProcessID")
    .agg(
        {
            "NH3_kg": "sum",
            "N2O_kg": "sum",
            "NO3_kg": "sum",
            "P2O5_kg": "sum",
        }
    )
    .reset_index()
)
# Diesel CO2 switched off for now, because the scope 1 emission is included in the ecoinvent flow
# emissions_df = emissions_df.merge(diesel_df, on="BetriebProcessID", how="left")
# emissions_df["CO2_Diesel_kg"] = emissions_df["CO2_Diesel_kg"].fillna(0)

emissions_df = emissions_df.merge(kalk_df, on="BetriebProcessID", how="left")
emissions_df["CO2_Kalk_kg"] = emissions_df["CO2_Kalk_kg"].fillna(0)

print(emissions_df.head())
print(emissions_df.tail(20))
print(emissions_df.sample(10))


# ============================================================================
# CONVERT TO LONG FORMAT
# ============================================================================

emissions_long = pd.melt(
    emissions_df,
    id_vars=["BetriebProcessID"],
    value_vars=[
        "NH3_kg",
        "N2O_kg",
        "NO3_kg",
        "P2O5_kg",
        # "CO2_Diesel_kg",
        "CO2_Kalk_kg",
    ],
    var_name="EmissionType",
    value_name="Wert",
)

# Keep only emissions > 0
emissions_long = emissions_long[emissions_long["Wert"] > 0]

print(emissions_long.head())
print(emissions_long.tail(20))
print(emissions_long.sample(10))

# ============================================================================
# Map calculated emissions to the emission flows (red DB and openLCA)
# ============================================================================

# Mapping between the calculated emissions and the flow names in Katalog_Flow
flow_mapping = {
    "NH3_kg": {
        "Flow": "Ammonia",
        "Category": "Elementary flows/Emission to air/low population density",
    },
    "N2O_kg": {
        "Flow": "Dinitrogen monoxide",
        "Category": "Elementary flows/Emission to air/low population density",
    },
    "NO3_kg": {
        "Flow": "Nitrate",
        "Category": "Elementary flows/Emission to water/unspecified",
    },
    "P2O5_kg": {
        "Flow": "Phosphorus pentoxide",
        "Category": "Elementary flows/Emission to soil/agricultural",
    },
    "CO2_Diesel_kg": {
        "Flow": "Carbon dioxide, fossil",
        "Category": "Elementary flows/Emission to air/low population density",
    },
    "CO2_Kalk_kg": {
        "Flow": "Carbon dioxide",
        "Category": "Elementary flows / Emission to air / low population density",
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
print(insert_emissions_df.sample(10))
print(f"{len(insert_emissions_df)} Emissionszeilen vorbereitet")


# ============================================================================
# Robust upsert: update existing emissions in the red DB, insert new emissions
# ============================================================================

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

    print(
        f"✓ {len(update_rows)} Emissionen aktualisiert und "
        f"{len(insert_rows)} neue Emissionen eingefügt."
    )

else:
    print("Abgebrochen – keine Änderungen geschrieben.")
