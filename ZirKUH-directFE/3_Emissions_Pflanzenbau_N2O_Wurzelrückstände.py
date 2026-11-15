import pyodbc
import pandas as pd
import numpy as np
from config import db_path

print(
    "Hello World!, Lets calculate some N2O emissions in the Pflanzenbaumodule for Ernte- and Wurzelrueckstände"
)


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

PROCESS_TO_CROP = {
    # =========================
    # PERMANENT GRASSLAND (DG)
    # =========================
    "P_DG_Weide": "perennial_grass",
    "P_DG_Wiese": "perennial_grass",
    "P_DG_Maehweide": "perennial_grass",
    "P_DG_WieseExtensiv": "perennial_grass",
    # =========================
    # CEREALS (grains)
    # =========================
    "P_Acker_WW": "grains",  # Winter wheat
    "P_Acker_WD": "grains",  # Winter durum (assumed)
    "P_Acker_WG": "grains",  # Winter barley
    "P_Acker_WR": "grains",  # Winter rye
    "P_Acker_Wtri": "grains",  # Triticale
    "P_Acker_SH": "grains",  # Spring oats
    "P_Acker_SD": "grains",  # Spring durum or similar
    "P_Acker_SW": "grains",  # Spring wheat
    "P_Acker_SG": "grains",  # Spring barley
    "P_Acker_SR": "grains",  # Spring rye
    # =========================
    # MAIZE
    # =========================
    "P_Acker_Silomais": "maize",
    "P_Acker_Körnermais": "maize",
    # =========================
    # ROOT AND TUBER CROPS
    # =========================
    "P_Acker_Kartoffel": "tubers",
    "P_Acker_Zuckerrübe": "root_crops",
    # =========================
    # LEGUMES (N-fixing)
    # =========================
    "P_Acker_Erbse": "pulses",
    "P_Acker_Ackerbohne": "pulses",
    "P_Acker_Soja": "pulses",
    # =========================
    # OILSEEDS
    # =========================
    "P_Acker_Oelsaaten": "oilseeds",
    "P_Acker_Wraps": "oilseeds",
    # =========================
    # FORAGE CROPS
    # =========================
    "P_Acker_LKG": "n_fixing_forage",  # Lucerne-clover grass
    "P_Acker_Ackergras": "grass",
    "P_Acker_GPS": "grass",
    # =========================
    # OTHER
    # =========================
    "P_Acker_Gemüse": "root_crops",  # simplification
    "P_Acker_Miscanthus": "perennial_grass",
    "P_Acker_Soblu": "oilseeds",  # Sunflower
}
# =======================================================
# Values from TABLE 11.1A (NEW) and Table 11.2 (Updated) of IPCC 2006 + 2019 Refinement
# ========================================================
CROP_PARAMS = {
    # =========================
    # GRAINS (cereals)
    # =========================
    "grains": {
        "DRY": 0.88,
        "slope": 1.09,
        "intercept": 0.88,
        "N_AG": 0.006,
        "RBG": 0.22,
        "N_BG": 0.009,
        "FracRenew": 1,
    },
    # =========================
    # BEANS & PULSES (legumes)
    # =========================
    "pulses": {
        "DRY": 0.91,
        "slope": 1.13,
        "intercept": 0.85,
        "N_AG": 0.008,
        "RBG": 0.19,
        "N_BG": 0.008,
        "FracRenew": 1,
    },
    # =========================
    # TUBERS (potato)
    # =========================
    "tubers": {
        "DRY": 0.22,
        "slope": 0.10,
        "intercept": 1.06,
        "N_AG": 0.019,
        "RBG": 0.20,
        "N_BG": 0.014,
        "FracRenew": 1,
    },
    # =========================
    # ROOT CROPS (sugar beet etc.)
    # =========================
    "root_crops": {
        "DRY": 0.94,
        "slope": 1.07,
        "intercept": 1.54,
        "N_AG": 0.016,
        "RBG": 0.20,
        "N_BG": 0.014,
        "FracRenew": 1,
    },
    # =========================
    # N-FIXING FORAGES (lucerne, clover grass)
    # =========================
    "n_fixing_forage": {
        "DRY": 0.90,
        "slope": 0.30,
        "intercept": 0.0,
        "N_AG": 0.027,
        "RBG": 0.40,
        "N_BG": 0.022,
        "FracRenew": 0.34,  # assumption: three-year clover grass for cutting.
    },
    # =========================
    # NON-N FIXING FORAGES (grass)
    # =========================
    "grass": {
        "DRY": 0.90,
        "slope": 0.30,
        "intercept": 0.0,
        "N_AG": 0.015,
        "RBG": 0.54,
        "N_BG": 0.012,
        "FracRenew": 1,
    },
    # =========================
    # PERENNIAL GRASSES (permanent grassland)
    # =========================
    "perennial_grass": {
        "DRY": 0.90,
        "slope": 0.30,
        "intercept": 0.0,
        "N_AG": 0.015,
        "RBG": 0.80,
        "N_BG": 0.012,
        "FracRenew": 0.01,  # Permanent grassland is hardly ever renewed, hence a very low FracRenew value
    },
    # =========================
    # GRASS-CLOVER MIX
    # =========================
    "grass_clover": {
        "DRY": 0.90,
        "slope": 0.30,
        "intercept": 0.0,
        "N_AG": 0.025,
        "RBG": 0.80,
        "N_BG": 0.016,
        "FracRenew": 0.34,  # assumption: three-year clover grass for cutting.
    },
    # =========================
    # MAIZE
    # =========================
    "maize": {
        "DRY": 0.87,
        "slope": 1.03,
        "intercept": 0.61,
        "N_AG": 0.006,
        "RBG": 0.22,
        "N_BG": 0.007,
        "FracRenew": 1,
    },
    # =========================
    # OILSEEDS (no own entry → = cereals)
    # =========================
    "oilseeds": {
        "DRY": 0.88,
        "slope": 1.09,
        "intercept": 0.88,
        "N_AG": 0.006,
        "RBG": 0.22,
        "N_BG": 0.009,
        "FracRenew": 1,
    },
}


# FIX 2026-10-06 (DRY): crop groups whose yield is recorded in tbl_4 as FRESH MATTER.
# The DRY factor from CROP_PARAMS is applied only to these; all other yields
# are already in kg DM. Currently only potatoes (P_Acker_Kartoffel -> "tubers").
CROPS_RECORDED_FRESH = {"tubers"}


# get Fraction of residues removed (0.9 if straw is removed, 0 if incorporated)
def get_frac_remove(flows, crop):

    # 🌱 Permanent grassland
    if crop == "perennial_grass":
        return 0  # Permanent grassland has no straw or crop residues removed, so 0 is always returned here.

    # 🌾 Arable farming (straw logic)
    has_straw = any("Stroh" in str(f) for f in flows)

    if has_straw:
        return 0.9
    else:
        return 0.0


# ========================
# Mapping
# ========================
process_df = data["tbl_3_Betrieb_Process"].merge(
    data["Katalog_Process"][["ProcessID", "Process"]], on="ProcessID", how="left"
)

process_df["crop"] = process_df["Process"].map(PROCESS_TO_CROP)

process_df = process_df[
    process_df["Process"].str.startswith(("P_Acker", "P_DG"), na=False)
].copy()

process_df = process_df[process_df["crop"].notna()].copy()

print("\n=== CROP MAPPING CHECK ===")
print(process_df[["Process", "crop"]].drop_duplicates().sort_values("Process"))

missing = process_df[process_df["crop"].isna()]["Process"].unique()
print("\nNicht gemappte Prozesse:", missing)

# ========================
# get yield
# ========================
flow_df = data["tbl_4_Betrieb_Process_FlowValue"].merge(
    data["Katalog_Flow"][["FlowID", "Flow"]], on="FlowID", how="left"
)

flow_df = flow_df[flow_df["Direction"] == "Output"]

# ========================
# Extract yield (arable + permanent grassland only)
# ========================

# relevant flows (yield)
yield_flows = data["Katalog_Flow"][
    (data["Katalog_Flow"]["Category"].isin(["Ackerbau", "DG"]))
    & (data["Katalog_Flow"]["Subcategory"] == "Ertrag")
][["FlowID"]]

# keep only these flows
yield_df = flow_df.merge(yield_flows, on="FlowID", how="inner")

# numeric values (comma as decimal separator)
yield_df["Wert"] = pd.to_numeric(
    yield_df["Wert"].astype(str).str.replace(",", ".", regex=False),
    errors="coerce",
)

# optional: remove straw (very important!)
yield_df = yield_df[~yield_df["Flow"].str.contains("Stroh", case=False, na=False)]

# sum up per process
yield_df = (
    yield_df.groupby("BetriebProcessID", as_index=False)["Wert"]
    .sum()
    .rename(columns={"Wert": "yield"})
)

# merge into process_df
process_df = process_df.merge(yield_df, on="BetriebProcessID", how="left")

# Aggregate flows per process as a list (for the later check whether straw is included)
flows_per_process = flow_df.groupby("BetriebProcessID")["Flow"].apply(list).to_dict()

process_df["flows"] = process_df["BetriebProcessID"].map(flows_per_process)

# ========================================
# Calculation of the emissions from crop residues and roots
# ========================================


def calc_residue_n(row):
    crop = row["crop"]
    yield_dm = row["yield"]
    flows = row["flows"]  # must be built beforehand!

    if crop not in CROP_PARAMS or pd.isna(yield_dm) or yield_dm == 0:
        return 0

    p = CROP_PARAMS[crop]

    # FIX 2026-10-06 (DRY): IPCC Eq. 11.7 Crop = Yield_fresh * DRY. In the red DB
    # all yields are already recorded in kg DM, EXCEPT potatoes (fresh matter).
    # So DRY is applied only to the groups in CROPS_RECORDED_FRESH.
    if crop in CROPS_RECORDED_FRESH:
        yield_dm = yield_dm * p["DRY"]

    # AGDM correct = above ground yield of residue in kg DM/ha, based on the formula from IPCC 2006 Tier 2
    agdm = (yield_dm / 1000) * p["slope"] + p["intercept"]
    agdm_kg = agdm * 1000

    rag = agdm_kg / yield_dm

    # 🔥 Determine FracRemove
    frac_remove = get_frac_remove(flows, crop)

    # Formula IPCC 2019 Refinement, Eq. 11.6 (updated), Area = 1 ha:
    #   AGR = AG_DM;  BGR = (Crop + AG_DM) * RS * FracRenew
    n_ag = yield_dm * rag * p["N_AG"] * (1 - frac_remove)
    # FIX 2026-10-06 (R_BG): root biomass refers to the total above-ground
    # biomass (yield + crop residue), not only to the yield.
    # Old: n_bg = yield_dm * p["FracRenew"] * p["RBG"] * p["N_BG"]
    n_bg = (yield_dm + agdm_kg) * p["FracRenew"] * p["RBG"] * p["N_BG"]

    return n_ag + n_bg


process_df["N_residues"] = process_df.apply(calc_residue_n, axis=1)

process_df["N2O_kg"] = process_df["N_residues"] * 0.01 * (44 / 28)

print("\n=== SAMPLE OUTPUT ===")
print(process_df[["Process", "crop", "yield", "N_residues", "N2O_kg"]].head(10))

emissions_df = process_df[["BetriebProcessID", "N2O_kg"]].copy()

# ============================================================================
# Bring emissions into long format (for mapping to flows)
# ============================================================================

emissions_long = pd.melt(
    emissions_df,
    id_vars=["BetriebProcessID"],
    value_vars=["N2O_kg"],
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
    "N2O_kg": {
        "Flow": "Dinitrogen monoxide",
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
insert_emissions_df["TierkategorieEmissionFutter"] = (
    "Emission aus Ernte- und Wurzelrückständen"
)

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
    new_emissions_display = insert_rows.merge(
        data["Katalog_Flow"][["FlowID", "Flow", "Category"]],
        on="FlowID",
        how="left",
    )

    summary = (
        new_emissions_display.groupby("Flow")["Wert"]
        .agg(["count", "sum"])
        .reset_index()
    )

    print("\n=== NEUE EMISSIONEN (ZUSAMMENFASSUNG) ===")
    print(summary.to_string(index=False))

else:
    print("\nKeine neuen Emissionen einzufügen.")

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
    print("Änderungen wurden in die Access-Datenbank geschrieben.")
else:
    print(
        "Änderungen wurden verworfen. Keine Updates oder Inserts wurden durchgeführt."
    )
print(f"\nUpdates: {len(update_rows)}")
print(f"Inserts: {len(insert_rows)}")
