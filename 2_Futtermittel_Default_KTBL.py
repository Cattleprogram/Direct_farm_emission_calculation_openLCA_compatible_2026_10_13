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
# LOAD KTBL DATA
# ============================================================================

ktbl_df = data["Katalog_Futtermittelgehalte_KTBL"]

mapping_df = data["tbl_Futtermittel_Mapping"]

# Defensive deduplication: ensure the mapping table doesn't create duplicate rows
# when merged later. Keep unique (FuttermittelID, FlowID) pairs.
if {"FuttermittelID", "FlowID"}.issubset(mapping_df.columns):
    mapping_df = mapping_df.drop_duplicates(subset=["FuttermittelID", "FlowID"])


# ============================================================================
# SIMPLIFY COLUMN NAMES
# ============================================================================

ktbl_df = ktbl_df.rename(
    columns={
        "TM (g/kg)": "TM",
        "XP (g/kg TM)": "XP",
        "XF (g/kg TM)": "XF",
        "XA (g/kg TM)": "XA",
        "XL (g/kg TM)": "XL",
        "NEL (MJ/kg TM)": "NEL",
        "Gross Energy (GE MJ/kg TM)": "GE",
        "Energie (ME MJ/kg TM)": "ME",
        "NFE  (g/kg TM)": "NFE",
        "aNDF om (g/kg TM)": "aNDFom",
        "ADF om (g/kg TM)": "ADFom",
        "XZ (g/kg TM)": "XZ",
        "nXP (g/kg TM)": "nXP",
        "UDP (%)": "UDP",
    }
)


# ============================================================================
# PROPERTY COLUMNS
# ============================================================================

property_cols = [
    "TM",
    "XP",
    "XF",
    "XA",
    "XL",
    "GE",
    "ME",
    "NEL",
    "NFE",
    "aNDFom",
    "ADFom",
    "XZ",
    "nXP",
    "UDP",
]


# ============================================================================
# CREATE MISSING FLOWPROPERTIES AUTOMATICALLY
# ============================================================================

cursor = conn.cursor()

existing_props = data["Katalog_Flowproperty"][["FlowProperty"]].copy()

UNIT_MAPPING = {
    "TM": "g/kg",
    "XP": "g/kg TM",
    "XF": "g/kg TM",
    "XA": "g/kg TM",
    "XL": "g/kg TM",
    "GE": "MJ/kg TM",
    "ME": "MJ/kg TM",
    "NEL": "MJ/kg TM",
    "UDP": "%",
    "NFE": "g/kg TM",
    "aNDFom": "g/kg TM",
    "ADFom": "g/kg TM",
    "XZ": "g/kg TM",
    "nXP": "g/kg TM",
}

properties_df = pd.DataFrame(
    {
        "FlowProperty": property_cols,
        "Category": "Tierhaltung",
        "Subcategory": "Futtermittel",
        "Description": (
            "Achilles, W., Anter, J., Belau, T., "
            "& Blankenburg, J. (2018). "
            "Faustzahlen für die Landwirtschaft "
            "(15. Auflage). KTBL, S.910"
        ),
    }
)

# Add units
properties_df["Unit"] = properties_df["FlowProperty"].map(UNIT_MAPPING)

print(properties_df)

new_props = properties_df[
    ~properties_df["FlowProperty"].isin(existing_props["FlowProperty"])
]

print("\n=== NEUE FLOWPROPERTIES ===")
print(new_props)

# ============================================================================
# FLOWPROPERTIES UPSERT:
# Update existing properties, add new ones
# ============================================================================

cursor = conn.cursor()

# Load existing properties
existing_props_df = data["Katalog_Flowproperty"][
    ["FlowPropertyID", "FlowProperty"]
].copy()

# Link with new properties
merge_props_df = properties_df.merge(
    existing_props_df,
    on="FlowProperty",
    how="left",
)

# Existing properties → UPDATE
update_props = merge_props_df[merge_props_df["FlowPropertyID"].notna()]

# Properties that do not exist yet → INSERT
insert_props = merge_props_df[merge_props_df["FlowPropertyID"].isna()]

# Summary
print(f"{len(update_props)} vorhandene FlowProperties werden aktualisiert.")

print(f"{len(insert_props)} neue FlowProperties werden eingefügt.")

# ============================================================================
# Preview of new properties
# ============================================================================

if len(insert_props) > 0:

    print("\n=== NEUE FLOWPROPERTIES ===")

    print(
        insert_props[
            [
                "FlowProperty",
                "Category",
                "Subcategory",
                "Unit",
            ]
        ]
        .head(30)
        .to_string(index=False)
    )

    print("=" * 50)

# ============================================================================
# Confirm writing
# ============================================================================

confirm_props = input("\nFlowProperties in die Access-Datenbank schreiben? (j/n): ")

if confirm_props.lower() == "j":

    # =========================================================================
    # UPDATE existing properties
    # =========================================================================

    for _, row in update_props.iterrows():

        cursor.execute(
            """
            UPDATE Katalog_Flowproperty
            SET
                Category = ?,
                Subcategory = ?,
                Unit = ?,
                Description = ?
            WHERE FlowPropertyID = ?
            """,
            row["Category"],
            row["Subcategory"],
            row["Unit"],
            row["Description"],
            int(row["FlowPropertyID"]),
        )

    # =========================================================================
    # INSERT new properties
    # =========================================================================

    if len(insert_props) > 0:

        rows_to_insert = list(
            insert_props[
                [
                    "FlowProperty",
                    "Category",
                    "Subcategory",
                    "Unit",
                    "Description",
                ]
            ].itertuples(index=False, name=None)
        )

        cursor.executemany(
            """
            INSERT INTO Katalog_Flowproperty
            (
                FlowProperty,
                Category,
                Subcategory,
                Unit,
                Description
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            rows_to_insert,
        )

    else:

        print("Keine neuen FlowProperties zum Einfügen.")

    conn.commit()

    print("\n✅ FlowProperties erfolgreich geschrieben.")

else:

    print("\nÄnderungen verworfen.")
# ============================================================================
# RELOAD FLOWPROPERTIES
# ============================================================================

prop_df = pd.read_sql(
    """
    SELECT
        [FlowPropertyID],
        [FlowProperty]
    FROM Katalog_Flowproperty
    """,
    conn,
)


# ============================================================================
# KTBL FROM WIDE → LONG
# ============================================================================

ktbl_long = pd.melt(
    ktbl_df,
    id_vars=[
        "FuttermittelID",
        "Literatur",
    ],
    value_vars=property_cols,
    var_name="FlowProperty",
    value_name="DefaultValue",
)

# ============================================================================
# MAP PROPERTY IDs
# ============================================================================

ktbl_long = ktbl_long.merge(
    prop_df,
    on="FlowProperty",
    how="left",
)


# ============================================================================
# ADD FLOWIDs
# ============================================================================

ktbl_long = ktbl_long.merge(
    mapping_df,
    on="FuttermittelID",
    how="left",
)

print("\n✅ Daten erfolgreich transformiert und gemappt!")

print("\nHier ein Blick auf die gemappten Daten:")

print(
    ktbl_long[
        [
            "FuttermittelID",
            "FlowProperty",
            "DefaultValue",
            "FlowPropertyID",
            "FlowID",
            "Literatur",
        ]
    ].head(20)
)


# ============================================================================
# CREATE INSERT DATA
# ============================================================================
insert_df = ktbl_long[
    [
        "FlowID",
        "FlowPropertyID",
        "DefaultValue",
        "Literatur",
    ]
].copy()

insert_df = insert_df.rename(columns={"Literatur": "DefaultValueQuelle"})
insert_df = insert_df.dropna(subset=["FlowID", "FlowPropertyID"])

insert_df["DefaultValue"] = pd.to_numeric(
    insert_df["DefaultValue"],
    errors="coerce",
)

insert_df = insert_df.dropna(subset=["DefaultValue"])

print("\n✅ Bereit für Insert-Statements!")

print(insert_df.head())


# ============================================================================
# LOAD EXISTING VALUES
# ============================================================================

existing_df = pd.read_sql(
    """
    SELECT
        [FlowID],
        [FlowPropertyID],
        [FlowPropertyDefaultID],
        [DefaultValueQuelle]
    FROM Katalog_FlowPropertyDefault
    """,
    conn,
)


# ============================================================================
# PREPARE UPSERT
# ============================================================================

merge_df = insert_df.merge(
    existing_df,
    on=["FlowID", "FlowPropertyID"],
    how="left",
    suffixes=("", "_existing"),
)

update_rows = merge_df[merge_df["FlowPropertyDefaultID"].notna()]

insert_rows = merge_df[merge_df["FlowPropertyDefaultID"].isna()]

print(f"\n{len(update_rows)} vorhandene Werte werden aktualisiert.")

print(f"{len(insert_rows)} neue Werte werden eingefügt.")


# ============================================================================
# PREVIEW
# ============================================================================

if len(insert_rows) > 0:

    print("\n=== NEUE PROPERTY-WERTE ===")

    preview_df = insert_rows.merge(
        data["Katalog_Flow"][["FlowID", "Flow"]],
        on="FlowID",
        how="left",
    ).merge(
        prop_df,
        on="FlowPropertyID",
        how="left",
    )

    print(
        preview_df[
            [
                "Flow",
                "FlowProperty",
                "DefaultValue",
            ]
        ]
        .head(30)
        .to_string(index=False)
    )


# ============================================================================
# WRITE TO ACCESS
# ============================================================================

confirm = input("\nProperty-Werte in Access schreiben? (j/n): ")

if confirm.lower() == "j":

    # UPDATE

    for _, row in update_rows.iterrows():

        cursor.execute(
            """
            UPDATE Katalog_FlowPropertyDefault
            SET DefaultValue = ?,
                DefaultValueQuelle = ?
            WHERE FlowPropertyDefaultID = ?
            """,
            row["DefaultValue"],
            row["DefaultValueQuelle"],
            int(row["FlowPropertyDefaultID"]),
        )

    # INSERT

    if len(insert_rows) > 0:

        rows_to_insert = list(
            insert_rows[
                [
                    "FlowID",
                    "FlowPropertyID",
                    "DefaultValue",
                    "DefaultValueQuelle",
                ]
            ].itertuples(index=False, name=None)
        )

        cursor.executemany(
            """
            INSERT INTO Katalog_FlowPropertyDefault
            (
                FlowID,
                FlowPropertyID,
                DefaultValue,
                DefaultValueQuelle
            )
            VALUES (?, ?, ?, ?)
            """,
            rows_to_insert,
        )

    conn.commit()

    print("\n✅ Property-Werte erfolgreich geschrieben.")

else:

    print("\nÄnderungen verworfen.")
