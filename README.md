# ZirKUH-directFE – Direct emissions calculation for dairy and suckler cattle farms - mixed crop livestock farms

Python scripts that calculate the direct field and livestock emissions of farms
(crop production, manure management, grazing, enteric fermentation). The farm data
is read from a Microsoft Access database. The results are written back as output flows of the farm processes, ready for the life cycle assessment in openLCA.

The repository includes an example dataset with two farms (**farm 2** and **farm 8**),
so you can reproduce the full calculation chain. A run on the example dataset gives the
same emissions for these two farms as the calculation behind the paper.

> Code comments are in English. Console output, database table/column names and
> catalogue entries (e.g. `P_Milchkuh`, `Kälber`) are in German, because the code
> matches on them.

## Repository structure

```
├── 1_ … 5_*.py              emission calculation (run in this order, or via run_all.py)
├── run_all.py               runs all scripts on a copy of the database, writes results/
├── emission_tables.py       Word tables for the example farms (called by run_all.py)
├── config.py                all paths, relative to the repository folder
├── Transfer_OpenLCA/        transfer of the farm processes and emissions to openLCA
├── data/                    example data – own licence, see data/LICENSE and data/README.md
├── results/                 output (created by the scripts, not part of the repository)
├── LICENSE                  MIT licence of the code
└── CITATION.cff             how to cite this repository
```

## Calculation modules

The scripts must run in this order, because later scripts use results that
earlier ones wrote to the database.

| # | Script | Content | Main sources |
|---|--------|---------|--------------|
| 1 | `1_Update_P_Milchkuh_Kaelberration.py` | Derives the calf ration in dairy processes (share of the cow ration plus milk intake) | Project assumptions |
| 2 | `2_Futtermittel_Default_KTBL.py` | Fills default feed properties (GE, XP, XF, …) from KTBL feed tables | Achilles et al. 2018; LfL 2024 |
| 3 | `3_Emissions_Pflanzenbau_N2O_CO2_NH3_NO3_P2O5.py` | Crop production: N₂O, CO₂, NH₃, NO₃⁻, P from fertiliser, lime and diesel | European Commission 2021 (PEF); IPCC 2006 |
| 3 | `3_Emissions_Pflanzenbau_N2O_Wurzelrückstände.py` | Crop production: N₂O from crop residues | IPCC 2006; IPCC 2019 |
| 3 | `3_Emissions_Pflanzenbau_NMVOC_PM.py` | Crop production: NMVOC and particulate matter | EMEP/EEA 2023 (Tier 1) |
| 4 | `4_Emissions_Duengemanagement_N2O_NOx_NH3_Processlevel.py` | Manure management, housing and grazing: N₂O, NO, NH₃; NO₃⁻ and P₂O₅ from pasture excreta (per ha on the pasture processes) | Arbeitsgruppe BEK 2024; European Commission 2021; Achilles et al. 2018 |
| 4 | `4_Emissions_Duengemanagement_Weide_CH4_Processlevel.py` | Manure management and grazing: CH₄ (VS, DOM) | Arbeitsgruppe BEK 2024; Arbeitsgruppe BEK 2026 |
| 4 | `4_Emission_Duengemanagement_PM_NMVOC_Processlevel.py` | Manure management: NMVOC and particulate matter | EMEP/EEA 2023 |
| 5 | `5_Emissions_Tierhaltung_CH4_Cattle.py` | Enteric CH₄ of cattle (Tier 2 / Tier 3) | Arbeitsgruppe BEK 2026 |
| 5 | `5_Emissions_Tierhaltung_CH4_Schwein_Schaf_Pferd.py` | Enteric CH₄ of pigs, sheep and horses (Tier 1) | IPCC 2006 |

Full references: see [References](#references) and the accompanying paper.

### Nitrate and phosphate emissions

**Fertiliser application (script 3).** Nitrate leaching is 1.33 kg NO₃⁻ per kg N
applied (EU PEF 2021). Phosphorus is recorded as the P₂O₅ applied, with a factor of 1,
as an emission to soil/agricultural; the impact assessment method then models leaching
and run-off. Both are booked per ha on the crop and grassland processes.

**Pasture excreta (script 4).** The N excreted on pasture (`mexcr_graz`) is first summed
per farm. Nitrate is `mexcr_graz` × 1.33 kg NO₃⁻/kg N, the same factor as for fertiliser.
Phosphorus is `mexcr_graz` × the P₂O₅/N ratio of the excreta for each animal type
(Achilles et al. 2018):

| Animal | kg P₂O₅ per kg N |
|--------|------------------|
| Dairy cows (p. 487) | 0.359 |
| Suckler cows with calf (p. 487) | 0.295 |
| Rearing and fattening cattle (p. 486) | 0.323 |
| Ewes (p. 491) | 0.284 |

The farm total is split across the pasture processes (`P_DG_Weide`, `P_DG_Maehweide`)
by their share of grazed feed (Weidefutteraufnahme × area). It is then divided by the
area, because these processes are modelled per ha. As with fertiliser, NO₃⁻ is an
emission to water and P₂O₅ an emission to soil/agricultural. Both are tagged
"Weide Exkremente" so they stay separate from the fertiliser rows.

## Requirements

- **Windows** with the **Microsoft Access Database Engine** (ODBC driver
  `Microsoft Access Driver (*.mdb, *.accdb)`). The driver and Python must have the
  same bitness (usually 64-bit).
- Python 3.13 (other versions ≥ 3.10 probably work but are untested)
- Only for the transfer to openLCA: openLCA 2.6 and a database with ecoinvent
  (see [Transfer to openLCA](#transfer-to-openlca))

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Data

The scripts expect the following files in `data/` (filenames set in `config.py`). The data
have their own licence and citation, see [data/README.md](data/README.md) and
[data/LICENSE](data/LICENSE).

| File | Content |
|------|---------|
| `2026_10_13_Alle_Daten_Example.accdb` | Example database with farms 2 and 8 (farm master data, processes, input flows, feed tables, flow catalogue) |
| `Parameterdatei_2024-12_Export.xlsx` | KTBL BEK parameter file, sheet `BEK-Parameterdatei` (TAN shares and emission factors, used by script 4) |

The example database contains almost only what the scripts and the transfer to openLCA need:

- **Farm data** (farms 2 and 8 only): `Katalog_Betriebe`, `tbl_2_Betrieb_Stammdaten`
  (master data), `tbl_3_Betrieb_Process` (processes per farm), `tbl_4_Betrieb_Process_FlowValue`
  (input and output flows), `tbl_5_Betrieb_FlowpropertyValue` (farm-specific feed properties),
  `tbl_6_Betrieb_ProcessParameterValue` (process parameters such as animal numbers).
  The farm locations (coordinates) were removed.
- **Catalogues**: `Katalog_Flow`, `Katalog_Flowproperty`, `Katalog_FlowPropertyDefault`,
  `Katalog_Process`, `Katalog_Process_Category`, `Katalog_Processparameter`,
  `Katalog_ProcessProcessparameter`, `Katalog_Futtermittelgehalte_KTBL`, `tbl_Futtermittel_Mapping`.
  The source of every feed value is stored in the columns `Literatur` / `DefaultValueQuelle`.
- **Queries**: `☀_N_P_Duengemittel_qry` and its sub-queries `☼_*` (fertiliser N and P₂O₅,
  used by scripts 3 and 4), `★_python_qry_Processes_Flows` (all flows per farm process,
  read by the transfer to openLCA)

To use your own data, put a database with the same structure in `data/` and change
`DB_FILENAME` in `config.py`.

## Usage

### Recommended: everything in one run

```powershell
python run_all.py
```

- Copies the database to a temporary folder and runs all scripts on the copy.
  **The original database stays unchanged.**
- Answers every `(j/n)` prompt of the scripts with "j" automatically.
- Writes to `results/`:
  - `emissions_<timestamp>.csv`: all emissions as a long table (farm, process, flow, compartment, amount in kg/yr, module)
  - `logs_<timestamp>/`: one log file per script
  - `table_reproducibility_<timestamp>.docx` and `table_SI_full_<timestamp>.docx`: tables for the example farms
- At the end it asks whether to save the copy with all results as a new dated database in
  `data/` (`<date>_Alle_Daten_EmissionenBerechnet.accdb`). Answer `j` if you want to
  transfer the results to openLCA afterwards.

`python run_all.py --tables-only` rebuilds the Word tables from the newest CSV without
recalculating anything.

### Running single scripts

Each script can also be started on its own, e.g. `python 3_Emissions_Pflanzenbau_NMVOC_PM.py`.
Before writing, every script asks for confirmation (`j/n`).

> **Caution:** single scripts write **directly into the database configured in
> `config.py`**. Work on a copy, and keep the order 1 → 5.

## Transfer to openLCA

The scripts in `Transfer_OpenLCA/` write the farm processes, including the calculated
emissions, to openLCA through its IPC interface. They read the saved Access query
`★_python_qry_Processes_Flows` from the database **with** the emissions (see above: save
with `j` at the end of `run_all.py`). By default they take the newest
`data/*_EmissionenBerechnet*.accdb`; another database can be given with `--db`.

**Requirements**

- openLCA 2.6 with the IPC server running (openLCA: *Tools → Developer tools → IPC server*,
  default port 8080)
- an openLCA database with the background data and the impact assessment method. The
  paper uses ecoinvent 3.11 (system model APOS) and Agribalyse 3.1.1 with ReCiPe 2016 Midpoint (H). ecoinvent is
  licensed and therefore **not** included; you need your own licence.

**Scripts**

| Script | What it does |
|--------|--------------|
| `integrate_data_to_open_lca.py` | Creates the flows and unit processes. Existing flows (ecoinvent, elementary flows of the method) are reused by name and category; only missing flows are created under `00001_ZirKUH/`. One process per farm process, named `<ZirKUH_ID>_<Process>`. Existing processes are updated. |
| `update_values_AccessDB_to_openLCA.py` | After a recalculation: updates **only the amounts** of exchanges that already exist (dry run by default, `--write` asks before writing, `--farms 2 8` limits the farms). Reports go to `results/update_reports/`. |
| `read_database.py` | Shared helper: finds the database and reads the query. |

```powershell
cd Transfer_OpenLCA
python integrate_data_to_open_lca.py                       # first transfer
python update_values_AccessDB_to_openLCA.py --farms 2 8    # later: dry run of value updates
python update_values_AccessDB_to_openLCA.py --farms 2 8 --write
```

**Unit conversions** (Access unit → openLCA reference unit)

| Access | openLCA | Factor |
|--------|---------|--------|
| dt | kg | × 100 |
| t | kg | × 1000 |
| l (not diesel) | m³ | ÷ 1000 |
| l diesel | kWh | × 11.8 kWh/kg × 0.83 kg/l |

**Manual steps in openLCA (not done by the scripts)**

The scripts transfer the inventory data. The following modelling decisions are made by hand
in openLCA and are described in the paper:

- **Reference flow:** a new process gets its first output row as quantitative reference,
  only so that openLCA accepts it. This is arbitrary (the query has no fixed row order) and
  can be a co-product, manure or even an emission. The script warns when the reference
  would be an elementary flow. **Check and set the reference flow of every new process.**
- allocation between co-products (e.g. milk, calves, cull cows)
- linking of the inputs to their providers and building the product systems
- the impact assessment (calculation of the product systems)

The script also warns when an elementary flow was not found in openLCA and had to be
created: such a flow has no characterisation factors, so the method ignores it. Map it to
the matching elementary flow of your method.

## Validation

`5_Emissions_Tierhaltung_CH4_Cattle.py` contains manual check calculations for the example
farms (farm 8: BetriebProcessID 31; farm 2: BetriebProcessID 494). These compare the script
result with a hand calculation using `assert`. The input values of these checks are fixed
to the example dataset. If you use other data, remove or adapt this block.

## Licence

- **Code:** MIT licence, see [LICENSE](LICENSE).
- **Data** (`data/`): not covered by the MIT licence. Farm data: CC BY 4.0. Third-party data
  (KTBL, LfL) remain under the rights of their publishers and are reproduced with
  permission. Details: [data/LICENSE](data/LICENSE).

## Citation

If you use this code, please cite the paper and the repository:

> Zeidler J., Gramsch J., Schiefen S., Knob D.-A., Wölfinger M., Zeller V., Häußermann U.,
> Gattinger A., Breuer L. (in preparation): ZirKUH-directFE: An open-source modular LCA tool for
> calculating direct on-farm emissions in complex agricultural production systems compatible
> with openLCA 2.6. Manuscript to be submitted.

> Zeidler J., Gramsch J., Schiefen S., Knob D.-A., Wölfinger M., Zeller V., Häußermann U.,
> Gattinger A., Breuer L. (2026): ZirKUH-directFE – Direct emissions calculation for dairy and
> suckler cattle farms, version <VERSION>. <ZENODO DOI>

GitHub shows the machine-readable version from [CITATION.cff](CITATION.cff) under
"Cite this repository". When you use the example data, also cite the data sources listed
in [data/README.md](data/README.md).

## References

- Achilles, W., Anter, J., Belau, T., Blankenburg, J. et al. (2018): Faustzahlen für die
  Landwirtschaft, 15. Auflage. Darmstadt: Kuratorium für Technik und Bauwesen in der
  Landwirtschaft (KTBL).
- Arbeitsgruppe BEK (2024): Berechnungsstandard für einzelbetriebliche Klimabilanzen (BEK) in
  der Landwirtschaft. Berechnungsparameter für einzelbetriebliche Klimabilanzen. Darmstadt:
  KTBL. https://daten.ktbl.de/bek/ (accessed 18.12.2024)
- Arbeitsgruppe BEK (2026): Methodenbeschreibung Sachbilanz Tierhaltung der BEK-Fachgruppen. Kuratorium für Technik und Bauwesen in der Landwirtschaft e. V. Handbuch, Darmstadt
- Bachinger, J., Klöble, U. (Eds.) (2019): Faustzahlen für den Ökologischen Landbau,
  2. Auflage. Darmstadt: KTBL.
- Bayerische Landesanstalt für Landwirtschaft (LfL) (Ed.) (2024): Gruber Tabelle zur Fütterung
  der Milchkühe, Zuchtrinder, Schafe, Ziegen. LfL-Information 48. Freising: LfL.
- EMEP/EEA (2023): EMEP/EEA air pollutant emission inventory guidebook 2023. Copenhagen:
  European Environment Agency.
- European Commission (2021): Commission Recommendation (EU) 2021/2279 on the use of the
  Environmental Footprint methods (PEF). Official Journal of the European Union L 471.
- IPCC (2006): 2006 IPCC Guidelines for National Greenhouse Gas Inventories, Volume 4:
  Agriculture, Forestry and Other Land Use. Hayama: IGES.
- IPCC (2019): 2019 Refinement to the 2006 IPCC Guidelines for National Greenhouse Gas
  Inventories, Volume 4: Agriculture, Forestry and Other Land Use. Geneva: IPCC.
