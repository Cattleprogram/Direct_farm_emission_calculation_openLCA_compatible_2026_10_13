# Data

This folder contains the input data for a test run with two example farms. The data have
their own licence terms (see [LICENSE](LICENSE)); the MIT licence of the repository applies
to the code only.

## Files

| File | Content | Licence |
|------|---------|---------|
| `2026_10_13_Alle_Daten_Example.accdb` | Access database: farm data of farms 2 and 8 plus catalogues and queries (structure: see the main README) | farm data CC BY 4.0; feed values: rights of KTBL / LfL |
| `Parameterdatei_2024-12_Export.xlsx` | KTBL BEK parameter file, sheet `BEK-Parameterdatei` (TAN shares and emission factors, read by script 4) | rights of KTBL, reproduced with permission |

Files written by `run_all.py` (`*_EmissionenBerechnet*.accdb`) are not part of the repository.

## How to cite the data

**Farm data (farms 2 and 8):**

- Zeidler J., Gramsch J., Schiefen S., Knob D.-A., Wölfinger M., Zeller V., Häußermann U., Gattinger A., Breuer L. (2027) (to be submitted): ZirKUH-directFE: An open-source modular LCA tool for calculating direct on-farm emissions in complex agricultural production systems compatible with openLCA 2.6. <DOI>. Licence: CC BY 4.0.

**Third-party sources used in the data** (cite them when you use these values):

- Arbeitsgruppe BEK (2024): Berechnungsstandard für einzelbetriebliche Klimabilanzen (BEK) in
  der Landwirtschaft. Berechnungsparameter für einzelbetriebliche Klimabilanzen. Darmstadt:
  Kuratorium für Technik und Bauwesen in der Landwirtschaft (KTBL).
  https://daten.ktbl.de/bek/ (accessed 13.10.2026). → `Parameterdatei_2024-12_Export.xlsx`
- Achilles, W., Anter, J., Belau, T., Blankenburg, J. et al. (2018): Faustzahlen für die
  Landwirtschaft, 15. Auflage. Darmstadt: KTBL. → feed values, pages are given per value in
  the column `Literatur` / `DefaultValueQuelle`
- Bayerische Landesanstalt für Landwirtschaft (LfL) (Ed.) (2024): Gruber Tabelle zur Fütterung
  der Milchkühe, Zuchtrinder, Schafe, Ziegen. LfL-Information 48. Freising: LfL, Institut für
  Tierernährung und Futterwirtschaft.
  https://www.lfl.bayern.de/mam/cms07/publikationen/daten/informationen/gruber_tabelle_milchvieh-2023_lfl-information.pdf
  → feed values
- LfL: Deckungsbeiträge und Kalkulationsdaten. https://www.stmelf.bayern.de/idb/winterweizen.html
  → default nutrient contents of farm manures (total N, P₂O₅, K₂O) in `Katalog_FlowPropertyDefault`
- Bachinger, J., Klöble, U. (Eds.) (2019): Faustzahlen für den Ökologischen Landbau,
  2. Auflage. Darmstadt: KTBL.

## Anonymisation

The farms are identified only by their project number (ZirKUH_ID). Coordinates and free-text
survey answers were removed. The calculation uses only the master data "GVE Rind" and
"Betriebsform" from `tbl_2_Betrieb_Stammdaten`.
