"""
Print-ready Word tables from the emission table of run_all.py.

  - Reproducibility table (paper): sum per module × flow, one column per example farm
  - Complete table (SI): every emission row of the example farms

Layout in scientific style (booktabs): only rules at the top, below the header and
at the bottom, no vertical lines.
"""

import re

import pandas as pd
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

TABLE_LABEL_PAPER = "Table X"
TABLE_LABEL_SI = "Table S3"
FONT = "Times New Roman"

MODULE_LABELS = {
    "crop_N_CO2_P": "Crop production: fertiliser and lime",
    "crop_residues_N2O": "Crop production: crop residues",
    "crop_NMVOC_PM": "Crop production: NMVOC and PM",
    "manure_N": "Manure management and grazing: N",
    "manure_CH4": "Manure management and grazing: CH4",
    "manure_PM_NMVOC": "Manure management: NMVOC and PM",
    "enteric_cattle": "Enteric fermentation: cattle",
    "enteric_other": "Enteric fermentation: pigs, sheep, horses",
}

RULE_THICK = 12  # eighths of a point -> 1.5 pt
RULE_THIN = 6  # 0.75 pt


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================


def compartment(category):
    """'Elementary flows / Emission to air / low population density' -> 'air, low population density'."""
    if pd.isna(category):
        return ""
    text = re.sub(r"\s*/\s*", "/", str(category))
    text = re.sub(r"^Elementary flows/", "", text)
    text = re.sub(r"^Emission to ", "", text)
    return text.replace("/", ", ")


def fmt(value):
    if pd.isna(value):
        return "–"
    if value != 0 and abs(value) < 0.005:
        return f"{value:.2g}"
    return f"{value:,.2f}"


def set_cell_border(cell, edge, size):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    existing = borders.find(qn(f"w:{edge}"))
    if existing is not None:
        borders.remove(existing)
    element = OxmlElement(f"w:{edge}")
    element.set(qn("w:val"), "single")
    element.set(qn("w:sz"), str(size))
    element.set(qn("w:space"), "0")
    element.set(qn("w:color"), "000000")
    borders.append(element)


def rule(row, edge, size):
    for cell in row.cells:
        set_cell_border(cell, edge, size)


def set_row_flag(row, flag):
    """w:tblHeader = repeat the header row on every page,
    w:cantSplit = do not split the row across a page break."""
    tr_pr = row._tr.get_or_add_trPr()
    element = OxmlElement(f"w:{flag}")
    element.set(qn("w:val"), "true")
    tr_pr.append(element)


def write_cell(cell, text, size, bold=False, italic=False, align="left"):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_before = Pt(1)
    paragraph.paragraph_format.space_after = Pt(1)
    paragraph.alignment = (
        WD_ALIGN_PARAGRAPH.RIGHT if align == "right" else WD_ALIGN_PARAGRAPH.LEFT
    )
    run = paragraph.add_run(str(text))
    run.font.size = Pt(size)
    run.font.name = FONT
    run.bold = bold
    run.italic = italic


def new_document(landscape):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = FONT
    style.font.size = Pt(10)
    section = doc.sections[0]
    if landscape:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = Cm(29.7), Cm(21.0)
    else:
        section.page_width, section.page_height = Cm(21.0), Cm(29.7)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, side, Cm(2.0))
    return doc


def save_docx(doc, path):
    """Saves; if the file is currently open in Word, saves under a new name (_2, _3, ...)."""
    candidate, counter = path, 2
    while True:
        try:
            doc.save(candidate)
            break
        except PermissionError:
            candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
            counter += 1
    if candidate != path:
        print(f"  ⚠️  {path.name} ist geöffnet - gespeichert als {candidate.name}")
    return candidate


def add_caption(doc, label, text):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(f"{label}. ")
    run.bold = True
    paragraph.add_run(text)


def add_note(doc, text):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(3)
    run = paragraph.add_run(text)
    run.font.size = Pt(8)
    run.italic = True


def build_table(doc, header, rows, widths, size, right_cols, group_rows=()):
    """Table with booktabs rules. group_rows: row indexes that run across the full
    width as subheadings (bold, rule above)."""
    table = doc.add_table(rows=1 + len(rows), cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    for j, text in enumerate(header):
        write_cell(
            table.rows[0].cells[j],
            text,
            size,
            bold=True,
            align="right" if j in right_cols else "left",
        )
    set_row_flag(table.rows[0], "tblHeader")

    for i, values in enumerate(rows, start=1):
        row = table.rows[i]
        set_row_flag(row, "cantSplit")
        if (i - 1) in group_rows:
            merged = row.cells[0].merge(row.cells[-1])
            write_cell(merged, values[0], size, bold=True)
            if i > 1:
                rule(row, "top", RULE_THIN)
            continue
        for j, text in enumerate(values):
            write_cell(
                row.cells[j], text, size, align="right" if j in right_cols else "left"
            )

    for grid_col, width in zip(table._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        grid_col.set(qn("w:w"), str(int(Cm(width).twips)))
    for i, row in enumerate(table.rows):
        if i > 0 and (i - 1) in group_rows:
            row.cells[0].width = Cm(sum(widths))
            continue
        for j, width in enumerate(widths):
            row.cells[j].width = Cm(width)

    rule(table.rows[0], "top", RULE_THICK)
    rule(table.rows[0], "bottom", RULE_THIN)
    rule(table.rows[-1], "bottom", RULE_THICK)
    return table


# ============================================================================
# TABLE 1: REPRODUCIBILITY (PAPER)
# ============================================================================


def reproducibility_table(result, farms, module_order, path, source_name, run_date):
    df = result[result["farm"].isin(farms)].copy()
    df["compartment"] = df["category"].map(compartment)

    pivot = df.pivot_table(
        index=["module", "flow", "compartment"],
        columns="farm",
        values="amount",
        aggfunc="sum",
    ).reindex(columns=farms)
    pivot = pivot.reset_index()
    pivot["_order"] = pivot["module"].map({m: i for i, m in enumerate(module_order)})
    pivot = pivot.sort_values(["_order", "flow", "compartment"])

    totals = (
        df.groupby(["flow", "compartment", "farm"])["amount"]
        .sum()
        .unstack("farm")
        .reindex(columns=farms)
        .reset_index()
        .sort_values(["flow", "compartment"])
    )

    rows, group_rows = [], []
    for module, block in pivot.groupby("_order", sort=True):
        module_name = block["module"].iloc[0]
        group_rows.append(len(rows))
        rows.append([MODULE_LABELS.get(module_name, module_name)])
        for _, r in block.iterrows():
            rows.append([r["flow"], r["compartment"]] + [fmt(r[f]) for f in farms])
    group_rows.append(len(rows))
    rows.append(["Total over all modules"])
    for _, r in totals.iterrows():
        rows.append([r["flow"], r["compartment"]] + [fmt(r[f]) for f in farms])

    header = ["Flow", "Compartment"] + [f"Farm {f}" for f in farms]
    farm_width = 2.6
    widths = [6.2, 5.0] + [farm_width] * len(farms)
    right_cols = set(range(2, 2 + len(farms)))

    doc = new_document(landscape=False)
    add_caption(
        doc,
        TABLE_LABEL_PAPER,
        "Direct emissions per calculation module and elementary flow for the example "
        f"farms {', '.join(str(f) for f in farms)} (kg yr⁻¹).",
    )
    build_table(doc, header, rows, widths, 9, right_cols, group_rows)
    add_note(
        doc,
        "Compartment = openLCA flow category without the prefix "
        "'Elementary flows/Emission to'. – = no emission calculated. "
        f"Generated by run_all.py from {source_name} on {run_date}. "
        f"Row-level values are given in {TABLE_LABEL_SI}.",
    )
    return save_docx(doc, path)


# ============================================================================
# TABLE 2: COMPLETE (SI)
# ============================================================================


def full_table(result, farms, module_order, path, source_name, run_date):
    df = result[result["farm"].isin(farms)].copy()
    df["compartment"] = df["category"].map(compartment)
    df["_farm_order"] = df["farm"].map({f: i for i, f in enumerate(farms)})
    df["_module_order"] = df["module"].map({m: i for i, m in enumerate(module_order)})
    df = df.sort_values(
        ["_farm_order", "_module_order", "process", "process_instance", "flow"]
    )

    rows, group_rows = [], []
    for farm, block in df.groupby("_farm_order", sort=True):
        group_rows.append(len(rows))
        rows.append([f"Farm {block['farm'].iloc[0]}"])
        for _, r in block.iterrows():
            process = r["process"]
            if r["process_instance"] > 1:
                process = f"{process} ({r['process_instance']})"
            rows.append(
                [
                    r["module"],
                    process,
                    "" if pd.isna(r["description"]) else r["description"],
                    "" if pd.isna(r["source"]) else r["source"],
                    r["flow"],
                    r["compartment"],
                    fmt(r["amount"]),
                ]
            )

    header = [
        "Module",
        "Process",
        "Description",
        "Source",
        "Flow",
        "Compartment",
        "Amount (kg yr⁻¹)",
    ]
    widths = [3.1, 5.0, 4.4, 2.9, 4.2, 3.6, 2.5]

    doc = new_document(landscape=True)
    add_caption(
        doc,
        TABLE_LABEL_SI,
        "Complete direct-emission inventory of the example farms "
        f"{', '.join(str(f) for f in farms)} as written to the foreground processes "
        "(one row per process × description × flow).",
    )
    build_table(doc, header, rows, widths, 8, {6}, group_rows)

    legend = "; ".join(
        f"{m} = {MODULE_LABELS.get(m, m)}" for m in module_order if m in set(df["module"])
    )
    add_note(
        doc,
        f"Module: {legend}. Description = exchange description in openLCA; "
        "Source = emission type in the calculation script. Compartment = openLCA flow "
        "category without the prefix 'Elementary flows/Emission to'; the exact category "
        "strings are in the machine-readable CSV. "
        f"Generated by run_all.py from {source_name} on {run_date}.",
    )
    return save_docx(doc, path)


def write_tables(result, farms, module_order, output_dir, stamp, source_name):
    """Writes both tables and returns their paths."""
    missing = [f for f in farms if f not in set(result["farm"])]
    if missing:
        print(f"  ⚠️  Beispielbetriebe ohne Emissionen: {missing}")
    farms = [f for f in farms if f not in missing]

    run_date = stamp.split("_")[0]
    paper_path = output_dir / f"table_reproducibility_{stamp}.docx"
    si_path = output_dir / f"table_SI_full_{stamp}.docx"
    paper_path = reproducibility_table(
        result, farms, module_order, paper_path, source_name, run_date
    )
    si_path = full_table(result, farms, module_order, si_path, source_name, run_date)
    return paper_path, si_path
