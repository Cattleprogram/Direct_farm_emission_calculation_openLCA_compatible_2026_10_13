"""Update ONLY the amounts of exchanges that already exist in openLCA from the Access DB.

Nothing is created: no flows, no processes, no exchanges. Direction, provider, costs,
unit, description, category and quantitative reference of an exchange stay untouched.

Matching (Access row <-> openLCA exchange), within the process "<ZirKUH_ID>_<Process>":
    flow name + flow category + description (TierkategorieEmissionFutter) + direction
Flows that were moved from input to output (or vice versa) in openLCA therefore do not
match and are left untouched.

Skipped and reported (never written):
    - Access row without a matching exchange (it is NOT created)
    - several Access rows or several openLCA exchanges for the same key (ambiguous)
    - empty or 0 value in Access
    - exchange with an amount formula (openLCA would overwrite the value)
    - exchange whose unit is not the reference unit of its flow property

Use it after the emissions were recalculated, when the processes already exist in openLCA
(created with integrate_data_to_open_lca.py).

Usage (openLCA must be running with the IPC server on port 8080):
    python update_values_AccessDB_to_openLCA.py                 # dry run, all farms
    python update_values_AccessDB_to_openLCA.py --farms 2 8     # dry run, only ZirKUH_ID 2 and 8
    python update_values_AccessDB_to_openLCA.py --farms 2 --write   # dry run, then ask y/n and write
    python update_values_AccessDB_to_openLCA.py --db path/to.accdb  # other database
"""

import argparse
import math
import re
from datetime import datetime

import olca_ipc as ipc
import olca_schema as o
import pandas as pd
import pyodbc

from integrate_data_to_open_lca import DIESEL_KG_PER_L, DIESEL_KWH_PER_KG
from read_database import QUERY_NAME, REPO_DIR, resolve_db_path

REPORT_DIR = REPO_DIR / "results" / "update_reports"
ZIRKUH_ROOT = "00001_ZirKUH"


# ---------------------------------------------------------------- Access
def read_access(db_path):
    conn = pyodbc.connect(rf"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};", timeout=5)
    try:
        df = pd.read_sql_query(f"SELECT * FROM [{QUERY_NAME}]", conn)
    finally:
        conn.close()
    return df


def process_name(zirkuh_id, process):
    # same naming as integrate_data_to_open_lca.py
    process = process[2:] if process.startswith("P_") else process
    return f"{str(zirkuh_id).zfill(2)}_{process}"


def convert_amount(value, unit, flow_name):
    """Same conversions as integrate_data_to_open_lca.py (create_processes). Returns None for empty values."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    amount = float(str(value).replace(",", "."))
    unit = str(unit).strip().lower() if unit is not None else ""
    if unit == "dt":
        amount = amount * 100  # dt -> kg
    elif unit == "l" and "diesel" in flow_name.lower():
        amount = amount * DIESEL_KWH_PER_KG * DIESEL_KG_PER_L  # l diesel -> kWh
    elif unit == "l":
        amount = amount / 1000  # l -> m3
    elif unit == "t":
        amount = amount * 1000  # t -> kg
    return amount


def normalize_category(category):
    if category is None or (isinstance(category, float) and math.isnan(category)):
        return ""
    return re.sub(r"\s*/\s*", "/", str(category).strip())


def normalize_text(text):
    if text is None or (isinstance(text, float) and math.isnan(text)):
        return ""
    return str(text).strip()


def category_variants(category):
    """An Access FlowCategory matches the flow's original category or the ZirKUH namespace copy of it."""
    c = normalize_category(category)
    return {c, f"{ZIRKUH_ROOT}/{c}"}


# ---------------------------------------------------------------- openLCA
class RefUnits:
    """Caches the reference unit id per flow property."""

    def __init__(self, client):
        self.client = client
        self.cache = {}

    def ref_unit_id(self, flow_property_ref):
        if flow_property_ref is None:
            return None
        fp_id = flow_property_ref.id
        if fp_id not in self.cache:
            fp = self.client.get(o.FlowProperty, fp_id)
            ug = self.client.get(o.UnitGroup, fp.unit_group.id)
            self.cache[fp_id] = next((u.id for u in ug.units if u.is_ref_unit), None)
        return self.cache[fp_id]


def exchange_key(ex):
    return (
        ex.flow.name,
        normalize_category(ex.flow.category),
        normalize_text(ex.description),
        "Input" if ex.is_input else "Output",
    )


# ---------------------------------------------------------------- planning (dry run)
def plan_updates(df, client, farms=None):
    df = df.copy()
    if farms:
        df = df[df["ZirKUH_ID"].astype(int).isin(farms)]
    df["olca_process"] = [process_name(z, p) for z, p in zip(df["ZirKUH_ID"], df["Process"])]
    df["new_amount"] = [convert_amount(v, u, f) for v, u, f in zip(df["Wert"], df["Unit"], df["Flow"])]

    proc_desc = {}
    for d in client.get_descriptors(o.Process):
        proc_desc.setdefault(d.name, []).append(d)

    ref_units = RefUnits(client)
    report, updates = [], {}  # updates: process id -> {internal_id: new amount}
    loaded = {}

    for pname, rows in df.groupby("olca_process", sort=True):
        base = dict(process=pname, zirkuh_id=int(rows["ZirKUH_ID"].iloc[0]))
        descs = proc_desc.get(pname, [])
        if len(descs) != 1:
            status = "skipped: process not in openLCA" if not descs else f"skipped: {len(descs)} processes with this name"
            for _, r in rows.iterrows():
                report.append({**base, "flow": r["Flow"], "flow_category": r["FlowCategory"], "description": normalize_text(r["TierkategorieEmissionFutter"]),
                               "direction": r["Direction"], "old_amount": None, "new_amount": r["new_amount"], "status": status})
            continue

        process = client.get(o.Process, descs[0].id)
        loaded[process.id] = process
        exchanges = {}
        for ex in process.exchanges or []:
            exchanges.setdefault(exchange_key(ex), []).append(ex)

        # Access rows per key
        access_keys = {}
        for _, r in rows.iterrows():
            access_keys.setdefault((r["Flow"], normalize_category(r["FlowCategory"]),
                                    normalize_text(r["TierkategorieEmissionFutter"]), r["Direction"]), []).append(r)

        matched_ex_ids = set()
        for (flow, cat, desc, direction), rs in access_keys.items():
            cands = [ex for c in category_variants(cat) for ex in exchanges.get((flow, c, desc, direction), [])]
            row = dict(base, flow=flow, flow_category=cat, description=desc, direction=direction,
                       old_amount=cands[0].amount if len(cands) == 1 else None,
                       new_amount=rs[0]["new_amount"] if len(rs) == 1 else None)
            matched_ex_ids.update(ex.internal_id for ex in cands)
            if len(rs) > 1:
                row["status"] = f"skipped: {len(rs)} Access rows for this key ({', '.join(str(x['Wert']) for x in rs)})"
            elif not cands:
                row["status"] = "skipped: no matching exchange in openLCA (not created)"
            elif len(cands) > 1:
                row["status"] = f"skipped: {len(cands)} matching exchanges in openLCA (duplicates?)"
            elif row["new_amount"] is None or row["new_amount"] == 0:
                row["status"] = "skipped: empty or 0 in Access"
            elif cands[0].amount_formula:
                row["status"] = f"skipped: exchange has amount formula '{cands[0].amount_formula}'"
            elif cands[0].unit is None or cands[0].unit.id != ref_units.ref_unit_id(cands[0].flow_property):
                row["status"] = f"skipped: exchange unit '{getattr(cands[0].unit, 'name', None)}' is not the reference unit"
            elif math.isclose(cands[0].amount or 0.0, row["new_amount"], rel_tol=1e-9, abs_tol=1e-12):
                row["status"] = "unchanged"
            else:
                row["status"] = "update"
                updates.setdefault(process.id, {})[cands[0].internal_id] = row["new_amount"]
            report.append(row)

        # exchanges in openLCA without an Access row stay untouched - listed for information
        for ex in process.exchanges or []:
            if ex.internal_id not in matched_ex_ids:
                k = exchange_key(ex)
                report.append(dict(base, flow=k[0], flow_category=k[1], description=k[2], direction=k[3],
                                   old_amount=ex.amount, new_amount=None, status="untouched: not in Access"))

    return pd.DataFrame(report), updates, loaded


# ---------------------------------------------------------------- writing
def write_updates(client, updates, loaded):
    results = []
    for pid, amounts in updates.items():
        process = loaded[pid]
        n_before = len(process.exchanges)
        for ex in process.exchanges:
            if ex.internal_id in amounts:
                ex.amount = float(amounts[ex.internal_id])
        ref = client.put(process)  # returns None on failure (does not raise)
        if ref is None:
            results.append((process.name, "FAILED: put() returned None"))
            continue
        check = client.get(o.Process, pid)
        got = {ex.internal_id: ex.amount for ex in check.exchanges}
        wrong = [i for i, a in amounts.items() if not math.isclose(got.get(i, float("nan")), a, rel_tol=1e-9)]
        if wrong:
            results.append((process.name, f"CHECK FAILED: {len(wrong)} amount(s) not stored"))
        elif len(check.exchanges) != n_before:
            results.append((process.name, f"WARNING: exchange count changed {n_before} -> {len(check.exchanges)} (IPC duplication?)"))
        else:
            results.append((process.name, f"OK: {len(amounts)} amount(s) updated"))
    return results


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--farms", type=int, nargs="+", help="ZirKUH_IDs to process (default: all)")
    ap.add_argument("--write", action="store_true", help="after the dry run, ask for confirmation and write")
    ap.add_argument("--port", type=int, default=8080)
    ap.add_argument("--db", help="Access database with the calculated emissions (default: newest "
                                 "data/*_EmissionenBerechnet*.accdb)")
    args = ap.parse_args()

    db_path = resolve_db_path(args.db)
    print(f"Access DB: {db_path}")
    df = read_access(db_path)
    print(f"{len(df)} rows from {QUERY_NAME}")

    client = ipc.Client(args.port)
    report, updates, loaded = plan_updates(df, client, args.farms)

    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    out = REPORT_DIR / f"{datetime.now():%Y_%m_%d_%H%M}_update_report.csv"
    report.to_csv(out, index=False, sep=";", decimal=",", encoding="utf-8-sig")

    # group statuses without their details, e.g. "skipped: 2 Access rows for this key (6400, 1600)" -> "skipped: N Access rows for this key"
    status_group = report["status"].str.replace(r" \(.*$| '.*$", "", regex=True).str.replace(r"\d+ ", "N ", regex=True)
    print("\nSummary (rows):")
    print(status_group.value_counts().to_string())
    upd = report[report["status"] == "update"]
    if not upd.empty:
        print("\nPlanned updates:")
        print(upd[["process", "flow", "direction", "description", "old_amount", "new_amount"]].to_string(index=False))
    print(f"\nFull report: {out}")

    n = sum(len(v) for v in updates.values())
    if not args.write:
        print(f"\nDry run only - {n} exchange(s) in {len(updates)} process(es) would be updated. Use --write to write.")
        return
    if n == 0:
        print("\nNothing to write.")
        return
    answer = input(f"\nWrite {n} amount(s) in {len(updates)} process(es) to openLCA? (y/n): ").strip().lower()
    if answer not in ("y", "j"):
        print("Aborted, nothing written.")
        return
    for name, msg in write_updates(client, updates, loaded):
        print(f"  {name}: {msg}")


if __name__ == "__main__":
    main()
