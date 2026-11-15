"""Create (or update) the farm flows and processes in openLCA from the Access database.

Reads the saved Access query ★_python_qry_Processes_Flows (all input/output flows per farm
process, including the emissions calculated by the scripts 1-5) and writes them to openLCA
via the IPC server:

    1. flows: an existing flow with the same name and category is reused (e.g. ecoinvent
       product flows and the elementary flows of the impact assessment method). Only flows
       that exist nowhere are created, under the category "00001_ZirKUH/<category>".
    2. processes: one unit process per farm process, named "<ZirKUH_ID>_<Process>" (the
       prefix "P_" of the Access process name is dropped), category
       "00001_ZirKUH/<Betriebsform>/Hof_<ZirKUH_ID>/<ProcessCategory>". Existing processes are
       updated: the amount of an existing exchange (same flow, flow category and description)
       is overwritten, missing exchanges are added.

NOT done by this script (manual modelling in openLCA): the reference flow of new processes
(see the comment in create_or_update_process), allocation, provider linking, product systems.

Requirements: openLCA 2.6 running with the IPC server (Tools > Developer tools > IPC server),
a database that contains the background data (ecoinvent) and the impact assessment method.

Usage:
    python integrate_data_to_open_lca.py                  # newest data/*_EmissionenBerechnet*.accdb
    python integrate_data_to_open_lca.py --db path/to.accdb --port 8080
"""

import argparse
import re

import olca_ipc as ipc
import olca_schema as o
from tqdm import tqdm

from read_database import query_db_data, resolve_db_path

CLIENT = None  # olca_ipc.Client, set in main()

# Access unit -> openLCA flow property of new flows (the amount is written in the reference unit
# of that property, see the conversions in create_processes)
UNIT_TO_PROPERY = {
    "kg": "Mass",
    "m3": "Volume",
    "h": "Time",
    "kwh": "Energy",
    "ha": "Area",
    "dt": "Mass",
    "t": "Mass",
    "l": "Volume",
}

# diesel: litre -> kWh with the lower heating value 11.8 kWh/kg and the density 0.83 kg/l
DIESEL_KWH_PER_KG = 11.8
DIESEL_KG_PER_L = 0.83

created_elementary = []  # elementary flows that had to be created (no characterisation factors!)


def create_flows(column_names, rows):
    flow_name_idx = column_names.index("Flow")
    unit_idx = column_names.index("Unit")
    flow_category_idx = column_names.index("FlowCategory")
    flow_type_idx = column_names.index("Flow type")
    unique_flows = set()
    for row in rows:
        flow_name = row[flow_name_idx]
        unit = row[unit_idx]
        flow_category = row[flow_category_idx]
        flow_type = row[flow_type_idx]
        if unit is None:
            continue
        unit = str(unit).strip()
        normalized_unit = unit.lower()
        if "kg" in normalized_unit:
            normalized_unit = "kg"
        unique_flows.add((flow_name, normalized_unit, flow_category, flow_type))

    for flow in unique_flows:
        create_flow_in_lca_if_not_exists(*flow)


def resolve_flow_type(flow_type):
    if flow_type == "product":
        return o.FlowType.PRODUCT_FLOW
    elif flow_type == "waste":
        return o.FlowType.WASTE_FLOW
    else:
        return o.FlowType.ELEMENTARY_FLOW


def normalize_category(category):
    """openLCA joins category levels with a bare '/' (e.g. "Elementary flows/Emission to
    soil/agricultural"), the Access FlowCategory values use " / ". Remove the whitespace
    around '/' so that both sides compare equal."""
    return re.sub(r"\s*/\s*", "/", category.strip())


def flow_category_path(category):
    return f"00001_ZirKUH/{normalize_category(category)}"


def find_flow(flow_name, category):
    """Look up a flow by name, preferring its original category (as given by Access), so that
    existing background flows (ecoinvent product flows, elementary flows) are reused. This
    matters for elementary flows in particular: the characterisation factors of the impact
    assessment method are linked to those flow IDs. Only flows that match nowhere are new and
    live under the ZirKUH namespace, so a second lookup checks there."""
    category = normalize_category(category)
    candidates = [d for d in CLIENT.get_descriptors(o.Flow) if d.name == flow_name]
    if not candidates:
        return None
    real_match = next(
        (d for d in candidates if d.category and normalize_category(d.category) == category),
        None,
    )
    if real_match:
        return real_match
    zirkuh_category = flow_category_path(category)
    return next(
        (d for d in candidates if d.category and normalize_category(d.category) == zirkuh_category),
        None,
    )


def create_flow_in_lca_if_not_exists(flow_name, unit, category, flow_type):
    desired_flow_type = resolve_flow_type(flow_type)

    existing_desc = find_flow(flow_name, category)
    if existing_desc:
        existing_flow = CLIENT.get(o.Flow, existing_desc.id)
        if existing_flow.flow_type != desired_flow_type:
            print(
                f"Flow '{flow_name}' ({existing_flow.category}) exists with type {existing_flow.flow_type}, "
                f"updating to {desired_flow_type}."
            )
            existing_flow.flow_type = desired_flow_type
            CLIENT.put(existing_flow)
        else:
            print(f"Flow '{flow_name}' ({existing_flow.category}) already exists in LCA, skipping creation.")
        return

    print(f"Creating flow '{flow_name}' with unit '{unit}' and category '{category}' in LCA...")
    if desired_flow_type == o.FlowType.ELEMENTARY_FLOW:
        created_elementary.append((flow_name, category))
    flow = o.Flow()
    flow.name = flow_name
    flow.category = flow_category_path(category)
    flow.flow_type = desired_flow_type
    prop = None
    property_name = UNIT_TO_PROPERY.get(unit)
    if property_name is None:
        raise ValueError(f"Unsupported unit '{unit}' for flow '{flow_name}'")
    for d in CLIENT.get_descriptors(o.FlowProperty):
        if d.name == property_name:
            prop = CLIENT.get(o.FlowProperty, d.id)
            break
    if prop is None:
        raise ValueError(f"Flow property for unit '{unit}' not found")

    ref = o.FlowPropertyFactor()
    ref.flow_property = prop
    ref.conversion_factor = 1.0
    ref.is_ref_flow_property = True
    flow.flow_properties = [ref]
    CLIENT.put(flow)


def create_processes(column_names, rows):
    process_name_idx = column_names.index("Process")
    zirkuh_id_idx = column_names.index("ZirKUH_ID")
    tierkategorie_idx = column_names.index("TierkategorieEmissionFutter")
    flow_category_idx = column_names.index("FlowCategory")
    flow_type_idx = column_names.index("Flow type")
    processes = {}
    flow_tierkategorien = {}  # (process, flow, flow category) -> set of TierkategorieEmissionFutter values
    first_output_type = {}  # process -> flow type of its first output (= quantitative reference)

    for row in rows:
        hof_id = str(row[zirkuh_id_idx]).zfill(2)
        process_name = row[process_name_idx]
        process_name = process_name[2:] if process_name.startswith("P_") else process_name
        process_name = f"{hof_id}_{process_name}"
        flow_name = row[column_names.index("Flow")]
        direction = row[column_names.index("Direction")]
        amount = row[column_names.index("Wert")]
        unit = row[column_names.index("Unit")].lower()
        process_form = row[column_names.index("Betriebsform")]
        tierkategorie = row[tierkategorie_idx]
        flow_category = row[flow_category_idx]
        category = f"00001_ZirKUH/{process_form}/Hof_{hof_id}/{row[column_names.index('ProcessCategory')]}"

        if amount is None:
            continue
        amount = float(str(amount).replace(",", "."))  # comma as decimal separator
        if unit == "dt":
            amount = amount * 100  # dt -> kg
        elif unit == "l" and "diesel" in flow_name.lower():
            amount = amount * DIESEL_KWH_PER_KG * DIESEL_KG_PER_L  # l diesel -> kWh
        elif unit == "l":
            amount = amount / 1000  # l -> m3
        elif unit == "t":
            amount = amount * 1000  # t -> kg

        if process_name not in processes:
            processes[process_name] = {"inputs": [], "outputs": [], "category": category}

        if direction == "Output":
            processes[process_name]["outputs"].append((flow_name, amount, tierkategorie, flow_category))
            first_output_type.setdefault(process_name, row[flow_type_idx])
        else:
            processes[process_name]["inputs"].append((flow_name, amount, tierkategorie, flow_category))

        flow_tierkategorien.setdefault((process_name, flow_name, flow_category), set()).add(tierkategorie)

    affected = {k: v for k, v in flow_tierkategorien.items() if len(v) > 1}
    if affected:
        print(
            f"\nWARNING: {len(affected)} process/flow/category combination(s) have multiple distinct "
            "'TierkategorieEmissionFutter' values. If these processes already exist in openLCA "
            "from a previous run (before exchanges were disambiguated by description), check them "
            "for stale/duplicate exchanges after this run:\n"
        )
        for (process_name, flow_name, flow_category), values in sorted(
            affected.items(), key=lambda kv: (kv[0][0], kv[0][1], str(kv[0][2]))
        ):
            print(f"  - Process '{process_name}', Flow '{flow_name}' [{flow_category}]: "
                  f"{sorted(v for v in values if v is not None)}")
        print()

    # The reference flow of a NEW process is simply its first output row. The Access query has no
    # ORDER BY, so this is not a modelling decision: set the reference flow manually in openLCA.
    elementary_ref = sorted(p for p, t in first_output_type.items() if str(t).lower().startswith("elementary"))
    if elementary_ref:
        print(f"\nWARNING: in {len(elementary_ref)} process(es) the first output is an ELEMENTARY flow and would "
              "become the reference flow of a new process. Set the reference flow manually in openLCA:")
        for p in elementary_ref:
            print(f"  - {p}")
        print()

    for process_name, data in tqdm(processes.items(), desc="Processing processes"):
        create_or_update_process(
            process_name, inputs=data["inputs"], outputs=data["outputs"], category=data["category"]
        )


def create_or_update_process(name, inputs=None, outputs=None, category=None):
    inputs = inputs or []
    outputs = outputs or []

    existing_ref = CLIENT.find(o.Process, name)

    if existing_ref:
        print(f"Process '{name}' exists → updating")
        process = CLIENT.get(o.Process, existing_ref.id)
    else:
        print(f"Process '{name}' does not exist → creating")
        process = o.Process()
        process.name = name
        process.process_type = o.ProcessType.UNIT_PROCESS
        process.exchanges = []

    process.category = category if category else "00001_ZirKUH/Uncategorized"

    # Existing exchanges by (flow name, flow category, description), so that rows with the same
    # flow name but another emission compartment or TierkategorieEmissionFutter do not collide.
    existing_exchanges = {(ex.flow.name, ex.flow.category, ex.description): ex for ex in process.exchanges}

    def get_or_create_exchange(flow_name, amount, is_input, flow_category, description=None, quantitative=False):
        flow = find_flow(flow_name, flow_category)
        if not flow:
            raise ValueError(
                f"Flow not found: {flow_name} [{flow_category}] — run create_flows first "
                "so it gets created under the ZirKUH namespace."
            )

        flow_full = CLIENT.get(o.Flow, flow.id)
        ref_factor = flow_full.flow_properties[0]

        key = (flow_name, flow.category, description)
        if key in existing_exchanges:
            ex = existing_exchanges[key]
            original_amount = ex.amount
            ex.amount = float(amount)
            if original_amount != ex.amount:
                print(f"Updated {'input' if is_input else 'output'} flow '{flow_name}' [{flow_category}] "
                      f"({description}): {original_amount} → {amount}")
            else:
                print(f"No change for {'input' if is_input else 'output'} flow '{flow_name}' [{flow_category}] "
                      f"({description}).")
            return None
        ex = o.Exchange()
        ex.is_input = is_input
        ex.amount = float(amount)
        ex.flow = flow
        ex.description = description
        flow_property = CLIENT.get(o.FlowProperty, ref_factor.flow_property.id)
        unit_group = CLIENT.get(o.UnitGroup, flow_property.unit_group.id)
        ref_unit = next(u for u in unit_group.units if u.is_ref_unit)
        ex.unit = o.Ref(ref_unit.id, name=ref_unit.name, ref_type=o.RefType.Unit)
        ex.flow_property = ref_factor.flow_property
        ex.is_quantitative_reference = quantitative
        print(f"Added flow '{flow_name}' [{flow_category}] ({description}) → {amount} to process '{name}'")
        return ex

    new_exchanges = []
    for flow_name, amount, tierkategorie, flow_category in inputs:
        ex = get_or_create_exchange(flow_name, amount, True, flow_category, description=tierkategorie)
        if ex:
            new_exchanges.append(ex)
    # ------------------------------------------------------------------------------------------
    # REFERENCE FLOW: NOT DEFINED BY THIS SCRIPT - SET IT MANUALLY IN openLCA
    # The first output exchange of a NEW process is marked as quantitative reference only so
    # that openLCA accepts the process. Which output comes first depends on the row order of
    # the Access query (no ORDER BY) and can be a co-product, manure or even an emission.
    # After the transfer, check and set the reference flow of every new process by hand (as
    # well as allocation, provider linking and the product systems). Existing processes keep
    # their reference flow, because only new exchanges get this flag.
    # ------------------------------------------------------------------------------------------
    for i, (flow_name, amount, tierkategorie, flow_category) in enumerate(outputs):
        ex = get_or_create_exchange(
            flow_name, amount, False, flow_category, description=tierkategorie, quantitative=(i == 0)
        )
        if ex:
            new_exchanges.append(ex)

    process.exchanges.extend(new_exchanges)
    CLIENT.put(process)
    return process


def main():
    global CLIENT
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--db", help="Access database with the calculated emissions (default: newest "
                                 "data/*_EmissionenBerechnet*.accdb)")
    ap.add_argument("--port", type=int, default=8080, help="port of the openLCA IPC server (default 8080)")
    args = ap.parse_args()

    CLIENT = ipc.Client(args.port)
    column_names, rows = query_db_data(resolve_db_path(args.db))
    create_flows(column_names, rows)
    create_processes(column_names, rows)

    if created_elementary:
        print(f"\nWARNING: {len(created_elementary)} elementary flow(s) were not found in openLCA and were "
              "CREATED under 00001_ZirKUH. They have no characterisation factors, so the impact assessment "
              "ignores them. Map them to the elementary flows of your method:")
        for flow_name, category in sorted(created_elementary):
            print(f"  - {flow_name} [{category}]")


if __name__ == "__main__":
    main()
