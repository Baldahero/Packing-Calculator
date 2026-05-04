import math
from dataclasses import dataclass
from io import BytesIO
from typing import Dict, List

import pandas as pd
import streamlit as st


# ============================================================
# SETTINGS
# ============================================================
MAX_GLAZED_HEIGHT = 2700
MAX_PACKING_HEIGHT = 2700
MAX_CONSTRUCTION_HEIGHT = 6600  # absolute max pallet size
MAX_PALLET_WEIGHT_KG = 1000.0
MAX_ITEMS_PER_PALLET = 6
MAX_ITEMS_PER_PALLET_HEAVY = 2  # for sliding/folding types

GLASS_BOX_PRICE_EUR = 180.0
GLASS_BOX_MAX_WEIGHT_KG = 1000.0
GLASS_PALLET_WIDTH_MM = 1200
TRUCK_WIDTH_M = 2.0

# Types with special glazing rule: glazed only if height <= 2700 and weight <= 1000 kg
# Also limited to MAX_ITEMS_PER_PALLET_HEAVY per pallet
HEAVY_GLAZING_TYPES = {
    "double sliding door",
    "triple sliding door",
    "2-leaf+2-fixed sliding door",
    "folding door",
}

# Facades: glass always packed separately regardless of height/weight
FACADE_TYPES = {
    "facade",
}



# ============================================================
# DATA MODEL
# ============================================================
@dataclass
class Construction:
    item_name: str
    item_type: str
    width_mm: float
    height_mm: float
    qty: int
    weight_kg: float
    glass_mode: str = "Glazed"  # "Glazed" | "Unglazed" | "Without glass"
    glass_weight_kg: float = 0.0


# ============================================================
# PACKING RULES
# ============================================================
def min_pallet_width_by_height(height_mm: float) -> int:
    if height_mm <= 1000:
        return 400
    if height_mm <= 2000:
        return 800
    if height_mm <= MAX_PACKING_HEIGHT:
        return 1200
    return 0


def round_up_pallet_width(size_mm: float) -> int:
    if size_mm <= 1000:
        return 1000
    if size_mm <= 1500:
        return 1500
    if size_mm <= 2500:
        return 2500
    if size_mm <= 3500:
        return 3500
    if size_mm <= 5000:
        return 5000
    if size_mm <= 6000:
        return 6000
    if size_mm <= 6600:
        return 6600
    return int(math.ceil(size_mm / 1000.0) * 1000)


def real_pallet_width(width_mm: float, height_mm: float) -> float:
    """Physical pallet width: height if packed sideways, else width+100."""
    if height_mm > MAX_GLAZED_HEIGHT:
        return height_mm
    return width_mm + 100


def pallet_price_eur(width_mm: float) -> float:
    w = float(width_mm or 0)
    if w <= 1000:
        return 31.0
    if w <= 1500:
        return 44.0
    if w <= 2500:
        return 60.0
    if w <= 3500:
        return 72.0
    if w <= 5000:
        return 94.0
    if w <= 6000:
        return 120.0
    if w <= 6600:
        return 120.0
    return 120.0


def ldm_from_width(width_mm: float, count: int = 1) -> float:
    return (float(width_mm) * float(count)) / TRUCK_WIDTH_M / 1000.0


def calculate_construction(construction: Construction) -> Dict[str, object]:
    real_width = real_pallet_width(construction.width_mm, construction.height_mm)
    packed_sideways = construction.height_mm > MAX_GLAZED_HEIGHT
    is_heavy_type = construction.item_type.lower() in HEAVY_GLAZING_TYPES
    is_facade = construction.item_type.lower() in FACADE_TYPES
    mode = construction.glass_mode  # "Glazed" | "Unglazed" | "Without glass"

    if construction.height_mm > MAX_CONSTRUCTION_HEIGHT:
        return {
            "Item": construction.item_name,
            "Type": construction.item_type,
            "Width (mm)": float(construction.width_mm),
            "Height (mm)": float(construction.height_mm),
            "Qty": int(construction.qty),
            "Unit weight (kg)": float(construction.weight_kg),
            "Glass weight (kg)": float(construction.glass_weight_kg),
            "Glass mode": mode,
            "Packed as": "NOT POSSIBLE",
            "Glass separate": "N/A",
            "Packed sideways": "N/A",
            "Max per pallet": "N/A",
            "Pallet width (mm)": "N/A",
            "Notes": "Construction size exceeds current pallet pricing ranges",
        }

    packed_as = "UNGLAZED"
    glass_separate = "NO"
    notes = "Packed without glass"

    if mode == "Without glass":
        # glass is not ours — no glass box, just pack the frame
        packed_as = "UNGLAZED"
        glass_separate = "NO"
        notes = "Without glass — frame only"

    elif mode == "Unglazed":
        # glass travels separately → glass box needed
        packed_as = "UNGLAZED"
        glass_separate = "YES"
        notes = "Glass packed separately"
        if packed_sideways:
            notes = "Glass packed separately; construction packed sideways"

    elif mode == "Glazed":
        # glass travels with frame — check if any rule forces separation
        if is_facade:
            packed_as = "UNGLAZED"
            glass_separate = "YES"
            notes = "Facade — glass always packed separately"
        elif packed_sideways:
            packed_as = "UNGLAZED"
            glass_separate = "YES"
            notes = "Glass must be packed separately; construction packed sideways"
        elif is_heavy_type and construction.weight_kg > MAX_PALLET_WEIGHT_KG:
            packed_as = "UNGLAZED"
            glass_separate = "YES"
            notes = f"Weight exceeds {MAX_PALLET_WEIGHT_KG:.0f} kg — packed without glass; glass packed separately"
        else:
            packed_as = "GLAZED"
            notes = "Can be packed with glass"

    if mode != "Glazed" and packed_sideways and glass_separate == "NO":
        notes += "; construction packed sideways"

    max_per_pallet = MAX_ITEMS_PER_PALLET_HEAVY if is_heavy_type else MAX_ITEMS_PER_PALLET

    # Always store glass weight for visibility; it's used for pallet weight when glazed together
    stored_glass_weight = float(construction.glass_weight_kg) if construction.glass_mode != "Without glass" else 0.0

    return {
        "Item": construction.item_name,
        "Type": construction.item_type,
        "Width (mm)": float(construction.width_mm),
        "Height (mm)": float(construction.height_mm),
        "Qty": int(construction.qty),
        "Unit weight (kg)": float(construction.weight_kg),
        "Glass weight (kg)": stored_glass_weight,
        "Glass mode": mode,
        "Packed as": packed_as,
        "Glass separate": glass_separate,
        "Packed sideways": "YES" if packed_sideways else "NO",
        "Max per pallet": int(max_per_pallet),
        "Pallet width (mm)": int(real_width),
        "Notes": notes,
    }


# ============================================================
# PALLET PACKING
# ============================================================
def expand_by_qty(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for _, row in df.iterrows():
        qty = int(row["Qty"])
        for unit_idx in range(1, qty + 1):
            rr = row.copy()
            rr["Qty"] = 1
            rr["Unit idx"] = unit_idx
            rows.append(rr)
    return pd.DataFrame(rows)


def pack_mixed(units: pd.DataFrame) -> List[Dict[str, object]]:
    if units.empty:
        return []

    u = units.copy()
    u["Unit weight (kg)"] = pd.to_numeric(u["Unit weight (kg)"], errors="coerce").fillna(0.0)
    u["Glass weight (kg)"] = pd.to_numeric(u.get("Glass weight (kg)", 0), errors="coerce").fillna(0.0)

    # Total weight on pallet = frame + glass if glass is packed together, else frame only
    u["_total_weight"] = u.apply(
        lambda r: r["Unit weight (kg)"] + r["Glass weight (kg)"]
        if r.get("Glass separate", "NO") == "NO" and r.get("Glass mode", "Without glass") == "Glazed"
        else r["Unit weight (kg)"],
        axis=1,
    )

    u = u.sort_values("_total_weight", ascending=False).reset_index(drop=True)

    pallets: List[Dict[str, object]] = []

    for _, item in u.iterrows():
        w = float(item["_total_weight"])
        try:
            item_max = int(item["Max per pallet"])
        except (KeyError, ValueError, TypeError):
            item_max = MAX_ITEMS_PER_PALLET
        placed = False

        for pallet in pallets:
            pallet_max = min(pallet["max_per_pallet"], item_max)
            if (
                pallet["weight_kg"] + w <= MAX_PALLET_WEIGHT_KG
                and pallet["items_count"] + 1 <= pallet_max
            ):
                pallet["weight_kg"] += w
                pallet["items_count"] += 1
                pallet["max_per_pallet"] = pallet_max
                pallet["items"].append(item)
                placed = True
                break

        if not placed:
            pallets.append(
                {
                    "weight_kg": w,
                    "items_count": 1,
                    "max_per_pallet": item_max,
                    "items": [item],
                }
            )

    return pallets

    return pallets


def _get_pallet_width(item) -> float:
    """Get pallet width from item, handling different column name versions."""
    for col in ("Pallet width (mm)", "Real pallet width (mm)"):
        try:
            v = item[col]
            f = float(v)
            if not math.isnan(f):
                return f
        except (KeyError, ValueError, TypeError):
            pass
    try:
        return real_pallet_width(float(item["Width (mm)"]), float(item["Height (mm)"]))
    except Exception:
        return 1200.0


def build_pallet_outputs(results_df: pd.DataFrame):
    valid_df = results_df[results_df["Packed as"] != "NOT POSSIBLE"].copy()
    if valid_df.empty:
        return pd.DataFrame(), pd.DataFrame(), 0.0, 0.0

    units = expand_by_qty(valid_df)
    pallets = pack_mixed(units)

    pallet_summary_rows = []
    plan_rows = []
    total_pallet_cost = 0.0
    total_pallet_ldm = 0.0

    for i, pallet in enumerate(pallets, start=1):
        items_df = pd.DataFrame(pallet["items"])

        pallet_real_width = float(items_df.apply(_get_pallet_width, axis=1).max())
        pallet_req_width = round_up_pallet_width(pallet_real_width)
        pallet_price = pallet_price_eur(pallet_req_width)
        pallet_ldm = ldm_from_width(pallet_real_width, 1)

        total_pallet_cost += pallet_price
        total_pallet_ldm += pallet_ldm

        pallet_summary_rows.append(
            {
                "Pallet no": i,
                "Pallet weight (kg)": round(float(pallet["weight_kg"]), 2),  # frame + glass if glazed together
                "Constructions count": int(items_df["Item"].nunique()),
                "Units count": int(len(items_df)),
                "Pallet width (mm)": round(pallet_real_width, 1),
                "Pallet price (EUR)": pallet_price,
                "Pallet LDM": round(pallet_ldm, 3),
            }
        )

        for _, item in items_df.iterrows():
            plan_rows.append(
                {
                    "Pallet no": i,
                    "Item": item["Item"],
                    "Type": item["Type"],
                    "Width (mm)": item["Width (mm)"],
                    "Height (mm)": item["Height (mm)"],
                    "Unit weight (kg)": item["Unit weight (kg)"],
                    "Packed as": item["Packed as"],
                    "Glass separate": item["Glass separate"],
                    "Packed sideways": item["Packed sideways"],
                    "Pallet width (mm)": _get_pallet_width(item),
                    "Unit idx": item["Unit idx"],
                }
            )

    pallet_summary_df = pd.DataFrame(pallet_summary_rows)
    plan_df = pd.DataFrame(plan_rows)
    return pallet_summary_df, plan_df, total_pallet_cost, total_pallet_ldm


def add_glass_to_pallet_summary(pallet_summary_df: pd.DataFrame, glass_boxes: int, glass_cost: float, glass_ldm: float, total_glass_weight: float) -> pd.DataFrame:
    """Append glass box rows to pallet summary."""
    if glass_boxes <= 0:
        return pallet_summary_df
    last_pallet_no = int(pallet_summary_df["Pallet no"].max()) if not pallet_summary_df.empty else 0
    glass_rows = []
    for i in range(glass_boxes):
        glass_rows.append({
            "Pallet no": last_pallet_no + i + 1,
            "Pallet weight (kg)": round(total_glass_weight / glass_boxes, 2),
            "Constructions count": "-",
            "Units count": "-",
            "Pallet width (mm)": 1200,
            "Pallet price (EUR)": glass_cost / glass_boxes,
            "Pallet LDM": round(glass_ldm / glass_boxes, 3),
            "Note": "GLASS BOX",
        })
    glass_df = pd.DataFrame(glass_rows)
    return pd.concat([pallet_summary_df, glass_df], ignore_index=True)

# ============================================================
# GLASS BOXES
# ============================================================
def calculate_glass_boxes(results_df: pd.DataFrame):
    if results_df.empty:
        return 0, 0.0, 0.0, 0.0

    separate_glass_df = results_df[results_df["Glass separate"] == "YES"].copy()
    if separate_glass_df.empty:
        return 0, 0.0, 0.0, 0.0

    # Use glass weight if provided, fallback to unit weight if glass weight is 0
    glass_w = pd.to_numeric(separate_glass_df.get("Glass weight (kg)", 0), errors="coerce").fillna(0.0)
    unit_w = pd.to_numeric(separate_glass_df["Unit weight (kg)"], errors="coerce").fillna(0.0)
    effective_glass_w = glass_w.where(glass_w > 0, unit_w)

    total_glass_weight = float(
        (
            effective_glass_w
            * pd.to_numeric(separate_glass_df["Qty"], errors="coerce").fillna(0)
        ).sum()
    )

    if total_glass_weight <= 0:
        return 0, 0.0, 0.0, 0.0

    glass_boxes = int(math.ceil(total_glass_weight / GLASS_BOX_MAX_WEIGHT_KG))
    glass_cost = glass_boxes * GLASS_BOX_PRICE_EUR
    glass_ldm = ldm_from_width(GLASS_PALLET_WIDTH_MM, glass_boxes)

    return glass_boxes, total_glass_weight, glass_cost, glass_ldm


# ============================================================
# EXPORT
# ============================================================
def make_excel_file(results_df, pallet_summary_df, plan_df, kpi_df) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        kpi_df.to_excel(writer, sheet_name="Summary", index=False)
        results_df.to_excel(writer, sheet_name="Constructions", index=False)
        pallet_summary_df.to_excel(writer, sheet_name="Pallet Summary", index=False)
        plan_df.to_excel(writer, sheet_name="Packing Plan", index=False)
    return output.getvalue()


# ============================================================
# SESSION HELPERS
# ============================================================
def add_result_to_session(result: Dict[str, object]) -> None:
    if "results" not in st.session_state:
        st.session_state.results = []
    st.session_state.results.append(result)


def clear_results() -> None:
    st.session_state.results = []


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="Packing Calculator", layout="wide")

header_left, header_right = st.columns([1, 4])
with header_left:
    try:
        st.image("nordan_logo1.png", use_container_width=True)
    except Exception:
        st.markdown("**NorDan**")

with header_right:
    st.title("Packing Calculator Pre-Alfa Version")
    st.caption("Manual packing calculation for constructions")

# with st.expander("Rules used", expanded=False):
#     st.markdown(...)  # rules hidden — uncomment to restore


if "results" not in st.session_state:
    st.session_state.results = []
if "project_name" not in st.session_state:
    st.session_state.project_name = "packing_calculation"
if "edit_idx" not in st.session_state:
    st.session_state.edit_idx = None

left, right = st.columns([1, 1])

TYPES = [
    "Door", "Window", "Fixed Window",
    "Double Sliding Door", "Triple Sliding Door",
    "2-leaf+2-fixed Sliding Door", "Folding Door",
    "Door + Sidelight", "Window + Sidelight", "Facade",
]

with left:
    # Prefill values from edit state if editing
    _e = st.session_state.edit_idx
    if _e is not None:
        _r = st.session_state.results[_e]
        _def_name     = _r.get("Item", "...")
        _def_type     = _r.get("Type", TYPES[0])
        _def_width    = float(_r.get("Width (mm)", 1000.0))
        _def_height   = float(_r.get("Height (mm)", 1000.0))
        _def_qty      = int(_r.get("Qty", 1))
        _def_weight   = float(_r.get("Unit weight (kg)", 0.0))
        _def_mode     = _r.get("Glass mode", "Glazed")
        _def_glass_w  = float(_r.get("Glass weight (kg)", 0.0))
        _title = f"✏️ Edit construction: {_def_name}"
    else:
        _def_name, _def_type, _def_width, _def_height = "...", TYPES[0], 1000.0, 1000.0
        _def_qty, _def_weight, _def_mode, _def_glass_w = 1, 0.0, "Glazed", 0.0
        _title = "Add construction"

    st.subheader(_title)

    with st.form(f"packing_form_{_e if _e is not None else 'new'}"):
        item_name = st.text_input("Item name", value=_def_name)

        item_type = st.selectbox("Type", TYPES,
            index=TYPES.index(_def_type) if _def_type in TYPES else 0)

        width_mm  = st.number_input("Width (mm)",  min_value=1.0,  value=_def_width,  step=1.0)
        height_mm = st.number_input("Height (mm)", min_value=1.0,  value=_def_height, step=1.0)
        qty       = st.number_input("Quantity",    min_value=1,    value=_def_qty,    step=1)
        weight_kg = st.number_input("Unit weight (kg)", min_value=0.0, value=_def_weight, step=0.01)

        glass_mode = st.selectbox("Glass", ["Glazed", "Unglazed", "Without glass"],
            index=["Glazed", "Unglazed", "Without glass"].index(_def_mode) if _def_mode in ["Glazed", "Unglazed", "Without glass"] else 0,
            help="Glazed = glass travels with frame | Unglazed = glass goes separately (glass box) | Without glass = no glass at all",
        )

        glass_weight_kg = st.number_input(
            "Glass weight (kg)",
            min_value=0.0, value=_def_glass_w, step=0.01,
            help="Weight of glass only — required when glass is packed separately",
        )

        _btn_label = "Save changes" if _e is not None else "Calculate and add"
        submitted = st.form_submit_button(_btn_label)

        if submitted:
            if weight_kg <= 0:
                st.warning("⚠️ Unit weight is 0 — please enter the actual weight before adding.")
            elif glass_mode in ("Glazed", "Unglazed") and glass_weight_kg <= 0:
                st.warning("⚠️ Glass weight is 0 — please enter the glass weight before adding.")
            else:
                construction = Construction(
                    item_name=item_name.strip() or "Unnamed",
                    item_type=item_type,
                    width_mm=float(width_mm),
                    height_mm=float(height_mm),
                    qty=int(qty),
                    weight_kg=float(weight_kg),
                    glass_mode=glass_mode,
                    glass_weight_kg=float(glass_weight_kg) if glass_mode != "Without glass" else 0.0,
                )
                result = calculate_construction(construction)
                if _e is not None:
                    st.session_state.results[_e] = result
                    st.session_state.edit_idx = None
                    st.success(f"Updated: {result['Item']}")
                else:
                    add_result_to_session(result)
                    st.success(f"Added: {result['Item']}")
                st.rerun()

with right:
    # ---- Order statistics ----
    if st.session_state.results:
        st.subheader("Order statistics")
        stats_df = pd.DataFrame(st.session_state.results)

        total_units = int(pd.to_numeric(stats_df["Qty"], errors="coerce").fillna(0).sum())
        total_weight = float(
            (pd.to_numeric(stats_df["Unit weight (kg)"], errors="coerce").fillna(0)
             * pd.to_numeric(stats_df["Qty"], errors="coerce").fillna(0)).sum()
        )
        glazed_units = int(
            pd.to_numeric(stats_df.loc[stats_df["Packed as"] == "GLAZED", "Qty"], errors="coerce").fillna(0).sum()
        )
        unglazed_units = int(
            pd.to_numeric(stats_df.loc[stats_df["Packed as"] == "UNGLAZED", "Qty"], errors="coerce").fillna(0).sum()
        )
        glass_separate_units = int(
            pd.to_numeric(stats_df.loc[stats_df["Glass separate"] == "YES", "Qty"], errors="coerce").fillna(0).sum()
        )
        sideways_units = int(
            pd.to_numeric(stats_df.loc[stats_df["Packed sideways"] == "YES", "Qty"], errors="coerce").fillna(0).sum()
        )
        _psdf, _, _, _ = build_pallet_outputs(stats_df)
        est_pallets = int(len(_psdf))

        s1, s2 = st.columns(2)
        s1.metric("Total units", total_units)
        s2.metric("Total weight", f"{total_weight:,.0f} kg")

        s3, s4 = st.columns(2)
        s3.metric("Glazed units", glazed_units)
        s4.metric("Unglazed units", unglazed_units)

        s5, s6 = st.columns(2)
        s5.metric("Glass separate", glass_separate_units)
        s6.metric("Packed sideways", sideways_units)

        st.metric("Estimated pallets", est_pallets)

        if sideways_units > 0:
            st.warning(f"⚠️ {sideways_units} unit(s) will be packed sideways (height > 2700 mm)")
        if glass_separate_units > 0:
            st.info(f"ℹ️ {glass_separate_units} unit(s) require separate glass boxes")

        st.divider()

    # ---- Preview ----
    st.subheader("Preview")

    preview = Construction(
        item_name=item_name.strip() or "Unnamed",
        item_type=item_type,
        width_mm=float(width_mm),
        height_mm=float(height_mm),
        qty=int(qty),
        weight_kg=float(weight_kg),
        glass_mode=glass_mode,
        glass_weight_kg=float(glass_weight_kg) if glass_mode != "Without glass" else 0.0,
    )

    preview_df = pd.DataFrame([calculate_construction(preview)])
    st.dataframe(preview_df, use_container_width=True)

st.divider()
st.subheader("Constructions")

if st.session_state.results:
    results_df = pd.DataFrame(st.session_state.results)
    st.dataframe(results_df, use_container_width=True)

    st.markdown("### Edit / Remove item")
    col_sel, col_edit, col_del = st.columns([6, 1, 1])

    with col_sel:
        item_to_manage = st.selectbox(
            "Select item",
            options=list(range(len(results_df))),
            format_func=lambda x: f"{results_df.iloc[x]['Item']} (row {x})",
            label_visibility="collapsed",
        )

    with col_edit:
        if st.button("✏️ Edit", use_container_width=True):
            st.session_state.edit_idx = item_to_manage
            st.rerun()

    with col_del:
        if st.button("🗑️ Delete", use_container_width=True):
            if st.session_state.edit_idx == item_to_manage:
                st.session_state.edit_idx = None
            st.session_state.results.pop(item_to_manage)
            st.rerun()

    pallet_summary_df, plan_df, total_pallet_cost, total_pallet_ldm = build_pallet_outputs(results_df)
    glass_boxes, total_glass_weight, glass_cost, glass_ldm = calculate_glass_boxes(results_df)
    total_packaging_cost = total_pallet_cost + glass_cost
    total_ldm = total_pallet_ldm + glass_ldm

    # Add glass boxes to pallet summary
    pallet_summary_with_glass_df = add_glass_to_pallet_summary(
        pallet_summary_df, glass_boxes, glass_cost, glass_ldm, total_glass_weight
    )

    kpi_df = pd.DataFrame(
        [
            {
                "Product pallets count": int(len(pallet_summary_df)),
                "Pallet cost (EUR)": round(total_pallet_cost, 2),
                "Glass boxes count": int(glass_boxes),
                "Glass weight total (kg)": round(total_glass_weight, 2),
                "Glass cost (EUR)": round(glass_cost, 2),
                "Total packaging cost (EUR)": round(total_packaging_cost, 2),
                "Product LDM": round(total_pallet_ldm, 3),
                "Glass LDM": round(glass_ldm, 3),
                "Total LDM": round(total_ldm, 3),
            }
        ]
    )

    st.subheader("Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Product pallets", int(len(pallet_summary_df)))
    c2.metric("Glass boxes", int(glass_boxes))
    c3.metric("Packaging cost", f"{total_packaging_cost:.2f} EUR")
    c4.metric("Total LDM", f"{total_ldm:.3f}")

    st.dataframe(kpi_df, use_container_width=True)

    if not pallet_summary_with_glass_df.empty:
        st.subheader("Pallet summary")
        st.dataframe(pallet_summary_with_glass_df, use_container_width=True)

    if not plan_df.empty:
        st.subheader("Packing plan")
        st.dataframe(plan_df, use_container_width=True)

    excel_data = make_excel_file(results_df, pallet_summary_with_glass_df, plan_df, kpi_df)

    dl_col, clear_col = st.columns([3, 1])
    with dl_col:
        name_c1, name_c2 = st.columns([4, 1])
        with name_c1:
            new_name = st.text_input(
                "Project name",
                value=st.session_state.project_name,
                key="project_name_input",
                label_visibility="collapsed",
                placeholder="Enter project name",
            )
        with name_c2:
            if st.button("💾 Save name", use_container_width=True):
                st.session_state.project_name = new_name.strip() or "packing_calculation"

        st.download_button(
            label="⬇️ Download Excel report",
            data=excel_data,
            file_name=f"{st.session_state.project_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with clear_col:
        st.write("")
        st.write("")
        st.write("")
        if st.button("🗑️ Clear all", use_container_width=True):
            clear_results()
            st.session_state.edit_idx = None
            st.rerun()

else:
    st.info("No constructions added yet.")

st.divider()
st.subheader("📂 Import from Excel")
st.caption("Upload a previously saved packing_calculation.xlsx to continue editing")

uploaded = st.file_uploader("Upload Excel file", type=["xlsx"], label_visibility="collapsed")

if uploaded is not None:
    try:
        xl = pd.read_excel(uploaded, sheet_name="Constructions")
        required_cols = {"Item", "Type", "Width (mm)", "Height (mm)", "Qty",
                         "Unit weight (kg)", "Glass weight (kg)", "Glass mode",
                         "Packed as", "Glass separate", "Packed sideways",
                         "Max per pallet", "Pallet width (mm)", "Notes"}
        missing = required_cols - set(xl.columns)
        if missing:
            st.error(f"❌ Missing columns in file: {missing}")
        else:
            st.success(f"✅ Found {len(xl)} construction(s) — ready to import")
            st.dataframe(xl[["Item", "Type", "Width (mm)", "Height (mm)",
                              "Qty", "Unit weight (kg)", "Glass weight (kg)", "Glass mode"]],
                         use_container_width=True)

            imp_col1, imp_col2 = st.columns([1, 3])
            with imp_col1:
                if st.button("⬆️ Import and replace all", use_container_width=True):
                    imported = []
                    for _, row in xl.iterrows():
                        imported.append({
                            "Item":              str(row.get("Item", "Unnamed")),
                            "Type":              str(row.get("Type", "Door")),
                            "Width (mm)":        float(row.get("Width (mm)", 1000)),
                            "Height (mm)":       float(row.get("Height (mm)", 1000)),
                            "Qty":               int(row.get("Qty", 1)),
                            "Unit weight (kg)":  float(row.get("Unit weight (kg)", 0)),
                            "Glass weight (kg)": float(row.get("Glass weight (kg)", 0)),
                            "Glass mode":        str(row.get("Glass mode", "Glazed")),
                            "Packed as":         str(row.get("Packed as", "UNGLAZED")),
                            "Glass separate":    str(row.get("Glass separate", "NO")),
                            "Packed sideways":   str(row.get("Packed sideways", "NO")),
                            "Max per pallet":    int(row.get("Max per pallet", 6)),
                            "Pallet width (mm)": float(row.get("Pallet width (mm)", 1000)),
                            "Notes":             str(row.get("Notes", "")),
                        })
                    st.session_state.results = imported
                    st.session_state.edit_idx = None
                    st.success(f"Imported {len(imported)} construction(s)!")
                    st.rerun()
            with imp_col2:
                if st.button("➕ Import and add to existing", use_container_width=True):
                    if "results" not in st.session_state:
                        st.session_state.results = []
                    for _, row in xl.iterrows():
                        st.session_state.results.append({
                            "Item":              str(row.get("Item", "Unnamed")),
                            "Type":              str(row.get("Type", "Door")),
                            "Width (mm)":        float(row.get("Width (mm)", 1000)),
                            "Height (mm)":       float(row.get("Height (mm)", 1000)),
                            "Qty":               int(row.get("Qty", 1)),
                            "Unit weight (kg)":  float(row.get("Unit weight (kg)", 0)),
                            "Glass weight (kg)": float(row.get("Glass weight (kg)", 0)),
                            "Glass mode":        str(row.get("Glass mode", "Glazed")),
                            "Packed as":         str(row.get("Packed as", "UNGLAZED")),
                            "Glass separate":    str(row.get("Glass separate", "NO")),
                            "Packed sideways":   str(row.get("Packed sideways", "NO")),
                            "Max per pallet":    int(row.get("Max per pallet", 6)),
                            "Pallet width (mm)": float(row.get("Pallet width (mm)", 1000)),
                            "Notes":             str(row.get("Notes", "")),
                        })
                    st.success(f"Added {len(xl)} construction(s) to existing list!")
                    st.rerun()
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
