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

# Types where glass is ALWAYS packed separately regardless of height/weight
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
    glazed: bool
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
    return int(math.ceil(size_mm / 1000.0) * 1000)


def real_pallet_width(width_mm: float, height_mm: float) -> float:
    """Physical pallet width: height+200 if packed sideways, else width+100."""
    if height_mm > MAX_GLAZED_HEIGHT:
        return height_mm + 200
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
    return 120.0


def ldm_from_width(width_mm: float, count: int = 1) -> float:
    return (float(width_mm) * float(count)) / TRUCK_WIDTH_M / 1000.0


def calculate_construction(construction: Construction) -> Dict[str, object]:
    real_width = real_pallet_width(construction.width_mm, construction.height_mm)
    packed_sideways = construction.height_mm > MAX_GLAZED_HEIGHT
    is_heavy_type = construction.item_type.lower() in HEAVY_GLAZING_TYPES
    is_facade = construction.item_type.lower() in FACADE_TYPES

    if construction.height_mm > 5000:
        return {
            "Item": construction.item_name,
            "Type": construction.item_type,
            "Width (mm)": float(construction.width_mm),
            "Height (mm)": float(construction.height_mm),
            "Qty": int(construction.qty),
            "Unit weight (kg)": float(construction.weight_kg),
            "Glass weight (kg)": float(construction.glass_weight_kg),
            "Input glazed": "YES" if construction.glazed else "NO",
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

    if construction.glazed:
        if is_facade:
            # facade: always unglazed, glass always in separate box
            packed_as = "UNGLAZED"
            glass_separate = "YES"
            notes = "Facade — glass always packed separately"
        elif packed_sideways:
            # height > 2700: always unglazed, glass in separate box
            packed_as = "UNGLAZED"
            glass_separate = "YES"
            notes = "Glass must be packed separately; construction packed sideways"
        elif is_heavy_type and construction.weight_kg > MAX_PALLET_WEIGHT_KG:
            # heavy type + weight > 1000 kg: too heavy to pack glazed → glass in separate box
            packed_as = "UNGLAZED"
            glass_separate = "YES"
            notes = f"Weight exceeds {MAX_PALLET_WEIGHT_KG:.0f} kg — packed without glass; glass packed separately"
        else:
            packed_as = "GLAZED"
            notes = "Can be packed with glass"

    if not construction.glazed and packed_sideways:
        notes = "Construction packed sideways"

    max_per_pallet = MAX_ITEMS_PER_PALLET_HEAVY if is_heavy_type else MAX_ITEMS_PER_PALLET

    return {
        "Item": construction.item_name,
        "Type": construction.item_type,
        "Width (mm)": float(construction.width_mm),
        "Height (mm)": float(construction.height_mm),
        "Qty": int(construction.qty),
        "Unit weight (kg)": float(construction.weight_kg),
        "Glass weight (kg)": float(construction.glass_weight_kg) if construction.glazed else 0.0,
        "Input glazed": "YES" if construction.glazed else "NO",
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
    u = u.sort_values("Unit weight (kg)", ascending=False).reset_index(drop=True)

    pallets: List[Dict[str, object]] = []

    for _, item in u.iterrows():
        w = float(item["Unit weight (kg)"])
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
                "Pallet weight (kg)": round(float(pallet["weight_kg"]), 2),
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

with st.expander("Rules used", expanded=True):
    st.markdown(
        f"""
        - Max glazed height: **{MAX_GLAZED_HEIGHT} mm**
        - If height is **more than 2700 mm**, construction is packed **sideways**
        - **Pallet width** (physical, used for LDM):
            - Normal: **construction width + 100 mm**
            - Sideways: **construction height + 200 mm**
        - Pallet price tier is based on rounded pallet width (internal):
            - Normal: **max(construction width, minimum pallet width by height)**
            - Sideways: **rounded(height + 200)**
        - Minimum pallet width by height:
            - **≤ 1000 mm → 400 mm**
            - **≤ 2000 mm → 800 mm**
            - **≤ 2700 mm → 1200 mm**
        - Max pallet weight = **{MAX_PALLET_WEIGHT_KG:.0f} kg**
        - Max items per pallet (standard) = **{MAX_ITEMS_PER_PALLET}**
        - Max items per pallet for **Double Sliding, Triple Sliding, 2-leaf+2-fixed, Folding door** = **{MAX_ITEMS_PER_PALLET_HEAVY}**
        - **Double Sliding, Triple Sliding, 2-leaf+2-fixed, Folding door**:
            - Glazed if height ≤ 2700 mm **and** unit weight ≤ {MAX_PALLET_WEIGHT_KG:.0f} kg
            - If weight > {MAX_PALLET_WEIGHT_KG:.0f} kg → packed **unglazed**, glass goes to **separate glass box**
        - If height > 2700 mm → packed **sideways**, glass goes to **separate glass box**
        - **Facade**: glass is **always** packed separately (regardless of height or weight)
        - Glass box price = **{GLASS_BOX_PRICE_EUR:.0f} EUR**
        - Glass box max weight = **{GLASS_BOX_MAX_WEIGHT_KG:.0f} kg**
        - **Glass weight** is entered manually per construction (used for glass box calculation when glass is packed separately)
        - **LDM** is calculated using **pallet width**
        - **Pallet price** is determined by rounded width tier (not shown in tables)
        """
    )

if "results" not in st.session_state:
    st.session_state.results = []

left, right = st.columns([1, 1])

with left:
    st.subheader("Add construction")

    with st.form("packing_form"):
        item_name = st.text_input("Item name", value="...")

        item_type = st.selectbox(
            "Type",
            [
                "Door",
                "Window",
                "Fixed Window",
                "Double Sliding Door",
                "Triple Sliding Door",
                "2-leaf+2-fixed Sliding Door",
                "Folding Door",
                "Door + Sidelight",
                "Window + Sidelight",
                "Facade",
            ],
        )

        width_mm = st.number_input("Width (mm)", min_value=1.0, value=1000.0, step=1.0)
        height_mm = st.number_input("Height (mm)", min_value=1.0, value=1000.0, step=1.0)
        qty = st.number_input("Quantity", min_value=1, value=1, step=1)
        weight_kg = st.number_input("Unit weight (kg)", min_value=0.0, value=0.0, step=1.0)
        glazed = st.checkbox("Glazed", value=True)
        glass_weight_kg = st.number_input(
            "Glass weight (kg)",
            min_value=0.0,
            value=0.0,
            step=1.0,
            help="Weight of glass only (used for glass box calculation when glass is packed separately)",
            disabled=not glazed,
        )

        submitted = st.form_submit_button("Calculate and add")

        if submitted:
            if weight_kg <= 0:
                st.warning("⚠️ Unit weight is 0 — please enter the actual weight before adding.")
            elif glazed and glass_weight_kg <= 0:
                st.warning("⚠️ Glass weight is 0 — please enter the glass weight before adding.")
            else:
                construction = Construction(
                    item_name=item_name.strip() or "Unnamed",
                    item_type=item_type,
                    width_mm=float(width_mm),
                    height_mm=float(height_mm),
                    qty=int(qty),
                    weight_kg=float(weight_kg),
                    glazed=glazed,
                    glass_weight_kg=float(glass_weight_kg) if glazed else 0.0,
                )
                result = calculate_construction(construction)
                add_result_to_session(result)
                st.success(f"Added: {result['Item']}")

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
        glazed=glazed,
        glass_weight_kg=float(glass_weight_kg) if glazed else 0.0,
    )

    preview_df = pd.DataFrame([calculate_construction(preview)])
    st.dataframe(preview_df, use_container_width=True)

st.divider()
st.subheader("Constructions")

if st.session_state.results:
    results_df = pd.DataFrame(st.session_state.results)
    st.dataframe(results_df, use_container_width=True)

    st.markdown("### Remove item")
    col_del1, col_del2 = st.columns([2, 1])

    with col_del1:
        item_to_delete = st.selectbox(
            "Select item to remove",
            options=list(range(len(results_df))),
            format_func=lambda x: f"{results_df.iloc[x]['Item']} (row {x})",
        )

    with col_del2:
        st.write("")
        st.write("")
        if st.button("Delete selected"):
            st.session_state.results.pop(item_to_delete)
            st.rerun()

    pallet_summary_df, plan_df, total_pallet_cost, total_pallet_ldm = build_pallet_outputs(results_df)
    glass_boxes, total_glass_weight, glass_cost, glass_ldm = calculate_glass_boxes(results_df)
    total_packaging_cost = total_pallet_cost + glass_cost
    total_ldm = total_pallet_ldm + glass_ldm

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

    if not pallet_summary_df.empty:
        st.subheader("Pallet summary")
        st.dataframe(pallet_summary_df, use_container_width=True)

    if not plan_df.empty:
        st.subheader("Packing plan")
        st.dataframe(plan_df, use_container_width=True)

    excel_data = make_excel_file(results_df, pallet_summary_df, plan_df, kpi_df)
    st.download_button(
        label="Download Excel report",
        data=excel_data,
        file_name="packing_calculation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if st.button("Clear all results"):
        clear_results()
        st.rerun()
else:
    st.info("No constructions added yet.")
