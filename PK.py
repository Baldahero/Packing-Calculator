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

MAX_GLAZED_WIDTH_HEAVY = 4500  # max width for glazed sliding/folding doors
MAX_SLIDING_WIDTH = 5960       # max width for fully assembled sliding door; above this → SPLIT
SPLIT_PALLET_WIDTH = 5960 + 200  # pallet width for split sliding door
GLASS_BOX_PRICE_EUR = 180.0
GLASS_BOX_MAX_WEIGHT_KG = 1000.0
GLASS_PALLET_WIDTH_MM = 1200
TRUCK_WIDTH_M = 2.0

# ============================================================
# FREIGHT RATES — IRELAND (Naujos +10%)
# Standard: all constructions <= 2700mm height
# Mega: at least one construction > 2700mm height
# ============================================================
IRELAND_RATES = {
    #  LDM: (Standard, Mega)
    0.4:  (407,    462),
    0.8:  (577.5,  633),
    1.2:  (693,    715),
    1.6:  (858,    924),
    2.0:  (1039.5, 1067),
    2.4:  (1188,   1188),
    2.8:  (1386,   1386),
    3.2:  (1595,   1595),
    3.6:  (1826,   1848),
    4.0:  (1963.5, 1964),
    4.4:  (2167,   2167),
    4.8:  (2365,   2426),
    5.2:  (2563,   2596),
    5.6:  (2761,   2805),
    6.0:  (2959,   3003),
    6.4:  (3190,   3179),
    6.8:  (3377,   3399),
    7.2:  (3608,   3619),
    7.6:  (3773,   3839),
    8.0:  (3982,   3982),
    8.4:  (4092,   4147),
    8.8:  (4273.5, 4334),
    9.2:  (4411,   4444),
    9.6:  (4620,   4675),
    10.0: (4675,   4906),
    10.4: (4735.5, 4967),
    10.8: (4906,   5137),
    11.2: (5082,   5313),
}
IRELAND_FTL = (5170, 5500)  # (Standard, Mega)


def get_ireland_freight(total_ldm: float, is_mega: bool) -> float:
    """Get Ireland freight cost based on LDM and trailer type.
    Splits into multiple trucks if total LDM exceeds max truck capacity (13.6 LDM).
    """
    MAX_TRUCK_LDM = 13.6
    ldm_keys = sorted(IRELAND_RATES.keys())

    if total_ldm <= 0:
        return 0.0

    # Split into trucks
    full_trucks = int(total_ldm // MAX_TRUCK_LDM)
    remaining_ldm = total_ldm % MAX_TRUCK_LDM

    # FTL cost per truck
    ftl_std, ftl_mega = IRELAND_FTL
    ftl_cost = ftl_mega if is_mega else ftl_std

    total_cost = full_trucks * ftl_cost

    # Remaining LDM
    if remaining_ldm > 0:
        for key in ldm_keys:
            if remaining_ldm <= key:
                std, mega = IRELAND_RATES[key]
                total_cost += mega if is_mega else std
                break
        else:
            total_cost += ftl_cost

    return total_cost

# Types with special glazing rule: glazed only if height <= 2700 and weight <= 1000 kg
# Also limited to MAX_ITEMS_PER_PALLET_HEAVY per pallet
HEAVY_GLAZING_TYPES = {
    "sliding door",
    "double sliding door",
    "triple sliding door",
    "quad sliding door",
    "double folding door",
    "triple folding door",
    "quad folding door",
    "5-leaf folding door",
}

# Number of parts for split sliding/folding doors
SLIDING_PARTS = {
    "sliding door": 1,
    "double sliding door": 2,
    "triple sliding door": 3,
    "quad sliding door": 4,
    "double folding door": 2,
    "triple folding door": 3,
    "quad folding door": 4,
    "5-leaf folding door": 5,
}

# Facades: glass always packed separately regardless of height/weight
FACADE_TYPES = {
    "facade",
}
