from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from emc_institutional_model.defaults import PRODUCT_ROWS

ProductKey = Literal[
    "AI_lightning_grid",
    "AI_battery_integrated_grid",
    "AI_plus_solar_offgrid",
]


@dataclass(frozen=True)
class ProductProfile:
    key: ProductKey
    # `delivered_f / baseline_f` are grid-kWh-style factors in PRODUCT_ROWS.
    delivered_factor: float
    baseline_factor: float
    # Cost drivers (relative comparisons used for scaling)
    capex_composite: float
    routine_maint_factor: float
    # Behavior flags
    is_grid: bool
    has_battery: bool

    def implied_grid_saving_rate(self) -> float:
        """
        Return implied electricity saving vs a grid baseline:
        saving = 1 - delivered_factor / baseline_factor

        For off-grid solar, delivered_factor is 0 → saving ~ 1.
        """
        bf = max(1e-9, float(self.baseline_factor))
        s = 1.0 - float(self.delivered_factor) / bf
        return float(max(0.0, min(0.999, s)))


def _row(key: ProductKey) -> tuple[int, int, int, float, float, float, float, float, float, float, int, float, float]:
    return PRODUCT_ROWS[str(key)]


def get_product_profile(key: ProductKey) -> ProductProfile:
    (
        is_grid,
        _req_ug,
        has_batt,
        fixture,
        trench,
        cable,
        labor,
        logistics,
        routine_maint,
        _batt_cost,
        _batt_cycle,
        delivered_f,
        baseline_f,
    ) = _row(key)

    capex_composite = float(fixture) + float(trench) + float(cable) + float(labor) + float(logistics)
    return ProductProfile(
        key=key,
        delivered_factor=float(delivered_f),
        baseline_factor=float(baseline_f),
        capex_composite=float(capex_composite),
        routine_maint_factor=float(routine_maint),
        is_grid=bool(int(is_grid)),
        has_battery=bool(int(has_batt)),
    )


def capex_scale_vs_reference(key: ProductKey, *, reference: ProductKey = "AI_lightning_grid") -> float:
    ref = get_product_profile(reference).capex_composite
    cur = get_product_profile(key).capex_composite
    if ref <= 0:
        return 1.0
    return float(max(0.5, min(3.0, cur / ref)))


def routine_om_scale(key: ProductKey, *, reference: ProductKey = "AI_lightning_grid") -> float:
    ref = get_product_profile(reference).routine_maint_factor
    cur = get_product_profile(key).routine_maint_factor
    if ref <= 0:
        return 1.0
    return float(max(0.4, min(2.0, cur / ref)))


def forces_grid_electricity_zero(key: ProductKey) -> bool:
    # Off-grid solar: no grid purchases (deliver factor 0 in PRODUCT_ROWS).
    return key == "AI_plus_solar_offgrid"


def provenance_note(key: ProductKey) -> str:
    return f"Source: emc_institutional_model/defaults.py PRODUCT_ROWS[{key}]"

