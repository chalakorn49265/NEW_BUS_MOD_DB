"""Enumerate product × tier for dashboard matrix (Scenarios sheet)."""

from __future__ import annotations

from emc_institutional_model.defaults import PRODUCT_ROWS
from emc_institutional_model.metrics import scenario_kpis
from emc_institutional_model.params import ModelParams

TIERS = ["city_center", "suburb", "rural"]
PRODUCTS = list(PRODUCT_ROWS.keys())


def scenario_matrix(base: ModelParams) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for prod in PRODUCTS:
        for tier in TIERS:
            p = base.model_copy(update={"product_type": prod, "location_tier": tier})
            k = scenario_kpis(p)
            rows.append(
                {
                    "product_type": prod,
                    "location_tier": tier,
                    **k,
                }
            )
    return rows
