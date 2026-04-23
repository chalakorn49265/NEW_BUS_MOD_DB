"""CAPEX (CAPEX sheet B4:B10 logic)."""

from __future__ import annotations

from emc_institutional_model.defaults import PRODUCT_ROWS
from emc_institutional_model.location import effective_multipliers
from emc_institutional_model.params import ModelParams


def total_capex_usd(p: ModelParams) -> float:
    pr = PRODUCT_ROWS[p.product_type]
    req_ug = pr[1]
    (
        _fixture,
        trench_u,
        cable_u,
        labor_u,
        log_u,
    ) = pr[3:8]
    labor_m, trans_m, trench_m, _sec_m = effective_multipliers(p)
    n_lights = p.number_of_lights
    n_poles = p.number_of_poles

    poles_fixtures = n_lights * pr[3]
    trenching = (n_lights * trench_u * trench_m) if req_ug else 0.0
    ug_cable = (n_lights * cable_u * trench_m) if req_ug else 0.0
    install = n_poles * labor_u * labor_m
    logistics = n_poles * log_u * trans_m
    return poles_fixtures + trenching + ug_cable + install + logistics
