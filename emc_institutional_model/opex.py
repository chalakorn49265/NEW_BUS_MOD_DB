"""Monthly OPEX (OPEX sheet B4:B10)."""

from __future__ import annotations

from dataclasses import dataclass

from emc_institutional_model.defaults import PRODUCT_ROWS
from emc_institutional_model.energy import delivered_kwh_month
from emc_institutional_model.params import ModelParams


@dataclass(frozen=True)
class OpexMonth1:
    """Month-1 USD/month before escalation."""

    kwh_consumption_month: float
    electrical_fee: float
    maintenance_routine: float
    battery_replacement_reserve: float
    maintenance_fee: float
    total_opex: float


def opex_month1(p: ModelParams) -> OpexMonth1:
    pr = PRODUCT_ROWS[p.product_type]
    is_grid = bool(pr[0])
    has_batt = bool(pr[2])
    routine = pr[8]
    batt_cost = pr[9]
    batt_cycle = pr[10]

    kwh_del = delivered_kwh_month(p)
    if is_grid:
        electrical = p.tariff_model.edm_cost.energy_usd(kwh_del) + p.aux_grid_fee_monthly_usd
    else:
        electrical = 0.0

    maint_routine = p.number_of_poles * routine
    batt_reserve = (
        (p.number_of_lights * batt_cost / max(1, batt_cycle)) if has_batt else 0.0
    )
    maint_fee = maint_routine + batt_reserve
    total = maint_fee + electrical
    return OpexMonth1(
        kwh_consumption_month=kwh_del,
        electrical_fee=electrical,
        maintenance_routine=maint_routine,
        battery_replacement_reserve=batt_reserve,
        maintenance_fee=maint_fee,
        total_opex=total,
    )
