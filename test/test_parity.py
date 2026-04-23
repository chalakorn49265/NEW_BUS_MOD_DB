"""Golden checks vs MOZ v2 Excel formulas (flat tariff degeneracy)."""

from __future__ import annotations

from datetime import date

import pytest

from emc_institutional_model.capex import total_capex_usd
from emc_institutional_model.energy import (
    avoided_kwh_month,
    baseline_kwh_month,
    delivered_kwh_month,
    kwh_value_basis_month,
)
from emc_institutional_model.opex import opex_month1
from emc_institutional_model.params import ModelParams, TariffModel, TouSide
from emc_institutional_model.runner import run_model


def _default_params() -> ModelParams:
    return ModelParams(
        project_start_date=date(2026, 1, 1),
        analysis_length_months=120,
        number_of_lights=1000,
        number_of_poles=1000,
        operating_hours_per_night=11.0,
        power_kw_per_light=0.10,
        location_tier="city_center",
        product_type="AI_lightning_grid",
        revenue_basis="avoided_kwh",
        escalation_pct_annual=0.03,
        aux_grid_fee_monthly_usd=0.0,
        custody_fee_usd_per_pole_month=0.0,
        custody_fee_enabled=False,
        tariff_model=TariffModel.flat(0.18, 0.10),
    )


def test_kwh_baseline_delivered_avoided():
    p = _default_params()
    assert baseline_kwh_month(p) == pytest.approx(1000 * 30 * 11 * 0.1 * 2.30)
    assert delivered_kwh_month(p) == pytest.approx(1000 * 30 * 11 * 0.1 * 0.95)
    assert avoided_kwh_month(p) == pytest.approx(baseline_kwh_month(p) - delivered_kwh_month(p))
    assert kwh_value_basis_month(p) == avoided_kwh_month(p)


def test_energy_savings_fraction_overrides_delivered():
    p = _default_params().model_copy(update={"energy_savings_fraction": 0.1})
    assert delivered_kwh_month(p) == pytest.approx(1000 * 30 * 11 * 0.1 * 2.30 * 0.9)
    assert avoided_kwh_month(p) == pytest.approx(baseline_kwh_month(p) - delivered_kwh_month(p))


def test_opex_month1_electrical_matches_flat_tariff():
    p = _default_params()
    o1 = opex_month1(p)
    assert o1.kwh_consumption_month == delivered_kwh_month(p)
    assert o1.electrical_fee == pytest.approx(delivered_kwh_month(p) * 0.10)


def test_tou_degenerates_to_flat():
    p = _default_params()
    p2 = p.model_copy(
        update={
            "tariff_model": TariffModel(
                gov_payment=TouSide(
                    bucket_ids=["a", "b"],
                    prices_usd_per_kwh=[0.18, 0.18],
                    load_weights=[0.4, 0.6],
                ),
                edm_cost=TouSide(
                    bucket_ids=["a", "b"],
                    prices_usd_per_kwh=[0.10, 0.10],
                    load_weights=[0.25, 0.75],
                ),
            )
        }
    )
    o0 = opex_month1(p)
    o2 = opex_month1(p2)
    assert o2.electrical_fee == pytest.approx(o0.electrical_fee)


def test_capex_reasonable_range():
    p = _default_params()
    c = total_capex_usd(p)
    assert 320_000 < c < 480_000


def test_run_model_smoke():
    r = run_model(_default_params())
    assert len(r.monthly) == 120
    assert r.monthly["cumulative_net_cashflow"].iloc[-1] > 0
