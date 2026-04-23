"""Typed project parameters (Pydantic)."""

from __future__ import annotations

from datetime import date
from typing import Literal, Optional

from pydantic import BaseModel, Field, model_validator

RevenueBasis = Literal["avoided_kwh", "delivered_kwh"]
ScenarioMode = Literal["Base", "Conservative", "Aggressive"]


def default_tariff_model() -> "TariffModel":
    from emc_institutional_model.defaults import DEFAULT_EDM_FLAT_USD_PER_KWH, DEFAULT_GOV_FLAT_USD_PER_KWH

    return TariffModel.flat(DEFAULT_GOV_FLAT_USD_PER_KWH, DEFAULT_EDM_FLAT_USD_PER_KWH)


def _traditional_benchmark_copy() -> dict[str, float]:
    from emc_institutional_model.defaults import TRADITIONAL_BENCHMARK

    return dict(TRADITIONAL_BENCHMARK)


class TouSide(BaseModel):
    """One tariff stack (e.g. contract gov payment or utility EDM cost).

    Monthly kWh allocated to bucket i is ``total_kwh * load_weights[i]``.
    Monthly energy USD is ``sum_i total_kwh * load_weights[i] * prices_usd_per_kwh[i]``.
    """

    bucket_ids: list[str] = Field(
        default_factory=lambda: ["offpeak", "shoulder", "peak"],
        description="TOU bucket labels",
    )
    prices_usd_per_kwh: list[float] = Field(
        ...,
        description="Price in USD/kWh for each bucket, same order as bucket_ids",
    )
    load_weights: list[float] = Field(
        ...,
        description="Share of monthly kWh in each bucket; must sum to 1",
    )

    @model_validator(mode="after")
    def _check(self) -> TouSide:
        if not (len(self.bucket_ids) == len(self.prices_usd_per_kwh) == len(self.load_weights)):
            raise ValueError("bucket_ids, prices_usd_per_kwh, load_weights must have equal length")
        s = sum(self.load_weights)
        if abs(s - 1.0) > 1e-9:
            raise ValueError(f"load_weights must sum to 1, got {s}")
        if any(w < 0 for w in self.load_weights):
            raise ValueError("load_weights must be non-negative")
        if any(p < 0 for p in self.prices_usd_per_kwh):
            raise ValueError("prices must be non-negative")
        return self

    def effective_flat_usd_per_kwh(self) -> float:
        """Blended flat rate implied by weights (for parity checks only)."""
        return sum(w * p for w, p in zip(self.load_weights, self.prices_usd_per_kwh))

    def energy_usd(self, total_kwh: float) -> float:
        return total_kwh * sum(w * p for w, p in zip(self.load_weights, self.prices_usd_per_kwh))


class TariffModel(BaseModel):
    """Contract (government payment) and operator (EDM-style) TOU stacks."""

    gov_payment: TouSide
    edm_cost: TouSide

    @classmethod
    def flat(cls, gov: float, edm: float) -> TariffModel:
        """Degenerate single-bucket model matching legacy Tariffs_Mozambique B4/B5."""
        one = TouSide(bucket_ids=["flat"], prices_usd_per_kwh=[gov], load_weights=[1.0])
        two = TouSide(bucket_ids=["flat"], prices_usd_per_kwh=[edm], load_weights=[1.0])
        return cls(gov_payment=one, edm_cost=two)


class SiteCostOverride(BaseModel):
    """Optional multipliers on tier-based labor / transport / trenching / security (default = no effect)."""

    labor_multiplier: float = Field(1.0, gt=0)
    transport_multiplier: float = Field(1.0, gt=0)
    trenching_multiplier: float = Field(1.0, gt=0)
    security_multiplier: float = Field(1.0, gt=0)


class EmcAdjustments(BaseModel):
    """Optional tax and fee/distribution items for Sources & Uses (no external borrowing).

    EMC here is equipment and contract cashflows only. When all rates are zero,
    ``net_cashflow_adjusted`` equals project ``net_cashflow``. Headline NPV/IRR stay on full equipment CAPEX.
    """

    corporate_tax_rate: float = Field(0.0, ge=0.0, le=0.5)
    depreciation_months: int = Field(
        120,
        ge=1,
        le=600,
        description="Straight-line depreciation on total CAPEX for taxable income (custody fee excluded from tax base).",
    )
    distribution_pct_of_gross_inflow: float = Field(0.0, ge=0.0, le=1.0)
    distribution_fixed_usd_month: float = Field(0.0, ge=0.0)


class ModelParams(BaseModel):
    """Full inputs for one deterministic run."""

    project_start_date: date = date(2026, 1, 1)
    analysis_length_months: int = Field(120, ge=1, le=600)
    number_of_lights: int = Field(1000, ge=1)
    number_of_poles: int = Field(1000, ge=1)
    operating_hours_per_night: float = Field(11.0, gt=0)
    power_kw_per_light: float = Field(0.10, gt=0)
    location_tier: Literal["city_center", "suburb", "rural"] = "city_center"
    product_type: str = "AI_lightning_grid"
    revenue_basis: RevenueBasis = "avoided_kwh"
    escalation_pct_annual: float = Field(0.03, ge=-0.5, le=0.5)
    aux_grid_fee_monthly_usd: float = Field(0.0, ge=0)
    custody_fee_usd_per_pole_month: float = Field(
        0.0,
        ge=0,
        description="EMC custody/trustee fixed fee (Excel: fixed_service_fee_usd_per_pole_month).",
    )
    custody_fee_enabled: bool = True

    scenario_mode: ScenarioMode = "Base"
    global_override_multiplier: float = Field(1.0, gt=0)
    site_cost_override: SiteCostOverride = Field(default_factory=SiteCostOverride)

    tariff_model: TariffModel = Field(default_factory=default_tariff_model)

    traditional: dict[str, float] = Field(default_factory=_traditional_benchmark_copy)

    discount_rate_annual: float = Field(0.12, ge=0, le=0.5)

    emc_performance_fee_pct_of_energy_savings: float = Field(
        0.0,
        ge=0,
        le=1,
        description="Share of max(0, trad_electrical - ai_electrical) as extra monthly inflow.",
    )

    energy_savings_fraction: Optional[float] = Field(
        None,
        description="If set, delivered kWh factor is baseline_f*(1-s). If None, use product row delivered_f.",
    )

    emc_adjustments: EmcAdjustments = Field(default_factory=EmcAdjustments)

    @model_validator(mode="after")
    def _product_ok(self) -> ModelParams:
        from emc_institutional_model.defaults import PRODUCT_ROWS

        if self.product_type not in PRODUCT_ROWS:
            raise ValueError(f"Unknown product_type {self.product_type!r}")
        return self

    @model_validator(mode="after")
    def _savings_fraction_ok(self) -> ModelParams:
        if self.energy_savings_fraction is None:
            return self
        s = self.energy_savings_fraction
        if s < 0.0 or s >= 1.0:
            raise ValueError("energy_savings_fraction must be in [0, 1) when provided")
        return self
