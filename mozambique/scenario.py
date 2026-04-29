from __future__ import annotations

from dataclasses import dataclass

from emc_institutional_model.laas import (
    CustomerBaselineInputs,
    CustomerLaaSInputs,
    ProviderLaaSInputs,
    customer_incremental_cashflows_monthly,
    irr_annual_from_monthly_cashflows,
    payback_month_from_monthly_cashflows,
    provider_cashflows_monthly,
)

from mozambique.baseline import BaselineTech, baseline_electricity_annual_usd, baseline_maintenance_annual_usd
from mozambique.deal import DealCharge, DealSplits, DealSubscription
from mozambique.distribution import annualize_monthly, cumulative


@dataclass(frozen=True)
class MozambiqueInputs:
    number_of_lights: int
    operating_hours_per_night: float
    days_per_year: float
    electricity_price_usd_per_kwh: float

    # Offered system (AI+Solar) economics (provider side).
    capex_usd_total: float
    provider_opex_annual_usd: float


@dataclass(frozen=True)
class ScenarioResult:
    baseline: BaselineTech
    subscription: DealSubscription
    splits: DealSplits
    inputs: MozambiqueInputs

    # Customer view
    baseline_energy_annual_usd: float
    baseline_maintenance_annual_usd: float
    customer_incremental_monthly: list[float]
    customer_irr_annual: float | str
    customer_payback_month: int | str

    # Provider view (after splits/undertable)
    provider_gross_monthly: list[float]
    provider_net_monthly: list[float]
    provider_net_irr_annual: float | str
    provider_net_payback_month: int | str

    # Annual tables for UI
    provider_net_annual_y0_to_yN: list[float]
    provider_net_cumulative_annual: list[float]
    stakeholder_revenue_annual: list[dict]
    traceability: dict


def _apply_charges_monthly(
    *,
    gross_provider_monthly: list[float],
    subscription: DealSubscription,
    charges: list[DealCharge],
) -> tuple[list[float], list[dict]]:
    """
    Returns (net_provider_monthly, stakeholder_rows_annual).

    Convention:
    - provider gross monthly is already (subscription inflow - provider base opex), with month0 including -capex + upfront.
    - charges are modeled as additional outflows to provider (reducing net), and as annual stakeholder revenues.
    """
    term_months = int(subscription.term_years) * 12
    flows = list(float(x) for x in gross_provider_monthly[: term_months + 1])

    # Precompute annual subscription totals (used for pct charges).
    annual_inflow = annualize_monthly(
        provider_cashflows_monthly(
            ProviderLaaSInputs(
                capex_usd=0.0,
                term_years=int(subscription.term_years),
                annual_fee_usd=float(subscription.annual_fee_usd),
                upfront_usd=float(subscription.upfront_usd),
                escalation_pct_annual=float(subscription.escalation_pct_annual),
                provider_opex_annual_usd=0.0,
            )
        )
    )

    stakeholder_rows: list[dict] = []

    def add_stakeholder(year: int, recipient: str, kind: str, amt: float) -> None:
        if abs(float(amt)) < 1e-9:
            return
        stakeholder_rows.append(
            {
                "year": int(year),
                "stakeholder": str(recipient),
                "kind": str(kind),
                "cash_in_usd": float(amt),
            }
        )

    # Apply charges as annual (spread evenly monthly for net-cashflow realism).
    for c in charges:
        timing = c.timing
        pct = float(c.pct_of_subscription or 0.0)
        fixed = float(c.fixed_usd or 0.0)
        kind = str(c.kind)
        recip = str(c.recipient)

        if timing == "upfront":
            base = float(annual_inflow[0] if annual_inflow else 0.0)
            amt0 = pct * base + fixed
            flows[0] -= amt0
            add_stakeholder(0, recip, kind, amt0)
            continue

        # annual
        for y in range(1, int(subscription.term_years) + 1):
            base = float(annual_inflow[y]) if y < len(annual_inflow) else 0.0
            amt = pct * base + fixed
            # spread across months 12*(y-1)+1 .. 12*y
            per_m = amt / 12.0
            for m in range(1 + (y - 1) * 12, 1 + y * 12):
                if m < len(flows):
                    flows[m] -= per_m
            add_stakeholder(y, recip, kind, amt)

    return flows, stakeholder_rows


def run_mozambique_scenario(
    *,
    baseline: BaselineTech,
    subscription: DealSubscription,
    splits: DealSplits,
    inputs: MozambiqueInputs,
) -> ScenarioResult:
    """
    LaaS-only scenario:
    - Customer: replaces (baseline electricity + baseline maintenance) with (subscription fee), residual costs ≈ 0.
    - Provider: earns subscription, pays CAPEX + provider OPEX, and pays out deal charges (gov/intermediary/undertable).
    """
    splits_n = splits.normalized()

    b_energy = baseline_electricity_annual_usd(
        number_of_lights=int(inputs.number_of_lights),
        watt_per_light=float(baseline.watt_per_light),
        operating_hours_per_night=float(inputs.operating_hours_per_night),
        days_per_year=float(inputs.days_per_year),
        electricity_price_usd_per_kwh=float(inputs.electricity_price_usd_per_kwh),
    )
    b_maint = baseline_maintenance_annual_usd(
        number_of_lights=int(inputs.number_of_lights),
        maintenance_usd_per_light_year=float(baseline.maintenance_usd_per_light_year),
    )

    # Customer view: baseline costs vs LaaS payments; AI+Solar residual costs default to 0.
    cust_base = CustomerBaselineInputs(
        term_years=int(subscription.term_years),
        baseline_energy_annual_usd=float(b_energy),
        baseline_maintenance_annual_usd=float(b_maint),
        baseline_escalation_pct_annual=float(subscription.escalation_pct_annual),
    )
    cust_laas = CustomerLaaSInputs(
        term_years=int(subscription.term_years),
        annual_fee_usd=float(subscription.annual_fee_usd),
        upfront_usd=float(subscription.upfront_usd),
        escalation_pct_annual=float(subscription.escalation_pct_annual),
        residual_energy_annual_usd=0.0,
        residual_maintenance_annual_usd=0.0,
        residual_escalation_pct_annual=float(subscription.escalation_pct_annual),
    )
    cust_flows = customer_incremental_cashflows_monthly(cust_base, cust_laas)
    cust_irr = irr_annual_from_monthly_cashflows(cust_flows)
    cust_pb = payback_month_from_monthly_cashflows(cust_flows)

    # Provider gross (before charges): subscription inflow - base provider opex, with capex at month0.
    provider_gross = provider_cashflows_monthly(
        ProviderLaaSInputs(
            capex_usd=float(inputs.capex_usd_total),
            term_years=int(subscription.term_years),
            annual_fee_usd=float(subscription.annual_fee_usd),
            upfront_usd=float(subscription.upfront_usd),
            escalation_pct_annual=float(subscription.escalation_pct_annual),
            provider_opex_annual_usd=float(inputs.provider_opex_annual_usd),
        )
    )

    provider_net, stakeholder_rows = _apply_charges_monthly(
        gross_provider_monthly=provider_gross,
        subscription=subscription,
        charges=splits_n.charges,
    )

    provider_net_irr = irr_annual_from_monthly_cashflows(provider_net)
    provider_net_pb = payback_month_from_monthly_cashflows(provider_net)

    provider_net_annual = annualize_monthly(provider_net)
    provider_net_cum = cumulative(provider_net_annual)

    traceability = {
        "conventions": {
            "offer_framing": "Before: client pays electricity + maintenance. After: client pays subscription; provider handles maintenance; grid electricity treated as ~0 for AI+Solar.",
            "cashflow_convention": "Monthly flows with month 0 containing CAPEX outflow and upfront inflow (provider) / upfront outflow (customer).",
            "splits_convention": "Government/intermediary/undertable items are modeled as provider outflows and reported as stakeholder revenues.",
        },
        "baseline_math": {
            "electricity_annual_usd": "lights * (watt_per_light/1000) * operating_hours_per_night * days_per_year * electricity_price_usd_per_kwh",
            "maintenance_annual_usd": "lights * maintenance_usd_per_light_year",
            "inputs": {
                "lights": int(inputs.number_of_lights),
                "watt_per_light": float(baseline.watt_per_light),
                "operating_hours_per_night": float(inputs.operating_hours_per_night),
                "days_per_year": float(inputs.days_per_year),
                "electricity_price_usd_per_kwh": float(inputs.electricity_price_usd_per_kwh),
                "maintenance_usd_per_light_year": float(baseline.maintenance_usd_per_light_year),
            },
        },
        "subscription": subscription.__dict__,
        "provider_inputs": {
            "capex_usd_total": float(inputs.capex_usd_total),
            "provider_opex_annual_usd": float(inputs.provider_opex_annual_usd),
        },
        "charges": [c.__dict__ for c in splits_n.charges],
        "engine": {
            "module": "emc_institutional_model.laas",
            "provider_cashflows": "ProviderLaaSInputs(capex, term, annual_fee, upfront, escalation, provider_opex_annual)",
            "customer_incremental_cashflows": "CustomerBaselineInputs(baseline_energy+baseline_maintenance) vs CustomerLaaSInputs(subscription + residual_costs(=0))",
        },
    }

    return ScenarioResult(
        baseline=baseline,
        subscription=subscription,
        splits=splits_n,
        inputs=inputs,
        baseline_energy_annual_usd=float(b_energy),
        baseline_maintenance_annual_usd=float(b_maint),
        customer_incremental_monthly=cust_flows,
        customer_irr_annual=cust_irr,
        customer_payback_month=cust_pb,
        provider_gross_monthly=provider_gross,
        provider_net_monthly=provider_net,
        provider_net_irr_annual=provider_net_irr,
        provider_net_payback_month=provider_net_pb,
        provider_net_annual_y0_to_yN=provider_net_annual,
        provider_net_cumulative_annual=provider_net_cum,
        stakeholder_revenue_annual=stakeholder_rows,
        traceability=traceability,
    )

