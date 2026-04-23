from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Literal, Optional

import numpy as np
import numpy_financial as npf


SolveFor = Literal["annual_fee_usd", "upfront_usd", "provider_opex_annual_usd", "term_years"]


def monthly_rate_from_annual(r_annual: float) -> float:
    return (1.0 + float(r_annual)) ** (1.0 / 12.0) - 1.0


def annual_rate_from_monthly(r_monthly: float) -> float:
    return (1.0 + float(r_monthly)) ** 12.0 - 1.0


def npv_monthly(cashflows_month0: list[float], annual_discount: float) -> float:
    r_m = monthly_rate_from_annual(annual_discount)
    total = 0.0
    for t, cf in enumerate(cashflows_month0):
        total += cf / ((1.0 + r_m) ** t)
    return float(total)


def irr_annual_from_monthly_cashflows(cashflows_month0: list[float]) -> float | Literal["NO_IRR"]:
    try:
        r_m = float(npf.irr(cashflows_month0))
    except Exception:
        return "NO_IRR"
    if not np.isfinite(r_m):
        return "NO_IRR"
    return float(annual_rate_from_monthly(r_m))


def payback_month_from_monthly_cashflows(cashflows_month0: list[float]) -> int | Literal["NO_PAYBACK"]:
    cum = 0.0
    for idx, cf in enumerate(cashflows_month0):
        cum += float(cf)
        if idx == 0:
            continue
        if cum >= 0:
            return int(idx)
    return "NO_PAYBACK"


@dataclass(frozen=True)
class ProviderLaaSInputs:
    capex_usd: float = 3_000_000.0
    term_years: int = 10
    annual_fee_usd: float = 600_000.0
    upfront_usd: float = 0.0
    escalation_pct_annual: float = 0.0
    provider_opex_annual_usd: float = 0.0


def provider_cashflows_monthly(i: ProviderLaaSInputs) -> list[float]:
    n = int(i.term_years) * 12
    esc_m = (1.0 + float(i.escalation_pct_annual)) ** (1.0 / 12.0) - 1.0
    fee_m1 = float(i.annual_fee_usd) / 12.0
    opex_m1 = float(i.provider_opex_annual_usd) / 12.0

    flows: list[float] = []
    flows.append(-float(i.capex_usd) + float(i.upfront_usd))
    for m in range(1, n + 1):
        esc = (1.0 + esc_m) ** (m - 1)
        inflow = fee_m1 * esc
        outflow = opex_m1 * esc
        flows.append(inflow - outflow)
    return flows


@dataclass(frozen=True)
class CustomerBaselineInputs:
    term_years: int = 10
    baseline_energy_annual_usd: float = 0.0
    baseline_maintenance_annual_usd: float = 0.0
    baseline_escalation_pct_annual: float = 0.0


@dataclass(frozen=True)
class CustomerLaaSInputs:
    term_years: int = 10
    annual_fee_usd: float = 600_000.0
    upfront_usd: float = 0.0
    escalation_pct_annual: float = 0.0
    residual_energy_annual_usd: float = 0.0
    residual_maintenance_annual_usd: float = 0.0
    residual_escalation_pct_annual: float = 0.0


def customer_incremental_cashflows_monthly(b: CustomerBaselineInputs, l: CustomerLaaSInputs) -> list[float]:
    years = int(min(b.term_years, l.term_years))
    n = years * 12

    base_esc_m = (1.0 + float(b.baseline_escalation_pct_annual)) ** (1.0 / 12.0) - 1.0
    laas_esc_m = (1.0 + float(l.escalation_pct_annual)) ** (1.0 / 12.0) - 1.0
    resid_esc_m = (1.0 + float(l.residual_escalation_pct_annual)) ** (1.0 / 12.0) - 1.0

    base_cost_m1 = (float(b.baseline_energy_annual_usd) + float(b.baseline_maintenance_annual_usd)) / 12.0
    pay_m1 = float(l.annual_fee_usd) / 12.0
    resid_cost_m1 = (float(l.residual_energy_annual_usd) + float(l.residual_maintenance_annual_usd)) / 12.0

    flows: list[float] = []
    flows.append(-float(l.upfront_usd))
    for m in range(1, n + 1):
        base = base_cost_m1 * ((1.0 + base_esc_m) ** (m - 1))
        laas_pay = pay_m1 * ((1.0 + laas_esc_m) ** (m - 1))
        resid = resid_cost_m1 * ((1.0 + resid_esc_m) ** (m - 1))
        flows.append(base - (laas_pay + resid))
    return flows


class SolveError(ValueError):
    pass


def solve_bisection(
    f: Callable[[float], float],
    lo: float,
    hi: float,
    tol: float = 1e-6,
    max_iter: int = 120,
) -> float:
    a = float(lo)
    b = float(hi)
    fa = float(f(a))
    fb = float(f(b))
    if not np.isfinite(fa) or not np.isfinite(fb):
        raise SolveError("Objective returned NaN/Inf at bounds.")
    if fa == 0.0:
        return a
    if fb == 0.0:
        return b
    if fa * fb > 0:
        raise SolveError("Bounds do not bracket a solution (objective has same sign at lo/hi).")

    for _ in range(int(max_iter)):
        m = (a + b) / 2.0
        fm = float(f(m))
        if not np.isfinite(fm):
            raise SolveError("Objective returned NaN/Inf during solve.")
        if abs(fm) < tol or abs(b - a) < 1e-12:
            return float(m)
        if fa * fm <= 0:
            b, fb = m, fm
        else:
            a, fa = m, fm
    return float((a + b) / 2.0)


def solve_provider_for_target_irr(
    base: ProviderLaaSInputs,
    target_irr_annual: float,
    solve_for: SolveFor = "annual_fee_usd",
    bounds: tuple[float, float] = (0.0, 5_000_000.0),
    term_year_bounds: tuple[int, int] = (1, 30),
) -> tuple[ProviderLaaSInputs, float]:
    target = float(target_irr_annual)

    if solve_for == "term_years":
        lo_y, hi_y = int(term_year_bounds[0]), int(term_year_bounds[1])
        best: Optional[tuple[ProviderLaaSInputs, float]] = None
        for y in range(lo_y, hi_y + 1):
            cand = ProviderLaaSInputs(
                capex_usd=base.capex_usd,
                term_years=int(y),
                annual_fee_usd=base.annual_fee_usd,
                upfront_usd=base.upfront_usd,
                escalation_pct_annual=base.escalation_pct_annual,
                provider_opex_annual_usd=base.provider_opex_annual_usd,
            )
            npv_v = npv_monthly(provider_cashflows_monthly(cand), target)
            if best is None or abs(npv_v) < abs(best[1]):
                best = (cand, float(npv_v))
        if best is None:
            raise SolveError("No candidates evaluated for term_years.")
        return best[0], best[1]

    lo, hi = float(bounds[0]), float(bounds[1])

    def with_var(x: float) -> ProviderLaaSInputs:
        if solve_for == "annual_fee_usd":
            return ProviderLaaSInputs(**{**base.__dict__, "annual_fee_usd": float(x)})
        if solve_for == "upfront_usd":
            return ProviderLaaSInputs(**{**base.__dict__, "upfront_usd": float(x)})
        if solve_for == "provider_opex_annual_usd":
            return ProviderLaaSInputs(**{**base.__dict__, "provider_opex_annual_usd": float(x)})
        raise SolveError(f"Unknown solve_for: {solve_for}")

    def obj(x: float) -> float:
        cand = with_var(x)
        flows = provider_cashflows_monthly(cand)
        return npv_monthly(flows, target)

    x_star = solve_bisection(obj, lo, hi)
    solved = with_var(x_star)
    return solved, float(obj(x_star))


def solve_customer_for_target_irr(
    baseline: CustomerBaselineInputs,
    laas: CustomerLaaSInputs,
    target_irr_annual: float,
    solve_for: SolveFor = "annual_fee_usd",
    bounds: tuple[float, float] = (0.0, 5_000_000.0),
    term_year_bounds: tuple[int, int] = (1, 30),
) -> tuple[CustomerLaaSInputs, float]:
    target = float(target_irr_annual)

    if solve_for == "term_years":
        lo_y, hi_y = int(term_year_bounds[0]), int(term_year_bounds[1])
        best: Optional[tuple[CustomerLaaSInputs, float]] = None
        for y in range(lo_y, hi_y + 1):
            cand = CustomerLaaSInputs(**{**laas.__dict__, "term_years": int(y)})
            flows = customer_incremental_cashflows_monthly(
                CustomerBaselineInputs(**{**baseline.__dict__, "term_years": int(y)}),
                cand,
            )
            npv_v = npv_monthly(flows, target)
            if best is None or abs(npv_v) < abs(best[1]):
                best = (cand, float(npv_v))
        if best is None:
            raise SolveError("No candidates evaluated for term_years.")
        return best[0], best[1]

    lo, hi = float(bounds[0]), float(bounds[1])

    def with_var(x: float) -> CustomerLaaSInputs:
        if solve_for == "annual_fee_usd":
            return CustomerLaaSInputs(**{**laas.__dict__, "annual_fee_usd": float(x)})
        if solve_for == "upfront_usd":
            return CustomerLaaSInputs(**{**laas.__dict__, "upfront_usd": float(x)})
        if solve_for == "provider_opex_annual_usd":
            raise SolveError("provider_opex_annual_usd is not a customer-side knob in this perspective.")
        raise SolveError(f"Unknown solve_for: {solve_for}")

    def obj(x: float) -> float:
        cand = with_var(x)
        flows = customer_incremental_cashflows_monthly(baseline, cand)
        return npv_monthly(flows, target)

    x_star = solve_bisection(obj, lo, hi)
    solved = with_var(x_star)
    return solved, float(obj(x_star))

