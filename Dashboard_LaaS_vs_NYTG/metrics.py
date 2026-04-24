from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

import numpy as np
import numpy_financial as npf


def npv_monthly(cashflows_month0: list[float], annual_discount: float) -> float:
    r_m = (1.0 + float(annual_discount)) ** (1.0 / 12.0) - 1.0
    total = 0.0
    for t, cf in enumerate(cashflows_month0):
        total += float(cf) / ((1.0 + r_m) ** t)
    return float(total)


def irr_annual_from_monthly_cashflows(cashflows_month0: list[float]) -> float | Literal["NO_IRR"]:
    try:
        r_m = float(npf.irr(cashflows_month0))
    except Exception:
        return "NO_IRR"
    if not np.isfinite(r_m):
        return "NO_IRR"
    return float((1.0 + r_m) ** 12.0 - 1.0)


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
class DebtMetrics:
    dscr_by_year: dict[int, float | None]  # None if debt service is zero
    dscr_min: float | None
    dscr_avg: float | None


def compute_dscr_by_year(
    *,
    cfads_rmb_y: dict[int, float],
    debt_service_rmb_y: dict[int, float],
) -> DebtMetrics:
    dscr: dict[int, float | None] = {}
    vals: list[float] = []
    for y in sorted(set(cfads_rmb_y.keys()) | set(debt_service_rmb_y.keys())):
        svc = float(debt_service_rmb_y.get(y, 0.0))
        if svc <= 0:
            dscr[y] = None
            continue
        v = float(cfads_rmb_y.get(y, 0.0)) / svc
        dscr[y] = v
        if np.isfinite(v):
            vals.append(float(v))
    return DebtMetrics(
        dscr_by_year=dscr,
        dscr_min=(min(vals) if vals else None),
        dscr_avg=(sum(vals) / len(vals) if vals else None),
    )

