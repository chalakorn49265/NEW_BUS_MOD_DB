from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class StakeholderYearlyRow:
    year: int  # 0..term
    stakeholder: str
    kind: str  # government/intermediary/undertable/provider/customer
    cash_in_usd: float


def annualize_monthly(flows_month0: list[float]) -> list[float]:
    """
    Convert month0-based monthly flows into year0..yearN totals.
    Year0 is month0 only; Year k aggregates months (12*(k-1)+1 .. 12*k).
    """
    if not flows_month0:
        return []
    out: list[float] = [float(flows_month0[0])]
    m = flows_month0[1:]
    n_years = (len(m) + 11) // 12
    for y in range(1, n_years + 1):
        start = (y - 1) * 12
        end = min(y * 12, len(m))
        out.append(float(sum(float(x) for x in m[start:end])))
    return out


def cumulative(flows: list[float]) -> list[float]:
    s = 0.0
    out: list[float] = []
    for x in flows:
        s += float(x)
        out.append(float(s))
    return out

