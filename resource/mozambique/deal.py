from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


ChargeTiming = Literal["upfront", "annual"]


@dataclass(frozen=True)
class DealSubscription:
    """
    Client payments under LaaS (only revenue source).
    """

    term_years: int
    annual_fee_usd: float
    upfront_usd: float = 0.0
    escalation_pct_annual: float = 0.0


@dataclass(frozen=True)
class DealCharge:
    """
    A transfer out of subscription inflow to a stakeholder.

    - If `pct_of_subscription` is set, it applies to subscription inflow (same timing).
    - If `fixed_usd` is set, it is an absolute amount per timing.
    """

    recipient: str
    timing: ChargeTiming = "annual"
    pct_of_subscription: float = 0.0
    fixed_usd: float = 0.0
    kind: Literal["government", "intermediary", "undertable"] = "intermediary"


@dataclass(frozen=True)
class DealSplits:
    """
    Collection of charges (government/intermediaries/undertable).

    These are modeled as **provider outflows** (reducing provider IRR),
    while remaining transparent as a stakeholder revenue list.
    """

    charges: list[DealCharge]

    def normalized(self) -> "DealSplits":
        # Clamp absurd inputs instead of throwing inside Streamlit.
        out: list[DealCharge] = []
        for c in self.charges:
            pct = float(c.pct_of_subscription)
            if pct != pct:
                pct = 0.0
            pct = max(0.0, min(1.0, pct))
            fixed = float(c.fixed_usd)
            if fixed != fixed:
                fixed = 0.0
            fixed = max(0.0, fixed)
            out.append(
                DealCharge(
                    recipient=str(c.recipient),
                    timing=c.timing,
                    pct_of_subscription=pct,
                    fixed_usd=fixed,
                    kind=c.kind,
                )
            )
        return DealSplits(charges=out)

