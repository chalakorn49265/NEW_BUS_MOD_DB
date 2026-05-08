"""Monte Carlo over `run_model` (independent draws)."""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from emc_institutional_model.params import ModelParams, TariffModel, TouSide
from emc_institutional_model.runner import run_model


@dataclass
class MonteCarloConfig:
    n_paths: int = 400
    seed: int | None = 42
    # relative std dev on flat tariff levels when degenerate TOU
    sigma_gov_price: float = 0.08
    sigma_edm_price: float = 0.10
    sigma_hours: float = 0.05
    sigma_escalation: float = 0.01


@dataclass
class MonteCarloResult:
    npv_samples: np.ndarray
    irr_samples: np.ndarray
    payback_samples: np.ndarray
    monthly_net_panel: np.ndarray  # shape (n_paths, n_months)

    def summary(self) -> pd.DataFrame:
        def q(x: np.ndarray, p: float) -> float:
            y = x[np.isfinite(x)]
            if y.size == 0:
                return float("nan")
            return float(np.quantile(y, p))

        irr = self.irr_samples
        return pd.DataFrame(
            [
                {
                    "metric": "npv_usd",
                    "p5": q(self.npv_samples, 0.05),
                    "p50": q(self.npv_samples, 0.50),
                    "p95": q(self.npv_samples, 0.95),
                    "mean": float(np.nanmean(self.npv_samples)),
                },
                {
                    "metric": "irr_annual",
                    "p5": q(irr, 0.05),
                    "p50": q(irr, 0.50),
                    "p95": q(irr, 0.95),
                    "mean": float(np.nanmean(irr)),
                },
            ]
        )


def _perturb_tou(side: TouSide, rng: np.random.Generator, sigma: float) -> TouSide:
    """Lognormal multiplicative shock per bucket, renormalized weights unchanged."""
    shocks = rng.lognormal(mean=-0.5 * sigma**2, sigma=sigma, size=len(side.prices_usd_per_kwh))
    new_prices = [max(1e-6, p * s) for p, s in zip(side.prices_usd_per_kwh, shocks)]
    return TouSide(bucket_ids=list(side.bucket_ids), prices_usd_per_kwh=new_prices, load_weights=list(side.load_weights))


def run_monte_carlo(base: ModelParams, cfg: MonteCarloConfig) -> MonteCarloResult:
    rng = np.random.default_rng(cfg.seed)
    n = cfg.n_paths
    months = base.analysis_length_months
    npvs = np.empty(n)
    irrs = np.empty(n)
    pbs = np.empty(n)
    panel = np.empty((n, months))

    for i in range(n):
        gov = _perturb_tou(base.tariff_model.gov_payment, rng, cfg.sigma_gov_price)
        edm = _perturb_tou(base.tariff_model.edm_cost, rng, cfg.sigma_edm_price)
        hours = max(0.5, rng.normal(base.operating_hours_per_night, cfg.sigma_hours))
        esc = max(-0.2, min(0.5, rng.normal(base.escalation_pct_annual, cfg.sigma_escalation)))
        p = base.model_copy(
            update={
                "tariff_model": TariffModel(gov_payment=gov, edm_cost=edm),
                "operating_hours_per_night": float(hours),
                "escalation_pct_annual": float(esc),
            }
        )
        r = run_model(p)
        npvs[i] = r.npv_usd
        irrs[i] = float(r.irr_annual) if isinstance(r.irr_annual, float) else float("nan")
        pbs[i] = float(r.payback) if isinstance(r.payback, int) else float("nan")
        panel[i, :] = r.monthly["net_cashflow"].to_numpy()

    return MonteCarloResult(
        npv_samples=npvs,
        irr_samples=irrs,
        payback_samples=pbs,
        monthly_net_panel=panel,
    )


def fan_chart_quantiles(panel: np.ndarray, qs: tuple[float, ...] = (0.05, 0.5, 0.95)) -> dict[str, np.ndarray]:
    out: dict[str, np.ndarray] = {}
    for q in qs:
        out[f"p{int(q * 100)}"] = np.quantile(panel, q, axis=0)
    return out
