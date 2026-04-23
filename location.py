"""Location multipliers (tier + scenario; optional site cost overrides)."""

from __future__ import annotations

from emc_institutional_model.defaults import SCENARIO_FACTORS, TIER_DEFAULTS
from emc_institutional_model.params import ModelParams, ScenarioMode


def _scenario_vector(mode: ScenarioMode) -> tuple[float, float, float, float, float, float]:
    return SCENARIO_FACTORS[mode]


def effective_multipliers(p: ModelParams) -> tuple[float, float, float, float]:
    """Returns (labor, transport, trenching, security) effective multipliers."""
    tier = p.location_tier
    l0, t0, tr0, s0, _p = TIER_DEFAULTS[tier]
    base_l = round(l0, 6)
    base_t = round(t0, 6)
    base_tr = round(tr0, 6)
    base_s = round(s0, 6)

    sf = _scenario_vector(p.scenario_mode)
    g = p.global_override_multiplier
    labor_f, trans_f, trench_f, sec_f, _perm_f, _inst_f = sf

    ov = p.site_cost_override
    labor = base_l * labor_f * g * ov.labor_multiplier
    transport = base_t * trans_f * g * ov.transport_multiplier
    trenching = base_tr * trench_f * g * ov.trenching_multiplier
    security = base_s * sec_f * g * ov.security_multiplier
    return labor, transport, trenching, security
