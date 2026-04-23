from emc_institutional_model.monte_carlo import MonteCarloConfig, run_monte_carlo
from emc_institutional_model.params import ModelParams


def test_monte_carlo_shapes():
    p = ModelParams(analysis_length_months=24)
    r = run_monte_carlo(p, MonteCarloConfig(n_paths=30, seed=1))
    assert r.npv_samples.shape == (30,)
    assert r.monthly_net_panel.shape == (30, 24)
