# LaaS solver technical notes (`emc_institutional_model/laas.py`)

This document is for **Quant/Strat + Engineering** to understand (and safely modify) the LaaS IRR targeting logic implemented in:

- `EMC_INSTITUTIONAL_MODEL/emc_institutional_model/laas.py`

The Streamlit pages are thin UIs over this module:

- Provider page: `EMC_INSTITUTIONAL_MODEL/pages/01_LaaS_Provider_IRR.py`
- Customer page: `EMC_INSTITUTIONAL_MODEL/pages/02_LaaS_Customer_IRR.py`

## Design goals (why it’s implemented this way)

- **Single “Solve-for” knob**: the product requirement says “adjust various parameters,” but that is underdetermined. We enforce **one adjustable parameter at a time** for explainability and to avoid arbitrary optimizer behavior.
- **Robust solve**: solve **NPV at target discount = 0** using **bisection**, not a derivative-based solver.
- **Monthly timeline**: Streamlit pages plot monthly series and payback; the math is consistent with monthly discounting.

## Module map

### Types and inputs

- `SolveFor`: allowed solve knobs
  - `"annual_fee_usd"`, `"upfront_usd"`, `"term_years"`
- `ProviderLaaSInputs`
  - `capex_usd` (fixed at 3,000,000 in UI)
  - `term_years`
  - `annual_fee_usd`
  - `upfront_usd`
  - `escalation_pct_annual`
  - `provider_opex_annual_usd` (still exists in the math model; the Streamlit provider page currently fixes this at **0** and does not expose it as a solve-for knob)
- `CustomerBaselineInputs`
  - `term_years`
  - `baseline_energy_annual_usd`
  - `baseline_maintenance_annual_usd`
  - `baseline_escalation_pct_annual`
- `CustomerLaaSInputs`
  - `term_years`
  - `annual_fee_usd`
  - `upfront_usd`
  - `escalation_pct_annual` (for LaaS fee)
  - residual costs: `residual_energy_annual_usd`, `residual_maintenance_annual_usd`, `residual_escalation_pct_annual`

### Core computations

- Discount conversions:
  - `monthly_rate_from_annual(r_annual)`
  - `annual_rate_from_monthly(r_monthly)`
- Valuation:
  - `npv_monthly(cashflows_month0, annual_discount)`
  - `irr_annual_from_monthly_cashflows(cashflows_month0)`
  - `payback_month_from_monthly_cashflows(cashflows_month0)`
- Cashflow builders:
  - `provider_cashflows_monthly(inputs)`
  - `customer_incremental_cashflows_monthly(baseline, laas)`
- Solvers:
  - `solve_bisection(f, lo, hi, ...)`
  - `solve_provider_for_target_irr(base, target_irr_annual, solve_for, ...)`
  - `solve_customer_for_target_irr(baseline, laas, target_irr_annual, solve_for, ...)`

## Cashflow definitions (exact conventions)

### Time indexing

All monthly cashflow arrays are **month-0 indexed**:

- index 0 = **Month 0**
- index 1..N = **Month 1..N**

This aligns with `npv_monthly(...)` which discounts by \(t\) starting at 0.

### Provider perspective cashflows

Produced by `provider_cashflows_monthly(i: ProviderLaaSInputs)`:

- Month 0:
  - \(\text{CF}_0 = -\text{capex} + \text{upfront}\)
- Month \(m \in [1, N]\):
  - \(\text{CF}_m = \text{fee}_m - \text{opex}_m\)
  - \(\text{fee}_{m1} = \text{annual_fee_usd} / 12\)
  - \(\text{opex}_{m1} = \text{provider_opex_annual_usd} / 12\)
  - escalation uses a monthly equivalent:
    - \(e_m = (1 + e_{annual})^{1/12} - 1\)
    - \(\text{fee}_m = \text{fee}_{m1} \cdot (1+e_m)^{m-1}\)
    - \(\text{opex}_m = \text{opex}_{m1} \cdot (1+e_m)^{m-1}\)

Notes:

- Escalation is applied to fee and (if non-zero) provider opex using the same escalation parameter (minimal coupling).
- In the current Streamlit provider UI, provider opex is fixed at **0**, so operationally only the fee escalates.
- If you need separate escalation paths, add new fields rather than overloading one.

### Customer perspective incremental cashflows (baseline-cost stream)

Produced by `customer_incremental_cashflows_monthly(b, l)`:

- Month 0:
  - \(\text{CF}_0 = -\text{upfront}\) (customer pays upfront → negative benefit)
- Month \(m \in [1, N]\):
  - \(\text{CF}_m = \text{baseline_cost}_m - (\text{laas_payment}_m + \text{residual_cost}_m)\)

Where:

- \(\text{baseline_cost}_{m1} = (\text{baseline_energy} + \text{baseline_maintenance})/12\)
- \(\text{laas_payment}_{m1} = \text{annual_fee_usd}/12\)
- \(\text{residual_cost}_{m1} = (\text{residual_energy} + \text{residual_maintenance})/12\)
- Each has its own escalation (baseline, fee, residual), converted annual→monthly as above.

Interpretation:

- Positive \(\text{CF}_m\) means **customer is better off** than baseline that month.
- Cumulative sum is cumulative benefit; “payback” is the first month cumulative ≥ 0.

## NPV and IRR math

### NPV at annual discount

`npv_monthly(cashflows, annual_discount)` does:

1) convert annual → monthly:

\[
r_m = (1+r_{annual})^{1/12} - 1
\]

2) compute:

\[
\text{NPV} = \sum_{t=0}^{T} \frac{\text{CF}_t}{(1+r_m)^t}
\]

### IRR output

`irr_annual_from_monthly_cashflows(...)`:

- uses `numpy_financial.irr(cashflows)` to get monthly IRR \(r_m\)
- converts to annual:

\[
r_{annual} = (1+r_m)^{12} - 1
\]

Edge cases:

- Returns `"NO_IRR"` if IRR fails or is non-finite.

## Solver logic (how “target IRR” is hit)

### Objective

Given a target annual IRR \(r\), define monthly discount \(r_m\) and solve:

\[
\text{NPV}(r, \text{cashflows}(x)) = 0
\]

Where \(x\) is the selected “solve-for” knob (fee, upfront, term).

Implementation detail:

- We do **not** solve `IRR(cashflows(x)) = r` directly.
- We solve `NPV(discount=r, cashflows(x)) = 0`, which is typically monotonic in \(x\) for these cashflow shapes and is more stable for bisection.

### Continuous knobs (bisection)

For solve-for knobs:

- `annual_fee_usd`
- `upfront_usd`

We define an objective `obj(x)` = NPV at target discount of the cashflows built with `x`, and apply:

- `solve_bisection(obj, lo, hi)`

#### Bracketing requirement

Bisection requires `obj(lo)` and `obj(hi)` to have opposite signs. If they do not:

- `SolveError("Bounds do not bracket a solution...")`

### Discrete knob (`term_years`)

`term_years` is discrete (integer years). For that knob we do:

- Evaluate candidates in `term_year_bounds` (default 1..30)
- Pick the year with minimum `abs(NPV_at_target)`

This is a best-effort discrete solve; it does not guarantee exact NPV=0.

## Payback logic

`payback_month_from_monthly_cashflows`:

- cumulative sum from month 0
- returns first month index **≥ 1** where cumulative ≥ 0
- returns `"NO_PAYBACK"` otherwise

This is consistent with the dashboard narrative: payback occurs after operations begin, not at time zero.

## Extending the model (safe patterns)

### Adding a new solve-for knob

1) Add it to `SolveFor`
2) Add a field to `ProviderLaaSInputs` / `CustomerLaaSInputs` (and page UI)
3) Update `with_var(...)` in the relevant solver(s)
4) Ensure the objective is monotonic in the solve variable for normal ranges; if not monotonic, bisection may fail.

### Adding new cashflow lines

Add them to the cashflow builder functions, not to the solver.

Keep cashflow sign conventions explicit:

- Provider inflows positive, outflows negative
- Customer “benefit vs baseline” positive

## Known limitations (intentional)

- No debt, taxes, depreciation, working capital, or terminal value.
- Escalation is simplified; provider fee and provider opex share the same escalation parameter (opex is typically **0** in the Streamlit provider UI).
- Customer page uses baseline-cost stream only (not “savings-only” mode).
- `term_years` solve is discrete best-fit, not root-finding.

