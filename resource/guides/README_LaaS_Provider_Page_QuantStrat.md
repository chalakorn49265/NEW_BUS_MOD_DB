# LaaS IRR Target — Provider Page (Quant/Strat → Sales Explainer)

This guide is written for **Quant/Strategy** to enable **Sales** to confidently present the **Provider / Project** perspective page:

- **`LaaS IRR Target — Provider / Project`** (`pages/01_LaaS_Provider_IRR.py`)

The provider page answers: **“If we need X% project IRR, what contract terms (fee/upfront/term) make that achievable, given CAPEX is fixed?”**

Note: the Streamlit provider UI currently holds **`provider_opex_annual_usd` fixed at 0** (OPEX is not exposed as an adjustable/solve-for knob) to avoid unphysical “negative OPEX subsidy” artifacts when targeting high IRRs.

## How to launch

From `EMC_INSTITUTIONAL_MODEL/`:

```bash
./.venv/bin/python -m streamlit run streamlit_app.py
```

Then open **LaaS IRR Target — Provider / Project** from the left navigation.

## Model definition (cashflow conventions)

### Timeline

- **Month 0** is the “contract start” cashflow date.
- **Months 1..N** are monthly operating periods where \(N = 12 \cdot \text{term_years}\).

### Provider cashflows

Provider cashflows are defined in `emc_institutional_model/laas.py` as:

- **Month 0**: \(-\text{CAPEX} + \text{upfront}\)
  - CAPEX is fixed at **$3,000,000** (locked in the UI)
  - `upfront_usd` is cash **received** at month 0 (positive)
- **Month m (1..N)**: \(\text{fee}_m - \text{opex}_m\)
  - fee is derived from `annual_fee_usd / 12`
  - opex is derived from `provider_opex_annual_usd / 12`
  - escalation is applied monthly using:
    - \(r_m = (1+r_{annual})^{1/12}-1\)
    - \(\text{amount}_m = \text{amount}_{m1}\cdot(1+r_m)^{m-1}\)

This is intentionally simple so Sales can explain it without ambiguity.

## IRR targeting logic (what “Solve-for” means)

### The core objective

When the user sets **Target IRR (annual)** = \(r\), we convert to a monthly discount rate:

- \(r_m = (1+r)^{1/12} - 1\)

Then we solve for the chosen knob \(x\) such that:

- \(\text{NPV}(r_m, \text{cashflows}(x)) \approx 0\)

This is implemented as a **bisection solve** on NPV, which is robust and monotonic for typical project cashflows.

### Why “Solve-for” is single-parameter

The product requirement says “everything can be adjusted except CAPEX,” but that yields infinitely many combinations. So the UI forces a single solve-for knob (fee/upfront/term) to keep the result explainable and repeatable.

## Sales talk track (recommended)

### 1) Set the return target

“Let’s assume we need **X% IRR** on the provider project.”

### 2) Pick the knob you want to negotiate

- **`annual_fee_usd`**: best for “budget-fit” conversations (default starts at **$600k/year**)
- **`upfront_usd`**: best if the customer can prepay part of deployment
- **`term_years`**: best for structuring (“what term makes the math work?”)

### 3) Read the KPIs

- **Achieved IRR (annual)**: should match the target once solved
- **Solved value**: the fee/upfront/term required to hit the IRR
- **NPV @ target IRR (USD)**: near **0** when solved (numerical tolerance)
- **Payback (months)**: first month when cumulative net cashflow ≥ 0

### 4) Use the charts for narrative

- **Cumulative net cashflow**: “how quickly we recover and how much value accumulates”
- **Unrecovered CAPEX**: \(\max(0, -\text{cumulative})\) — hits 0 at payback
- **Payback marker**: vertical line at payback month

## Solver operational details (for Sales support)

### Bounds and “no solution” cases

The solver requires **bounds** \([lo, hi]\) where NPV changes sign:

- If both NPV(lo) and NPV(hi) are positive (or both negative), the UI will show:
  - “Bounds do not bracket a solution…”

What to do:

- Increase the **upper bound** for fee or upfront (if solving for those)
- Increase the **term** (if not the solve-for) to allow longer recovery
- If the target IRR is infeasible at the current fee/upfront/term, adjust **fee/upfront/term/target** (not OPEX in the UI)

### Discrete `term_years` solve-for

When solving for `term_years`, the page evaluates integer terms (default range 1..30) and chooses the term that gets **NPV closest to 0** at the target IRR (best-effort discrete fit).

## Where the logic lives (for engineering handoff)

- Provider cashflows + IRR/NPV + solver:
  - `EMC_INSTITUTIONAL_MODEL/emc_institutional_model/laas.py`
- Provider page UI:
  - `EMC_INSTITUTIONAL_MODEL/pages/01_LaaS_Provider_IRR.py`

