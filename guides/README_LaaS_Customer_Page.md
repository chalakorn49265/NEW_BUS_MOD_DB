# LaaS IRR Target — Customer Page (Sales Playbook)

This guide is for **Sales** to present the **Customer** perspective page:

- **`LaaS IRR Target — Customer`** (`pages/02_LaaS_Customer_IRR.py`)

The customer page answers: **“Compared to doing nothing, what return does the customer earn, and when do they break even?”**

## How to launch

From `EMC_INSTITUTIONAL_MODEL/`:

```bash
./.venv/bin/python -m streamlit run streamlit_app.py
```

Then, in the left navigation, open **LaaS IRR Target — Customer**.

## What to enter (talk track)

### 1) Baseline (existing system)

Ask the customer for their current annual costs and enter:

- **Baseline energy cost (USD/year)**
- **Baseline maintenance cost (USD/year)**
- **Baseline escalation (annual %)**

This is the “do nothing” cost stream.

### 2) LaaS (new system)

Enter the proposed contract:

- **Annual LaaS fee (USD/year)** (default starting value is 600k/year)
- **Upfront payment (USD)** (paid at month 0; often 0)
- **LaaS fee escalation (annual %)**

### 3) Residual customer costs (post-upgrade)

If the customer will still pay any energy/maintenance after LaaS, enter:

- **Residual energy cost (USD/year)**
- **Residual maintenance cost (USD/year)**
- **Residual escalation (annual %)**

If LaaS is “all-in”, keep these at **0**.

## What the model computes (plain-English)

The model computes **incremental monthly benefit** versus baseline:

- Monthly net benefit = **(baseline cost) − (LaaS payments + residual costs)**

So:

- **Positive** means the customer is **better off** than the baseline that month.
- **Cumulative net benefit** shows how benefits accumulate over time.

## Controls that matter

- **Target IRR (annual)**: the customer IRR you want to demonstrate.
- **Solve-for parameter**: the **single** knob the solver adjusts to hit the target IRR:
  - `annual_fee_usd` (most common: “What fee still gives customer X% IRR?”)
  - `upfront_usd`
  - `term_years`
- **Solver bounds**: widen bounds if the solver can’t find a solution.

## KPIs (what to say in the room)

- **Achieved IRR (annual)**: the customer’s return on the incremental cashflows.
- **Solved value**: the fee / upfront / term that achieves the target.
- **NPV @ target IRR (USD)**: should be close to **0** when solved.
- **Payback (months)**: first month when cumulative net benefit reaches **≥ 0**.

## Charts

- **Cumulative net benefit**: the “value creation over time” story.
- **Payback marker**: vertical marker at the payback month.
- **Unrecovered (benefit gap)**: \(\max(0, -\text{cumulative})\), hits 0 at payback.

## Troubleshooting

- **“Bounds do not bracket a solution”**
  - Increase the **upper bound** for `annual_fee_usd`, or switch the solve-for knob, or re-check baseline/residual assumptions.

