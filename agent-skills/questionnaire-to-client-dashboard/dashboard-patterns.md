# Dashboard patterns (layout, UX, and minimum charts)

Use these as **interaction models** when scaffolding new client apps. Copy **structure** (sidebar defaults, tabs, Plotly usage); adapt **math** per engagement.

**Repo traceability:** see **[`streamlit-inventory.md`](streamlit-inventory.md)** for a **full table** of each Streamlit entrypoint in NEW_BUS_MOD_DB — Plotly chart types, `st.dataframe` tables, and `st.metric` KPI strips — so agents can mirror real implementations.

---

## Minimum viable economics dashboard (MVED)

When routing leads to **retrofit / fee-split / payback / EMC savings**, treat the following as **mandatory** unless the user explicitly requests a **wireframe**:

### 1) KPI strip (“big metrics”)

Place immediately after intro copy (before primary charts). Typical layout: **`st.columns(3)`–`st.columns(6)`** with **`st.metric`**.

| Metric | Role |
|--------|------|
| Client investment / CAPEX share | Money at risk (may include extra one-time costs). |
| Annual cash-in to client | From annual benefit × fee split (or subscription net of cost—match model). |
| Simple payback | Investment ÷ annual cash-in when both positive; show **—** if undefined. |
| Horizon cumulative net | −CAPEX + sum of annual inflows over analysis period (state horizon in label). |
| Simple ROI or cumulative return | Optional but encouraged when horizon metric exists. |
| Annual IRR | Optional; only if you build a proper yearly cashflow vector (`numpy_financial.irr` or equivalent). |

Reference wiring (same ideas): **`wenzhou/wencheng_client_dashboard.py`** builds a multi-column KPI row after computing `client_investment`, `client_annual_cashflow_in`, payback, cumulative net, IRR.

### 2) Cumulative cashflow chart + payback + annual table

**Chart (Plotly):**

- X-axis: **year** (0 = investment year if cashflow convention uses year 0 outflow).
- Y-axis: **cumulative net cashflow** (same currency as survey **INT4**).
- **`go.Scatter`** `mode="lines+markers"` for cumulative series.
- **`fig.add_hline(y=0)`** dashed — breakeven.
- **`fig.add_vline`** at **simple payback** (fractional years allowed as x-position) **or** annotate first calendar year where cumulative ≥ 0.

**Table:**

- Columns at minimum: **year**, **net cashflow that year**, **cumulative net**.
- Put full table in **`st.expander("Cashflow detail")`** or directly under the chart with **`column_config`** number formatting.

Reference wiring: same file — section **「累计现金流与回收期」** (`compute_cashflow_table`, Plotly figure, expander with `cash_df`).

### 3) CAPEX vs OPEX breakdown (“cost stack”)

Stakeholders expect to **see capital separately from recurring costs**.

| Presentation | When to use |
|----------------|---------------|
| **Tabs**: “CAPEX lines”, “OPEX / O&M assumptions”, “Aggregates” | Matches CSV-backed builds (`record_type` slices). |
| **Stacked bar** (single year): segments = CAPEX buckets vs annual O&M | Strong visual when few categories. |
| **Pie / donut** | Use sparingly; good for **share of total CAPEX** by scope, not for time series. |
| **Waterfall** | Optional: bridge from baseline bill → savings → fees → net. |

Minimum rule: **one chart** + **one dataframe** that lists line items (even if simplified). If data only has totals, derive illustrative splits from **E***, **D***, survey answers and label **“illustrative allocation”**.

Reference wiring: **`wenzhou/wencheng_client_dashboard.py`** — **`st.tabs`**「改造成本分项 / 汇总与比例 / 运维假设」, sourcing `savings_line`, `aggregate`, `om_assumption` rows from CSV.

### 4) Comparison tables (not optional)

Include at least **one** of:

- **Before vs after electricity** (kWh and/or annual bill), or  
- **Baseline vs scenario** columns side by side.

Use **`st.dataframe`** with formatted strings or **`column_config`** for thousands separators.

Reference wiring: **`fee_compare_tbl`** + metrics block explaining **节费** vs **应付电费**口径.

### 5) Data plumbing when no CSV exists

If there is **no** `project_capex_pack` or engineering CSV yet:

1. Build a **minimal annual series** from questionnaire + sliders (e.g. implied savings from **C***, **B*** annual spend, **E*** rough CAPEX).
2. Still render **all MVED sections** with disclosed assumptions.
3. Prefer loading normalized inputs via **`scripts/read_questionnaire_input.py`** → JSON, then map IDs to model parameters.

---

## Entrypoints vs multipage apps

| Role | File | Notes |
|------|------|------|
| EMC institutional cockpit | `streamlit_app.py` | Larger multipage app at repo root. |
| Minimal workbook viewer | `streamlit_viewer_app.py` | Imports loading logic only; targets tier comparison page—smaller cold start. Do **not** confuse with `streamlit_app.py` when wiring deployment. |

`streamlit_viewer_app.py` exists because `streamlit_app.py` is a separate, heavier EMC cockpit; the viewer keeps imports minimal for deployment.

## Pattern inventory (reference repo)

| Pattern | Location | Use when |
|---------|----------|----------|
| Single-page client economics, CSV-driven, Plotly cashflow, KPI strip, CAPEX/OPEX tabs | `wenzhou/wencheng_client_dashboard.py`; run **`python3 -m streamlit run wenzhou/run_dashboard.py`** from repo root | **Primary reference for MVED** — retrofit / fee split / payback / IRR strip |
| LaaS provider IRR | `pages/01_LaaS_Provider_IRR.py` | User explores provider-side LaaS economics — mirror **IRR + cashflow strip** patterns |
| LaaS customer IRR | `pages/02_LaaS_Customer_IRR.py` | Customer-side LaaS |
| Trust vs LaaS feasible envelope | `pages/03_Trust_vs_LaaS_Feasible_Envelope.py` | **托管 → LaaS** envelope / constraints — charts required for envelope, not text-only |
| Tier comparison / workbook-backed viewer | `streamlit_viewer_app.py` → loads `pages/04_Tier_Comparison_Dashboard.py` | Excel/model tier comparison cockpit |
| Alternative layouts | `mozambique_app/`, `viewer_app/` | Institutional / pitch variants |

## UX conventions

1. **Defaults from survey** — Load fixed `#` IDs once at startup; display in expander “Submitted baseline”.
2. **Sensitivity** — Sidebar sliders with labels tied to acquisition doc semantics (e.g. λ from **C4**, horizon aligned with **H10** when relevant).
3. **Disclosure** — If **J1–J6** indicate incumbent EMC, show a short “existing arrangement” panel before savings claims.
4. **Warnings** — Surface reconcile flags if backing CSV/model exposes them (e.g. `reconcile_required` in wenzhou pipeline).
5. **Never ship chart-empty economics pages** — See MVED section above.

## Optional structured export

If you persist inputs for reuse across dashboards, align optional JSON with **`docs/MODEL_INPUTS_DASHBOARD_MINIMUM.md`** and **`schemas/project_capex_pack.v2026_01.schema.json`** when available.

See **[`reference-implementation.md`](reference-implementation.md)** for paths relative to the lighting monorepo root.
