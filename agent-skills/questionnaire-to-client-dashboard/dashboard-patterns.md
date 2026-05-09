# Dashboard patterns (layout & UX, not formulas)

Use these as **interaction models** when scaffolding new client apps. Copy **structure** (sidebar defaults, tabs, Plotly usage); adapt **math** per engagement.

## Entrypoints vs multipage apps

| Role | File | Notes |
|------|------|------|
| EMC institutional cockpit | `streamlit_app.py` | Larger multipage app at repo root. |
| Minimal workbook viewer | `streamlit_viewer_app.py` | Imports loading logic only; targets tier comparison page—smaller cold start. Do **not** confuse with `streamlit_app.py` when wiring deployment. |

`streamlit_viewer_app.py` exists because `streamlit_app.py` is a separate, heavier EMC cockpit; the viewer keeps imports minimal for deployment.

## Pattern inventory (reference repo)

| Pattern | Location | Use when |
|---------|----------|----------|
| Single-page client economics, CSV-driven, Plotly cashflow, Chinese copy | `wenzhou/wencheng_client_dashboard.py`; run **`python3 -m streamlit run wenzhou/run_dashboard.py`** from repo root | Retrofit / fee split / payback / IRR strip for one jurisdiction |
| LaaS provider IRR | `pages/01_LaaS_Provider_IRR.py` | User explores provider-side LaaS economics |
| LaaS customer IRR | `pages/02_LaaS_Customer_IRR.py` | Customer-side LaaS |
| Trust vs LaaS feasible envelope | `pages/03_Trust_vs_LaaS_Feasible_Envelope.py` | **托管 → LaaS** envelope / constraints |
| Tier comparison / workbook-backed viewer | `streamlit_viewer_app.py` → loads `pages/04_Tier_Comparison_Dashboard.py` | Excel/model tier comparison cockpit |
| Alternative layouts | `mozambique_app/`, `viewer_app/` | Institutional / pitch variants |

## UX conventions

1. **Defaults from survey** — Load fixed `#` IDs once at startup; display in expander “Submitted baseline”.
2. **Sensitivity** — Sidebar sliders with labels tied to acquisition doc semantics (e.g. λ from **C4**, horizon aligned with **H10** when relevant).
3. **Disclosure** — If **J1–J6** indicate incumbent EMC, show a short “existing arrangement” panel before savings claims.
4. **Warnings** — Surface reconcile flags if backing CSV/model exposes them (e.g. `reconcile_required` in wenzhou pipeline).

## Optional structured export

If you persist inputs for reuse across dashboards, align optional JSON with **`docs/MODEL_INPUTS_DASHBOARD_MINIMUM.md`** and **`schemas/project_capex_pack.v2026_01.schema.json`** when available.

See **[`reference-implementation.md`](reference-implementation.md)** for paths relative to the lighting monorepo root.
