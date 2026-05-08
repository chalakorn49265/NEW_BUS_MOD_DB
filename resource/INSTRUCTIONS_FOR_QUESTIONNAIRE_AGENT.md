# Instructions for Claude: Build sales-team questionnaires from this `resource/` bundle

## Purpose

Use the files in **`resource/`** (copies of the repo’s Streamlit dashboards and financial model code) to produce **country-specific questionnaires** for sales representatives. Each questionnaire should collect **only the data needed** to parameterize dashboards like the ones in this bundle—clear labels, units, and “why we ask” where helpful.

**Do not assume** the sales team knows Python. Questions should be business-language (USD/MZN, lamps, tariffs, contract terms).

---

## What this folder contains (high-level)

| Area | Path under `resource/` | Role |
|------|-------------------------|------|
| Main EMC institutional cockpit | `streamlit_app.py`, `pages/*.py` | Sidebar inputs → `ModelParams`, tariffs, Monte Carlo, etc. |
| Institutional model schema & math | `emc_institutional_model/*.py` | Field names, defaults, revenue/OPEX/capex logic |
| Workbook-backed tier comparison | `pages/04_Tier_Comparison_Dashboard.py`, `Dashboard_LaaS_vs_NYTG/*.py` | Excel extraction keys, product keys, what-if knobs |
| Isolated workbook viewer | `viewer_app/` | Same tier dashboard, multipage wrapper |
| Mozambique LaaS pitch prototype | `mozambique_app/`, `mozambique/` | Subscription vs baseline LED/HPS, deal splits, undertable framing |
| Trust → LaaS envelope (RMB / roadlight) | `pages/03_Trust_vs_LaaS_Feasible_Envelope.py`, `business_model_comparison/*.py`, `data/*.csv` | Different paradigm—questionnaire section optional per country |
| Human-readable guides | `guides/*.md` | Extra context on LaaS pages and Excel |

---

## Your deliverable

Produce **one questionnaire document per country** (or one master questionnaire with **country annexes**), structured so sales can fill it in a meeting or send back as a form.

### Required structure

1. **Cover metadata** — Country, currency for client conversation, contact, project name (optional).
2. **Baseline inventory** — Lights/poles, incumbent technology (e.g. LED vs HPS), wattages, who pays electricity today.
3. **Operating assumptions** — Hours/night, days/year, outages/dimming if relevant.
4. **Tariffs & energy** — Effective electricity price (and source/confidence), any fixed charges.
5. **Commercial offer** — Term, fee structure (annual, per-light/month, upfront), escalation.
6. **Provider economics** (if sales can estimate or obtain from internal teams) — CAPEX scope, annual OPEX, major replacements (e.g. batteries).
7. **Stakeholders & distribution** — Government fees/taxes, intermediaries, commissions—**percent vs fixed**, timing (upfront vs annual). Include a row for **explicit “undertable” or sensitive transfers** only if legal/commercial policy allows; phrase neutrally (“other contractual distributions”).
8. **Client messaging constraints** — Numbers not to show publicly; preferred framing (savings vs subscription).
9. **Data quality** — Each critical numeric field: source (utility bill, proposal, estimate) and confidence (high/medium/low).

### Method (how to derive questions)

1. **Scan all `resource/pages/*.py` and `resource/mozambique_app/pages/*.py`**  
   Extract every **`st.sidebar` / `st.number_input` / `st.slider` / `st.selectbox` / `st.radio` / `st.checkbox`** label and default. Turn each into a questionnaire field with **unit** and **definition**.

2. **Cross-check `resource/emc_institutional_model/params.py` and `defaults.py`**  
   Map UI labels to **`ModelParams`** fields (and enums/literals). Ensure questionnaire terminology matches the model (e.g. USD/kWh, poles vs lights).

3. **For `04_Tier_Comparison_Dashboard.py` + `Dashboard_LaaS_vs_NYTG/`**  
   Note inputs that come from **Excel workbooks** (`new_models`) vs sidebar what-ifs (product key, electricity bearer). Add questions: “Do we have a filled workbook path?” and “Which product_key / tier name?” if relevant.

4. **For Mozambique-style LaaS-only pitches**  
   Use `resource/mozambique/` and `mozambique_app/` to add sections on **baseline LED vs HPS**, **subscription-only revenue** (no energy-savings-as-revenue), and **split of subscription cash**.

5. **De-duplicate**  
   Many dashboards overlap (tariffs, lights, term). Merge into one logical flow; use **“If using EMC cockpit…”** / **“If using tier workbook viewer…”** branches where needed.

### Output format

- Prefer **Markdown** or **Google Doc–friendly** headings and tables.
- Use **tables** for numeric fields: `Field | Unit | Example | Notes | Required (Y/N)`.
- End with a **“Minimum viable dataset”** checklist (smallest set of fields to run one scenario).

---

## Constraints and caveats

- **Multiple paradigms**: The repo mixes **USD institutional EMC model**, **RMB roadlight / envelope** (`03_*`), and **Mozambique LaaS subscription** prototype. Do not merge incompatible assumptions in one numeric column—label sections clearly.
- **Legal/compliance**: Phrase sensitive distribution questions professionally; avoid accusatory language.
- **Files are snapshots**: Paths refer to **`resource/`** copies; line numbers may drift if the main repo changes—always grep/read from this bundle when generating the questionnaire.

---

## Optional: one-paragraph prompt to paste into Claude

You can paste this below plus attach or zip the **`resource/`** folder:

> Read everything under `resource/` (Streamlit pages, `emc_institutional_model`, Mozambique app, Dashboard_LaaS_vs_NYTG helpers, guides). Extract every user-adjustable parameter from the Streamlit UIs and align them with `ModelParams` / deal logic where applicable. Output a **sales-facing questionnaire** (Markdown) to collect data per country, with units, required vs optional fields, and a minimum viable dataset checklist. Organize by: baseline inventory, energy/tariffs, commercial terms, provider costs, stakeholder splits, and messaging constraints. Call out where different dashboards (main `streamlit_app.py`, tier comparison, Mozambique LaaS, Trust→LaaS envelope) need different questions.

---

## File index (quick reference)

- `streamlit_app.py` — EMC financial cockpit  
- `pages/01_LaaS_Provider_IRR.py`, `02_LaaS_Customer_IRR.py` — LaaS IRR solvers  
- `pages/03_Trust_vs_LaaS_Feasible_Envelope.py` — Feasible envelope (RMB)  
- `pages/04_Tier_Comparison_Dashboard.py` — Tier/workbook comparison  
- `emc_institutional_model/params.py`, `defaults.py` — Core inputs  
- `emc_institutional_model/laas.py` — Subscription-style cashflows  
- `viewer_app/` — Isolated workbook viewer entry  
- `Dashboard_LaaS_vs_NYTG/tier_dashboard_data.py`, `workbook_extract.py`, `product_profiles.py` — Workbook data path  
- `mozambique_app/`, `mozambique/` — Country-style LaaS pitch + `SALES_DATA_REQUEST.md` (example stub)  
- `guides/*.md` — Narrative help  
- `business_model_comparison/*.py`, `data/*.csv` — Envelope page inputs  

---

*Generated for forwarding to Claude or any other agent. Update this file if the bundle contents change.*
