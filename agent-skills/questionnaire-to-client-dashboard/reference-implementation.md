# Reference implementation (NEW_BUS_MOD_DB)

Use this appendix **when** the lighting monorepo is checked out beside or above your work. The skill core stays valid **without** these paths.

## Questionnaire & semantics

| Topic | Path |
|--------|------|
| Questionnaire workbook | `questionnaire/questionnaire_01_input.xlsx` |
| ZH labels / helper map (`ZH_MAP`) | `questionnaire/questionnaire_01_zh_map.py` |
| Regenerate input workbook | `questionnaire/build_questionnaire_input_xlsx.py`, `questionnaire/questionnaire_input_row_ids.txt` |
| Acquisition layers (fixed vs flexible, **J1–J6**, **G/H/I/K**) | `docs/GENERALIZED_MODEL_DATA_ACQUISITION.md` |
| Dashboard minimum inputs | `docs/MODEL_INPUTS_DASHBOARD_MINIMUM.md` |
| Optional capex pack schema | `schemas/project_capex_pack.v2026_01.schema.json` |
| Flatten XLSX → JSON for agents | `scripts/read_questionnaire_input.py` |

## Wenzhou file as MVED template

For **KPI strip**, **Plotly cumulative cashflow + payback**, **fee comparison table**, and **CAPEX / O&M tabs**, trace **`wenzhou/wencheng_client_dashboard.py`** end-to-end (`compute_cashflow_table`, metrics columns, `go.Scatter` cumulative chart, `st.tabs` for cost lines). CSV input: **`wenzhou/wencheng_retrofit_and_om_baseline.csv`** (same directory).

## Full Streamlit graph/table inventory

See **`streamlit-inventory.md`** in this skill folder for **every** dashboard: `streamlit_app.py`, `pages/01`–`04`, `wenzhou/`, `mozambique_app/pages`, `questionnaire/dashboard_mapper.py`, and viewer entrypoints.

## Streamlit entry commands (from repo root)

```bash
python3 -m streamlit run wenzhou/run_dashboard.py
python3 -m streamlit run streamlit_viewer_app.py
python3 -m streamlit run streamlit_app.py
```

## Files by pattern

| Pattern | Paths |
|---------|--------|
| Wenzhou client economics | `wenzhou/wencheng_client_dashboard.py`, `wenzhou/run_dashboard.py` |
| LaaS IRR | `pages/01_LaaS_Provider_IRR.py`, `pages/02_LaaS_Customer_IRR.py` |
| Feasible envelope | `pages/03_Trust_vs_LaaS_Feasible_Envelope.py` |
| Tier viewer | `streamlit_viewer_app.py` → `pages/04_Tier_Comparison_Dashboard.py` |
| Other apps | `mozambique_app/`, `viewer_app/` |

## Splitting this skill to a standalone repo

After `git subtree split` or copy of `agent-skills/questionnaire-to-client-dashboard/`:

- Core markdown remains self-contained.
- Link here only when consumers also clone **NEW_BUS_MOD_DB** (submodule or sibling directory).
- Alternatively vendor short excerpts from `docs/GENERALIZED_MODEL_DATA_ACQUISITION.md` under CC-BY/your license—do not assume monorepo availability.
