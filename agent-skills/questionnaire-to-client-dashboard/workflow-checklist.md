# Workflow checklist: ingest → map → scaffold → validate

## 1. Ingest

- [ ] Open **`questionnaire_01_input.xlsx`**; use sheet **填写** or **English** per stakeholder language.
- [ ] Read row keys from column **`#`** (e.g. `INT4`, `B2`, `J1`).
- [ ] Apply **Lists** / validation semantics from **`docs/GENERALIZED_MODEL_DATA_ACQUISITION.md`** when interpreting coded answers.

## 2. Lock (fixed facts)

- [ ] Build a **fixed facts** structure: INT*, A*, B*, C*, D*, E*, **J***, G/H/I/K as answered.
- [ ] Resolve human labels via **`questionnaire_01_zh_map.py`** (`ZH_MAP`) when needed for UI copy.

## 3. Route

- [ ] **J1–J6**: If incumbent EMC/ESCO, plan UI copy and baseline electricity payer (**J4**).
- [ ] **H/I/K**: If LaaS-style acceptance (**I2**, **H10**, etc.), prefer LaaS / feasible-envelope patterns ([`survey-routing.md`](survey-routing.md)).
- [ ] Choose [`dashboard-patterns.md`](dashboard-patterns.md) template(s).

## 4. Scaffold

- [ ] New or forked Streamlit page: **defaults from survey**; **`st.sidebar`** for tariffs, horizons, fee splits, scenarios.
- [ ] Optional: persist **`project_capex_pack`** or project CSV (wenzhou-style) per **`docs/MODEL_INPUTS_DASHBOARD_MINIMUM.md`** / schema.

## 5. Validate

- [ ] **B1** vs **B2+B3+B4** per Definitions in acquisition doc.
- [ ] **C4** consistent with **C4a–C4e** if TOU is used.
- [ ] Flag upstream **`reconcile_required`** or similar if present in loaded CSV/model output.

## 6. Ship

- [ ] Document run command (`streamlit run …`) and which root entrypoint matches the story (`streamlit_app.py` vs `streamlit_viewer_app.py`).
