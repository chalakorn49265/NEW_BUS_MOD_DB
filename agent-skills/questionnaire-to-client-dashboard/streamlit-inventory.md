# Repo Streamlit inventory — graphs, tables, KPIs

Canonical paths are relative to **NEW_BUS_MOD_DB** repo root. Duplicates under `resource/` mirror this layout and are not listed separately.

Use this document when generating **client dashboards from questionnaires**: mirror the **artifact types** (cumulative cashflow, OPEX stack, KPI strip, comparison tables) even when numbers come from survey-derived stubs.

---

## Entry points (how apps are launched)

| Command | What loads |
|---------|------------|
| `python3 -m streamlit run streamlit_app.py` | EMC institutional model (tabs inside one script; **not** repo `pages/`). |
| `python3 -m streamlit run streamlit_viewer_app.py` | Minimal viewer → executes `pages/04_Tier_Comparison_Dashboard.py` only. |
| `python3 -m streamlit run wenzhou/run_dashboard.py` | Single-page **文成** client economics (`wenzhou/wencheng_client_dashboard.py`). |
| `python3 -m streamlit run mozambique_app/streamlit_app.py` | Multipage: `mozambique_app/pages/*.py`. |
| `python3 -m streamlit run viewer_app/streamlit_app.py` | Multipage: wraps `pages/04_Tier_Comparison_Dashboard.py` via `viewer_app/pages/01_单方案查看器.py`. |

From **repo root**, `python3 -m streamlit run streamlit_app.py` loads the EMC institutional UI **and** Streamlit multipage discovery adds sibling **`pages/`** (`01_`–`04_` scripts) to the sidebar. The heavy EMC charts live **inside** `streamlit_app.py` tabs; **LaaS/envelope/tier** pages are separate files under **`pages/`**.

---

## 1. `streamlit_app.py` — EMC institutional sales cockpit

**Library:** `emc_institutional_model`, Plotly Express + Graph Objects.

**Top KPI row (`st.metric` ×4):** NPV (USD), IRR (annual), Payback (months), Required custody fee ($/pole/mo).

| Tab | Charts | Tables / other |
|-----|--------|----------------|
| **Cashflow** | **Area:** cumulative net cashflow vs month (`px.area`). Optional **MC:** monthly net cashflow with **P5–P95 fan** (`go.Scatter` band). | — |
| **Sources & Uses** | **Composite:** sources/uses bars + cumulative net + shaded negative periods (`_sources_uses_figure`); payback vertical alignment. | Caption explains orange “unrecovered CAPEX” vs blue cumulative. |
| **Tornado** | **Horizontal grouped bars:** ΔNPV upside/downside per driver (`go.Bar` overlay). | — |
| **TCO bridge** | **Grouped bar:** AI vs Traditional snapshot — CAPEX pair, net OPEX yr1 pair, revenue yr1 pair (`go.Bar`). | Caption: negative revenue = inflow. |
| **OPEX** | **Stacked bar:** electrical vs maintenance fee outflows, first 60 months (`go.Bar` stack). | — |
| **TOU** | **Bar:** load weight by bucket; **Line:** gov USD/kWh by bucket (`px.bar`, `px.line`). | — |
| **Monte Carlo** | **Histogram:** NPV distribution; **Histogram:** IRR distribution (`px.histogram`). | **`st.dataframe(mc.summary())`** |

---

## 2. `wenzhou/wencheng_client_dashboard.py` (via `wenzhou/run_dashboard.py`)

**Data:** `wenzhou/wencheng_retrofit_and_om_baseline.csv`.

**KPI / metrics:** Three-column **节费拆分** metrics; six-column strip — client investment, annual cash-in, simple payback, horizon cumulative net, simple ROI, annual IRR; optional large IRR callout.

| Section | Charts | Tables |
|---------|--------|--------|
| **电费对照** | — | **Before/after** fee compare (`st.dataframe`); expander with **post-retrofit bill** estimate. |
| **累计现金流与回收期** | **Line+markers:** cumulative net cashflow; **hline** y=0; **vline** simple payback or first nonnegative year (`go.Scatter`). | **Cashflow detail:** year, annual net, cumulative (`st.dataframe` + `column_config`). |
| **成本与明细溯源** | — | **Tab 改造成本分项:** savings_line detail; **Tab 汇总与比例:** aggregates; **Tab 运维假设:** om_assumption; **metric** 节电量/改造前表尾. |

---

## 3. `pages/01_LaaS_Provider_IRR.py`

**KPI row (`st.metric` ×4):** Achieved IRR, NPV @ target, Payback (months), solved parameter (fee/upfront/term).

| Chart | Spec |
|-------|------|
| **Cumulative cashflow** | `go.Scatter`: cumulative net + **unrecovered CAPEX** line; **vline** + annotation at payback month. |
| **Monthly bars** | `px.bar`: net cashflow, **first 36 months** after month 0. |

---

## 4. `pages/02_LaaS_Customer_IRR.py`

Same layout as Provider IRR with **customer** semantics: cumulative **net benefit**, “unrecovered (benefit gap)”, **monthly net benefit** bars (first 36 months).

---

## 5. `pages/03_Trust_vs_LaaS_Feasible_Envelope.py`

**KPI row:** Evaluated scenarios, Provider-feasible count, Everyone-feasible count, Baseline payback (months).

| Tab | Charts | Tables |
|-----|--------|--------|
| **Feasible envelope** | **Scatter:** x = average client payment RMB/yr, y = term years, color = client_gap RMB, facets = opex_mode (`px.scatter`). | Sorted **scenario grid** dataframe. |
| **Best feasible offer** | **Grouped bars:** baseline 托管费 vs LaaS 服务费 by year; **Lines+markers:** cumulative net cashflow **托管 vs LaaS** (`go.Bar`, `go.Scatter`). | **simple_cashflow_comparison_table** full detail. |
| **Traceability** | — | JSON bundle; **baseline_summary_table**. |

---

## 6. `pages/04_Tier_Comparison_Dashboard.py` (also loaded by `streamlit_viewer_app.py` / `viewer_app`)

**Data:** `Dashboard_LaaS_vs_NYTG/new_models` workbooks via `build_tier_tables`.

**CSS:** Shrinks metric font for dense KPIs.

**Sidebar:** Product comparison **dataframe** (product_key, 节电率, CAPEX scale, 运维缩放, 电费规则).

**Main:** Styled **wide comparison table** (`pandas.Styler`) — EMC vs LaaS columns, deltas, highlights.

| Section | Charts | Tables |
|---------|--------|--------|
| **累计现金流（服务商）** | **Two lines:** EMC vs LaaS **provider** cumulative cashflow; **hline** y=0 (`go.Scatter`). | — |
| **OPEX拆解（Y1）** | **Stacked bar** by 方案: 电费, 人工/维修, 平台, 备件 (`px.bar`, `melt`). | Captions for workbook inputs. |
| **10年经营侧拆解** | **Single-category bars** with text labels: revenue gap, electricity cost gap, non-electric OPEX gap, reconcile net (`go.Bar`). | Caption with reconciliation numbers. |
| **Tab 业主视角** | **Lines:** EMC vs LaaS **owner annual spend**; **Lines:** EMC vs LaaS **owner annual net savings** (`go.Scatter` ×2). | Optional annual dataframe. |
| **Tab 服务商视角** | **Grouped bar** NPV vs IRR (IRR scaled for display — note title). | **`st.metric`** EMC/LaaS NPV & IRR ×4. |
| **Tab 成本拆解** | **Pie:** non-electric OPEX structure (人工/平台/备件) (`px.pie`). | Markdown bullets: lamps, CAPEX EMC/LaaS, baseline bill, saving rates. |
| **Tab 为什么更好** | — | Evidence **expanders** (`cards_for_selected_tier`). |
| **Tab 可追溯性** | — | JSON tier trace. |

---

## 7. `mozambique_app/pages/*` — AI+Solar pitch package

| Page | KPIs | Charts | Tables |
|------|------|--------|--------|
| **01_Pitch_1min** | Before: elec, maint, total; After: subscription; Provider IRR, payback, cumulative, upfront | **Line+markers:** provider **cumulative net** annual; **hline** 0 | Pivot **where subscription money goes** |
| **02_Deal_Splits** | IRR, payback, end cumulative | **Stacked bar** by year: split kinds (`px.bar`) | Stakeholder revenue list; net cashflow dataframe |
| **03_Audit_Model** | — | — | **Tabs:** inputs JSON; customer incremental annual; provider net + cumulative; splits list; trace JSON |
| **04_Export_Audit_Workbook** | — | — | Export/download oriented (no Plotly in file) |

---

## 8. `questionnaire/dashboard_mapper.py` (non-Streamlit helper)

**ChartBundle** maps `project_capex_pack` → **`bundle_to_plotly_dicts`:** **bar** annual OPEX by scenario; **stacked bar** maintenance categories per scenario (`go.Bar`). Agents can reuse this when JSON pack exists.

---

## Skill takeaway — cross-cutting “big 3” patterns

| Pattern | Representative locations |
|---------|-------------------------|
| **Cumulative cashflow + payback** | `wenzhou/wencheng_client_dashboard.py`; `pages/01_*`, `pages/02_*`; `pages/04_*` EMC/LaaS curves; `mozambique_app/01_Pitch_1min.py`; `streamlit_app.py` Cashflow tab |
| **OPEX / cost stack** | `pages/04_*` stacked OPEX + pie; `streamlit_app.py` OPEX tab; `wenzhou` tabs; `dashboard_mapper` bundle |
| **KPI strip** | Almost every page: `st.metric` rows before charts |
| **Sensitivity / risk** | `streamlit_app.py` Tornado; Monte Carlo histograms |
| **Comparison tables** | `pages/04_*` styled wide table; `wenzhou` fee compare; `mozambique` stakeholder tables |

When building **questionnaire-driven** apps without these files, **reproduce the artifact types**, not necessarily the file paths.
