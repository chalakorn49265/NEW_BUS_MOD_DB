---
name: questionnaire-to-client-dashboard
description: >-
  Maps questionnaire_01_input.xlsx (fixed client answers) to Streamlit dashboards with tunable sidebar parameters;
  routes EMC vs LaaS vs retrofit economics using survey IDs J/H/I/K. Use when building or extending client-facing
  Streamlit apps from surveys, questionnaire_01_input, EMC, ESCO, LaaS, energy savings dashboards, or fee-split scenarios.
disable-model-invocation: true
---

# Questionnaire → client-specific Streamlit dashboards

## When to use this skill

- User provides **`questionnaire_01_input.xlsx`** (or answers keyed by `#` column IDs).
- Goal is a **client-facing Streamlit** dashboard with **survey facts locked** and **negotiation knobs** in the sidebar.
- Commercial context is **EMC/ESCO**, **LaaS**, **retrofit economics**, or **tier/model comparison**.

## Core rules

1. **Fixed inputs** — Values taken from the submitted survey (and reconciled bills/O&M where noted). Do not treat these as default sliders unless the user explicitly wants “what-if” override mode.
2. **Flexible parameters** — Tariffs, horizons, fee-split percentages, savings assumptions, escalation, product-side CAPEX—expose via **`st.sidebar`** (or tabs) with defaults that **initialize from survey** where possible.
3. **Routing** — Use [`survey-routing.md`](survey-routing.md) to choose narrative and which dashboard **pattern** to follow ([`dashboard-patterns.md`](dashboard-patterns.md)).
4. **Validation** — Cross-check spend and energy lines per acquisition checklist ([`workflow-checklist.md`](workflow-checklist.md)).

## Minimum dashboard deliverables (client economics)

A sidebar that reads survey JSON is **not** a finished dashboard. Unless the user explicitly asks for a **wireframe only**, every **client economics** Streamlit page built under this skill **must** render **quantitative outputs** tied to the locked survey + scenario sliders:

| Layer | Required |
|--------|----------|
| **KPI strip (“big metrics”)** | At minimum: **capital at risk / client CAPEX share**, **annual benefit or recurring cash-in** (aligned to your model), **simple payback** (years), and either **horizon cumulative net** or **IRR** when the cashflow series supports it. Use `st.columns` + `st.metric` (or equivalent). |
| **Cumulative cashflow + payback** | **Plotly** line chart: **year index** vs **cumulative net cashflow**, horizontal line at **0**, vertical marker or annotation at **simple payback** (and/or first year cumulative crosses zero). Pair with an **annual cashflow table** (`st.dataframe` / `column_config`) in an expander or below. |
| **Cost stack: CAPEX vs OPEX** | At least one **visual** (stacked bar, waterfall, or pie) **and** a **table** splitting **one-time CAPEX** vs **annual OPEX / fees / service** (labels match the engagement). If the model only has totals, still show the split from assumptions and label uncertainty. |
| **Scenario / fee tables** | At least one **comparison table**: e.g. before vs after electricity cost, or baseline vs scenario A/B—so stakeholders see numbers, not only charts. |

**Anti-patterns to avoid:** pages that only display raw survey fields; empty charts; `st.success` with no figures; placeholder text where numbers should be. **Implement simplified metrics** from questionnaire + sliders when a full CSV/model hook is missing—then disclose simplifications in captions.

Full patterns, chart specs, and mapping to the Wenzhou reference implementation are in [`dashboard-patterns.md`](dashboard-patterns.md).

## Workflow (short)

1. **Ingest** — Parse sheet **填写** or **English**; keys are `#` column IDs (see [`fixed-vs-flexible.md`](fixed-vs-flexible.md)).
2. **Lock** — Build a **fixed facts** dict: INT*, A*, B*, C*, D*, E*, **J***, G/H/I/K as applicable.
3. **Route** — EMC incumbent → emphasize **J1–J6** in copy and baseline; LaaS appetite → LaaS/feasible-envelope style layouts.
4. **Build** — Streamlit: `st.session_state` or module defaults from fixed facts; sliders for flexible params; **implement KPI strip, cumulative cashflow chart + table, CAPEX/OPEX breakdown, and scenario tables** per [`dashboard-patterns.md`](dashboard-patterns.md) (not optional for client economics).
5. **Validate** — B1 vs B2+B3+B4; C4 vs TOU fields **C4a–C4e**; surface reconcile warnings if upstream CSV/model uses `reconcile_required`-style flags.

## Detailed docs in this folder

| File | Purpose |
|------|---------|
| [`fixed-vs-flexible.md`](fixed-vs-flexible.md) | What stays fixed vs tunable |
| [`survey-routing.md`](survey-routing.md) | Question IDs → EMC / LaaS / generic |
| [`dashboard-patterns.md`](dashboard-patterns.md) | UX patterns (no formula copying) |
| [`reference-implementation.md`](reference-implementation.md) | Paths in lighting monorepo when present |
| [`workflow-checklist.md`](workflow-checklist.md) | Ordered checklist |
| [`streamlit-inventory.md`](streamlit-inventory.md) | **Every Streamlit app in the repo:** charts, tables, KPIs by file |

## Canonical references in the monorepo (when checked out)

| Topic | Source |
|--------|--------|
| Question IDs → meaning (ZH) | `questionnaire/questionnaire_01_zh_map.py` (`ZH_MAP`) |
| Fixed vs flexible layers, **J1–J6**, **G/H/I/K** | `docs/GENERALIZED_MODEL_DATA_ACQUISITION.md` |
| Optional JSON pack for dashboards | `docs/MODEL_INPUTS_DASHBOARD_MINIMUM.md`, `schemas/project_capex_pack.v2026_01.schema.json` |
| Input workbook regeneration | `questionnaire/build_questionnaire_input_xlsx.py`, `questionnaire/questionnaire_input_row_ids.txt` |
| Survey workbook | `questionnaire/questionnaire_01_input.xlsx` (sheets **填写** / **English**, key column **`#`**) |
| XLSX → JSON helper | `scripts/read_questionnaire_input.py` |

## Portability (standalone repo / Claude Code / Cursor)

This directory is intended to be **copied or published as its own Git repository** for Claude Code or other agents. The **workflow above does not require** the lighting monorepo to be checked out. When working **inside** the reference repo, use [`reference-implementation.md`](reference-implementation.md) for concrete file paths.

**Cursor:** Copy or symlink this folder to `.cursor/skills/questionnaire-to-client-dashboard/` (project) or under user skills per Cursor docs.

**Claude Code:** Clone this repo (or submodule) and register the skill path per current Claude Code / Anthropic documentation for project or user skills.

**Embedding in monorepo:** Keep this folder under `agent-skills/`; relative links to `questionnaire/` and `docs/` work when the parent repo is `NEW_BUS_MOD_DB`.

See [`README.md`](README.md) for human-facing install notes.
