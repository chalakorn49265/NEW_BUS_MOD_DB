# Feishu / Lark field mapping → `project_capex_pack` (schema `2026.01`)

This document describes how to wire **Feishu Base** columns or **Feishu Form** fields to the canonical JSON consumed by [`schemas/project_capex_pack.v2026_01.schema.json`](../schemas/project_capex_pack.v2026_01.schema.json).

## Sales-facing minimum field set (dashboard parity)

For distribution to sales teams or sheet builders who only need **CAPEX triplet + payback / OPEX / 运维费** dashboards (not full project finance), use:

- **[`MODEL_INPUTS_DASHBOARD_MINIMUM.md`](MODEL_INPUTS_DASHBOARD_MINIMUM.md)** — human-readable checklist with keys, labels, units, and JSON mapping notes.
- **[`csv/model_inputs_dashboard_minimum.csv`](csv/model_inputs_dashboard_minimum.csv)** — wide CSV template (header row = flattened JSON paths; second row = illustrative example). Import into Google Sheets or Feishu Base and align column names to those headers when generating `project_capex_pack` JSON.
- **Feishu Form (short list)** — [`../questionnaire/feishu_form_simple.xlsx`](../questionnaire/feishu_form_simple.xlsx) and [`../questionnaire/feishu_form_simple.csv`](../questionnaire/feishu_form_simple.csv): ~21 fields, Chinese labels + suggested 题型/选项, for copy into the form designer or Bitable import; see sheet **使用说明** in the xlsx. Regenerate from [`../questionnaire/build_feishu_form_simple.py`](../questionnaire/build_feishu_form_simple.py) if you change the field set.
- **Excel quick capture (CN + EN + validation)** — [`../questionnaire/questionnaire_01_input.xlsx`](../questionnaire/questionnaire_01_input.xlsx): sheets **填写** and **English** share the same `#` list, driven by [`../questionnaire/questionnaire_input_row_ids.txt`](../questionnaire/questionnaire_input_row_ids.txt) (optional subset). Answer cells use list / numeric / date validation per question type; hidden **Lists** / **Lists_EN**. See [`GENERALIZED_MODEL_DATA_ACQUISITION.md`](GENERALIZED_MODEL_DATA_ACQUISITION.md). Regenerate from [`../questionnaire/build_questionnaire_input_xlsx.py`](../questionnaire/build_questionnaire_input_xlsx.py) after changing the master workbook, ZH map, or row-id file.

## Principles

1. **One submission row per project version** — append-only revisions via `submission_id` (outside schema; store in DB).
2. **Numbers** — store raw numeric cells; handle currency conversion in the intake service before validation.
3. **Attachments** — Feishu file tokens → download server-side → upload to **OSS/COS** → store `storage_uri` in JSON.

## Core mapping (Phase 1 — CAPEX triplet)

| Feishu column label (suggested CN) | JSON path | Type |
|-----------------------------------|-----------|------|
| 项目编号 | `identity.project_id` | text |
| 客户名称 | `identity.client_name` | text |
| 国家/地区 | `identity.country` | text |
| 填写人 | `identity.submitted_by` | text |
| 提交时间 | `identity.submitted_at` | ISO 8601 |
| 我方方案总投资 | `capex_triplet.capex_ours` | number |
| 现状/基准总投资 | `capex_triplet.capex_baseline_incumbent` | number |
| 普通LED方案总投资 | `capex_triplet.capex_normal_led` | number |
| 币种 | `capex_triplet.currency` | ISO 4217 |
| 灯具数量 | `scale.number_of_lights` | integer |
| 灯杆数量（可选） | `scale.number_of_poles` | integer |

Set `identity.source_system` to e.g. `feishu_base_v1`.

## Phase 2 — OPEX + 运维费

Use **fixed scenario keys**: `baseline`, `normal_led`, `emc`, `laas`.

### Annual OPEX (`opex_annual_by_scenario`)

For each scenario row group (or prefixed columns):

| Column pattern | JSON |
|----------------|------|
| `{scenario}_年运维总成本` | `opex_annual_by_scenario.{scenario}.total_annual` |
| `{scenario}_年电费` | `.electricity_annual` |
| `{scenario}_年非电费` | `.non_electric_annual` |

### Maintenance breakdown (`maintenance_breakdown_by_scenario`)

| Column suffix | JSON key |
|---------------|----------|
| 人工费 | `labor` |
| 材料费 | `materials` |
| 其他 | `other` |
| 运维费合计 | `total_om` |
| 巡检费 | `inspection` |
| 清洗费 | `cleaning` |
| 检测费 | `testing` |
| 平台/软件 | `platform_software` |
| 备件 | `spares` |
| 电池/储备 | `battery_reserve` |

## Webhook payload (recommended)

POST JSON body:

```json
{
  "feishu_record_id": "...",
  "payload": { "...": "canonical project_capex_pack object" },
  "attachment_tokens": ["file_token_1"]
}
```

The intake service expands tokens, uploads binaries to object storage, and validates `payload` against the schema.

## Appendix A alignment

Human-readable field inventory (Chinese form → English) lives in the plan appendix **Appendix A**; extend Feishu columns to match optional KPI blocks (`calculated_kpis`, Section 6–7 of the paper form).
