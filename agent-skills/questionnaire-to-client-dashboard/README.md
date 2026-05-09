# Questionnaire → client Streamlit dashboard (Agent Skill)

Portable skill pack: teach agents to turn **`questionnaire_01_input.xlsx`** answers into **client-specific Streamlit dashboards**, with **survey = fixed facts** and **sidebar = flexible negotiation parameters**. Supports **EMC/ESCO**, **LaaS**, and retrofit-style economics.

## Contents

- **`SKILL.md`** — Main agent instructions (YAML frontmatter for Cursor).
- **`fixed-vs-flexible.md`**, **`survey-routing.md`**, **`dashboard-patterns.md`**, **`workflow-checklist.md`**, **`reference-implementation.md`**.

## Use as a standalone GitHub repo

1. Use this folder as the **repository root** (e.g. rename repo `questionnaire-dashboard-skill`).
2. Push to GitHub; consume from **Claude Code**, **Cursor**, or other agents by pointing the tool at this directory.

## Cursor

Copy or symlink this directory to:

`.cursor/skills/questionnaire-to-client-dashboard/`

at the **project** root (or use personal skills per Cursor settings).

## Claude Code

Follow Anthropic / Claude Code documentation for **project or user skill paths**, then clone this repository (or add it as a submodule) and register that path so `SKILL.md` is discoverable.

## Optional: reference lighting monorepo

When building dashboards that mirror internal examples, clone or submodule your copy of **NEW_BUS_MOD_DB** next to this skill and read **`reference-implementation.md`** for paths like `wenzhou/`, `pages/`, `questionnaire/`.
