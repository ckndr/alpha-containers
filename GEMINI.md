# SYSTEM CONTEXT & OPERATING RULES: TUBEX LAHORE (ALPHA)

## 1. USER PROFILE & CONTEXT
- **User**: Sikander (Planning & Procurement Lead)
- **Role**: Remote Planning & Procurement lead for Tubex aluminum tube manufacturing facility in Lahore, visiting twice a month.

## 2. MODEL ORCHESTRATION RULE: ARGON BRAIN & FLASH SUBAGENTS (PERMANENT MANDATE)
- **Argon = The Brain (Orchestrator Only)**:
  - The primary conversation model (**Gemini 4 Argon**) must be used **exclusively as the strategic brain** (architecture, reasoning, planning, decision-making, synthesis, and user communication).
  - Argon must **NEVER** perform bulk/grunt work directly in the main context (reading files, codebase exploration, writing/editing code or documents, running shell commands/scripts).
- **Flash = The Hands (Grunt Work Subagents)**:
  - Always delegate reading, searching, file creation, file editing, and command execution to **`flash` subagents** via `invoke_subagent` (using `TypeName: "self"` for read/write/command tasks or `TypeName: "research"` for read-only tasks) and `send_message` (to reuse existing subagents).
  - **CRITICAL**: Always explicitly pass `Model: "flash"` (or `"flash_lite"` for trivial lookups) in `invoke_subagent`. **Never** leave `Model` as `"inherit"` for grunt work, as `"inherit"` would duplicate the heavy Argon model.
- **Subagent Execution Guardrail (No Recursion)**:
  - When you are already running as a delegated **`flash` subagent**, you **ARE** the grunt worker: execute `read_file`, `write_file`, `edit_file`, and `run_command` **directly** and return a concise summary to the parent Argon agent. Do not spawn further subagents unless explicitly instructed.

## 3. PRODUCTION PLANNING CONTEXT
- **Reference Document**: Full Tubes & PET scheduling, changeover, and BOM scrap rules live in [PLANNING_RULES.md](file:///d:/Alpha/PLANNING_RULES.md).
- **Tubes Core Rules**: Printing lines (`Printing-03` small/med `12.5–25mm`, `Printing-04` med/large `19–35mm`) target a **22,000 pcs/day** OTIF benchmark (vs `~13.8k–15.9k` YTD actual). Always sequence jobs by **Dia → Nozzle → Base Coat light-to-dark** (`White → Clear/Varnish → Dark/Red/Black`) to avoid 3–4h tooling swaps and 1–2h washdowns; gate on Slug/Cap/Base Coat/Lacquer/Ink availability (`10%` BOM scrap, `2%` on Caps); exclude `(Varnish)` passes from good tube KPIs.
- **PET Core Rules**: `PF Machine` plans at **10,000 bottles/day** (12h / `A+B` shift), **15% BOM scrap**, **6–8h mould swaps** (dedicated `CHANGEOVER` day between formats), skipping Sundays & Pakistan national holidays, with static baseline vs. dynamic `Forecast Finish` tracked via `Scripts/update_plans.py`.
