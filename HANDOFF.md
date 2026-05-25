# HANDOFF.md

> **Read this FIRST every session.** This file is the project's working memory.
> Aliaskar updates the "Last cycle" section at the end of each work session.
> Claude updates "Next cycle" section after diagnosing.

---

## Current state

- **Date:** 2026-05-25
- **Cycle:** 001 (Day 2 complete, awaiting Aliaskar+Claude review before Day 3)
- **Day in 7-day plan:** 2 / 7
- **Current `.leap` file:** `data/snapshots/cycle_000_colleague_baseline.leap` (57 MB, in place)
- **LEAP version target:** 2024.4.0.8 (Aliaskar's local), note colleague used 2024.4.0.6
- **Active blockers (confirmed by colleague, not yet re-verified in this project):**
  1. Nodal distribution sums to 0% (NEMO transmission optimization), diagnosed, fix planned Day 3
  2. 5 broken emission factor unit definitions, diagnosed, fix planned Day 4
  3. 64 scenarios, 6 regions, bloat, to be reduced Day 5

## Day 1 verification (confirmed by agent 2026-05-22)

- LEAP COM connection works (`oLEAP.WorkingDirectory` returned)
- LEAP cold-start takes 3-4 minutes, factor this into all script timings
- One issue fixed during init: em-dash in PowerShell file caused Windows-1252 encoding error
  - **PROJECT RULE:** No Unicode em-dashes, curly quotes, or non-ASCII chars in any `.ps1` file
  - Use `--` not the em-dash char, use straight `"` not curly quotes
- Repo live at `https://github.com/Alyasska/RE_LEAP_Alaskar_22_05` (2 commits)
- Energy balance xlsx in `data/ground_truth/`
- Colleague's .leap file was named `KAZ_workshop exercise.leap` in his folder (not `_for2024`)

## Quick context (for Claude resuming session)

- This is the **prototype phase** (v0.1). Goal: Calculate runs to completion. No calibration yet.
- Aliaskar works on Windows, VS Code, Git. He runs PowerShell/VBScript. He tests in LEAP GUI.
- Claude (me) reads from GitHub repo + uploaded files, writes scripts and docs.
- **Existing proven patterns:** PowerShell wraps VBScript; VBScript uses `CreateObject("LEAP.LEAPApplication")`; bulk writes need `oLEAP.Verbose=0` first; VBS files must be ASCII no BOM.
- **Encoding rule:** all `.ps1` files must be pure ASCII (no Unicode dashes, quotes, or symbols).
- Previous attempt repo (lessons): https://github.com/Alyasska/BTR_LEAP_01

## Last cycle (what just happened)

**Cycle 001 (Day 2 scout), complete 2026-05-25:**

What worked:
- `pip install pyzipper` -- used Anaconda Python at `C:\Users\User\anaconda3\python.exe` (system `python`/`py` not on PATH, both shadowed by MS Store stubs).
- `S01_extract_leap.py` ran clean: **195 files extracted** to `data/extracted/cycle_000_colleague_baseline/`. Breakdown: 5 VBS hooks, 94 NexusDB tables, 13 INI/text, 41 Excel, 42 other. `beforeCalculation.vbs_Safe` confirmed present.
- Colleague's area was already installed in LEAP at `C:\Users\User\Documents\LEAP_16_04\LEAP Areas\kaz_workshop exercise\` (installed 2026-05-21). The folder-level name is `kaz_workshop exercise` (lowercase, space) but the **internal LEAP display name is `KAZ_2024`** -- this is what `oLEAP.ActiveArea.Name` returns.
- `.\scripts\01_scout\S02_audit_model.ps1 -AreaName "kaz_workshop exercise"` completed in **94 seconds** (not the 3-7 min the script predicted). LEAP was probably already warm.
- Report: `data/audit_reports/audit_cycle_001_20260525_104946.md`. Extract summary: `data/audit_reports/extract_cycle_000_colleague_baseline_20260525_104808.md`.

Headline findings (confirmed reliably):
- **BaseYear=2024, FirstScenarioYear=2025, EndYear=2045** -- ALL match prototype targets. Colleague's base-year migration is real.
- 6 regions, **65 scenarios** (PLAN said 64 -- one off, no concern).
- All key transmission branches exist: `Transformation\Electricity Production\Transmission Nodes\KAZ_{North,West,South}` and `Demand\Electricity_Exports\Russia`.

Findings to treat with skepticism (script bugs, see below):
- **ISSUE-001 (Simulation Type read FAILED silently):** Every region returned `ERROR:object required` (`требуется объект` -- mojibake in report). The audit reports "0 scenarios use NetworkSimulation" but **this is a script bug, not a real zero**. Root cause: `Set var = oLEAP.Branches(...).Variable("Simulation Type")` returned Nothing (probably because `On Error Resume Next` was already active and a prior error left state); subsequent `var.ExpressionRS(r.Id, 0)` then throws "Object required". Also `ExpressionRS(r.Id, 0)` uses scenario_id=0 but actual CA id is 1.
- **ISSUE-002 (5 broken-unit branches all reported MISSING):** Could be real (colleague restructured/deleted them) or a path mismatch. Need manual verification in LEAP UI -- search the tree for `Avg Environmental Loading` rows under `Demand\Agriculture` and `Demand\Commercial`.

Cosmetic issues (not blocking):
- Cyrillic strings (region `ResultsShown`, VBS error text) come through as `����` mojibake -- cscript Echo writes in OEM codepage, but PS reads as UTF-8. Will fix in next audit revision (write VBS output to UTF-8 file directly rather than via Echo).

Files committed in this cycle:
- `scripts/01_scout/S02_audit_model.ps1` (new)
- `docs/day2_scout_instructions.md` (new)
- `HANDOFF.md` (overwritten)
- `data/audit_reports/extract_cycle_000_colleague_baseline_20260525_104808.md`
- `data/audit_reports/audit_cycle_001_20260525_104946.md` + `.raw.log`

## Next cycle (what Claude prescribes next)

**Cycle 001, Day 2 scout (estimated 20-30 min wall time):**

Read `docs/day2_scout_instructions.md`, it has full step-by-step.

Quick summary:
1. `pip install pyzipper`
2. `python scripts/01_scout/S01_extract_leap.py data/snapshots/cycle_000_colleague_baseline.leap`
3. Install area in LEAP UI if not already installed (`Area > Install` from .leap file). **Note the exact installed area name.**
4. `.\scripts\01_scout\S02_audit_model.ps1 -AreaName "<exact name>"`
5. Open the resulting audit report and paste its "Verdict" section to Claude

**After Day 2 we will know:**
- The actual installed area name in LEAP -- ANSWERED: collection key = `kaz_workshop exercise`, display name = `KAZ_2024`
- Whether BaseYear is really 2024 in the file -- ANSWERED: YES (also FirstScenarioYear=2025, EndYear=2045 confirmed)
- Exact count of scenarios using NetworkSimulation (Day 3 fix scope) -- **UNANSWERED, S02 SimType reader has a bug, needs script fix before Day 3 reduce scope can be sized**
- Whether 5 broken-unit branches still exist -- INCONCLUSIVE: script says missing, need UI confirmation
- Total scenario/region counts -- ANSWERED: 65 scenarios, 6 regions

## Open questions (waiting on Aliaskar)

- [ ] Exact installed area name in LEAP (will be answered by S02)
- [ ] Confirm BaseYear was actually saved as 2024 in the file (will be answered by S02)
- [ ] Which "Other" scenario should we keep alive besides S0 for the prototype, if any?

## Glossary (avoid re-defining each session)

| Term | Meaning |
|---|---|
| BTR | Biennial Transparency Report, UNFCCC requirement Kazakhstan submits every 2 years |
| BaseYear | The "current accounts" year in LEAP. Historical data ends here. |
| FirstScenarioYear | The first year LEAP projects forward. Usually BaseYear+1. |
| EndYear | Last year of LEAP projection. |
| `.leap` file | AES-encrypted ZIP (password = "LEAP") containing NexusDB tables + VBScript hooks. |
| NEMO | Julia-based optimization solver LEAP uses for transformation/transmission problems. |
| Nodal distribution | Share of demand assigned to each transmission node (KAZ_North/West/South). Must sum to 100% per region. |
| S0 Baseline Historical | The single scenario we care about for the prototype. |
| Scout / Reduce / Fix / Run | Our 4 script tiers, see scripts/ subdirectories. |
| ASCII-only rule | All .ps1 files must be pure ASCII. No em-dash, curly quotes, etc. |
