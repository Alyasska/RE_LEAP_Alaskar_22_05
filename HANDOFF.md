# HANDOFF.md

> **Read this FIRST every session.** This file is the project's working memory.
> Aliaskar updates the "Last cycle" section at the end of each work session.
> Claude updates "Next cycle" section after diagnosing.

---

## Current state

- **Date:** 2026-05-22
- **Cycle:** 000 (initial setup)
- **Day in 7-day plan:** 1 / 7
- **Current `.leap` file:** `data/snapshots/cycle_000_colleague_baseline.leap` (to be added)
- **LEAP version target:** 2024.4.0.8 (Aliaskar's local) — note colleague used 2024.4.0.6
- **Active blockers:**
  1. 🔴 Nodal distribution sums to 0% (NEMO transmission optimization) — **diagnosed**, fix planned Day 3
  2. 🔴 5 broken emission factor unit definitions — **diagnosed**, fix planned Day 4
  3. ⚠️ 64 scenarios, 6 regions — bloat, to be reduced Day 5

## Quick context (for Claude resuming session)

- This is the **prototype phase** (v0.1). Goal: Calculate runs to completion. No calibration yet.
- Aliaskar works on Windows, VS Code, Git. He runs PowerShell/VBScript. He tests in LEAP GUI.
- Claude (me) reads from GitHub repo + uploaded files, writes scripts and docs.
- **Existing proven patterns:** PowerShell wraps VBScript; VBScript uses `CreateObject("LEAP.LEAPApplication")`; bulk writes need `oLEAP.Verbose=0` first; VBS files must be ASCII no BOM.
- Previous attempt repo: https://github.com/Alyasska/BTR_LEAP_01 — has working M-modules pattern, energy balance xlsx, broken-unit branch list. We keep its lessons, not its code structure.

## Last cycle (what just happened)

**Cycle 000 (Day 1 setup) — completed 2026-05-22:**
- `init_env.ps1` ran; all folders confirmed/created. **One issue fixed:** line 111 had a UTF-8 em dash `—` that PowerShell read as Windows-1252, causing a string termination error. Replaced with `--`.
- LEAP COM connection test (`cscript` + VBScript `CreateObject("LEAP.LEAPApplication")`) was still in progress at handoff time — LEAP may be slow to start or waiting on a splash/license dialog. **Verify by running `.\init_env.ps1` yourself and confirming you see `OK: LEAP COM responsive`.**
- Colleague's `.leap` file found at `C:\Users\User\Documents\Aliaskar Bekishev _LEAP_2005\KAZ_workshop exercise.leap` (note: NOT named `kaz_workshop_exercise_for2024.leap` — the working file is `KAZ_workshop exercise.leap`, 57 MB). Copied to `data/snapshots/cycle_000_colleague_baseline.leap`.
- `kaz_energy_balance_2024_leap.xlsx` (19 952 bytes) downloaded from `Alyasska/BTR_LEAP_01` repo (`enery balance/` folder — note typo in folder name) → placed in `data/ground_truth/`.
- `git init` done; `main` branch created; remote set to existing GitHub repo `https://github.com/Alyasska/RE_LEAP_Alaskar_22_05` (repo already existed, was empty). First commit `50bb5f3` pushed successfully.
- **Day 1 is complete.** All checklist items from PROTOTYPE_PLAN.md Day 1 table are done.

## Next cycle (what Claude prescribes next)

> _Claude: update after each diagnosis._

**Cycle 001 — Day 2 scout:**
- Run `scripts/01_scout/S01_extract_leap.py` to decrypt and dump the colleague's .leap
- Upload result `data/audit_reports/audit_cycle_001.md` for review
- Confirm 5 broken-unit branches still exist + nodal distribution branches still present

## Open questions (waiting on Aliaskar)

- [ ] Which "Other" scenario should we keep alive besides S0 for the prototype, if any?
- [ ] Is the LEAP COM API guaranteed available with his current install? (Repo confirms YES based on prior connection test, but re-verify with fresh project)
- [ ] What's the exact LEAP install path? Some scripts may hardcode it.

## Glossary (avoid re-defining each session)

| Term | Meaning |
|---|---|
| BTR | Biennial Transparency Report — UNFCCC requirement Kazakhstan submits every 2 years |
| BaseYear | The "current accounts" year in LEAP. Historical data ends here. |
| FirstScenarioYear | The first year LEAP projects forward. Usually BaseYear+1. |
| EndYear | Last year of LEAP projection. |
| `.leap` file | AES-encrypted ZIP (password = "LEAP") containing NexusDB tables + VBScript hooks. |
| NEMO | Julia-based optimization solver LEAP uses for transformation/transmission problems. |
| Nodal distribution | Share of demand assigned to each transmission node (KAZ_North/West/South). Must sum to 100% per region. |
| S0 Baseline Historical | The single scenario we care about for the prototype. |
| Scout / Reduce / Fix / Run | Our 4 script tiers — see scripts/ subdirectories. |
