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

> _Aliaskar: write 3-5 lines after each work session._

**Cycle 000 (initial setup) — placeholder, update after Day 1 done:**
- Created `RE_LEAP_Alaskar_22_05` folder
- Dropped Claude's starter pack
- Git initialized, pushed to GitHub
- Copied colleague's `.leap` to `data/snapshots/`
- No errors.

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
