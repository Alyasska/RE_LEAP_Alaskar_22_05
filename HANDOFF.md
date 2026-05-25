# HANDOFF.md

> **Read this FIRST every session.**

---

## Current state

- **Date:** 2026-05-22
- **Cycle:** 002 (Day 2 audit revision, ready to run)
- **Day in 7-day plan:** 2 / 7 (still on Day 2, not slipping, v2 finishes Day 2 properly)
- **Current `.leap`:** `data/snapshots/cycle_000_colleague_baseline.leap`
- **Area name (collection key):** `kaz_workshop exercise` (display name `KAZ_2024`)
- **LEAP install:** Russian-localized, version 2024.4.0.8

## Confirmed facts from cycle 001 audit (v1)

- BaseYear=2024, FirstScenarioYear=2025, EndYear=2045 -- colleague's migration worked
- 6 regions, 65 scenarios (PLAN said 64, off by one, no problem)
- Transmission node branches all exist (KAZ_North, KAZ_West, KAZ_South)
- Russia exports branch exists
- Demand root branch exists

## Open questions after cycle 001 (v1 had bugs)

- ?? Does any scenario actually use NetworkSimulation? v1 returned 0, but v1 had bugs.
- ?? Do the 5 broken-unit branches still exist? v1 said no, but v1 didn't check fuzzy paths.

**These two questions are what cycle 002 must answer.**

## Cycle 001 (v1 audit) -- bugs found by agent

1. SimType reader returned all errors. Causes: missing Is Nothing check after Variable() Set; missing Err.Clear before each ExpressionRS call; assumed scenario_id=0 is Current Accounts (colleague's own VBS never uses scenario_id=0).
2. cscript stdout in OEM cp866 caused Cyrillic mojibake when PowerShell parsed as UTF-8.
3. Branch existence used exact path match only; no fuzzy fallback if colleague restructured.

Agent diagnosis was correct on all three. Fixes applied in v2.

## Next cycle (cycle 002) -- Day 2 audit v2

Run:
```powershell
.\scripts\01_scout\S02_audit_model_v2.ps1 -AreaName "kaz_workshop exercise"
```

Note: collection key is `kaz_workshop exercise` (with the space, no underscore), per cycle 001.

v2 fixes:
1. VBS writes UTF-16 LE data file (`data/audit_reports/audit_cycle_002_*.data.txt`), PowerShell reads with `-Encoding Unicode`. No mojibake.
2. SimType: `Set var` then `If var Is Nothing`; iterate all 65 scenarios (no scenario_id=0 reliance); `Err.Clear` before each ExpressionRS; per-cell errors logged separately from values.
3. Fuzzy branch walk: dumps every branch under Demand whose FullName contains "Lubricants" / "Methane" / "Nitrous Oxide" / "LPG" (at depth >= 4). If colleague restructured rather than deleted, the new paths appear here.

Expected runtime: 5-10 minutes.

After cycle 002 we should know with confidence:
- NetworkSimulation usage count (real number, not script bug)
- Whether 5 broken branches exist verbatim, exist under different paths, or are truly gone
- Day 4 may become a no-op if colleague already deleted them, which compresses the plan

## Project rules (still in force)

- All `.ps1` files must be pure ASCII (no em-dash, no curly quotes, no Cyrillic)
- `.md` files can have UTF-8 freely
- VBScript files written as ASCII no BOM
- Audit and scout scripts are READ-ONLY; never `.Save`
- One cycle = one commit = one log file

## Glossary (kept identical)

| Term | Meaning |
|---|---|
| BTR | Biennial Transparency Report, UNFCCC requirement, Kazakhstan every 2 years |
| BaseYear | The "current accounts" year. Historical data ends here. |
| FirstScenarioYear | First projection year. BaseYear+1 typically. |
| EndYear | Last projection year. |
| `.leap` file | AES-encrypted ZIP, password "LEAP". |
| NEMO | Julia optimizer for transformation/transmission. |
| Nodal distribution | Per-node share of demand. Must sum to 100% per region. |
| S0 Baseline Historical | The one scenario we care about for the prototype. |
| Collection key vs display name | LEAP areas have an indexing key (folder-derived) and a display name (set in UI). Pass the COLLECTION KEY to oLEAP.Areas(...). |
