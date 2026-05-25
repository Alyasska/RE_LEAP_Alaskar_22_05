# HANDOFF.md

> **Read this FIRST every session.**

---

## Current state

- **Date:** 2026-05-25
- **Cycle:** 004 (Day 2 audit revision 4, COMPLETE; clean numbers + scope clarity for Day 3)
- **Day in 7-day plan:** 2 / 7 (Day 2 done after cycle 004; Day 3 scope known)
- **Current `.leap`:** `data/snapshots/cycle_000_colleague_baseline.leap`
- **LEAP install:** Russian-localized, 2024.4.0.8

## Confirmed from cycle 003

- **Area index 6** is `KAZ_2024` (the prototype target, BaseYear=2024)
- **Area index 7** is `kaz_workshop exercise` (the pre-migration parent, BaseYear=2010)
- LEAP install also has 10 other unrelated areas (Asiana, Freedonia, GHG Mitigation Exercise, etc.)
- Index addressing is reliable; `Open sanity: PASSED` per v3
- The 5 broken-unit branches do not exist at their original paths in KAZ_2024 (confirmed twice independently). They were either deleted or restructured.
- Demand subtree has 1590 leaves matching Lubricants/Methane/Nitrous Oxide/LPG patterns. Most are noise (every region x fuel x gas combo). Real fix scope = subset with Avg Environmental Loading and bad units.

## Cycle 003 bug (mine to fix in v4)

- v3 used `oLEAP.Branches(path).Variable(name)` for SimType. This accessor returns `Nothing` silently on this LEAP install.
- Canonical accessor (per `docs/known_issues.md` ISSUE-001 and colleague's own VBS): `oLEAP.BranchVariable("Path:VarName")`.
- v4 switches to the canonical accessor.

## Cycle 003 also revealed two dead-weight areas

- Index 8: `Kazakhstan_new` -- old colleague snapshot, status unknown
- Index 9: `Kazakhtsan_new2` (sic, typo in name) -- ditto

Not investigating these in cycle 004. Will probe with a quick S00 v2 later if relevant to Day 5 cleanup.

## Last cycle (cycle 004 results) -- 2026-05-25

S02 v4 -AreaIndex 6, completed in **41.2 seconds** (faster than v3's 93.5s -- fewer COM operations because fuzzy walk dropped). Clean teardown, no popup.

### Accessor fix CONFIRMED working

`SIMTYPE_STATUS = VAR_OK`. The canonical `BranchVariable("path:var")` accessor reads cleanly. All 390 region-scenario pairs returned non-empty values without errors.

Curiosity for the record: the hook file itself (which we now have) shows the colleague uses `leap.branches(path).Variable(name).expressionrs(...) = expr` in **chained-call write context** and that pattern works. So `Branches.Variable` is not fundamentally broken -- it only returns Nothing when extracted via a standalone `Set var = ...`. Use `BranchVariable` for reads, the chained pattern for writes (matches colleague's idiom).

### Day 3 scope is huge but uniform

**All 390 (region x scenario) pairs are explicitly set to `NetworkSimulation(Pipeline)`.** Every single one. 6 regions x 65 scenarios = 390. Zero errors, zero non-empty-non-Network values, zero inheritance from CA.

Important sub-finding: the argument is **`Pipeline`, not `Transmission`**. This is gas/oil pipeline NetworkSimulation, NOT the electricity transmission optimization. The nodal-distribution bug in `beforeCalculation.vbs` is a **separate** electricity-side issue. We now have two transmission-flavored things to deal with in Day 3, not one:
1. `Simulation Type = NetworkSimulation(Pipeline)` at module level, 390 explicit pairs
2. `beforeCalculation` hook writing `Nodal Distribution` on KAZ_North/West/South electricity nodes

Day 3 options:
- **A**: clear/set-to-Standard the SimType at all 390 pairs (large but mechanical -- BatchSecondVar pattern)
- **B**: neutralize the beforeCalculation hook (rename / empty body / `Exit Sub` early)
- **Probably both** are needed. Either alone may not be sufficient.

### Hook file IS PRESENT (despite v4 verdict line saying otherwise)

`beforeCalculation.vbs` exists in `C:\Users\User\Documents\LEAP_16_04\LEAP Areas\kaz_workshop exercise\KAZ_2024\` -- 19,045 bytes, 322 lines. The file's first comment block is the same one cited in `docs/known_issues.md` ISSUE-001.

Confirmed by grep on the actual file:
- Contains `nodal` at lines 3, 79, 80, 81, 180, 182, 184, 189, 194, 195, ...
- Contains `KAZ_North` at lines 53, 60, 61, 62, 69, 79, ...
- Comment line 1: "Temporary before calculation script designed to work around some LEAP bugs, including: Improper populating of nodal distribution variables in NEMO"

The full hook is preserved and active in KAZ_2024.

### v4 parser bug to flag (PROJECT RULE candidate)

The verdict line at the bottom of the v4 report reads:
> "ISSUE-001 hook: NOT detected as a separate file. Possibly inactive in KAZ_2024 already."

**This is wrong.** The table above the verdict correctly shows `HOOK_FOUND` with `Has 'nodal' = Истина` and `Has 'KAZ_North' = Истина`. The PowerShell parser at line ~459 does:

```powershell
$anyNodal = ($hookFound | Where-Object { $_.HasNodal -eq "True" }).Count -gt 0
```

VBS `CStr(True)` on a Russian-localized install returns **`Истина`**, not `"True"`. The comparison is False, the wrong-branch verdict fires. **Project rule for the future:** when VBS writes booleans, write the integer (`If x Then writeline "1" Else writeline "0"`), never `CStr(boolean)`. Or alternatively the PS side compares against `Истина OR True`. Adding this to the project rules section below.

### ISSUE-002 (broken paths) -- consistent

0 / 5 verbatim paths present. **Third independent confirmation.** Day 4 needs a fuzzy walk with filter on `VariableExists("Avg Environmental Loading")` -- per v5 plan.

### Cycle 004 artifacts

- `scripts/01_scout/S02_audit_model_v4.ps1`
- `data/audit_reports/audit_cycle_004_idx6_20260525_125131.md` + `.data.txt`
- `logs/S02v4_idx6_20260525_125131.log`

## v5 plan (next cycle if v4 unblocks Day 3 sizing)

Filtered fuzzy walk: walk Demand subtree, but at each branch test `VariableExists("Avg Environmental Loading")` AND try to read its `Unit` / `UnitDenominator` properties (whatever the COM exposes -- we will probe property names like in S00). Output a small list of branches with the loading variable, annotated with whether their unit metadata looks broken. This replaces the 1590-row noise dump.

We do NOT yet know whether the unit metadata is readable via COM. v5 will find out.

## Project rules (updated)

- All `.ps1` and VBS source: pure ASCII, verified at write time
- VBS writes UTF-16 LE data files; PowerShell reads with -Encoding Unicode
- Address LEAP areas by INDEX (avoid name-collision ambiguity)
- Audit scripts READ-ONLY; if save popup appears, click No
- **Variable reads:** use `oLEAP.BranchVariable("path:var")` accessor (canonical)
- **Variable writes:** use chained `oLEAP.Branches("path").Variable("name").ExpressionRS(r,s) = expr` (matches colleague's idiom and works)
- **Never use `Set var = oLEAP.Branches(path).Variable(name)`:** returns Nothing on this install. Use BranchVariable for the intermediate object.
- **NEW (cycle 004):** When VBS writes booleans to the data file, write `1`/`0` (integers), never `CStr(boolean)`. On Russian-localized LEAP installs `CStr(True) = "Истина"` and PowerShell `-eq "True"` comparisons fail. Either: write integers, or compare against both English+localized strings.

## Cycle log

- 000: scaffolding
- 001: S02 v1 audit (accidentally hit KAZ_2024 via name collision, had 3 bugs)
- 002: S02 v2 audit (hit 2010 parent via name collision, popup hang)
- 003: S00 + S02 v3 (index addressing, accessor bug exposed by agent)
- 004: S02 v4 -- BranchVariable accessor confirmed working; 390 NetworkSimulation(Pipeline) pairs found (all region-scenario combos); beforeCalculation.vbs hook confirmed PRESENT (322 lines, has nodal + KAZ_North refs); parser bug in verdict line (Cyrillic Истина vs "True")

## Open questions

- [x] On KAZ_2024 with the right accessor: how many SimType expressions are Network? -- **390 / 390 (every (region, scenario) pair) are `NetworkSimulation(Pipeline)`**
- [x] Does the area folder contain a beforeCalculation hook with nodal distribution refs? -- **YES**, 19045 bytes / 322 lines at `kaz_workshop exercise\KAZ_2024\beforeCalculation.vbs`, contains both nodal-distribution logic and KAZ_North/West/South references
- [ ] How many branches under Demand actually have Avg Environmental Loading? (v5 -- next cycle)
- [ ] Can we read unit metadata (numerator/denominator) via COM? (v5 probe)
- [ ] Does NetworkSimulation(Pipeline) actually trigger the same nodal-distribution error, or only the SimulationType=NetworkSimulation electricity path? (Day 3 design decision -- if pipeline is benign, maybe only the hook needs neutralizing)
