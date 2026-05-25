# HANDOFF.md

> **Read this FIRST every session.**

---

## Current state

- **Date:** 2026-05-25
- **Cycle:** 003 (Day 2 audit COMPLETE, clean numbers obtained; awaiting Claude+Aliaskar before Day 3)
- **Day in 7-day plan:** 2 / 7 (Day 2 finished; Day 3 shape to be redesigned by Claude based on findings below)
- **Current `.leap`:** `data/snapshots/cycle_000_colleague_baseline.leap`
- **LEAP install:** Russian-localized, 2024.4.0.8

## CRITICAL FINDING from cycle 002

**Two installed areas share the same Areas-collection key** `kaz_workshop exercise`:

| Area (best guess) | BaseYear | What |
|---|---|---|
| Parent area | 2010 | Original SEI workshop file, pre-migration |
| Nested `KAZ_2024` | 2024 | Colleague's migrated version (this is our prototype target) |

`oLEAP.Areas("kaz_workshop exercise").Open` resolves non-deterministically:
- v1 (cycle 001) accidentally opened the 2024 nested area
- v2 (cycle 002) opened the 2010 parent

**Resolution strategy:** address areas by INDEX into `oLEAP.Areas`, not by name. v3 does this.

Bonus: v2 also got stuck on a "save changes?" popup during teardown. Read-only audits should not need to save, but LEAP may flag the area as modified anyway (view state, etc.). The agent dismissed the popup manually. v3 may hit the same; we will harden v4 only if it actually blocks.

## Cycle 002 also revealed

- v1's "0 of 5 broken-unit branches" reading was on the 2024 nested area, suggests colleague may have deleted them, but we cannot trust it because we did not run fuzzy search on the right area
- v2 read the parent area (2010) and found 5/5 still have HAS_LOADING, but that is the pre-migration state, not what we need

We will know the actual state of the 2024 area only after v3.

## Confirmed across both reads

- 6 regions, 65 scenarios in both areas (same model topology, expected)
- BaseYear of nested area really is 2024 (per v1's reading)
- BaseYear of parent area really is 2010 (per v2's reading)
- Transmission node branches exist in both
- Russian-localized LEAP, expect Cyrillic in some text fields

## Last cycle (cycle 003 results) -- 2026-05-25

### Step A: S00_list_areas.ps1 (18.3s, clean)

**12 areas installed.** All live under `C:\Users\User\Documents\LEAP_16_04\LEAP Areas\kaz_workshop exercise\` -- LEAP appears to treat every subdirectory of that folder as a separate area. The `kaz_workshop exercise` directory is functioning as the LEAP Areas root, not as a single area.

`a.BaseYear` etc. are NOT exposed on closed `Area` COM objects on this install (every `AREA_PROP_MISS` for BaseYear/FirstScenarioYear/EndYear/LastModified/etc.). Only `Name` and `Directory` worked. The "Recommended next call" auto-detection at the bottom of the S00 report consequently could not pick a target; I picked manually.

| Index | Name | Directory (relative to LEAP Areas root) |
|---|---|---|
| 1 | Asiana | kaz_workshop exercise/Asiana/ |
| 2 | central asia | kaz_workshop exercise/central asia/ |
| 3 | Freedonia | kaz_workshop exercise/Freedonia/ |
| 4 | Freedonia (Recovered 05-21-26) | kaz_workshop exercise/Freedonia (Recovered 05-21-26)/ |
| 5 | GHG Mitigation Exercise | kaz_workshop exercise/GHG Mitigation Exercise/ |
| **6** | **KAZ_2024** | **kaz_workshop exercise/KAZ_2024/** -- prototype target |
| 7 | kaz_workshop exercise | kaz_workshop exercise/kaz_workshop exercise/ -- 2010 parent (what v2 hit) |
| 8 | Kazakhstan_new | kaz_workshop exercise/Kazakhstan_new/ |
| 9 | Kazakhtsan_new2 | kaz_workshop exercise/Kazakhtsan_new2/ (sic, typo in folder) |
| 10 | Optimization Exercise | kaz_workshop exercise/Optimization Exercise/ |
| 11 | Transport Exercise | kaz_workshop exercise/Transport Exercise/ |
| 12 | WEAP-LEAP Tutorial | kaz_workshop exercise/WEAP-LEAP Tutorial/ |

Two other interesting entries to flag: `Kazakhstan_new` (#8) and `Kazakhtsan_new2` (#9, with typo) -- could be earlier colleague snapshots. Not investigated.

### Step B: S02_audit_model_v3.ps1 -AreaIndex 6 (93.5s, clean teardown, no popup hang)

**Open sanity: PASSED.** Resolved area `KAZ_2024`, ActiveArea=`KAZ_2024`, BaseYear=**2024**.

Headline:
- BaseYear=2024, 6 regions, 65 scenarios. UTF-16 fix worked: `ResultsShown` shows clean Cyrillic (`Ложь`/`Истина`).
- ISSUE-001 SimType: `VAR_IS_NOTHING`. The branch exists, `VariableExists("Simulation Type")` returned True, but `Set var = Branches(path).Variable("Simulation Type")` produced a Nothing object **without throwing**. We never iterated scenarios because of the short-circuit. **Likely root cause:** wrong accessor API. Per `docs/known_issues.md` and the colleague's own VBS, the canonical pattern is `oLEAP.BranchVariable("Transformation\Electricity Production:Simulation Type")` (single accessor, colon-separated path) -- NOT `Branches(path).Variable(name)`. v4 should switch to `BranchVariable`. We still do not know NetworkSimulation pair count.
- ISSUE-002 exact paths: all 5 MISSING in KAZ_2024 (same as v1 saw -- consistent now that we are confidently on the right area).
- ISSUE-002 fuzzy walk: **1590 candidate branches** under Demand. Colleague restructured rather than deleted. Notable:
  - `Demand\Agriculture\Amu Darya\Other\Lubricants\Methane` exists (close cousin of one original path)
  - Agriculture subtree now has Amu Darya, Other, AND Syr Darya region splits
  - `Demand\Commercial` has no `Lubricants` subtree at all -- has Biomass/Bitumen/Coal/LPG/etc. as direct children. The original `Demand\Commercial\Lubricants\Methane` was restructured away.
  - Industry/Iron and Steel still exists; `Top down` subbranch needs re-verification.

### v3 verdict (verbatim from report)

> This audit ran against the correct prototype-target area (BaseYear=2024).
>
> **Day 3 sizing:** 0 NetworkSimulation pairs to neutralize
> **Day 4 sizing:** 0 / 5 exact broken paths, 1590 fuzzy candidates

**Caveat on Day 3 number:** "0 NetworkSimulation" is still untrustworthy because SimType was never read (VAR_IS_NOTHING short-circuited the loop). v4 with `BranchVariable` accessor needed before Day 3 can be sized.

### Cycle 003 artifacts

- `scripts/01_scout/S00_list_areas.ps1` (new)
- `scripts/01_scout/S02_audit_model_v3.ps1` (new)
- `data/audit_reports/area_listing_20260525_123700.md` + `.data.txt`
- `data/audit_reports/audit_cycle_003_idx6_20260525_123832.md` + `.data.txt`

## Recommendation to Claude before Day 3

The audit pipeline now reliably reads ANY indexed area without ambiguity. Two follow-ups before Day 3 can be safely scripted:

1. **v4 SimType fix:** swap `Branches(path).Variable(name)` for `oLEAP.BranchVariable("path:var")`. Re-run on index 6. This finally answers the NetworkSimulation question.
2. **Day 4 path remap:** the 1590 fuzzy matches need filtering to identify branches that have `Avg Environmental Loading` with bad unit metadata. The fuzzy walk currently dumps everything matching name patterns; we want it to filter by `VariableExists("Avg Environmental Loading")` and surface only those, with their unit denominator/numerator if exposed.

Once both are done, Day 3 + Day 4 fixes can be scripted against known paths and known counts.

## Project rules (still in force)

- All `.ps1` files must be pure ASCII (verified per write)
- VBS sources must also be pure ASCII (verified per write)
- VBS writes data to a UTF-16 LE file; PowerShell reads with -Encoding Unicode
- Audit scripts are READ-ONLY; if a save-changes popup appears, choose No
- Address LEAP areas by INDEX when name collisions are possible

## Cycle log so far

- Cycle 000: scaffolding (committed)
- Cycle 001: S02 v1 audit, hit 2024 area accidentally, script had 3 bugs (committed)
- Cycle 002: S02 v2 audit, fixed v1 bugs but hit 2010 area, exposed name collision (committed `2cddf37`)
- Cycle 003: S00 enumerated 12 areas, S02 v3 cleanly audited index 6 (KAZ_2024), BaseYear=2024 confirmed; SimType still uninspected (VAR_IS_NOTHING from wrong accessor); 5 broken paths restructured (1590 fuzzy candidates)

## Open questions

- [x] Which area index corresponds to BaseYear=2024? -- **6** (`KAZ_2024`)
- [ ] On KAZ_2024, how many NetworkSimulation scenarios? -- still unknown, blocked on v4 `BranchVariable` accessor fix (v3 short-circuited at VAR_IS_NOTHING)
- [x] Do the 5 broken branches exist verbatim? -- **NO**, all 5 MISSING. Colleague restructured.
- [ ] Where did the 5 broken paths get remapped to? -- 1590 fuzzy candidates; need filtered re-walk (only those with `Avg Environmental Loading` + bad units)
- [ ] What are `Kazakhstan_new` (idx 8) and `Kazakhtsan_new2` (idx 9)? -- likely colleague's intermediate snapshots, not investigated. Flag for later.
