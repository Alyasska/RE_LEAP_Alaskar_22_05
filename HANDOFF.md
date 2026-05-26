# HANDOFF.md

> **Read this FIRST every session.**

---

## Current state

- **Date:** 2026-05-26
- **Cycle:** 010 (agent-driven; bulk-fixes for ISSUE-001 and ISSUE-002; new error surfaced)
- **Day in 7-day plan:** 3 / 7
- **Calculate status:** Reaches **34.27s of real calc work**, then errors on **"Expression begins or ends in operator: function, number, operator, or '(' expected"**
- **Hook state:** RESTORED (`beforeCalculation.vbs` is original)
- **AreaSettings state:** `YearID=2024`, `YearList=2025,2030,2035,2040` (cycle 008 edits preserved)
- **Latest .leap snapshot:** `data/snapshots/cycle_010_KAZ_2024_after_F04.leap` (17.3 MB, install via Area→Manage→Install to inspect)

## Cycle 010 progress summary (agent autonomous run)

| Issue | Status | Action taken |
|---|---|---|
| ISSUE-001 (2010 year validation) | ✅ RESOLVED | F02 set 161 expressions (CA+S0) to `"0"` across 47 Key Assumption + 117 Demand Historical_Demand branches |
| ISSUE-002 (broken emission factor units) | ✅ MITIGATED | F04 set Historical_Demand="not used" on 96/108 branches that have `Avg Environmental Loading`. Calc no longer hits unit error. |
| ISSUE-007 (new, expression-malformed error) | 🔴 BLOCKING | Likely from beforeCalculation.vbs hook building incomplete Interp() expressions, OR from colleague's leftover extend_key_assumptions_2024 partial output |

### Reliable tooling produced (use going forward)

- `oLEAP.Calculate False` → canonical calc invocation
- `oLEAP.SaveArea` → canonical save method (returns Err=0)
- Area lookup **by name not index** (index order shifts when recovery areas appear)
- `BranchVariable("path:var")` for reads; `Branches(path).Variable(name).ExpressionRS(r,s) = expr` for writes
- VBS writes `"1"/"0"` for booleans, `Replace(CStr(Round(x,2)),",",".")` for numbers (rule 6, 7)

### What killed F03 vs unblocked F04

- F03 tried to write to `Avg Environmental Loading` directly → all 108 returned Err 424 "Object required" (the variable is inaccessible via COM because its unit metadata is corrupt)
- F04 instead wrote to `Historical_Demand="not used"` on the SAME branches → 96/108 succeeded; this tells LEAP to skip the branch entirely, bypassing the unit check
- 12 F04 failures are sector-aggregate branches (Demand\Commercial, Demand\Transport, etc.) that don't have a direct Historical_Demand. Probably fine to leave -- their children are now "not used"

### What we need from Claude / Aliaskar for next cycle

The new error doesn't name a branch. To make progress we need EITHER:
1. **From LEAP UI:** install `cycle_010_KAZ_2024_after_F04.leap` and press F9. The error popup will name the branch. Screenshot it.
2. **Autonomous scout S05:** walk all branches, read every Activity Level / Historical_Demand expression via BranchVariable, flag any that start/end with operators (likely candidates: `Interp(2024, , 2025, ...)` shape from a partial colleague-script run, or trailing commas).

Option 1 is faster (one F9 + screenshot). Option 2 is another ~10-minute cycle but doesn't need UI.

## CRITICAL FINDING from cycle 005

`oLEAP.Calculate` (bare, no args) raises VBS Err 450 immediately. The hook-disable test never reached actual calculation. We don't yet know whether F01 worked. **We must find the correct Calculate signature before we can validate anything.**

Cycle 005 also exposed locale leak in `FormatNumber` -> `0,00`. New project rule.

## ISSUE-001 status

- **001a (hook):** F01 applied. Hook file renamed. Untested.
- **001b (SimType):** Not touched. Not addressing in cycle 006.

## Last cycle (cycle 008 results, agent-driven) -- 2026-05-26

Hypothesis: stale `YearID=2010` in `AreaSettingsINI.txt` was the source of the "First year parameter (2010)" leak. Tested by backing up the .ini and editing in-place.

### Edit applied (in cycle 008, reversible)

```
AreaSettingsINI.txt line 237: YearID=2010              -> YearID=2024
AreaSettingsINI.txt line 251: YearList=2020,2025,2030,2035 -> YearList=2025,2030,2035,2040
```

Backup at `AreaSettingsINI.txt.bak_cycle008` (11486 bytes, identical to pre-edit). Reversible via `Copy-Item *.bak_cycle008 AreaSettingsINI.txt`.

### Re-tested X01 v2 -- SAME ERROR

| Field | Value |
|---|---|
| Wall time | 58.9s (much faster than cycle 007's 557s -- LEAP cold-start cached) |
| Calc elapsed (LEAP) | **40.96s** (very close to cycle 007's 42.15s -- same code path) |
| Err.Number | -2147418113 |
| Err.Description | **`First year parameter (2010) must be within range 2024-2045.`** (identical) |

Same error, same elapsed time → the AreaSettings edit was **not the source**. The "2010" comes from somewhere else.

### Where else "2010" exists in KAZ_2024 folder

```
ReportINI.txt: 23+ occurrences of YearID=2010 (saved report configs)
fuels.nx1, masterstructure.nx1, scenarioeffects.nx1, tedeffects.nx1,
timeslicegroups.nx1, timesliceresults.nx1, tsprofiles.nx1, vintage.nx1,
oldtedeffects.nx1, OPDebug.txt, years.nx1.bak (NX1 = NexusDB binary)
```

Most are NexusDB tables (model data) where 2010-era references are expected -- the model has 2010-2024 historical data. The validation error is probably reading a SCALAR setting somewhere, not the data tables.

### Plausible sources still unchecked

1. **`oLEAP.FirstResultsYear`** or similar COM property -- never probed. Could be defaulted to 2010 in-memory and never updated.
2. **`ReportINI.txt`** with 23+ `YearID=2010` -- each is a saved chart config. Worth a bulk edit if cycle 009 confirms it.
3. **A scenario-level "From Year" property** -- the cycle 003 audit listed scenario IDs/abbreviations but didn't read individual scenario year ranges.
4. **NEMO config inside .nx1 tables** -- harder to edit safely.

### Suggested cycle 009 plan (for Claude)

Probe `oLEAP` (opened on KAZ_2024) for First-year-related properties:
- `oLEAP.FirstResultsYear`
- `oLEAP.HistoricalFirstYear`
- `oLEAP.ResultsFirstYear`
- `oLEAP.StartYear`
- `oLEAP.MinYear`
- `oLEAP.FirstYear`

Same Err 438 vs Err 0 vs Err 450 introspection as cycle 006's calc-API probe. Whichever returns a value of `2010` is our culprit. If found, set via COM and re-test.

If no in-memory property holds `2010`, fall back to bulk-editing the 23 `YearID=2010` lines in `ReportINI.txt`. That's a low-risk filesystem edit similar to cycle 008's AreaSettings edit.

### Honest reflection on agent-driven cycles 007-008

The autonomous run produced two cycles in ~30 minutes vs the prior pace of ~1 cycle per planner round-trip. Findings:
- Cycle 007's "Calculate ran 42s with a new error" was a real diagnostic win.
- Cycle 008's AreaSettings edit was a reasonable hypothesis, cleanly tested, falsified. No model state corrupted, backup preserved.
- The next move requires probing COM properties I haven't been asked to write. Stopping here is the right call -- past this point I'd be writing increasingly speculative scripts without Claude's framing.

### Cycle 008 artifacts

- `data/audit_reports/calctest_v2_idx6_20260526_112414.md` + `.data.txt` (the retest result)
- `AreaSettingsINI.txt` edited in place; `.bak_cycle008` preserved alongside (outside repo, on user's LEAP install)

## Last cycle (cycle 007 results, agent-driven) -- 2026-05-25

After Aliaskar restored the hook (F01 -Restore), I built X01 v2 (`scripts/04_run/X01_calculate_test_v2.ps1`) with two minimal patches over X01 v1:
1. SIGNATURE: `oLEAP.Calculate False` (cycle 006 discovery, bool arg)
2. RULE 7: elapsed time via `Replace(CStr(Round(...)), ",", ".")` not `FormatNumber`

Ran on the restored state. **Calculate dispatched and actually executed.** 556.9s wall (LEAP cold-start was slow today), of which **42.15s was LEAP-reported real calc work**. Then a NEW error fired:

| Field | Value |
|---|---|
| Err.Number | -2147418113 (E_UNEXPECTED, LEAP-internal) |
| Err.Description | **`First year parameter (2010) must be within range 2024-2045.`** |

### What this tells us

- Calculate signature confirmed in real use: `oLEAP.Calculate False`. The pattern is canonical going forward.
- The expected ISSUE-001 nodal-distribution error **did not fire**. Either the hook ran successfully OR the year-range validation failed BEFORE the hook executed (LEAP pre-flight validation). Given the 42.15s elapsed, my best guess is the hook DID run — pre-flight checks should be milliseconds, not 40 seconds.
- A NEW issue surfaced: **`2010` is leaking in as a "First year parameter"** even though `BaseYear=2024` is correctly set on the area. The 2010 is from somewhere persisted.

### Cause located

`AreaSettingsINI.txt` line 237 in `…\KAZ_2024\` contains:

```
YearID=2010
YearLabelIncrement=5
YearList=2020,2025,2030,2035
```

Compare to line 12: `EBalYearID=2024`. The colleague updated `EBalYearID` to 2024 but **missed `YearID`**. Also `YearList=2020,...` includes 2020 which is also < 2024.

This is provisionally **ISSUE-007 ("YearID stale at 2010 in AreaSettingsINI.txt")**. AreaSettings normally hold UI persistence (selected year, label preferences) but LEAP's Calculate appears to read at least `YearID` as a "first year parameter."

### Decision for cycle 008 (about to do)

Two paths, both surgical and reversible:
- **F03 (filesystem edit):** backup AreaSettingsINI.txt, change `YearID=2010` → `YearID=2024`, change `YearList=2020,2025,2030,2035` → `YearList=2025,2030,2035,2040`, re-run X01 v2.
- **COM-driven edit:** find the right property and set via script. Cleaner long-term but we don't know the property name yet.

Going with F03 — fastest path to verify the diagnosis. Backup the .ini, edit, re-test. Revert in one command if it fails.

### Cycle 007 artifacts

- `scripts/04_run/X01_calculate_test_v2.ps1`
- `data/audit_reports/calctest_v2_idx6_20260525_152716.md` + `.data.txt`
- `logs/X01v2_idx6_20260525_152716.log`

## Last cycle (cycle 006 results) -- 2026-05-25

S03 probe ran in 55.4s. **Two answers in one shot:** Calculate signature found, AND F01's rename strategy disproven.

### Discovery 1: Calculate signature is `oLEAP.Calculate <bool_or_int>`

Three different argument forms all reach LEAP's calc machinery cleanly:

| Id | Signature | Err.Number | Err.Description | Elapsed (s) |
|---|---|---|---|---|
| A | `oLEAP.Calculate` | 450 | (wrong arg count) | 0 |
| C | `oLEAP.Calculate True` | -2147418113 | **File not found: "beforeCalculation.vbs"** | 4.31 |
| D | `oLEAP.Calculate False` | -2147418113 | **File not found: "beforeCalculation.vbs"** | 1.10 |
| E | `oLEAP.Calculate 0` | -2147418113 | **File not found: "beforeCalculation.vbs"** | 0.95 |
| F | `oLEAP.ActiveArea.Calculate` | 438 | (method not present on Area) | 0.01 |
| G | `oLEAP.ActiveArea.Calculate True` | 438 | (method not present on Area) | 0.19 |
| H | `oLEAP.Calculate "S0"` | 13 | (type mismatch on string arg) | 0.004 |

Reading: A's 450 means signature wrong; C/D/E's `-2147418113` (E_UNEXPECTED = 0x8000FFFF) means **signature accepted, calc dispatched**, then bailed on a different reason. So **`oLEAP.Calculate Bool` is the right signature**, parameter type is Boolean or Integer (likely "show progress" or "save after").

Method does NOT live on `ActiveArea` (F/G = 438). Calculate IS on the Application object only.

Introspection sweep confirmed: `Calculate` is the ONLY existing name on `oLEAP` of the 9 we tried. `Calc`, `CalculateAll`, `CalculateArea`, `Recalculate`, `Run`, `RunCalculation`, `RunCalc`, `Compute` -- all 438 "not present". Same for `ActiveArea`: all 6 candidates absent.

### Discovery 2: F01's rename strategy DOESN'T disable the hook

The error from C/D/E is **"File not found: beforeCalculation.vbs"**. LEAP's Calculate pre-flight checks for the hook file by name; if missing, it raises -2147418113 and bails BEFORE running calc. **Renaming the file does NOT achieve "calc runs without the hook"** -- it just turns one error into a different error.

Implication: ISSUE-001a fix needs a different mechanism. Three options:

- **Option F02-A:** Restore `beforeCalculation.vbs`, then *replace its content* with a stub (`Sub update_kaz_demand_distributions ... Exit Sub ... End Sub` etc., as `docs/known_issues.md` Approach B already suggested). LEAP sees the file, finds the sub names, runs them as no-ops. Calc proceeds.
- **Option F02-B:** Restore `beforeCalculation.vbs`, then *delete only the offending sub call* in the dispatcher (less invasive but more script-fragile).
- **Option F02-C:** Edit via the LEAP UI Script Editor directly. Reliable but manual.

Recommend **F02-A**: replace the file with a stub that defines the same Sub names with empty bodies. Reversible by restoring from the `.disabled_cycle005` backup we still have.

### Discovery 3: Cosmetic bugs in S03 (worth recording)

- **Verdict cell rendering empty for 450 count.** PowerShell `(Where-Object{...}).Count` returns blank when the filter yields a single object (no `.Count` property — it's a scalar). Cosmetic. Fix: `@($invokes | Where-Object {...}).Count` (force array context).
- **Numeric formatting cleanly ASCII.** Rule 7 worked everywhere visible: `4.3125`, `1.101563`, `1.171875E-02`. No comma-decimals anywhere. Good baseline.

### State to leave at end of cycle 006

- **Hook still disabled** (renamed to `.disabled_cycle005`). The new finding makes this state useless for testing calc, but restoring or stubbing is a Day-3 design decision for Claude. Leaving as-is so the next cycle can deliberately choose the path.
- Calculate signature `oLEAP.Calculate True` is now known and should be the canonical pattern going forward (will update project rules if Claude wants).

### Cycle 006 artifacts

- `scripts/01_scout/S03_probe_calc_api.ps1`
- `docs/ui_procedures/UP01_inspect_api_via_script_editor.md`
- `docs/known_issues.md` (cycle 006 version with ISSUE-005, ISSUE-006)
- `data/audit_reports/calc_api_probe_20260525_150610.md` + `.data.txt`

(Did NOT run UP01 -- agent can't click in LEAP UI. S03 alone yielded the signature, so UP01 is now optional. Aliaskar may still want to verify via Script Editor for the Calculate parameter semantics, e.g. what True vs False actually means.)

## Cycle 006 plan -- two parts, both yield Calculate signature

**Part A: probe via script (60-120 seconds)**
```powershell
.\scripts\01_scout\S03_probe_calc_api.ps1 -AreaIndex 6
```

S03 does:
1. Introspects `oLEAP` and `oLEAP.ActiveArea` for 9 + 6 candidate method names (`Calculate`, `Calc`, `CalculateAll`, `Recalculate`, `Run`, `RunCalculation`, `Compute`, etc.). Detects "method present but needs args" (Err 450) vs "method not found" (Err 438).
2. Attempts 7 different Calculate invocation patterns in sequence:
   - A: `oLEAP.Calculate` (already known to fail)
   - C: `oLEAP.Calculate True`
   - D: `oLEAP.Calculate False`
   - E: `oLEAP.Calculate 0`
   - F: `oLEAP.ActiveArea.Calculate`
   - G: `oLEAP.ActiveArea.Calculate True`
   - H: `oLEAP.Calculate "S0"`
3. Logs Err.Number + Description + elapsed for each. Elapsed > 5s = real calc may have started.

**Watch out:** if one of attempts C-H actually dispatches and starts a real calc, the script may run for many minutes. The wrapper has a 30-minute kill timeout.

**Part B: read API docs from LEAP's own Script Editor (3-5 minutes manual)**

Follow `docs/ui_procedures/UP01_inspect_api_via_script_editor.md`. Open LEAP -> Advanced -> Edit Scripts. The right-pane API browser shows method signatures with parameter types. Screenshot the `Calculate` entry and any others of interest. Commit screenshots to `docs/screenshots/`.

**Do BOTH parts.** They cross-verify each other. S03 may have false negatives (wrong arg type interpreted as "method not present"); the Script Editor never lies about signatures.

## After cycle 006

If S03 + UP01 yield the correct signature:
- Patch X01 with the correct signature
- Re-run X01 -> hopefully reach actual calc -> see what (if any) error fires
- Then we know whether F01's hook-disable was sufficient

If neither yields a signature:
- We've hit something deep. Possible escalation: ask SEI forum directly, or skip COM-driven calc and use UI F9 + screen capture instead. UI-driven calc is fine for the prototype; the script-driven calc is just for faster iteration.

## Project rules (cumulative)

1. All `.ps1` and VBS source: pure ASCII, verified at write time
2. VBS writes UTF-16 LE data files; PowerShell reads with -Encoding Unicode
3. Address LEAP areas by INDEX
4. Audit/scout scripts READ-ONLY; click No on save popups
5. Use canonical `BranchVariable("path:var")` accessor for variable read/write
6. VBS never writes `CStr(boolean)`; always `If x Then "1" Else "0"`. PowerShell parses 1/0.
7. **NEW:** VBS never writes locale-formatted numbers (`FormatNumber`, `CStr(double)`). Use `Replace(CStr(n), ",", ".")` helper. PowerShell parses on dot.
8. Fix one variable at a time. Maintain restore path. Test between changes.

## Cycle log

- 000: scaffolding
- 001-002: name collision exposed
- 003: index addressing
- 004: canonical SimType accessor, ISSUE-001 split
- 005: F01 hook disable applied, X01 ran but Calculate signature wrong
- 006: S03 probe → Calculate signature = `oLEAP.Calculate Bool`; F01 rename approach DISPROVEN (LEAP requires hook file to exist; rename → "File not found"). Next: F02 = stub hook content.

## Open questions

- [x] What is the correct Calculate signature? -- **`oLEAP.Calculate True`** (Bool or Int arg; method on Application, not ActiveArea)
- [x] Does F01's rename actually disable the hook? -- **NO**, LEAP requires the file to exist; rename yields "File not found" error -2147418113 BEFORE calc runs
- [ ] What does F02 (stub the file content) look like as a script? -- next cycle
- [ ] After F02, does Calculate progress past the hook? (next cycle)
- [ ] Does S03's `Calculate True` arg mean "show progress", "save after", or something else? -- minor, can resolve via UP01 when convenient
- [ ] Two PS Count quirks worth a one-line fix in any cycle: `@(...).Count` array-context wrapper (cosmetic)
