# HANDOFF.md

> **Read this FIRST every session.**

---

## Current state

- **Date:** 2026-05-25
- **Cycle:** 005 (Day 3 fix B attempted; F01 succeeded, X01 hit a script bug not a model error)
- **Day in 7-day plan:** 3 / 7 (need a one-line X01 patch for cycle 006 before we can test the fix)
- **Current `.leap`:** `data/snapshots/cycle_000_colleague_baseline.leap` (KAZ_2024 area is index 6)
- **LEAP install:** Russian-localized, 2024.4.0.8

## CRITICAL FINDINGS from cycle 004

1. **SimType reader works with `BranchVariable("path:var")`** -- 390/390 pairs are uniformly `NetworkSimulation(Pipeline)`. Zero variation. This is **gas/oil pipeline** optimization, not electricity nodal distribution.
2. **Hook file is live**: `…\KAZ_2024\beforeCalculation.vbs` (19045 bytes, 322 lines, contains `KAZ_North` references and nodal distribution logic).
3. **The two parts of ISSUE-001 are independent failure modes.**

## ISSUE-001 split into 001a and 001b

- **001a: hook computes invalid nodal distributions** -- this is what produces the "sums to 0%" error. Filesystem fix (rename `beforeCalculation.vbs` to disabled). Day 3 fix B.
- **001b: 390 SimType=NetworkSimulation(Pipeline) for transformation** -- separate problem, may or may not block calc. We don't touch this in cycle 005. If calc still fails after 001a fix, we address 001b in cycle 006.

## Cycle 004 also produced a parser bug

v4 PowerShell parser checked `HasNodal -eq "True"` but VBS `CStr(boolean)` returns localized `Истина` on Russian LEAP install. Verdict line was wrong despite data being correct.

**New project rule:** VBScript NEVER writes `CStr(boolean)`. Always: `If x Then "1" Else "0"`. PowerShell parses on "1"/"0" only.

## Last cycle (cycle 005 results) -- 2026-05-25

### F01 -- SUCCEEDED
- Probed area dir via COM in 20.1s, resolved area `KAZ_2024` at `C:\Users\User\Documents\LEAP_16_04\LEAP Areas\kaz_workshop exercise\KAZ_2024\`
- Renamed `beforeCalculation.vbs` (19045 bytes) → `beforeCalculation.vbs.disabled_cycle005`
- **Reversible**: `.\scripts\03_fix\F01_neuter_before_calc_hook.ps1 -AreaIndex 6 -Restore`
- (Skipped LEAP-UI manual verification step -- agent can't click in UI; Aliaskar to confirm if needed)

### X01 -- SCRIPT BUG, NOT MODEL ERROR

22.1s wall time (mostly LEAP cold start). The hook fix could NOT be tested because Calculate itself never ran.

| Field | Value |
|---|---|
| Calc elapsed (LEAP-reported) | **0,00s** (note Russian-locale decimal comma; we hit 0 ms of actual calc) |
| Err.Number | **450** |
| Err.Description (RU) | `Недопустимое число аргументов или присвоение значения свойства` |
| Translation | **"Wrong number of arguments or invalid property assignment"** |
| Phase reached | `calc_done` (so the VBS got past `Err.Clear ; oLEAP.Calculate ; capture Err`) |

VBS error 450 is the **VBScript runtime** error for bad method dispatch -- it's not a LEAP error, it's COM saying our call signature is wrong. The X01 VBS does:

```vbscript
Err.Clear
oLEAP.Calculate   ' <-- bare call, no args, no parens
err_num = Err.Number
```

The `Calculate` method on `LEAP.LEAPApplication` is not callable this way on this install. Three plausible patterns to try in cycle 006:

1. `oLEAP.Calculate True` (or `False`) -- common LEAP signature takes a boolean (e.g. SaveAfter or ShowProgress)
2. `oLEAP.Calculate(scenarios)` -- requires scenario id list/object
3. `oLEAP.ActiveArea.Calculate` -- method may live on Area, not Application

`Err.Number 450` plus 0,00s LEAP-reported elapsed = method dispatch failed immediately. We never reached the hook, so we still don't know whether F01 actually fixes ISSUE-001a.

### State to leave at end of cycle 005

- F01 change is **left in place** (hook disabled). Don't restore yet; we'll need it disabled when X01 v2 actually runs.
- No model state changed via COM. Only one filesystem rename (the hook), reversible.
- Project state is clean: cycle 006 needs only an X01 patch.

### Cycle 005 artifacts committed

- `scripts/03_fix/F01_neuter_before_calc_hook.ps1`
- `scripts/04_run/X01_calculate_test.ps1`
- `docs/known_issues.md` (cycle 005 version with ISSUE-001a/b split + ISSUE-004 boolean rule)
- `data/audit_reports/calctest_idx6_20260525_145302.md` + `.data.txt`
- `logs/F01_cycle005_20260525_145227.log`

## Cycle 005 plan -- "do one thing, measure"

Two scripts, run in order. Reversible at every step.

**Step A: Snapshot the current state**
```powershell
Copy-Item data\snapshots\cycle_000_colleague_baseline.leap data\snapshots\cycle_005_before_F01.leap.bak
# (not strictly needed since F01 only touches one file, but cheap insurance)
```

**Step B: Disable the hook**
```powershell
.\scripts\03_fix\F01_neuter_before_calc_hook.ps1 -AreaIndex 6
```

Renames `…\KAZ_2024\beforeCalculation.vbs` to `…\KAZ_2024\beforeCalculation.vbs.disabled_cycle005`. Reversible:
```powershell
.\scripts\03_fix\F01_neuter_before_calc_hook.ps1 -AreaIndex 6 -Restore
```

**Step C: Test Calculate**
```powershell
.\scripts\04_run\X01_calculate_test.ps1 -AreaIndex 6
```

Runs `oLEAP.Calculate`, captures result + diagnostics. Timeout 30 min.

**Three possible outcomes:**

| Calc result | Next step |
|---|---|
| OK | Open LEAP UI, view Results, screenshot a chart. We've passed Day 3 and Day 4 in one cycle. Skip to Day 5 scope reduction. |
| Fails on nodal/distribution | F01 didn't work as expected. Re-read F01 logic, possibly LEAP cached the script. Manual UI verification needed. |
| Fails on different error | Capture the error verbatim. Sequence: investigate via cycle 006. ISSUE-001b (SimType) and/or ISSUE-002 (broken units) likely in scope. |

## Project rules (cumulative)

1. All `.ps1` and VBS source: pure ASCII, verified at write time
2. VBS writes UTF-16 LE data files; PowerShell reads with -Encoding Unicode
3. Address LEAP areas by INDEX (avoid name-collision ambiguity)
4. Audit scripts READ-ONLY; if save popup appears, click No
5. Use canonical `BranchVariable("path:var")` accessor
6. **NEW:** VBS never writes `CStr(boolean)`; always `If x Then "1" Else "0"`. PowerShell parses 1/0.
7. **NEW:** When making model changes, fix one thing at a time. Test between each. Maintain restore path.

## Cycle log

- 000: scaffolding
- 001-002: name collision exposed, two areas with same key
- 003: index addressing solved that
- 004: canonical accessor fix, hook file confirmed live, ISSUE-001 split
- 005: F01 disabled hook SUCCESS; X01 hit VBS Err 450 on `oLEAP.Calculate` -- script harness bug, not a model error. Hook still disabled at cycle 005 close, ready for cycle 006 X01 retry.

## Open questions

- [ ] Does disabling the hook unblock calc? -- **STILL UNKNOWN** (X01 hit script bug before Calculate could run; F01 unverified end-to-end). Reanswer in cycle 006.
- [ ] What is the correct `oLEAP.Calculate` invocation signature on this install? -- **NEW**, needs cycle 006 X01 v2 patch. Try `Calculate True/False`, `Calculate(scenarios)`, `ActiveArea.Calculate`. Reference: check if the colleague's pre-migration scripts in BTR_LEAP_01 have a Calculate call.
- [ ] If hook-only fix works, what does the next failure look like? (answered when X01 v2 actually runs)
- [ ] Is NetworkSimulation(Pipeline) on transformation actually NEMO-required, or benign? (still open, may not matter if 001a fix is enough)
- [ ] How many branches under Demand actually have Avg Environmental Loading + bad units? (v5 audit, not in cycle 005)
