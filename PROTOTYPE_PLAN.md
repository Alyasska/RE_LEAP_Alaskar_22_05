# PROTOTYPE_PLAN.md

**Target:** A `.leap` file that opens, runs Calculate without error, and shows some Results chart.
**Deadline:** 7 days from 2026-05-22.
**Definition of "prototype done":** F9 → no error popup → click a Results chart → numbers display.

This is NOT the production model. It is a proof that our pipeline works end-to-end. Calibration to ground truth is **v0.2**, after this prototype.

---

## Day 1 — Foundation (today, 2026-05-22)

| Task | Type | Output |
|---|---|---|
| Create new project folder `RE_LEAP_Alaskar_22_05` | Manual | Folder exists |
| Drop in starter pack from Claude | Manual | README, PLAN, HANDOFF, known_issues.md |
| `git init` + first commit | Manual | Repo initialized |
| Copy colleague's `.leap` → `data/snapshots/cycle_000_colleague_baseline.leap` | Manual | Baseline snapshot saved |
| Push to GitHub | Manual | Remote up to date |

**End-of-day check:** Repo is on GitHub, HANDOFF.md says "Day 1 complete".

---

## Day 2 — Scout the Colleague's Model

Goal: know exactly what's inside the file before we touch anything.

| Task | Type | Script |
|---|---|---|
| Extract `.leap` (decrypt, dump contents) | Python | `01_scout/S01_extract_leap.py` |
| Audit via LEAP COM (count branches, scenarios, find broken units) | PowerShell + VBS | `01_scout/S02_audit_model.ps1` |
| Run beforeCalculation script outside calc, see what it actually outputs | VBS | `01_scout/S03_dry_run_hooks.vbs` |

**End-of-day check:** `data/audit_reports/audit_cycle_001.md` exists, lists every blocker by name.

---

## Day 3 — Kill the Transmission Blocker (DANGER ZONE 1)

The colleague's "nodal distribution sums to 0%" error comes from NEMO transmission optimization. We disable it entirely for the prototype.

| Task | Type | Script |
|---|---|---|
| Disable `Simulation Type` for Electricity Production module → `Standard` | VBS | `02_reduce/R03_disable_transmission.vbs` |
| Comment out or no-op the `beforeCalculation` hook for nodal distribution | Manual + VBS | UI procedure UP02 |
| Verify Calculate progresses further than before | LEAP UI | Screenshot |

**End-of-day check:** Calculate runs past where it used to fail. Maybe new error, that's OK — that means progress.

---

## Day 4 — Surgical Fixes (DANGER ZONE 2)

Address the 5 broken emission factor unit definitions (from `Status_and_Blockers.md`).

| Task | Type | Script |
|---|---|---|
| Verify 5 broken branches are still broken in current file | VBS | `01_scout/S04_verify_known_blockers.vbs` |
| Delete `Avg Environmental Loading` row from those 5 branches | VBS + UI fallback | `03_fix/F01_remove_broken_loadings.vbs` |
| Any new blockers found day 3 — fix them | Ad-hoc | (depends on what we find) |

**End-of-day check:** Calculate progresses past emission factor step.

---

## Day 5 — Reduce Scope (DANGER ZONE 3)

64 scenarios, 6 regions. For the prototype we want **1 scenario × 1 region**.

| Task | Type | Script |
|---|---|---|
| Hide / disable all scenarios except S0 Baseline Historical | VBS | `02_reduce/R01_isolate_baseline.vbs` |
| Hide / disable all regions except Kazakhstan | VBS | `02_reduce/R02_isolate_kazakhstan.vbs` |
| Set Kazakhstan as default region for results | UI | UI procedure UP03 |

**Note:** "Hide" is safer than "delete" — reversible, doesn't cascade-break references. Use the `Visible` API.

**End-of-day check:** Only S0 and Kazakhstan show in views.

---

## Day 6 — First Successful Calculation

| Task | Type | Script |
|---|---|---|
| Run Calculate, capture log | VBS | `04_run/X01_calculate.vbs` |
| If errors → diagnose, fix, repeat | Ad-hoc | (cycle) |
| When Calculate completes → take screenshots of Results | LEAP UI | Saved to `docs/screenshots/` |

**End-of-day check:** At least one Results chart renders numbers, even if those numbers are nonsense.

---

## Day 7 — Tag, Document, Plan v0.2

| Task | Type | Output |
|---|---|---|
| Commit final `.leap` file as `data/snapshots/v0.1_prototype.leap` | Manual | |
| `git tag v0.1-prototype` + push | Manual | GitHub release |
| Write `docs/retrospective_v0.1.md` — what worked, what didn't, what's next | Manual | |
| Draft `PROTOTYPE_PLAN_v0.2.md` — calibration & forecasts roadmap | Manual | |

**End-of-prototype check:** Prototype delivered. Decision: continue toward v0.2 or refactor.

---

## Risk register

| Risk | Mitigation |
|---|---|
| Day 3 disable-transmission breaks more than it fixes | Keep `cycle_000` snapshot, rollback |
| 5 broken-unit branches turn out to be 50 broken branches | Day 4 expands into Days 4-5, push reduce to Day 6 |
| LEAP COM crashes mid-script | Use `On Error Resume Next` + verbose logging in every script |
| Old hardware (i5-3470, 8GB) can't handle 64 scenarios calc | Day 5 reduce comes BEFORE Day 6 calc — should be fine |
| You get pulled into other work | Days 6-7 are buffer. Plan absorbs 1-2 days of slippage. |

---

## What we are explicitly NOT doing in v0.1

These are deferred to v0.2 (post-prototype):

- ❌ Calibrating to 2024 energy balance (just need *some* numbers)
- ❌ Forecasting logic for 2025-2045 (whatever LEAP defaults to is fine)
- ❌ Removing the 4 non-KZ regions permanently (just hide them)
- ❌ Removing the 63 non-baseline scenarios permanently (just hide them)
- ❌ Climate change scenarios, mitigation scenarios
- ❌ Pretty charts, custom indicators, BTR-format export
- ❌ Solving the 5 broken-unit branches "properly" — we just delete the rows
