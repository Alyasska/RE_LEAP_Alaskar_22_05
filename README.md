# RE_LEAP_Alaskar_22_05

Reverse-engineering approach to migrating the Kazakhstan LEAP model from BaseYear 2010 → 2024.

**Owner:** Aliaskar Bekishev — Climate Change Coordination Centre, Kazakhstan
**Started:** 2026-05-22
**LEAP version:** 2024.4.0.8 (Windows)

## Goal (1-week prototype)

Produce a LEAP `.leap` file where:
1. BaseYear = 2024, FirstScenarioYear = 2025, EndYear = 2045
2. Pressing F9 (Calculate) runs to completion **without error**
3. Clicking any Results chart displays numbers (any numbers — calibration comes later)

That's it. Reasonable forecasts and ground-truth calibration are for **v0.2**, not v0.1.

## Working principles

1. **Keep what works** — the colleague's `.leap` file with BaseYear=2024 already done is our starting point. Don't redo what's already done.
2. **Read-only first** — every cycle starts with a scout script (`01_scout/`) before any writes.
3. **One change per commit** — easy rollback, clear history.
4. **Manual UI is OK** — if a fix takes 30 seconds in the UI vs 2 hours of scripting, do it in the UI and document the procedure in `docs/ui_procedures/`.
5. **Prototype ≠ production** — kill features for the prototype that we'll restore in v0.2 (e.g. transmission optimization, climate change scenarios).

## How to start a session

```powershell
cd C:\path\to\RE_LEAP_Alaskar_22_05
git pull
cat HANDOFF.md          # what state are we in?
cat PROTOTYPE_PLAN.md   # which day are we on?
```

Then run the script for today's cycle.

## Directory map

```
docs/             ← all reasoning, diagnoses, manual procedures
scripts/
  0_setup/        ← one-time environment setup
  01_scout/       ← read-only audit & inspection (safe anytime)
  02_reduce/      ← strip scenarios/regions/modules we don't need for prototype
  03_fix/         ← surgical fixes for specific blocking errors
  04_run/         ← Calculate, export results
  core/           ← shared helpers (LEAP_Connection_Test, etc.)
data/
  snapshots/      ← .leap files at each cycle (gitignored — too big)
  audit_reports/  ← scout output, one file per audit
  extracted/      ← decrypted .leap contents (gitignored, regenerable)
  ground_truth/   ← KZ energy balance, GDP, etc.
logs/             ← per-cycle logs (cycle_001.md, cycle_002.md, …)
HANDOFF.md        ← living state document — read this FIRST every session
PROTOTYPE_PLAN.md ← 7-day plan
```

## Companion repos / files

- Previous attempt repo (lessons learned): https://github.com/Alyasska/BTR_LEAP_01
- LEAP source areas folder: `C:\Users\User\Documents\LEAP_16_04\LEAP Areas\`
