# Day 2 -- Scout the colleague's model

**Goal:** Know exactly what's inside the .leap file before we touch anything.
**Outputs:** Two audit reports in `data/audit_reports/`.
**Time:** ~15-20 minutes (most of it waiting for LEAP COM).

---

## Step 1 -- Extract the .leap archive (file-level inspection)

This decrypts the AES-encrypted ZIP, dumps contents to `data/extracted/`,
and writes a markdown summary.

### First, install pyzipper

```powershell
pip install pyzipper
```

### Then run the extractor

```powershell
python scripts\01_scout\S01_extract_leap.py data\snapshots\cycle_000_colleague_baseline.leap
```

### Expected output

- `data/extracted/cycle_000_colleague_baseline/` -- ~30-40 files (VBS hooks, NexusDB tables, INI files, Excel files)
- `data/audit_reports/extract_cycle_000_colleague_baseline_<timestamp>.md` -- summary

### What to check in the summary

- Confirm `beforeCalculation.vbs_Safe` is listed (this is where the nodal distribution bug lives)
- Confirm there are ~40-50 `.nx1` files (NexusDB tables -- the actual model data)
- Note the total uncompressed size

---

## Step 2 -- Install the area in LEAP (if not already installed)

The audit script needs to open the area via COM, which means LEAP needs to know about it.

**Skip this if the area is already in `oLEAP.Areas` collection.**

### Manual UI procedure

1. Open LEAP
2. Menu: `Area > Manage Areas` (or `Area > Install`)
3. Find `data\snapshots\cycle_000_colleague_baseline.leap`
4. Accept default install location (`C:\Users\User\Documents\LEAP Areas\`)
5. The area will appear with a name -- **note the exact name**, that's what Step 3 needs

Common installed names seen so far:
- `kaz_workshop_exercise_for2024` (colleague's)
- `KAZ_workshop exercise`
- `KAZ_2024`

---

## Step 3 -- Run the COM-based audit

This opens the area via COM and reads its full state. **READ-ONLY** -- never saves.

```powershell
.\scripts\01_scout\S02_audit_model.ps1 -AreaName "kaz_workshop_exercise_for2024"
```

If the area name is different, pass the correct one. Available areas are listed if the script fails to open.

### Expected output

- `data/audit_reports/audit_cycle_001_<timestamp>.md` -- the audit report (open this)
- `data/audit_reports/audit_cycle_001_<timestamp>.raw.log` -- raw cscript output
- `logs/S02_audit_<timestamp>.log` -- run log

### Expected runtime

3-7 minutes:
- 30s -- LEAP COM startup
- 1-2 min -- area open (large model)
- 2-4 min -- iterate 64 scenarios x 6 regions for SimType check
- 10s -- close

---

## End-of-Day 2 checklist

- [ ] `data/extracted/cycle_000_colleague_baseline/` exists with ~30+ files
- [ ] First audit report (file-level) exists in `data/audit_reports/`
- [ ] Second audit report (COM-level) exists in `data/audit_reports/`
- [ ] Open the COM audit report -- verify:
  - BaseYear = 2024 (or note actual value)
  - 5 broken-unit branches: check how many still have `HAS_LOADING_VAR`
  - Count of scenarios using NetworkSimulation (will be ISSUE-001 fix scope on Day 3)
- [ ] Update `HANDOFF.md` "Last cycle" section
- [ ] Commit: `git add . && git commit -m "cycle 001: Day 2 scout complete, audit reports generated"`
- [ ] Push to GitHub
- [ ] Report back to Aliaskar (and Claude): paste the **Verdict** section of the audit report

---

## Troubleshooting

### S01 (extract) -- "pyzipper not installed"
```powershell
pip install pyzipper
```
If pip itself missing, install Python from python.org first.

### S02 (audit) -- FATAL ERROR: AREA_OPEN
The script will list available areas. Pick the right one and re-run with `-AreaName "<exact name>"`.

### S02 (audit) -- LEAP hangs at "PHASE:area_opened"
The model is large. Wait up to 10 minutes. If still hung:
- Kill LEAP and cscript via Task Manager
- Open LEAP UI manually first -- accept any popups (license, splash, "do you want to upgrade")
- Close LEAP UI cleanly
- Re-run the script

### S02 -- LEAP shows a popup the agent can't see
LEAP sometimes shows dialogs (e.g. "Do you want to convert this area?") that block COM.
**Workaround:** open the area once in LEAP UI manually, dismiss all popups, save, close,
then run S02.

### S02 -- "Verbose" or other property errors
Some LEAP installs report different versions of the API. If a specific line fails,
remove that single property check and re-run. Note the failure in HANDOFF.md.
