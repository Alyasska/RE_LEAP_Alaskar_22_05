# known_issues.md

> A permanent catalog of LEAP-specific issues we've encountered, root-caused, and fixed.
> Append-only. Don't delete entries even after fixes — future-you will hit them again.

---

## ISSUE-001: "Nodal distribution shares sum to 0%. They must sum to 100% in each region."

**First observed:** Colleague's attempt, cycle 0 (2026-05-21)
**Severity:** 🔴 Blocking — Calculate aborts
**Affected scenarios:** Any scenario where NEMO optimization runs with transmission modeling enabled (most of the 64)

### Symptoms
- LEAP shows error popup on Calculate: *"nodal distribution shares sum to 0%. they must sum to 100% in each region"*
- Calculation aborts. No partial results.
- Error appears regardless of whether you've touched the transmission section.

### Root cause
The model contains a VBScript hook in `beforeCalculation.vbs_Safe` (inside the encrypted `.leap` zip). The script's own first comment block reads:

```vbscript
' Temporary before calculation script designed to work around some LEAP bugs, including:
'    - Failure to execute before transformation scripts
'    - Improper populating of nodal distribution variables in NEMO
```

It tries to dynamically populate the `Nodal Distribution` variable for three branches:
- `Transformation\Electricity Production\Transmission Nodes\KAZ_North`
- `Transformation\Electricity Production\Transmission Nodes\KAZ_West`
- `Transformation\Electricity Production\Transmission Nodes\KAZ_South`

The calculation uses Russia electricity export demand. The formula (simplified):

```
KAZ_North_share = base_north + (100 - base_north) * russia_exports / total_electricity_demand
KAZ_West_share  = base_west  -        base_west  * russia_exports / total_electricity_demand
KAZ_South_share = Remainder(100)
```

It depends on **ALL** of these conditions:
1. Branch `Demand\Electricity_Exports\Russia` exists and is populated
2. Scenario "S1" (note: NOT S0 baseline) has valid `Base Distribution` values for those three nodes
3. All three transmission node branches are `Visible` for the Kazakhstan region
4. `total_electricity_demand > 0` (else divide-by-zero or NaN)

If any of those break, distributions become 0 → sum is 0% → NEMO rejects → error fires.

### Why the colleague's attempt triggered this
- He changed BaseYear → demand calculations may have produced 0 or NaN for some 2024+ values
- "S1" scenario's `Base Distribution` may not have been extended to 2024
- Possibly some demand branches were modified in ways that broke the totals

### Prototype fix (v0.1)
**Disable transmission optimization entirely.** Reversible, fast, gets us to a runnable model.

Two approaches, in order of preference:

#### Approach A: Disable Simulation Type on Electricity Production module
1. Open model in LEAP
2. Navigate to `Transformation\Electricity Production` (module-level, not a child branch)
3. In Analysis view, find variable `Simulation Type`
4. Current value is something like `NetworkSimulation(...)` or contains node references
5. Change to `Standard` for current scenario AND for `<no scenario>` (Current Accounts column)
6. Save

Scriptable via VBS:
```vbscript
oLEAP.BranchVariable("Transformation\Electricity Production:Simulation Type").ExpressionRS(r_id, scenario_id) = "Standard"
```

#### Approach B: Neuter the beforeCalculation hook for nodal distribution
If Approach A is not enough, also modify the `update_kaz_demand_distributions` sub in `beforeCalculation` to early-exit. Editable via LEAP's Script Editor inside the GUI (Advanced → Edit Scripts → beforeCalculation).

Replace the body of `Sub update_kaz_demand_distributions` with `Exit Sub` on its first line.

### Production fix (v0.2+ — deferred)
- Properly recompute `Base Distribution` for KAZ_North/West/South for 2024+ in scenario S1
- Or replace the dynamic VBScript logic with static `Nodal Distribution` expressions in each scenario
- Or split each transmission node into its own region (a major refactor)

### Files affected
- `beforeCalculation.vbs_Safe` (inside `.leap` zip, password "LEAP")
- `Transformation\Electricity Production:Simulation Type` variable

---

## ISSUE-002: "Invalid numerator unit for emissions factor" — 5 broken branches

**First observed:** Aliaskar's previous attempt, week of 2026-05-15
**Severity:** 🔴 Blocking — Calculate aborts when reaching emission step
**Affected branches:** 5 specific branches with `Avg Environmental Loading` variables that have empty or malformed unit metadata

### Symptoms
- LEAP error on Calculate referring to emission factor units
- Setting the expression to `0` or `"not used"` does NOT fix it — LEAP validates unit metadata before reading expression value

### Root cause
The variable `Avg Environmental Loading` on each of these branches has a unit field that is empty or malformed in LEAP's NexusDB. LEAP requires the numerator unit (e.g. `Kilogramme`) and denominator unit (e.g. `Terajoule`) both be valid to perform emissions calculations. Missing numerator → unresolvable unit chain → abort.

### Affected branches
| # | Branch path | Unit issue |
|---|---|---|
| 1 | `Demand\Agriculture\Syr Darya\Other\Lubricants\Methane` | Unit field empty |
| 2 | `Demand\Agriculture\Other\Lubricants\Methane` | Unit field empty |
| 3 | `Demand\Industry\Iron and Steel\Top down\LPG\Nitrous Oxide` | Unit field empty |
| 4 | `Demand\Industry\Other\Top Down\All Other\LPG\Nitrous Oxide` | Unit field empty |
| 5 | `Demand\Commercial\Lubricants\Methane` | Unit = `"/Terajoule"` (missing numerator `"Kilogramme"`) |

### Prototype fix (v0.1)
Delete the `Avg Environmental Loading` variable row entirely from each of the 5 branches.

#### Manual UI procedure (~30 seconds per branch)
1. Navigate to the branch in the tree
2. In Analysis view's data table, find the row labeled `Avg Environmental Loading`
3. Right-click the row header (the leftmost cell of that row)
4. Choose "Delete Variable" / "Delete Row" (exact label depends on context)
5. Confirm

#### Scripted attempt
LEAP's COM API may or may not support deleting variable rows cleanly. The fallback is the UI procedure above. Script attempt:
```vbscript
' May not work — needs testing
oLEAP.Branches("Demand\Commercial\Lubricants\Methane").Variables("Avg Environmental Loading").Delete
```

If scripted deletion fails, fallback to manual UI for all 5 branches (~3 minutes total).

### Production fix (v0.2+)
Properly define the unit metadata: numerator = `Kilogramme` for methane/N2O, denominator = `Terajoule`, with valid emission factor values from IPCC defaults or Kazakhstan-specific inventories.

---

## ISSUE-003: VBScript file encoding causes silent script failures

**First observed:** Aliaskar's previous attempt
**Severity:** 🟡 Annoying — script appears to run, does nothing or errors cryptically
**Affected:** Any VBScript file written by PowerShell or VS Code with UTF-8 BOM or wrong encoding

### Symptoms
- `cscript` runs the file without error
- But: LEAP COM operations have no effect, or you get "Microsoft VBScript compilation error" on weird lines
- Or: cyrillic comments / strings appear as garbage

### Root cause
The `cscript` engine expects **ASCII** or **UTF-16 LE with BOM** for VBScript files. UTF-8 (with or without BOM) is not reliably supported. PowerShell's default `Set-Content` writes UTF-8 with BOM on older versions, or UTF-8 no-BOM on PS Core.

### Fix
Always write VBS files via:
```powershell
[System.IO.File]::WriteAllText($vbsPath, $content, [System.Text.Encoding]::ASCII)
```

In VS Code:
- Click the encoding indicator in the bottom-right status bar (shows "UTF-8")
- Choose "Save with Encoding" → "Windows 1252" or "ASCII"
- For files with Cyrillic content, use "Windows 1251" instead

### Production fix
Same as prototype fix. This is a permanent rule. Add to project style guide.

---

## Template for new issues

```markdown
## ISSUE-NNN: <short title>

**First observed:** <cycle / date>
**Severity:** 🔴 / 🟡 / 🟢
**Affected:** <branches / scenarios / area>

### Symptoms
### Root cause
### Prototype fix (v0.1)
### Production fix (v0.2+ — deferred)
### Files affected
```
