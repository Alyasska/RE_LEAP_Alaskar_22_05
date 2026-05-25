# S02_audit_model.ps1
# Day 2 audit: open the colleague's area in LEAP via COM, read its full state,
# verify known blockers (ISSUE-001, ISSUE-002), report findings.
#
# READ-ONLY. Never calls .Save. Closes without saving.
#
# Usage:
#   .\scripts\01_scout\S02_audit_model.ps1
#   .\scripts\01_scout\S02_audit_model.ps1 -AreaName "kaz_workshop_exercise_for2024"
#
# Output:
#   - data\audit_reports\audit_cycle_001_<timestamp>.md  (human-readable)
#   - data\audit_reports\audit_cycle_001_<timestamp>.raw.log (raw cscript output)
#
# Prerequisites:
#   - The area must be INSTALLED in LEAP (not just sitting as a .leap file).
#     If you only have the .leap file, open LEAP UI once, Area > Install from File,
#     pick data\snapshots\cycle_000_colleague_baseline.leap, accept default name.
#
# Expected runtime: 3 to 7 minutes (LEAP COM startup + scenario iteration).

param(
    [string]$AreaName = "kaz_workshop_exercise_for2024",
    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"

# Resolve project root (two levels up from this script)
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$rawLogPath    = Join-Path $projectRoot "data\audit_reports\audit_cycle_001_${timestamp}.raw.log"
$reportPath    = Join-Path $projectRoot "data\audit_reports\audit_cycle_001_${timestamp}.md"
$runLogPath    = Join-Path $projectRoot "logs\S02_audit_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $rawLogPath -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== S02: Audit model (READ-ONLY) ==="
Log "Target area: $AreaName"
Log "Report will land at: $reportPath"
Log ""

# ----------------------------------------------------------------------------
# Build the VBScript. ASCII-only. No unicode punctuation. No em-dashes.
# Embedded $AreaName via placeholder.
# ----------------------------------------------------------------------------
$vbsTemplate = @'
Option Explicit

' Audit script. READ-ONLY. Closes LEAP without saving.

Dim oLEAP, r, s, p, var
Dim found_any_active
Dim broken_paths(4)
Dim node_paths(2)
Dim demand_paths(0)
Dim i

On Error Resume Next

WScript.Echo "PHASE:starting"

Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    WScript.Echo "ERROR:COM_CREATE:" & Err.Number & ":" & Err.Description
    WScript.Quit 1
End If
Err.Clear

WScript.Echo "PHASE:com_created"
WScript.Sleep 4000

oLEAP.Areas("{{AREA_NAME}}").Open
If Err.Number <> 0 Then
    WScript.Echo "ERROR:AREA_OPEN:" & Err.Number & ":" & Err.Description
    ' Try to list available areas to help diagnose
    WScript.Echo "AVAILABLE_AREAS_FOLLOW"
    Dim a
    For Each a In oLEAP.Areas
        WScript.Echo "AREA_AVAILABLE:" & a.Name
    Next
    WScript.Echo "AVAILABLE_AREAS_END"
    oLEAP.Quit
    WScript.Quit 2
End If
Err.Clear

WScript.Sleep 3000
WScript.Echo "PHASE:area_opened:" & oLEAP.ActiveArea.Name

' Set verbose=0 to skip UI updates during read loop
oLEAP.Verbose = 0

' --- Metadata ---
WScript.Echo "META:ActiveArea=" & oLEAP.ActiveArea.Name
WScript.Echo "META:BaseYear=" & oLEAP.BaseYear
WScript.Echo "META:FirstScenarioYear=" & oLEAP.FirstScenarioYear
WScript.Echo "META:EndYear=" & oLEAP.EndYear
WScript.Echo "META:WorkingDirectory=" & oLEAP.WorkingDirectory

' --- Regions ---
WScript.Echo "PHASE:regions"
For Each r In oLEAP.Regions
    WScript.Echo "REGION:" & r.Id & "|" & r.Name & "|ResultsShown=" & r.ResultsShown
Next

' --- Scenarios ---
WScript.Echo "PHASE:scenarios"
For Each s In oLEAP.Scenarios
    WScript.Echo "SCENARIO:" & s.Id & "|" & s.Abbreviation & "|" & s.Name
Next

' --- Branch existence checks ---
WScript.Echo "PHASE:branch_checks"

Dim key_branches(5)
key_branches(0) = "Transformation\Electricity Production"
key_branches(1) = "Transformation\Electricity Production\Transmission Nodes\KAZ_North"
key_branches(2) = "Transformation\Electricity Production\Transmission Nodes\KAZ_West"
key_branches(3) = "Transformation\Electricity Production\Transmission Nodes\KAZ_South"
key_branches(4) = "Demand\Electricity_Exports\Russia"
key_branches(5) = "Demand"

For i = 0 To 5
    If oLEAP.Branches.Exists(key_branches(i)) Then
        WScript.Echo "BRANCH_EXISTS:" & key_branches(i) & "|YES"
    Else
        WScript.Echo "BRANCH_EXISTS:" & key_branches(i) & "|NO"
    End If
Next

' --- ISSUE-001 verification: Electricity Production Simulation Type ---
WScript.Echo "PHASE:issue_001_simtype"

If oLEAP.Branches.Exists("Transformation\Electricity Production") Then
    If oLEAP.Branches("Transformation\Electricity Production").VariableExists("Simulation Type") Then
        Set var = oLEAP.Branches("Transformation\Electricity Production").Variable("Simulation Type")
        For Each r In oLEAP.Regions
            ' Current Accounts (scenario id = 0)
            Dim ca_val
            ca_val = var.ExpressionRS(r.Id, 0)
            If Err.Number <> 0 Then
                ca_val = "ERROR:" & Err.Description
                Err.Clear
            End If
            WScript.Echo "SIMTYPE:" & r.Id & "|" & r.Name & "|CA|" & ca_val

            ' Per scenario - only output scenarios with Network in expression
            For Each s In oLEAP.Scenarios
                Dim sval
                sval = var.ExpressionRS(r.Id, s.Id)
                If Err.Number <> 0 Then
                    Err.Clear
                ElseIf InStr(sval, "Network") > 0 Then
                    WScript.Echo "SIMTYPE_NET:" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & sval
                End If
            Next
        Next
    Else
        WScript.Echo "SIMTYPE:VAR_MISSING"
    End If
Else
    WScript.Echo "SIMTYPE:BRANCH_MISSING"
End If

' --- ISSUE-002 verification: 5 broken-unit branches ---
WScript.Echo "PHASE:issue_002_broken_units"

broken_paths(0) = "Demand\Agriculture\Syr Darya\Other\Lubricants\Methane"
broken_paths(1) = "Demand\Agriculture\Other\Lubricants\Methane"
broken_paths(2) = "Demand\Industry\Iron and Steel\Top down\LPG\Nitrous Oxide"
broken_paths(3) = "Demand\Industry\Other\Top Down\All Other\LPG\Nitrous Oxide"
broken_paths(4) = "Demand\Commercial\Lubricants\Methane"

For i = 0 To 4
    If oLEAP.Branches.Exists(broken_paths(i)) Then
        If oLEAP.Branches(broken_paths(i)).VariableExists("Avg Environmental Loading") Then
            WScript.Echo "BROKEN_UNIT:" & broken_paths(i) & "|BRANCH_EXISTS|HAS_LOADING_VAR"
        Else
            WScript.Echo "BROKEN_UNIT:" & broken_paths(i) & "|BRANCH_EXISTS|NO_LOADING_VAR"
        End If
    Else
        WScript.Echo "BROKEN_UNIT:" & broken_paths(i) & "|BRANCH_MISSING|N/A"
    End If
Next

' --- Counts ---
WScript.Echo "PHASE:counts"
WScript.Echo "COUNT:Regions=" & oLEAP.Regions.Count
WScript.Echo "COUNT:Scenarios=" & oLEAP.Scenarios.Count

' --- Close without saving ---
WScript.Echo "PHASE:closing"
oLEAP.Verbose = 4
oLEAP.Quit
WScript.Echo "DONE"
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_NAME}}', $AreaName)

# Write as ASCII no BOM (ISSUE-003 rule)
$vbsPath = Join-Path $env:TEMP "S02_audit_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath ($((Get-Item $vbsPath).Length) bytes)"

# ----------------------------------------------------------------------------
# Execute audit
# ----------------------------------------------------------------------------
Log ""
Log "Running cscript audit. LEAP cold start can take 3 to 5 minutes."
Log "Iterating 64 scenarios per region also takes time. Be patient."
Log ""

$startTime = Get-Date
$out = & cscript //NoLogo $vbsPath 2>&1
$elapsed = (Get-Date) - $startTime
Log "Audit completed in $([math]::Round($elapsed.TotalSeconds, 1)) seconds"

Remove-Item $vbsPath -ErrorAction SilentlyContinue

# Save raw output
$out | Out-File -FilePath $rawLogPath -Encoding UTF8
Log "Raw output saved: $rawLogPath"

# ----------------------------------------------------------------------------
# Parse output and build markdown report
# ----------------------------------------------------------------------------
Log ""
Log "Parsing audit output..."

$meta = @{}
$regions = @()
$scenarios = @()
$branches = @{}
$simtypeCA = @()
$simtypeNet = @()
$brokenUnits = @()
$counts = @{}
$availableAreas = @()
$phase = "init"
$fatalError = $null

foreach ($line in $out) {
    $s = "$line".Trim()
    if ($s -match "^PHASE:(.+)$") {
        $phase = $matches[1]
    }
    elseif ($s -match "^ERROR:(.+)$") {
        $fatalError = $matches[1]
    }
    elseif ($s -match "^META:([^=]+)=(.*)$") {
        $meta[$matches[1]] = $matches[2]
    }
    elseif ($s -match "^REGION:([^|]+)\|([^|]+)\|ResultsShown=(.+)$") {
        $regions += [pscustomobject]@{ Id = $matches[1]; Name = $matches[2]; ResultsShown = $matches[3] }
    }
    elseif ($s -match "^SCENARIO:([^|]+)\|([^|]+)\|(.+)$") {
        $scenarios += [pscustomobject]@{ Id = $matches[1]; Abbreviation = $matches[2]; Name = $matches[3] }
    }
    elseif ($s -match "^BRANCH_EXISTS:([^|]+)\|(.+)$") {
        $branches[$matches[1]] = $matches[2]
    }
    elseif ($s -match "^SIMTYPE:([^|]+)\|([^|]+)\|CA\|(.+)$") {
        $simtypeCA += [pscustomobject]@{ RegionId = $matches[1]; RegionName = $matches[2]; Value = $matches[3] }
    }
    elseif ($s -match "^SIMTYPE_NET:([^|]+)\|([^|]+)\|([^|]+)\|(.+)$") {
        $simtypeNet += [pscustomobject]@{ RegionId = $matches[1]; RegionName = $matches[2]; Scenario = $matches[3]; Value = $matches[4] }
    }
    elseif ($s -match "^BROKEN_UNIT:([^|]+)\|([^|]+)\|(.+)$") {
        $brokenUnits += [pscustomobject]@{ Branch = $matches[1]; Exists = $matches[2]; LoadingVar = $matches[3] }
    }
    elseif ($s -match "^COUNT:([^=]+)=(.+)$") {
        $counts[$matches[1]] = $matches[2]
    }
    elseif ($s -match "^AREA_AVAILABLE:(.+)$") {
        $availableAreas += $matches[1]
    }
}

# ----------------------------------------------------------------------------
# Generate markdown report
# ----------------------------------------------------------------------------
$md = [System.Collections.Generic.List[string]]::new()

$md.Add("# Audit Report -- Cycle 001 -- $timestamp")
$md.Add("")
$md.Add("**Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$md.Add("**Script:** scripts/01_scout/S02_audit_model.ps1")
$md.Add("**Target area:** ``$AreaName``")
$md.Add("**Runtime:** $([math]::Round($elapsed.TotalSeconds, 1)) seconds")
$md.Add("**Last phase reached:** ``$phase``")
$md.Add("")

if ($fatalError) {
    $md.Add("## FATAL ERROR")
    $md.Add("")
    $md.Add("``````")
    $md.Add($fatalError)
    $md.Add("``````")
    $md.Add("")
    if ($availableAreas.Count -gt 0) {
        $md.Add("### Available areas on this LEAP install")
        $md.Add("")
        foreach ($a in $availableAreas) { $md.Add("- ``$a``") }
        $md.Add("")
        $md.Add("**Action:** Re-run with one of the area names above, e.g. ``.\S02_audit_model.ps1 -AreaName ""<exact name>""``")
    } else {
        $md.Add("**Action:** Open LEAP UI, install area from data\snapshots\cycle_000_colleague_baseline.leap, then re-run.")
    }
} else {
    # Metadata section
    $md.Add("## Model metadata")
    $md.Add("")
    $md.Add("| Property | Value |")
    $md.Add("|---|---|")
    foreach ($k in $meta.Keys | Sort-Object) {
        $md.Add("| $k | ``$($meta[$k])`` |")
    }
    $md.Add("")

    # Expected vs actual
    $md.Add("### Expected vs actual")
    $md.Add("")
    $md.Add("| Property | Expected | Actual | OK? |")
    $md.Add("|---|---|---|---|")
    $rows = @(
        @("BaseYear", "2024", $meta["BaseYear"]),
        @("FirstScenarioYear", "2025", $meta["FirstScenarioYear"]),
        @("EndYear", "2045", $meta["EndYear"])
    )
    foreach ($row in $rows) {
        $ok = if ("$($row[1])" -eq "$($row[2])") { "OK" } else { "MISMATCH" }
        $md.Add("| $($row[0]) | $($row[1]) | $($row[2]) | $ok |")
    }
    $md.Add("")

    # Counts
    $md.Add("## Counts")
    $md.Add("")
    $md.Add("| Object | Count |")
    $md.Add("|---|---|")
    foreach ($k in $counts.Keys | Sort-Object) {
        $md.Add("| $k | $($counts[$k]) |")
    }
    $md.Add("")

    # Regions
    $md.Add("## Regions ($($regions.Count))")
    $md.Add("")
    $md.Add("| Id | Name | ResultsShown |")
    $md.Add("|---|---|---|")
    foreach ($r in $regions) {
        $md.Add("| $($r.Id) | $($r.Name) | $($r.ResultsShown) |")
    }
    $md.Add("")

    # Scenarios
    $md.Add("## Scenarios ($($scenarios.Count))")
    $md.Add("")
    if ($scenarios.Count -le 10) {
        $md.Add("| Id | Abbr | Name |")
        $md.Add("|---|---|---|")
        foreach ($s in $scenarios) {
            $md.Add("| $($s.Id) | ``$($s.Abbreviation)`` | $($s.Name) |")
        }
    } else {
        $md.Add("_Showing all $($scenarios.Count) scenarios:_")
        $md.Add("")
        $md.Add("| Id | Abbr | Name |")
        $md.Add("|---|---|---|")
        foreach ($s in $scenarios) {
            $md.Add("| $($s.Id) | ``$($s.Abbreviation)`` | $($s.Name) |")
        }
    }
    $md.Add("")

    # Branch existence
    $md.Add("## Key branch existence checks")
    $md.Add("")
    $md.Add("| Branch | Exists? |")
    $md.Add("|---|---|")
    foreach ($b in $branches.Keys | Sort-Object) {
        $md.Add("| ``$b`` | $($branches[$b]) |")
    }
    $md.Add("")

    # ISSUE-001
    $md.Add("## ISSUE-001 verification: Transformation Simulation Type")
    $md.Add("")
    $md.Add("Looking for scenarios where ``Transformation\Electricity Production:Simulation Type`` contains ""Network"" (means nodal distribution must work).")
    $md.Add("")
    $md.Add("### Current Accounts (no scenario) values per region")
    $md.Add("")
    if ($simtypeCA.Count -gt 0) {
        $md.Add("| Region | CA value |")
        $md.Add("|---|---|")
        foreach ($x in $simtypeCA) {
            $md.Add("| $($x.RegionName) | ``$($x.Value)`` |")
        }
    } else {
        $md.Add("_No data captured._")
    }
    $md.Add("")
    $md.Add("### Scenarios with Network-mode simulation (need disabling for prototype)")
    $md.Add("")
    if ($simtypeNet.Count -gt 0) {
        $md.Add("**$($simtypeNet.Count) scenario-region pairs use NetworkSimulation.** All must be set to ``Standard`` in Day 3.")
        $md.Add("")
        $md.Add("| Region | Scenario | Value |")
        $md.Add("|---|---|---|")
        foreach ($x in $simtypeNet) {
            $md.Add("| $($x.RegionName) | $($x.Scenario) | ``$($x.Value)`` |")
        }
    } else {
        $md.Add("_No scenarios using NetworkSimulation found. ISSUE-001 may already be resolved or never triggered in this state._")
    }
    $md.Add("")

    # ISSUE-002
    $md.Add("## ISSUE-002 verification: 5 broken-unit branches")
    $md.Add("")
    $md.Add("| Branch | Branch Exists | Has Avg Env Loading? |")
    $md.Add("|---|---|---|")
    foreach ($x in $brokenUnits) {
        $md.Add("| ``$($x.Branch)`` | $($x.Exists) | $($x.LoadingVar) |")
    }
    $md.Add("")
    $brokenStillPresent = ($brokenUnits | Where-Object { $_.LoadingVar -eq "HAS_LOADING_VAR" }).Count
    if ($brokenStillPresent -eq 5) {
        $md.Add("**All 5 broken-unit branches still have ``Avg Environmental Loading``. Day 4 fix is required.**")
    } elseif ($brokenStillPresent -eq 0) {
        $md.Add("**No broken branches have the loading variable. Either already fixed or branches moved.**")
    } else {
        $md.Add("**$brokenStillPresent of 5 broken branches still have the loading variable. Partial state.**")
    }
    $md.Add("")

    # Summary verdict
    $md.Add("## Verdict and next steps")
    $md.Add("")
    $networkCount = $simtypeNet.Count
    $md.Add("- ISSUE-001 (transmission): **$networkCount scenario-region pairs** need fix in Day 3")
    $md.Add("- ISSUE-002 (broken units): **$brokenStillPresent of 5 branches** need fix in Day 4")
    $md.Add("- Total scenarios: **$($scenarios.Count)** (will hide all but S0 in Day 5)")
    $md.Add("- Total regions: **$($regions.Count)** (will hide all but Kazakhstan in Day 5)")
    $md.Add("")
    $md.Add("**Recommended next action:** Proceed to Day 3 (``R03_disable_transmission.ps1``).")
}

$md.Add("")
$md.Add("---")
$md.Add("")
$md.Add("_Raw cscript output: ``$($rawLogPath.Replace($projectRoot + '\', ''))``_")

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Markdown report: $reportPath"

# Save run log
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log ""
Log "=== S02 complete ==="
Log ""

if ($fatalError) {
    Log "FATAL ERROR encountered: $fatalError"
    Log "See report for available areas. Re-run with -AreaName <correct name>."
    exit 1
} else {
    Log "Report ready. Open: $reportPath"
    exit 0
}
