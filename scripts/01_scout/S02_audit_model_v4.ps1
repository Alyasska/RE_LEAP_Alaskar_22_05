# S02_audit_model_v4.ps1
# Day 2 audit, REVISION 4.
# Single-purpose change from v3: fix SimType accessor.
#
# Why v4 (cycle 003 finding):
#   v3 used oLEAP.Branches(path).Variable(name) which returned Nothing silently.
#   Colleague's own VBS and docs/known_issues.md both use the canonical pattern:
#     oLEAP.BranchVariable("Path:VariableName").ExpressionRS(r.Id, s.Id)
#   v4 switches to that pattern.
#
# Also adds: read beforeCalculation.vbs_Safe content from the area's folder
# (we know the .leap layout: it's a zip; extracted on install, the file lives
# under the area's directory). This lets us see whether the nodal-distribution
# hook is actually present in KAZ_2024 independent of NetworkSimulation use.
#
# Index-based area addressing (v3 pattern, retained).
#
# Usage:
#   .\scripts\01_scout\S02_audit_model_v4.ps1 -AreaIndex 6

param(
    [Parameter(Mandatory=$true)]
    [int]$AreaIndex
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\audit_cycle_004_idx${AreaIndex}_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\audit_cycle_004_idx${AreaIndex}_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\S02v4_idx${AreaIndex}_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== S02 v4: Audit with BranchVariable accessor ==="
Log "Area index: $AreaIndex"

# ----------------------------------------------------------------------------
# VBS. Placeholders: {{AREA_INDEX}}, {{DATA_FILE}}
# ----------------------------------------------------------------------------
$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile
Dim r, s, b
Dim simVar, sval, err_num, err_desc
Dim i
Dim n_scenarios, n_regions, n_pairs_checked, n_network_found, n_nonempty
Dim target_area, current_idx, a
Dim path_arr, broken_path, area_dir

Sub W(key, val)
    outFile.WriteLine(key & "|" & val)
End Sub

On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.OpenTextFile("{{DATA_FILE}}", 2, True, -1)
If Err.Number <> 0 Then
    WScript.Echo "FATAL_FILE:" & Err.Number & ":" & Err.Description
    WScript.Quit 10
End If
Err.Clear

W "PHASE", "starting"

Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    outFile.WriteLine "FATAL|COM_CREATE|" & Err.Number & ":" & Err.Description
    outFile.Close
    WScript.Quit 1
End If
Err.Clear

W "PHASE", "com_created"
WScript.Sleep 4000

' --- Resolve area by index ---
W "PHASE", "resolving_area_by_index"
W "TARGET_INDEX", "{{AREA_INDEX}}"

current_idx = 0
Set target_area = Nothing
For Each a In oLEAP.Areas
    current_idx = current_idx + 1
    If current_idx = {{AREA_INDEX}} Then
        Set target_area = a
        Exit For
    End If
Next

If target_area Is Nothing Then
    outFile.WriteLine "FATAL|INDEX_OUT_OF_RANGE|requested={{AREA_INDEX}}|total=" & current_idx
    outFile.Close
    oLEAP.Quit
    WScript.Quit 2
End If

outFile.WriteLine "RESOLVED_AREA|{{AREA_INDEX}}|" & target_area.Name

' Capture area directory for filesystem reads BEFORE Open (Directory works on closed objects)
area_dir = target_area.Directory
W "AREA_DIRECTORY", area_dir

' --- Open ---
target_area.Open
If Err.Number <> 0 Then
    outFile.WriteLine "FATAL|AREA_OPEN|" & Err.Number & ":" & Err.Description
    outFile.Close
    oLEAP.Quit
    WScript.Quit 3
End If
Err.Clear

WScript.Sleep 3000
W "PHASE", "area_opened"
oLEAP.Verbose = 0

W "OPENED_AREA_NAME", oLEAP.ActiveArea.Name
W "OPENED_AREA_BASE_YEAR", oLEAP.BaseYear
W "OPENED_AREA_FIRST_SCENARIO_YEAR", oLEAP.FirstScenarioYear
W "OPENED_AREA_END_YEAR", oLEAP.EndYear

outFile.WriteLine "META|ActiveArea|" & oLEAP.ActiveArea.Name
outFile.WriteLine "META|BaseYear|" & oLEAP.BaseYear
outFile.WriteLine "META|FirstScenarioYear|" & oLEAP.FirstScenarioYear
outFile.WriteLine "META|EndYear|" & oLEAP.EndYear

' Regions
W "PHASE", "regions"
n_regions = 0
For Each r In oLEAP.Regions
    n_regions = n_regions + 1
    outFile.WriteLine "REGION|" & r.Id & "|" & r.Name & "|" & CStr(r.ResultsShown)
Next
outFile.WriteLine "COUNT|Regions|" & n_regions

' Scenarios
W "PHASE", "scenarios"
n_scenarios = 0
For Each s In oLEAP.Scenarios
    n_scenarios = n_scenarios + 1
    outFile.WriteLine "SCENARIO|" & s.Id & "|" & s.Abbreviation & "|" & s.Name & "|"
Next
outFile.WriteLine "COUNT|Scenarios|" & n_scenarios

' --- ISSUE-001 v4: SimType via canonical BranchVariable accessor ---
W "PHASE", "issue_001_simtype_v4"

Err.Clear
Set simVar = oLEAP.BranchVariable("Transformation\Electricity Production:Simulation Type")
err_num = Err.Number
err_desc = Err.Description
Err.Clear

If err_num <> 0 Then
    outFile.WriteLine "SIMTYPE_STATUS|ACCESSOR_ERROR|" & err_num & ":" & err_desc
ElseIf simVar Is Nothing Then
    W "SIMTYPE_STATUS", "VAR_IS_NOTHING_via_BranchVariable"
Else
    W "SIMTYPE_STATUS", "VAR_OK"

    n_pairs_checked = 0
    n_network_found = 0
    n_nonempty = 0

    For Each r In oLEAP.Regions
        For Each s In oLEAP.Scenarios
            Err.Clear
            sval = ""
            sval = simVar.ExpressionRS(r.Id, s.Id)
            err_num = Err.Number
            err_desc = Err.Description
            Err.Clear
            n_pairs_checked = n_pairs_checked + 1
            If err_num <> 0 Then
                outFile.WriteLine "SIMTYPE_ERR|" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & err_num & "|" & err_desc
            ElseIf Len(sval) > 0 Then
                n_nonempty = n_nonempty + 1
                If InStr(sval, "Network") > 0 Then
                    outFile.WriteLine "SIMTYPE_NET|" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & sval
                    n_network_found = n_network_found + 1
                Else
                    outFile.WriteLine "SIMTYPE_VAL|" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & sval
                End If
            End If
        Next
    Next

    outFile.WriteLine "SIMTYPE_PAIRS_CHECKED|" & n_pairs_checked
    outFile.WriteLine "SIMTYPE_NONEMPTY|" & n_nonempty
    outFile.WriteLine "SIMTYPE_NETWORK_COUNT|" & n_network_found
End If

' --- ISSUE-001 supplement: probe beforeCalculation script file in area folder ---
W "PHASE", "probe_before_calc_hook"

Dim hookCandidates
Dim hookPath, hookFound
hookFound = False
hookCandidates = Array( _
    area_dir & "\beforeCalculation.vbs", _
    area_dir & "\beforeCalculation.vbs_Safe", _
    area_dir & "\beforeCalculation.txt", _
    area_dir & "\Scripts\beforeCalculation.vbs", _
    area_dir & "\Scripts\beforeCalculation.vbs_Safe" _
)

For i = 0 To UBound(hookCandidates)
    hookPath = hookCandidates(i)
    If fso.FileExists(hookPath) Then
        Dim hookFile, hookSize, hookContains
        Set hookFile = fso.GetFile(hookPath)
        hookSize = hookFile.Size
        outFile.WriteLine "HOOK_FOUND|" & hookPath & "|" & hookSize
        ' Read first 200 lines and check for nodal distribution refs
        Dim ts, ln, line_count, has_nodal, has_kaz_north
        Set ts = fso.OpenTextFile(hookPath, 1, False, 0)
        line_count = 0
        has_nodal = False
        has_kaz_north = False
        Do While Not ts.AtEndOfStream And line_count < 500
            ln = ts.ReadLine
            line_count = line_count + 1
            If InStr(LCase(ln), "nodal") > 0 Then has_nodal = True
            If InStr(ln, "KAZ_North") > 0 Then has_kaz_north = True
        Loop
        ts.Close
        outFile.WriteLine "HOOK_LINES|" & hookPath & "|" & line_count
        outFile.WriteLine "HOOK_NODAL_REF|" & hookPath & "|" & has_nodal
        outFile.WriteLine "HOOK_KAZ_NORTH_REF|" & hookPath & "|" & has_kaz_north
        hookFound = True
    End If
Next

If Not hookFound Then
    W "HOOK_STATUS", "NO_HOOK_FILE_FOUND"
    outFile.WriteLine "HOOK_SEARCHED_PATHS|" & Join(hookCandidates, ";")
End If
Err.Clear

' --- ISSUE-002 exact ---
W "PHASE", "issue_002_exact"
Dim paths_str
paths_str = "Demand\Agriculture\Syr Darya\Other\Lubricants\Methane|" & _
            "Demand\Agriculture\Other\Lubricants\Methane|" & _
            "Demand\Industry\Iron and Steel\Top down\LPG\Nitrous Oxide|" & _
            "Demand\Industry\Other\Top Down\All Other\LPG\Nitrous Oxide|" & _
            "Demand\Commercial\Lubricants\Methane"
path_arr = Split(paths_str, "|")
For i = 0 To UBound(path_arr)
    broken_path = path_arr(i)
    If oLEAP.Branches.Exists(broken_path) Then
        If oLEAP.Branches(broken_path).VariableExists("Avg Environmental Loading") Then
            outFile.WriteLine "BROKEN_EXACT|" & broken_path & "|EXISTS|HAS_LOADING"
        Else
            outFile.WriteLine "BROKEN_EXACT|" & broken_path & "|EXISTS|NO_LOADING"
        End If
    Else
        outFile.WriteLine "BROKEN_EXACT|" & broken_path & "|MISSING|N/A"
    End If
Next

W "PHASE", "closing"
oLEAP.Verbose = 4
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_INDEX}}', "$AreaIndex").Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))

$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: VBS source contains non-ASCII"; exit 99
}

$vbsPath = Join-Path $env:TEMP "S02v4_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath"
Log ""
Log "Running cscript. 90-120 seconds expected (v3 took 93.5s)."

$startTime = Get-Date
$cscriptOut = & cscript //NoLogo $vbsPath 2>&1
$elapsed = (Get-Date) - $startTime
Log "cscript finished in $([math]::Round($elapsed.TotalSeconds, 1))s"
Remove-Item $vbsPath -ErrorAction SilentlyContinue
foreach ($l in $cscriptOut) { if ($l -is [string] -and $l.Trim().Length -gt 0) { Log "cscript: $l" } }

if (-not (Test-Path $dataFile)) {
    Log "FATAL: data file missing"
    exit 11
}

$dataLines = Get-Content $dataFile -Encoding Unicode
Log "Read $($dataLines.Count) data lines"

# Parse
$meta = @{}
$counts = @{}
$regions = @()
$scenarios = @()
$simtypeStatus = $null
$simtypeNet = @()
$simtypeVal = @()
$simtypeErr = @()
$simtypePairs = 0
$simtypeNonempty = 0
$simtypeNetwork = 0
$brokenExact = @()
$hookFound = @()
$hookSearched = $null
$fatal = $null
$lastPhase = $null
$resolvedAreaName = $null
$openedAreaName = $null
$openedBaseYear = $null
$areaDir = $null

foreach ($raw in $dataLines) {
    $line = "$raw".Trim()
    if ($line.Length -eq 0) { continue }
    $parts = $line -split '\|'
    switch ($parts[0]) {
        "PHASE"               { $lastPhase = $parts[1] }
        "FATAL"               { $fatal = "$($parts[1]):$($parts[2..($parts.Count-1)] -join ':')" }
        "RESOLVED_AREA"       { $resolvedAreaName = $parts[2] }
        "AREA_DIRECTORY"      { $areaDir = $parts[1] }
        "OPENED_AREA_NAME"    { $openedAreaName = $parts[1] }
        "OPENED_AREA_BASE_YEAR" { $openedBaseYear = $parts[1] }
        "META"                { $meta[$parts[1]] = $parts[2] }
        "COUNT"               { $counts[$parts[1]] = $parts[2] }
        "REGION"              { $regions += [pscustomobject]@{ Id=$parts[1]; Name=$parts[2]; ResultsShown=$parts[3] } }
        "SCENARIO"            { $scenarios += [pscustomobject]@{ Id=$parts[1]; Abbr=$parts[2]; Name=$parts[3] } }
        "SIMTYPE_STATUS"      {
            $simtypeStatus = if ($parts.Count -ge 3) { "$($parts[1]):$($parts[2])" } else { $parts[1] }
        }
        "SIMTYPE_PAIRS_CHECKED" { $simtypePairs = [int]$parts[1] }
        "SIMTYPE_NONEMPTY"    { $simtypeNonempty = [int]$parts[1] }
        "SIMTYPE_NETWORK_COUNT" { $simtypeNetwork = [int]$parts[1] }
        "SIMTYPE_NET"         { $simtypeNet += [pscustomobject]@{ Region=$parts[2]; Scenario=$parts[3]; Value=$parts[4] } }
        "SIMTYPE_VAL"         { $simtypeVal += [pscustomobject]@{ Region=$parts[2]; Scenario=$parts[3]; Value=$parts[4] } }
        "SIMTYPE_ERR"         { $simtypeErr += [pscustomobject]@{ Region=$parts[2]; Scenario=$parts[3]; ErrNum=$parts[4]; ErrDesc=$parts[5] } }
        "BROKEN_EXACT"        { $brokenExact += [pscustomobject]@{ Path=$parts[1]; Exists=$parts[2]; Loading=$parts[3] } }
        "HOOK_FOUND"          { $hookFound += [pscustomobject]@{ Path=$parts[1]; Size=$parts[2]; Lines=""; HasNodal=""; HasKazNorth="" } }
        "HOOK_LINES"          {
            $h = $hookFound | Where-Object { $_.Path -eq $parts[1] } | Select-Object -First 1
            if ($h) { $h.Lines = $parts[2] }
        }
        "HOOK_NODAL_REF"      {
            $h = $hookFound | Where-Object { $_.Path -eq $parts[1] } | Select-Object -First 1
            if ($h) { $h.HasNodal = $parts[2] }
        }
        "HOOK_KAZ_NORTH_REF"  {
            $h = $hookFound | Where-Object { $_.Path -eq $parts[1] } | Select-Object -First 1
            if ($h) { $h.HasKazNorth = $parts[2] }
        }
        "HOOK_STATUS"         { $hookSearched = "STATUS:$($parts[1])" }
        "HOOK_SEARCHED_PATHS" { $hookSearched = $parts[1] }
    }
}

# Report
$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# Audit Report v4 -- index $AreaIndex -- $timestamp")
$md.Add("")
$md.Add("**Script:** S02_audit_model_v4.ps1")
$md.Add("**Fixes from v3:** SimType accessor via BranchVariable (canonical pattern), beforeCalculation hook file probe added.")
$md.Add("**Requested AreaIndex:** $AreaIndex")
$md.Add("**Resolved area name:** ``$resolvedAreaName``")
$md.Add("**Area directory:** ``$areaDir``")
$md.Add("**Opened ActiveArea:** ``$openedAreaName`` (BaseYear=$openedBaseYear)")
$md.Add("**Runtime:** $([math]::Round($elapsed.TotalSeconds, 1))s")
$md.Add("**Last phase:** $lastPhase")
$md.Add("")

if ($openedBaseYear -eq "2024") {
    $md.Add("**Open sanity: PASSED.** BaseYear=2024 matches prototype target.")
} else {
    $md.Add("**Open sanity: FAILED.** BaseYear=$openedBaseYear, expected 2024.")
}
$md.Add("")

if ($fatal) {
    $md.Add("## FATAL")
    $md.Add("``````"); $md.Add($fatal); $md.Add("``````")
} else {
    $md.Add("## Counts")
    $md.Add("")
    $md.Add("| Object | Count |")
    $md.Add("|---|---|")
    foreach ($k in $counts.Keys | Sort-Object) { $md.Add("| $k | $($counts[$k]) |") }
    $md.Add("")

    # SimType
    $md.Add("## ISSUE-001 (v4): SimType via BranchVariable")
    $md.Add("")
    $md.Add("**Accessor status:** ``$simtypeStatus``")
    $md.Add("**Pairs checked:** $simtypePairs / **Non-empty:** $simtypeNonempty / **Network:** $simtypeNetwork / **Errors:** $($simtypeErr.Count)")
    $md.Add("")
    if ($simtypeNet.Count -gt 0) {
        $md.Add("### NetworkSimulation usage")
        $md.Add("")
        $md.Add("| Region | Scenario | Expression |")
        $md.Add("|---|---|---|")
        foreach ($x in $simtypeNet) {
            $md.Add("| $($x.Region) | ``$($x.Scenario)`` | ``$($x.Value)`` |")
        }
        $md.Add("")
    }
    if ($simtypeVal.Count -gt 0) {
        $md.Add("### Non-Network values (first 30)")
        $md.Add("")
        $md.Add("| Region | Scenario | Expression |")
        $md.Add("|---|---|---|")
        foreach ($x in $simtypeVal | Select-Object -First 30) {
            $md.Add("| $($x.Region) | ``$($x.Scenario)`` | ``$($x.Value)`` |")
        }
        if ($simtypeVal.Count -gt 30) {
            $md.Add("")
            $md.Add("_$($simtypeVal.Count - 30) more not shown._")
        }
        $md.Add("")
    }
    if ($simtypeErr.Count -gt 0) {
        $md.Add("### SimType accessor errors (first 20)")
        $md.Add("")
        $md.Add("| Region | Scenario | Err | Desc |")
        $md.Add("|---|---|---|---|")
        foreach ($x in $simtypeErr | Select-Object -First 20) {
            $md.Add("| $($x.Region) | $($x.Scenario) | $($x.ErrNum) | $($x.ErrDesc) |")
        }
        $md.Add("")
    }

    # Hook
    $md.Add("## ISSUE-001 supplement: beforeCalculation hook file probe")
    $md.Add("")
    if ($hookFound.Count -gt 0) {
        $md.Add("| Path | Size | Lines | Has 'nodal' | Has 'KAZ_North' |")
        $md.Add("|---|---|---|---|---|")
        foreach ($h in $hookFound) {
            $md.Add("| ``$($h.Path)`` | $($h.Size) | $($h.Lines) | $($h.HasNodal) | $($h.HasKazNorth) |")
        }
        $md.Add("")
        $anyNodal = ($hookFound | Where-Object { $_.HasNodal -eq "True" }).Count -gt 0
        if ($anyNodal) {
            $md.Add("**Nodal distribution hook IS present** in the area folder. ISSUE-001 logic is loaded by LEAP at calc time.")
        } else {
            $md.Add("**No nodal distribution references found in any hook file.** ISSUE-001 may be inactive in KAZ_2024.")
        }
    } else {
        $md.Add("**No hook file found at any of the candidate paths.**")
        $md.Add("")
        $md.Add("Paths searched:")
        $md.Add("``````"); $md.Add($hookSearched); $md.Add("``````")
        $md.Add("")
        $md.Add("Possible: hooks live inside the .leap zip (not extracted to area folder), or under a different name. Either way ISSUE-001 hook may not be active in KAZ_2024.")
    }
    $md.Add("")

    # Broken exact
    $md.Add("## ISSUE-002: 5 broken-unit branches (exact)")
    $md.Add("")
    $md.Add("| Path | Branch | Loading |")
    $md.Add("|---|---|---|")
    foreach ($x in $brokenExact) {
        $md.Add("| ``$($x.Path)`` | $($x.Exists) | $($x.Loading) |")
    }
    $md.Add("")
    $stillBroken = ($brokenExact | Where-Object { $_.Loading -eq "HAS_LOADING" }).Count
    $md.Add("**$stillBroken / 5 still have Avg Environmental Loading at the exact original path.**")
    $md.Add("")
    $md.Add("(Fuzzy search dropped in v4; will land in v5 with proper filtering.)")
    $md.Add("")

    # Verdict
    $md.Add("## Verdict")
    $md.Add("")
    if ($simtypeNetwork -eq 0 -and $simtypeNonempty -eq 0) {
        $md.Add("- ISSUE-001 (SimType): 0 explicit values at any (region, scenario). Variable inherits from Current Accounts. Day 3 fix must target CA expression specifically, OR neutralize the beforeCalculation hook directly (option B).")
    } elseif ($simtypeNetwork -gt 0) {
        $md.Add("- ISSUE-001 (SimType): $simtypeNetwork NetworkSimulation pairs identified for Day 3 neutralization.")
    } else {
        $md.Add("- ISSUE-001 (SimType): $simtypeNonempty non-empty expressions, 0 with Network keyword. Probably already Standard. Day 3 may be a no-op.")
    }

    if ($hookFound.Count -gt 0 -and ($hookFound | Where-Object { $_.HasNodal -eq "True" }).Count -gt 0) {
        $md.Add("- ISSUE-001 hook: PRESENT in area folder. Even if SimType is not Network, the hook itself may trigger the error if it computes invalid distributions. Plan to neutralize the hook in Day 3.")
    } else {
        $md.Add("- ISSUE-001 hook: NOT detected as a separate file. Possibly inactive in KAZ_2024 already.")
    }

    $md.Add("- ISSUE-002 exact: $stillBroken/5 verbatim. Fuzzy filtering (v5) needed to find remapped paths.")
}

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"
if ($fatal) { exit 1 } else { exit 0 }
