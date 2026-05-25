# S02_audit_model_v3.ps1
# Day 2 audit, REVISION 3.
# Identical to v2 in audit content, but with index-based area addressing.
#
# Why v3 (cycle 002 finding):
#   Two areas resolve under the same key "kaz_workshop exercise" -- the parent
#   (BaseYear=2010, pre-migration) and a nested area "KAZ_2024" (post-migration).
#   oLEAP.Areas("string").Open is non-deterministic when keys collide.
#   v3 takes a numeric INDEX into oLEAP.Areas instead, eliminating ambiguity.
#
# Usage:
#   .\scripts\01_scout\S02_audit_model_v3.ps1 -AreaIndex 2
#
# To find the right index, run S00_list_areas.ps1 first; it prints a table
# of every installed area with Index, Name, BaseYear, etc.
#
# READ-ONLY. Never calls Save.

param(
    [Parameter(Mandatory=$true)]
    [int]$AreaIndex,
    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\audit_cycle_003_idx${AreaIndex}_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\audit_cycle_003_idx${AreaIndex}_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\S02v3_idx${AreaIndex}_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== S02 v3: Audit by area INDEX (READ-ONLY) ==="
Log "Area index: $AreaIndex"
Log "Data file: $dataFile"
Log ""

# ----------------------------------------------------------------------------
# VBS template. Index addressing replaces name-string lookup.
# Placeholders: {{AREA_INDEX}}, {{DATA_FILE}}
# ----------------------------------------------------------------------------
$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile
Dim r, s, b, var
Dim sval, expr, err_num, err_desc
Dim i
Dim n_scenarios, n_regions, n_pairs_checked, n_network_found
Dim target_area, current_idx, a
Dim path_arr, broken_path

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

' --- Resolve area by INDEX, not by name string ---
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
    outFile.WriteLine "FATAL|INDEX_OUT_OF_RANGE|requested=" & {{AREA_INDEX}} & "|total=" & current_idx
    outFile.Close
    oLEAP.Quit
    WScript.Quit 2
End If

outFile.WriteLine "RESOLVED_AREA|" & {{AREA_INDEX}} & "|" & target_area.Name

' --- Open it ---
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

' --- Sanity log: confirm what we actually got ---
W "OPENED_AREA_NAME", oLEAP.ActiveArea.Name
W "OPENED_AREA_BASE_YEAR", oLEAP.BaseYear
W "OPENED_AREA_FIRST_SCENARIO_YEAR", oLEAP.FirstScenarioYear
W "OPENED_AREA_END_YEAR", oLEAP.EndYear

' === Standard audit body (identical to v2) ===

' Metadata (redundant with above, kept for parser compatibility with v2)
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

' SimType (v2 logic, unchanged)
W "PHASE", "issue_001_simtype"

If Not oLEAP.Branches.Exists("Transformation\Electricity Production") Then
    W "SIMTYPE_STATUS", "BRANCH_MISSING"
Else
    If Not oLEAP.Branches("Transformation\Electricity Production").VariableExists("Simulation Type") Then
        W "SIMTYPE_STATUS", "VAR_NOT_EXISTS"
    Else
        Err.Clear
        Set var = oLEAP.Branches("Transformation\Electricity Production").Variable("Simulation Type")
        If Err.Number <> 0 Then
            outFile.WriteLine "SIMTYPE_STATUS|VAR_GET_ERROR|" & Err.Number & ":" & Err.Description
            Err.Clear
        ElseIf var Is Nothing Then
            W "SIMTYPE_STATUS", "VAR_IS_NOTHING"
        Else
            W "SIMTYPE_STATUS", "VAR_OK"

            n_pairs_checked = 0
            n_network_found = 0

            For Each r In oLEAP.Regions
                For Each s In oLEAP.Scenarios
                    Err.Clear
                    sval = ""
                    sval = var.ExpressionRS(r.Id, s.Id)
                    err_num = Err.Number
                    err_desc = Err.Description
                    Err.Clear
                    n_pairs_checked = n_pairs_checked + 1
                    If err_num <> 0 Then
                        outFile.WriteLine "SIMTYPE_ERR|" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & err_num & "|" & err_desc
                    ElseIf Len(sval) > 0 Then
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
            outFile.WriteLine "SIMTYPE_NETWORK_COUNT|" & n_network_found
        End If
    End If
End If

' Broken paths exact
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

' Fuzzy walk
W "PHASE", "issue_002_fuzzy"

Sub WalkAndMatch(branch, depth)
    If depth > 10 Then Exit Sub
    Dim full_name, child
    full_name = branch.FullName
    If InStr(full_name, "Lubricants") > 0 Or _
       InStr(full_name, "Methane") > 0 Or _
       InStr(full_name, "Nitrous Oxide") > 0 Or _
       (InStr(full_name, "LPG") > 0 And depth >= 4) Then
        outFile.WriteLine "FUZZY|" & branch.Id & "|" & full_name & "|" & branch.BranchType
    End If
    Err.Clear
    For Each child In branch.Children
        WalkAndMatch child, depth + 1
    Next
    Err.Clear
End Sub

If oLEAP.Branches.Exists("Demand") Then
    WalkAndMatch oLEAP.Branches("Demand"), 0
End If
Err.Clear

W "PHASE", "closing"
oLEAP.Verbose = 4
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_INDEX}}', "$AreaIndex").Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))

# ASCII check
$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: VBS source contains non-ASCII"; exit 99
}

$vbsPath = Join-Path $env:TEMP "S02v3_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath"
Log ""
Log "Running cscript. 5-10 minutes expected."
Log ""

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

# ----------------------------------------------------------------------------
# Parse (same parser as v2, plus RESOLVED_AREA / OPENED_AREA_* sanity fields)
# ----------------------------------------------------------------------------
$dataLines = Get-Content $dataFile -Encoding Unicode
Log "Read $($dataLines.Count) data lines"

$meta = @{}
$counts = @{}
$regions = @()
$scenarios = @()
$simtypeStatus = $null
$simtypeNet = @()
$simtypeErr = @()
$simtypeVal = @()
$simtypePairsChecked = 0
$simtypeNetworkCount = 0
$brokenExact = @()
$fuzzy = @()
$fatal = $null
$lastPhase = $null
$resolvedAreaName = $null
$openedAreaName = $null
$openedBaseYear = $null

foreach ($raw in $dataLines) {
    $line = "$raw".Trim()
    if ($line.Length -eq 0) { continue }
    $parts = $line -split '\|'
    switch ($parts[0]) {
        "PHASE"      { $lastPhase = $parts[1] }
        "FATAL"      { $fatal = "$($parts[1]):$($parts[2..($parts.Count-1)] -join ':')" }
        "RESOLVED_AREA"       { $resolvedAreaName = $parts[2] }
        "OPENED_AREA_NAME"    { $openedAreaName = $parts[1] }
        "OPENED_AREA_BASE_YEAR" { $openedBaseYear = $parts[1] }
        "META"       { $meta[$parts[1]] = $parts[2] }
        "COUNT"      { $counts[$parts[1]] = $parts[2] }
        "REGION"     { $regions += [pscustomobject]@{ Id=$parts[1]; Name=$parts[2]; ResultsShown=$parts[3] } }
        "SCENARIO"   { $scenarios += [pscustomobject]@{ Id=$parts[1]; Abbr=$parts[2]; Name=$parts[3] } }
        "SIMTYPE_STATUS" {
            $simtypeStatus = if ($parts.Count -ge 3) { "$($parts[1]):$($parts[2])" } else { $parts[1] }
        }
        "SIMTYPE_PAIRS_CHECKED" { $simtypePairsChecked = [int]$parts[1] }
        "SIMTYPE_NETWORK_COUNT" { $simtypeNetworkCount = [int]$parts[1] }
        "SIMTYPE_NET" { $simtypeNet += [pscustomobject]@{ RegionId=$parts[1]; Region=$parts[2]; Scenario=$parts[3]; Value=$parts[4] } }
        "SIMTYPE_VAL" { $simtypeVal += [pscustomobject]@{ RegionId=$parts[1]; Region=$parts[2]; Scenario=$parts[3]; Value=$parts[4] } }
        "SIMTYPE_ERR" { $simtypeErr += [pscustomobject]@{ RegionId=$parts[1]; Region=$parts[2]; Scenario=$parts[3]; ErrNum=$parts[4]; ErrDesc=$parts[5] } }
        "BROKEN_EXACT" { $brokenExact += [pscustomobject]@{ Path=$parts[1]; Exists=$parts[2]; Loading=$parts[3] } }
        "FUZZY"      { $fuzzy += [pscustomobject]@{ Id=$parts[1]; Path=$parts[2]; BranchType=$parts[3] } }
    }
}

# ----------------------------------------------------------------------------
# Report
# ----------------------------------------------------------------------------
$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# Audit Report v3 -- index $AreaIndex -- $timestamp")
$md.Add("")
$md.Add("**Script:** S02_audit_model_v3.ps1")
$md.Add("**Area resolution method:** by INDEX (eliminates same-key ambiguity)")
$md.Add("**Requested AreaIndex:** $AreaIndex")
$md.Add("**Resolved area name:** ``$resolvedAreaName``")
$md.Add("**Opened ActiveArea name:** ``$openedAreaName``")
$md.Add("**Opened area BaseYear:** $openedBaseYear")
$md.Add("**Runtime:** $([math]::Round($elapsed.TotalSeconds, 1))s")
$md.Add("**Last phase:** $lastPhase")
$md.Add("")

# Sanity verdict on the open
if ($openedBaseYear -eq "2024") {
    $md.Add("**Open sanity: PASSED.** BaseYear=2024 matches prototype target.")
} elseif ($openedBaseYear -eq "2010") {
    $md.Add("**Open sanity: FAILED.** BaseYear=2010, this is the pre-migration parent area, not the prototype target. Use a different AreaIndex.")
} elseif ($openedBaseYear) {
    $md.Add("**Open sanity: UNEXPECTED.** BaseYear=$openedBaseYear, not 2024 or 2010. Investigate.")
} else {
    $md.Add("**Open sanity: UNKNOWN.** BaseYear not captured. Probably FATAL above.")
}
$md.Add("")

if ($fatal) {
    $md.Add("## FATAL")
    $md.Add("``````")
    $md.Add($fatal)
    $md.Add("``````")
} else {
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
    $md.Add("<details><summary>All scenarios</summary>")
    $md.Add("")
    $md.Add("| Id | Abbr | Name |")
    $md.Add("|---|---|---|")
    foreach ($s in $scenarios) { $md.Add("| $($s.Id) | ``$($s.Abbr)`` | $($s.Name) |") }
    $md.Add("")
    $md.Add("</details>")
    $md.Add("")

    # SimType
    $md.Add("## ISSUE-001: SimType verification")
    $md.Add("")
    $md.Add("**Variable status:** ``$simtypeStatus``")
    $md.Add("**Pairs checked:** $simtypePairsChecked / **Network found:** $simtypeNetworkCount / **Errors:** $($simtypeErr.Count) / **Non-Network values:** $($simtypeVal.Count)")
    $md.Add("")
    if ($simtypeNet.Count -gt 0) {
        $md.Add("### NetworkSimulation usage (Day 3 fix scope)")
        $md.Add("")
        $md.Add("| Region | Scenario | Expression |")
        $md.Add("|---|---|---|")
        foreach ($x in $simtypeNet) {
            $md.Add("| $($x.Region) | ``$($x.Scenario)`` | ``$($x.Value)`` |")
        }
        $md.Add("")
    }
    if ($simtypeVal.Count -gt 0 -and $simtypeVal.Count -le 10) {
        $md.Add("### Other SimType values (non-Network, non-empty)")
        $md.Add("")
        $md.Add("| Region | Scenario | Expression |")
        $md.Add("|---|---|---|")
        foreach ($x in $simtypeVal) {
            $md.Add("| $($x.Region) | ``$($x.Scenario)`` | ``$($x.Value)`` |")
        }
        $md.Add("")
    }

    # Broken
    $md.Add("## ISSUE-002: 5 broken-unit branches (exact)")
    $md.Add("")
    $md.Add("| Path | Branch | Loading Var |")
    $md.Add("|---|---|---|")
    foreach ($x in $brokenExact) {
        $md.Add("| ``$($x.Path)`` | $($x.Exists) | $($x.Loading) |")
    }
    $md.Add("")
    $exactWithLoading = ($brokenExact | Where-Object { $_.Loading -eq "HAS_LOADING" }).Count
    $md.Add("**$exactWithLoading / 5 still have Avg Environmental Loading.**")
    $md.Add("")

    # Fuzzy
    $md.Add("## ISSUE-002: Fuzzy search ($($fuzzy.Count) matches)")
    $md.Add("")
    if ($fuzzy.Count -gt 0 -and $fuzzy.Count -le 100) {
        $md.Add("| Id | Full path | BranchType |")
        $md.Add("|---|---|---|")
        foreach ($x in $fuzzy) {
            $md.Add("| $($x.Id) | ``$($x.Path)`` | $($x.BranchType) |")
        }
    } elseif ($fuzzy.Count -gt 100) {
        $md.Add("Too many matches ($($fuzzy.Count)) to inline. See raw data file.")
    } else {
        $md.Add("_No matches under Demand._")
    }
    $md.Add("")

    # Verdict
    $md.Add("## Verdict")
    $md.Add("")
    if ($openedBaseYear -eq "2024") {
        $md.Add("This audit ran against the correct prototype-target area (BaseYear=2024).")
        $md.Add("")
        $md.Add("**Day 3 sizing:** $simtypeNetworkCount NetworkSimulation pairs to neutralize")
        $md.Add("**Day 4 sizing:** $exactWithLoading / 5 exact broken paths, $($fuzzy.Count) fuzzy candidates")
        if ($simtypeNetworkCount -eq 0 -and $simtypeStatus -eq "VAR_OK" -and $simtypeVal.Count -eq 0) {
            $md.Add("")
            $md.Add("**Note:** Zero NetworkSimulation expressions found at scenario level. SimType is likely inherited from Current Accounts. Day 3 must scan/modify CA expression instead, or neutralize the beforeCalculation hook directly (ISSUE-001 fix option B).")
        }
    } else {
        $md.Add("Wrong area opened (BaseYear=$openedBaseYear). Re-run with different AreaIndex.")
    }
}

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"
if ($fatal) { exit 1 } else { exit 0 }
