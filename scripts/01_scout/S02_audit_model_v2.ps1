# S02_audit_model_v2.ps1
# Day 2 audit, REVISION 2.
# Fixes from v1 (cycle 001 feedback):
#   1. VBS writes UTF-16 LE file instead of cscript Echo (kills cp866 mojibake)
#   2. SimType reader: Is Nothing check, Err.Clear per call, no scenario_id=0 reliance
#   3. Branch fuzzy search added (Lubricants, Methane, Nitrous Oxide leaves)
#
# READ-ONLY. Never calls Save.
#
# Usage:
#   .\scripts\01_scout\S02_audit_model_v2.ps1 -AreaName "kaz_workshop exercise"
#
# Note: pass the COLLECTION KEY (what oLEAP.Areas indexes by), not the display name.
# Per cycle 001: collection key = "kaz_workshop exercise", display name = "KAZ_2024".

param(
    [string]$AreaName = "kaz_workshop exercise",
    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"

$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\audit_cycle_002_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\audit_cycle_002_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\S02v2_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== S02 v2: Audit model (READ-ONLY, UTF-16 output) ==="
Log "Target area key: $AreaName"
Log "Data file: $dataFile"
Log ""

# ----------------------------------------------------------------------------
# VBScript. ASCII-only source. Writes UTF-16 LE to $dataFile.
# Placeholders: {{AREA_NAME}}, {{DATA_FILE}}
# ----------------------------------------------------------------------------
$vbsTemplate = @'
Option Explicit

' Audit v2. READ-ONLY. Writes UTF-16 LE to data file.

Dim oLEAP, fso, outFile
Dim r, s, b, var
Dim sval, expr, err_num, err_desc
Dim i, broken_path
Dim n_scenarios, n_regions, n_pairs_checked, n_network_found
Dim broken_paths

' Reusable: write a key-value line to output file
Sub W(key, val)
    outFile.WriteLine(key & "|" & val)
End Sub

' Reusable: write a multi-field line
Sub W2(key, v1, v2)
    outFile.WriteLine(key & "|" & v1 & "|" & v2)
End Sub
Sub W3(key, v1, v2, v3)
    outFile.WriteLine(key & "|" & v1 & "|" & v2 & "|" & v3)
End Sub
Sub W4(key, v1, v2, v3, v4)
    outFile.WriteLine(key & "|" & v1 & "|" & v2 & "|" & v3 & "|" & v4)
End Sub

On Error Resume Next

' --- Open output file (UTF-16 LE) ---
Set fso = CreateObject("Scripting.FileSystemObject")
' OpenTextFile(path, 2=ForWriting, True=Create, -1=TristateTrue=Unicode/UTF-16)
Set outFile = fso.OpenTextFile("{{DATA_FILE}}", 2, True, -1)
If Err.Number <> 0 Then
    WScript.Echo "FATAL_FILE:" & Err.Number & ":" & Err.Description
    WScript.Quit 10
End If
Err.Clear

W "PHASE", "starting"

' --- Create LEAP ---
Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    W2 "FATAL", "COM_CREATE", Err.Number & ":" & Err.Description
    outFile.Close
    WScript.Quit 1
End If
Err.Clear

W "PHASE", "com_created"
WScript.Sleep 4000

' --- Open area ---
oLEAP.Areas("{{AREA_NAME}}").Open
If Err.Number <> 0 Then
    W2 "FATAL", "AREA_OPEN", Err.Number & ":" & Err.Description
    W "AVAILABLE_AREAS_BEGIN", ""
    Dim a
    Err.Clear
    For Each a In oLEAP.Areas
        W "AREA", a.Name
    Next
    W "AVAILABLE_AREAS_END", ""
    outFile.Close
    oLEAP.Quit
    WScript.Quit 2
End If
Err.Clear

WScript.Sleep 3000
W "PHASE", "area_opened"
oLEAP.Verbose = 0

' --- Metadata ---
W2 "META", "ActiveArea", oLEAP.ActiveArea.Name
W2 "META", "BaseYear", oLEAP.BaseYear
W2 "META", "FirstScenarioYear", oLEAP.FirstScenarioYear
W2 "META", "EndYear", oLEAP.EndYear

' --- Regions ---
W "PHASE", "regions"
n_regions = 0
For Each r In oLEAP.Regions
    n_regions = n_regions + 1
    W3 "REGION", r.Id, r.Name, CStr(r.ResultsShown)
Next
W2 "COUNT", "Regions", n_regions

' --- Scenarios ---
W "PHASE", "scenarios"
n_scenarios = 0
For Each s In oLEAP.Scenarios
    n_scenarios = n_scenarios + 1
    W4 "SCENARIO", s.Id, s.Abbreviation, s.Name, ""
Next
W2 "COUNT", "Scenarios", n_scenarios

' --- ISSUE-001: SimType reader (v2, robust) ---
W "PHASE", "issue_001_simtype_v2"

If Not oLEAP.Branches.Exists("Transformation\Electricity Production") Then
    W "SIMTYPE_STATUS", "BRANCH_MISSING"
Else
    If Not oLEAP.Branches("Transformation\Electricity Production").VariableExists("Simulation Type") Then
        W "SIMTYPE_STATUS", "VAR_NOT_EXISTS"
    Else
        Err.Clear
        Set var = oLEAP.Branches("Transformation\Electricity Production").Variable("Simulation Type")
        If Err.Number <> 0 Then
            W2 "SIMTYPE_STATUS", "VAR_GET_ERROR", Err.Number & ":" & Err.Description
            Err.Clear
        ElseIf var Is Nothing Then
            W "SIMTYPE_STATUS", "VAR_IS_NOTHING"
        Else
            W "SIMTYPE_STATUS", "VAR_OK"

            n_pairs_checked = 0
            n_network_found = 0

            ' Iterate every (region, scenario) pair, capture expression
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
                        ' Log error rows individually
                        outFile.WriteLine("SIMTYPE_ERR|" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & err_num & "|" & err_desc)
                    ElseIf Len(sval) > 0 Then
                        ' Non-empty expression. Flag if Network in expression.
                        If InStr(sval, "Network") > 0 Then
                            outFile.WriteLine("SIMTYPE_NET|" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & sval)
                            n_network_found = n_network_found + 1
                        Else
                            outFile.WriteLine("SIMTYPE_VAL|" & r.Id & "|" & r.Name & "|" & s.Abbreviation & "|" & sval)
                        End If
                    End If
                    ' If sval is empty string, scenario inherits, don't log
                Next
            Next

            W2 "SIMTYPE_PAIRS_CHECKED", n_pairs_checked, ""
            W2 "SIMTYPE_NETWORK_COUNT", n_network_found, ""
        End If
    End If
End If

' --- ISSUE-002: 5 expected broken paths (exact check) ---
W "PHASE", "issue_002_exact_check"

' Build broken_paths array as a string to iterate (VBScript array literal pain)
Dim paths_str
paths_str = "Demand\Agriculture\Syr Darya\Other\Lubricants\Methane|" & _
            "Demand\Agriculture\Other\Lubricants\Methane|" & _
            "Demand\Industry\Iron and Steel\Top down\LPG\Nitrous Oxide|" & _
            "Demand\Industry\Other\Top Down\All Other\LPG\Nitrous Oxide|" & _
            "Demand\Commercial\Lubricants\Methane"
Dim path_arr
path_arr = Split(paths_str, "|")

For i = 0 To UBound(path_arr)
    broken_path = path_arr(i)
    If oLEAP.Branches.Exists(broken_path) Then
        If oLEAP.Branches(broken_path).VariableExists("Avg Environmental Loading") Then
            W3 "BROKEN_EXACT", broken_path, "EXISTS", "HAS_LOADING"
        Else
            W3 "BROKEN_EXACT", broken_path, "EXISTS", "NO_LOADING"
        End If
    Else
        W3 "BROKEN_EXACT", broken_path, "MISSING", "N/A"
    End If
Next

' --- ISSUE-002: Fuzzy search (walk tree, find candidates) ---
W "PHASE", "issue_002_fuzzy_walk"

' Recursive walk of Demand subtree, log any branch with Lubricants/Methane/Nitrous Oxide in path
Sub WalkAndMatch(branch, depth)
    If depth > 10 Then Exit Sub  ' safety
    Dim full_name, child
    full_name = branch.FullName

    If InStr(full_name, "Lubricants") > 0 Or _
       InStr(full_name, "Methane") > 0 Or _
       InStr(full_name, "Nitrous Oxide") > 0 Or _
       (InStr(full_name, "LPG") > 0 And depth >= 4) Then
        outFile.WriteLine("FUZZY|" & branch.Id & "|" & full_name & "|" & branch.BranchType)
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

W "PHASE", "fuzzy_done"

' --- Close ---
W "PHASE", "closing"
oLEAP.Verbose = 4
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

# Substitute placeholders
$vbs = $vbsTemplate.Replace('{{AREA_NAME}}', $AreaName).Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))

# Sanity: VBS source must be pure ASCII
$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
$nonAscii = $bytes | Where-Object { $_ -gt 127 }
if ($nonAscii) {
    Log "ABORT: VBS source contains non-ASCII bytes. Refusing to write."
    exit 99
}

$vbsPath = Join-Path $env:TEMP "S02v2_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath ($((Get-Item $vbsPath).Length) bytes)"

# ----------------------------------------------------------------------------
# Execute
# ----------------------------------------------------------------------------
Log ""
Log "Running cscript. LEAP cold start 3-4 min, then iteration over 65 scenarios x 6 regions."
Log "Expected: 5 to 10 minutes total."
Log ""

$startTime = Get-Date
$cscriptOut = & cscript //NoLogo $vbsPath 2>&1
$elapsed = (Get-Date) - $startTime
Log "cscript finished in $([math]::Round($elapsed.TotalSeconds, 1)) seconds"
Remove-Item $vbsPath -ErrorAction SilentlyContinue

# cscript output for v2 should be empty (almost everything goes to the data file).
# Only WScript.Echo we kept is for early COM_CREATE / FILE failures.
foreach ($l in $cscriptOut) {
    if ($l -is [string] -and $l.Trim().Length -gt 0) {
        Log "cscript stdout: $l"
    }
}

# ----------------------------------------------------------------------------
# Read data file (UTF-16 LE)
# ----------------------------------------------------------------------------
if (-not (Test-Path $dataFile)) {
    Log "FATAL: data file not produced. Check cscript output above."
    exit 11
}

Log "Reading data file (UTF-16)..."
$dataLines = Get-Content $dataFile -Encoding Unicode
Log "  $($dataLines.Count) lines"

# ----------------------------------------------------------------------------
# Parse
# ----------------------------------------------------------------------------
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
$availableAreas = @()
$fatal = $null
$lastPhase = $null

foreach ($raw in $dataLines) {
    $line = "$raw".Trim()
    if ($line.Length -eq 0) { continue }
    $parts = $line -split '\|'
    $tag = $parts[0]

    switch ($tag) {
        "PHASE"      { $lastPhase = $parts[1] }
        "FATAL"      { $fatal = "$($parts[1]):$($parts[2])" }
        "META"       { $meta[$parts[1]] = $parts[2] }
        "COUNT"      { $counts[$parts[1]] = $parts[2] }
        "REGION"     { $regions += [pscustomobject]@{ Id=$parts[1]; Name=$parts[2]; ResultsShown=$parts[3] } }
        "SCENARIO"   { $scenarios += [pscustomobject]@{ Id=$parts[1]; Abbr=$parts[2]; Name=$parts[3] } }
        "AREA"       { $availableAreas += $parts[1] }
        "SIMTYPE_STATUS" { $simtypeStatus = if ($parts.Count -ge 3) { "$($parts[1]):$($parts[2])" } else { $parts[1] } }
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
# Build markdown report
# ----------------------------------------------------------------------------
$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# Audit Report v2 -- Cycle 002 -- $timestamp")
$md.Add("")
$md.Add("**Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$md.Add("**Script:** scripts/01_scout/S02_audit_model_v2.ps1")
$md.Add("**Area key:** ``$AreaName``")
$md.Add("**Runtime:** $([math]::Round($elapsed.TotalSeconds, 1)) seconds")
$md.Add("**Last phase reached:** ``$lastPhase``")
$md.Add("**Data file:** ``$($dataFile.Replace($projectRoot + '\',''))``")
$md.Add("")
$md.Add("v2 fixes applied: UTF-16 output (no mojibake), Is Nothing check on Variable, Err.Clear per ExpressionRS call, no scenario_id=0 reliance, fuzzy branch search added.")
$md.Add("")

if ($fatal) {
    $md.Add("## FATAL")
    $md.Add("")
    $md.Add("``````")
    $md.Add($fatal)
    $md.Add("``````")
    if ($availableAreas.Count -gt 0) {
        $md.Add("")
        $md.Add("Available areas: " + ($availableAreas -join ", "))
    }
} else {
    # Metadata
    $md.Add("## Metadata")
    $md.Add("")
    $md.Add("| Property | Value |")
    $md.Add("|---|---|")
    foreach ($k in $meta.Keys | Sort-Object) {
        $md.Add("| $k | ``$($meta[$k])`` |")
    }
    $md.Add("")
    $md.Add("| Object | Count |")
    $md.Add("|---|---|")
    foreach ($k in $counts.Keys | Sort-Object) {
        $md.Add("| $k | $($counts[$k]) |")
    }
    $md.Add("")

    # Regions (now with correct UTF-16, Cyrillic preserved)
    $md.Add("## Regions ($($regions.Count))")
    $md.Add("")
    $md.Add("| Id | Name | ResultsShown |")
    $md.Add("|---|---|---|")
    foreach ($r in $regions) {
        $md.Add("| $($r.Id) | $($r.Name) | $($r.ResultsShown) |")
    }
    $md.Add("")

    # Scenarios (compact)
    $md.Add("## Scenarios ($($scenarios.Count))")
    $md.Add("")
    $md.Add("<details><summary>Click to expand all $($scenarios.Count) scenarios</summary>")
    $md.Add("")
    $md.Add("| Id | Abbr | Name |")
    $md.Add("|---|---|---|")
    foreach ($s in $scenarios) {
        $md.Add("| $($s.Id) | ``$($s.Abbr)`` | $($s.Name) |")
    }
    $md.Add("")
    $md.Add("</details>")
    $md.Add("")

    # ISSUE-001
    $md.Add("## ISSUE-001 verification (v2): Transformation Simulation Type")
    $md.Add("")
    $md.Add("**Variable status:** ``$simtypeStatus``")
    $md.Add("**Pairs checked (region x scenario):** $simtypePairsChecked")
    $md.Add("**Pairs with explicit Network expression:** $simtypeNetworkCount")
    $md.Add("**Pairs with read errors:** $($simtypeErr.Count)")
    $md.Add("**Pairs with non-empty non-Network value:** $($simtypeVal.Count)")
    $md.Add("")

    if ($simtypeNet.Count -gt 0) {
        $md.Add("### Scenarios using NetworkSimulation (Day 3 fix scope)")
        $md.Add("")
        $md.Add("| Region | Scenario | Expression |")
        $md.Add("|---|---|---|")
        foreach ($x in $simtypeNet) {
            $md.Add("| $($x.Region) | ``$($x.Scenario)`` | ``$($x.Value)`` |")
        }
        $md.Add("")
    } else {
        $md.Add("**No explicit NetworkSimulation expressions found.** Interpretation: either (a) the SimType is inherited from Current Accounts which we did not scan in v2, or (b) the model genuinely uses Standard simulation everywhere.")
        $md.Add("")
        $md.Add("Sanity check: if pairs_checked > 0 and val_count = 0 and net_count = 0, every scenario inherits and we cannot see CA without scanning it specifically. We add a CA scan in v3 if needed.")
        $md.Add("")
    }

    if ($simtypeErr.Count -gt 0) {
        $md.Add("### SimType read errors (first 20)")
        $md.Add("")
        $md.Add("| Region | Scenario | Err | Desc |")
        $md.Add("|---|---|---|---|")
        foreach ($x in $simtypeErr | Select-Object -First 20) {
            $md.Add("| $($x.Region) | $($x.Scenario) | $($x.ErrNum) | $($x.ErrDesc) |")
        }
        $md.Add("")
    }

    # ISSUE-002 exact
    $md.Add("## ISSUE-002 verification: 5 broken-unit branches (exact paths)")
    $md.Add("")
    $md.Add("| Path | Branch | Loading Var |")
    $md.Add("|---|---|---|")
    foreach ($x in $brokenExact) {
        $md.Add("| ``$($x.Path)`` | $($x.Exists) | $($x.Loading) |")
    }
    $md.Add("")
    $exactExist = ($brokenExact | Where-Object { $_.Exists -eq "EXISTS" }).Count
    $exactWithLoading = ($brokenExact | Where-Object { $_.Loading -eq "HAS_LOADING" }).Count
    $md.Add("**$exactExist / 5 paths exist verbatim. $exactWithLoading / 5 still have Avg Environmental Loading.**")
    $md.Add("")

    # Fuzzy
    $md.Add("## ISSUE-002 fuzzy search: Lubricants / Methane / Nitrous Oxide / LPG leaves")
    $md.Add("")
    $md.Add("**Branches found in Demand subtree matching patterns:** $($fuzzy.Count)")
    $md.Add("")
    if ($fuzzy.Count -gt 0) {
        $md.Add("| Id | Full path | BranchType |")
        $md.Add("|---|---|---|")
        foreach ($x in $fuzzy) {
            $md.Add("| $($x.Id) | ``$($x.Path)`` | $($x.BranchType) |")
        }
        $md.Add("")
        $md.Add("Interpretation: cross-reference these against the 5 expected exact paths. If a matching branch exists under a different parent, the colleague restructured the tree but the broken-unit issue likely still applies, with a new path. If no Lubricants/Methane/Nitrous Oxide branches exist anywhere, the colleague deleted them entirely.")
    } else {
        $md.Add("**No matching branches found anywhere in Demand.** Colleague almost certainly deleted these branches.")
    }
    $md.Add("")

    # Verdict
    $md.Add("## Verdict and recommended next step")
    $md.Add("")
    $md.Add("**Day 3 (transmission disable) sizing:**")
    if ($simtypeNetworkCount -gt 0) {
        $md.Add("- $simtypeNetworkCount explicit NetworkSimulation expressions to neutralize")
    } else {
        $md.Add("- No explicit NetworkSimulation found. Day 3 may need to also handle CA / the beforeCalculation hook directly (see ISSUE-001 fix option B).")
    }
    $md.Add("")
    $md.Add("**Day 4 (broken units) sizing:**")
    $md.Add("- Exact paths: $exactWithLoading / 5 need fix")
    $md.Add("- Fuzzy matches: $($fuzzy.Count) candidate branches identified")
    if ($exactWithLoading -eq 0 -and $fuzzy.Count -eq 0) {
        $md.Add("- **Day 4 likely NO-OP.** Colleague deleted the affected branches. Day 4 becomes a 5-minute confirmation, then skip to Day 5.")
    } elseif ($exactWithLoading -eq 0 -and $fuzzy.Count -gt 0) {
        $md.Add("- **Day 4 needs path remap.** Colleague restructured. Use the fuzzy list to find actual broken branches.")
    } else {
        $md.Add("- **Day 4 proceeds as originally planned.**")
    }
}

$md.Add("")
$md.Add("---")
$md.Add("")
$md.Add("_Raw data: ``$($dataFile.Replace($projectRoot + '\',''))``_")

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report written: $reportPath"

[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log ""
Log "=== S02 v2 complete ==="
Log "Open the report: $reportPath"

if ($fatal) { exit 1 } else { exit 0 }
