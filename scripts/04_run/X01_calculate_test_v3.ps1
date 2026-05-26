# X01_calculate_test_v2.ps1
# v2 patches X01 v1 with two fixes:
#   1. SIGNATURE: oLEAP.Calculate False (cycle 006 discovery; bare call -> Err 450)
#   2. RULE 7:   elapsed time uses Replace(CStr(Round...), ",", ".") not FormatNumber
#
# Usage:
#   .\scripts\04_run\X01_calculate_test_v2.ps1 -AreaIndex 6
#   .\scripts\04_run\X01_calculate_test_v2.ps1 -AreaIndex 6 -OnlyScenario "S0"
#
# Note: Calculate is called with False (likely no progress dialog / no save).
#       If LEAP prompts on close, click No.

param(
    [string]$AreaName = "KAZ_2024",
    [string]$OnlyScenario = "",
    [int]$CalcTimeoutMinutes = 30
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$safeName = ($AreaName -replace '[^A-Za-z0-9_]','_')
$dataFile   = Join-Path $projectRoot "data\audit_reports\calctest_v3_${safeName}_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\calctest_v3_${safeName}_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\X01v3_${safeName}_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== X01 v3: Calculate test (name-based area lookup) ==="
Log "Area name: $AreaName"
Log "Only scenario filter: $(if ($OnlyScenario) { $OnlyScenario } else { '(all enabled)' })"
Log "Timeout: $CalcTimeoutMinutes min"

# ----------------------------------------------------------------------------
# VBS that opens area, runs Calculate, captures everything to data file.
# Placeholders: {{AREA_INDEX}}, {{DATA_FILE}}, {{ONLY_SCENARIO}}
# ----------------------------------------------------------------------------
$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile
Dim a, target_area, current_idx, s
Dim err_num, err_desc
Dim only_scen, start_time, end_time

' Bool serialization rule: write "1"/"0", never CStr(boolean)
Sub WBool(key, b)
    If b Then outFile.WriteLine(key & "|1") Else outFile.WriteLine(key & "|0")
End Sub

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
W "TARGET_NAME", "{{AREA_NAME}}"
only_scen = "{{ONLY_SCENARIO}}"
W "ONLY_SCENARIO", only_scen

Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    outFile.WriteLine "FATAL|COM_CREATE|" & Err.Number & ":" & Err.Description
    outFile.Close
    WScript.Quit 1
End If
Err.Clear

W "PHASE", "com_created"
WScript.Sleep 4000

' Resolve area by NAME (after cycle 010 finding: index ordering is not stable
' across sessions when new recovery areas appear)
Set target_area = Nothing
Dim found_names
found_names = ""
For Each a In oLEAP.Areas
    found_names = found_names & "," & a.Name
    If a.Name = "{{AREA_NAME}}" Then
        Set target_area = a
        Exit For
    End If
Next

If target_area Is Nothing Then
    outFile.WriteLine "FATAL|NAME_NOT_FOUND|wanted={{AREA_NAME}}|seen=" & found_names
    outFile.Close
    oLEAP.Quit
    WScript.Quit 2
End If

W "RESOLVED_AREA_NAME", target_area.Name

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

W "OPENED_AREA_NAME", oLEAP.ActiveArea.Name
W "OPENED_AREA_BASE_YEAR", oLEAP.BaseYear

oLEAP.Verbose = 0

' Optional: enable only one scenario via ResultsShown
If Len(only_scen) > 0 Then
    W "PHASE", "filtering_scenarios"
    Dim n_enabled, n_disabled
    n_enabled = 0
    n_disabled = 0
    For Each s In oLEAP.Scenarios
        Err.Clear
        If s.Abbreviation = only_scen Or s.Name = only_scen Then
            s.ResultsShown = True
            n_enabled = n_enabled + 1
        Else
            s.ResultsShown = False
            n_disabled = n_disabled + 1
        End If
        Err.Clear
    Next
    W "SCENARIOS_ENABLED", n_enabled
    W "SCENARIOS_DISABLED", n_disabled
End If

' Start calculation
W "PHASE", "calc_starting"
start_time = Timer

Err.Clear
oLEAP.Calculate False
err_num = Err.Number
err_desc = Err.Description
Err.Clear

end_time = Timer
W "CALC_ELAPSED_SECONDS", Replace(CStr(Round(end_time - start_time, 2)), ",", ".")

If err_num <> 0 Then
    outFile.WriteLine "CALC_RESULT|FAILED|" & err_num & "|" & err_desc
Else
    outFile.WriteLine "CALC_RESULT|OK|0|"
End If

W "PHASE", "calc_done"

' Try to read any diagnostics LEAP exposes
Err.Clear
Dim diag_text
diag_text = ""
diag_text = oLEAP.Diagnostics
If Err.Number = 0 And Len(diag_text) > 0 Then
    ' Split into lines, write each
    Dim diag_lines, di
    diag_lines = Split(Replace(diag_text, vbCrLf, vbLf), vbLf)
    For di = 0 To UBound(diag_lines)
        outFile.WriteLine "DIAG|" & diag_lines(di)
    Next
    W "DIAG_LINE_COUNT", UBound(diag_lines) + 1
Else
    W "DIAG_AVAILABLE", "0"
End If
Err.Clear

' Last result year / first year accessible?
Err.Clear
Dim rys
rys = oLEAP.ResultsYears
If Err.Number = 0 Then W "RESULTS_YEARS", rys
Err.Clear

' Discard any pending changes; do NOT save
W "PHASE", "closing_without_save"

' Try various close-without-save patterns
Err.Clear
oLEAP.ActiveArea.Close False
If Err.Number <> 0 Then
    Err.Clear
    ' Fallback: just quit, LEAP may prompt on screen
End If

W "PHASE", "quitting"
oLEAP.Verbose = 4
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_NAME}}', $AreaName).Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\')).Replace('{{ONLY_SCENARIO}}', $OnlyScenario)

$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: VBS source contains non-ASCII"; exit 99
}

$vbsPath = Join-Path $env:TEMP "X01_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath"
Log ""
Log "Running Calculate. Could take many minutes; timeout=$CalcTimeoutMinutes min."

$startTime = Get-Date

# Run with timeout
$proc = Start-Process -FilePath "cscript" -ArgumentList "//NoLogo", "`"$vbsPath`"" -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\X01_stdout_${timestamp}.txt" -RedirectStandardError "$env:TEMP\X01_stderr_${timestamp}.txt"
$timedOut = $false
if (-not $proc.WaitForExit($CalcTimeoutMinutes * 60 * 1000)) {
    Log "TIMEOUT after $CalcTimeoutMinutes minutes. Killing cscript and LEAP."
    try { $proc.Kill() } catch {}
    Get-Process -Name "LEAP" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    $timedOut = $true
}
$elapsed = (Get-Date) - $startTime
Log "Process ended after $([math]::Round($elapsed.TotalSeconds, 1))s (timeout=$timedOut)"

$cscriptStdout = Get-Content "$env:TEMP\X01_stdout_${timestamp}.txt" -ErrorAction SilentlyContinue
$cscriptStderr = Get-Content "$env:TEMP\X01_stderr_${timestamp}.txt" -ErrorAction SilentlyContinue
Remove-Item "$env:TEMP\X01_stdout_${timestamp}.txt", "$env:TEMP\X01_stderr_${timestamp}.txt", $vbsPath -ErrorAction SilentlyContinue

foreach ($l in $cscriptStdout) { if ($l -and $l.Trim().Length -gt 0) { Log "stdout: $l" } }
foreach ($l in $cscriptStderr) { if ($l -and $l.Trim().Length -gt 0) { Log "stderr: $l" } }

if (-not (Test-Path $dataFile)) {
    Log "FATAL: data file missing. Calc may have crashed LEAP before any output."
    [System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
    exit 11
}

# ----------------------------------------------------------------------------
# Parse and report
# ----------------------------------------------------------------------------
$dataLines = Get-Content $dataFile -Encoding Unicode
Log "Data lines: $($dataLines.Count)"

$meta = @{}
$calcResult = $null
$calcErrNum = $null
$calcErrDesc = $null
$calcElapsed = $null
$diagLines = @()
$fatal = $null
$lastPhase = $null
$openedBaseYear = $null
$scenariosEnabled = $null

foreach ($raw in $dataLines) {
    $line = "$raw"
    if ($line.Trim().Length -eq 0) { continue }
    $parts = $line -split '\|', 4
    switch ($parts[0]) {
        "PHASE"            { $lastPhase = $parts[1] }
        "FATAL"            { $fatal = "$($parts[1]):$($parts[2..($parts.Count-1)] -join ':')" }
        "RESOLVED_AREA_NAME" { $meta["ResolvedAreaName"] = $parts[1] }
        "OPENED_AREA_NAME" { $meta["OpenedAreaName"] = $parts[1] }
        "OPENED_AREA_BASE_YEAR" { $openedBaseYear = $parts[1] }
        "SCENARIOS_ENABLED" { $scenariosEnabled = $parts[1] }
        "CALC_RESULT"      {
            $calcResult = $parts[1]
            $calcErrNum = $parts[2]
            $calcErrDesc = $parts[3]
        }
        "CALC_ELAPSED_SECONDS" { $calcElapsed = $parts[1] }
        "DIAG"             { $diagLines += $parts[1] }
    }
}

$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# Calculate test report -- idx $AreaIndex -- $timestamp")
$md.Add("")
$md.Add("**Script:** X01_calculate_test.ps1")
$md.Add("**Area index:** $AreaIndex")
$md.Add("**Only scenario filter:** ``$OnlyScenario``")
$md.Add("**Resolved area:** ``$($meta['ResolvedAreaName'])``")
$md.Add("**Opened active area:** ``$($meta['OpenedAreaName'])`` (BaseYear=$openedBaseYear)")
$md.Add("**Wall time:** $([math]::Round($elapsed.TotalSeconds, 1))s ($(if ($timedOut) {'TIMED OUT'} else {'completed'}))")
$md.Add("**Calc elapsed (LEAP-reported):** ${calcElapsed}s")
$md.Add("**Last phase reached:** $lastPhase")
if ($scenariosEnabled) {
    $md.Add("**Scenarios enabled for this run:** $scenariosEnabled")
}
$md.Add("")

if ($fatal) {
    $md.Add("## FATAL")
    $md.Add("``````"); $md.Add($fatal); $md.Add("``````")
    $md.Add("")
} elseif ($timedOut) {
    $md.Add("## TIMED OUT")
    $md.Add("")
    $md.Add("Calculation did not complete within $CalcTimeoutMinutes minutes. LEAP was force-killed.")
    $md.Add("")
    $md.Add("Possible causes: very slow calc, infinite loop in hook, LEAP popup that the script can't dismiss.")
    $md.Add("")
} elseif ($calcResult -eq "OK") {
    $md.Add("## CALC RESULT: SUCCESS")
    $md.Add("")
    $md.Add("Calculate completed without raising an exception.")
    $md.Add("")
    $md.Add("**Next step:** open LEAP UI, navigate to Results view, verify charts display numbers.")
    $md.Add("")
} else {
    $md.Add("## CALC RESULT: FAILED")
    $md.Add("")
    $md.Add("| Field | Value |")
    $md.Add("|---|---|")
    $md.Add("| Err.Number | $calcErrNum |")
    $md.Add("| Err.Description | $calcErrDesc |")
    $md.Add("")
    $md.Add("**Failure mode classification (best-guess):**")
    if ($calcErrDesc -match "nodal|distribution") {
        $md.Add("- ISSUE-001 (nodal distribution) still active.")
    } elseif ($calcErrDesc -match "unit|loading|emission") {
        $md.Add("- ISSUE-002 (broken units) probably active.")
    } elseif ($calcErrDesc -match "Pipeline|network|NEMO") {
        $md.Add("- Probable NEMO/Pipeline optimization issue.")
    } else {
        $md.Add("- Unknown / new failure mode. Investigate diagnostics.")
    }
    $md.Add("")
}

if ($diagLines.Count -gt 0) {
    $md.Add("## LEAP Diagnostics ($($diagLines.Count) lines)")
    $md.Add("")
    $md.Add('```')
    foreach ($d in $diagLines | Select-Object -First 200) { $md.Add($d) }
    if ($diagLines.Count -gt 200) {
        $md.Add("... ($($diagLines.Count - 200) more lines truncated, see data file)")
    }
    $md.Add('```')
    $md.Add("")
}

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"

if ($fatal -or $timedOut) { exit 1 } else { exit 0 }
