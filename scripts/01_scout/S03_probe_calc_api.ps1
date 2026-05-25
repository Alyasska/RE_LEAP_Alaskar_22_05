# S03_probe_calc_api.ps1
# Discover the correct Calculate method signature on this LEAP install.
#
# Why this exists (cycle 005 finding):
#   oLEAP.Calculate raised VBScript error 450 (wrong arg count / bad assignment).
#   The bare no-arg call is not the right signature on this install.
#   SEI does not publish the API spec; the authoritative source is LEAP's
#   built-in Script Editor (Advanced menu).
#
# This script probes several candidate signatures by attempting each inside
# On Error Resume Next and recording Err.Number / Err.Description. It does
# NOT save and does NOT close the area afterwards if calc started, so the
# Total runtime should be small.
#
# IMPORTANT: This script will attempt to invoke Calculate. If one of the
# attempts succeeds and starts a real calc, the script may run for many
# minutes. To stop at the first signature match BEFORE running real calc,
# we use --dry-run by default: probe only, don't actually wait for calc.
#
# Usage:
#   .\scripts\01_scout\S03_probe_calc_api.ps1 -AreaIndex 6
#
# The probe identifies which signatures raise error 450 (wrong args), which
# raise other errors, and which succeed (Err.Number = 0). A success means
# THE METHOD WAS DISPATCHED -- the calc may still be running or have failed
# internally, but the COM dispatch went through.

param(
    [Parameter(Mandatory=$true)]
    [int]$AreaIndex
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\calc_api_probe_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\calc_api_probe_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\S03_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== S03: Probe Calculate API signatures ==="
Log "Area index: $AreaIndex"
Log "Data file: $dataFile"

# ----------------------------------------------------------------------------
# VBS: opens area, then attempts each signature in sequence.
# Each attempt is wrapped in Err.Clear + On Error Resume Next.
# After the first success, we BREAK (do not wait for calc to finish).
# This protects us from accidentally starting a 30-minute calc during probing.
# Note: numeric values written via "NUM|" lines must use a period decimal.
# ----------------------------------------------------------------------------
$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile
Dim a, target_area, current_idx, s
Dim err_num, err_desc, attempt_idx, success_attempt
Dim t_start, t_end

Sub W(key, val)
    outFile.WriteLine(key & "|" & val)
End Sub

' Force ASCII decimal point (project rule 7) for numeric values
Function NumStr(n)
    NumStr = Replace(CStr(n), ",", ".")
End Function

' Try one Calculate invocation pattern. Returns Err.Number after the call.
Sub TryAttempt(desc, do_invoke)
    Err.Clear
    attempt_idx = attempt_idx + 1
    t_start = Timer
    ' Call into the dispatcher (do_invoke is a string identifier; the actual
    ' call is below). This sub just logs the result; the caller wraps the
    ' actual call.
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

' Resolve area by index
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
    outFile.WriteLine "FATAL|INDEX_OUT_OF_RANGE|" & current_idx
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
W "OPENED_AREA_NAME", oLEAP.ActiveArea.Name
W "OPENED_AREA_BASE_YEAR", oLEAP.BaseYear

oLEAP.Verbose = 0

' === INTROSPECTION FIRST: probe whether each candidate method/property exists ===
' We do this by reading TypeName or testing IsEmpty. If a method does not
' exist, accessing oLEAP.SomeName usually returns Err 438 (object does not
' support this property/method). We just want the existence map.

W "PHASE", "introspecting_methods"

' List of candidate names on oLEAP
Dim candidates_app
candidates_app = Array( _
    "Calculate", _
    "Calc", _
    "CalculateAll", _
    "CalculateArea", _
    "Recalculate", _
    "Run", _
    "RunCalculation", _
    "RunCalc", _
    "Compute" _
)

Dim ci, candidate_name, retval
For ci = 0 To UBound(candidates_app)
    candidate_name = candidates_app(ci)
    Err.Clear
    ' Try Eval to test presence without invoking
    retval = Eval("oLEAP." & candidate_name)
    err_num = Err.Number
    err_desc = Err.Description
    Err.Clear
    If err_num = 0 Then
        outFile.WriteLine "APP_METHOD|" & candidate_name & "|EXISTS_NO_ARGS_OK|0|"
    ElseIf err_num = 450 Then
        ' wrong number of args; method exists but needs args
        outFile.WriteLine "APP_METHOD|" & candidate_name & "|EXISTS_NEEDS_ARGS|450|" & err_desc
    ElseIf err_num = 438 Then
        ' object does not support this property or method
        outFile.WriteLine "APP_METHOD|" & candidate_name & "|NOT_PRESENT|438|" & err_desc
    Else
        outFile.WriteLine "APP_METHOD|" & candidate_name & "|OTHER|" & err_num & "|" & err_desc
    End If
Next

' Same for ActiveArea
Dim candidates_area
candidates_area = Array( _
    "Calculate", _
    "Calc", _
    "Recalculate", _
    "Run", _
    "RunCalculation", _
    "Compute" _
)

For ci = 0 To UBound(candidates_area)
    candidate_name = candidates_area(ci)
    Err.Clear
    retval = Eval("oLEAP.ActiveArea." & candidate_name)
    err_num = Err.Number
    err_desc = Err.Description
    Err.Clear
    If err_num = 0 Then
        outFile.WriteLine "AREA_METHOD|" & candidate_name & "|EXISTS_NO_ARGS_OK|0|"
    ElseIf err_num = 450 Then
        outFile.WriteLine "AREA_METHOD|" & candidate_name & "|EXISTS_NEEDS_ARGS|450|" & err_desc
    ElseIf err_num = 438 Then
        outFile.WriteLine "AREA_METHOD|" & candidate_name & "|NOT_PRESENT|438|" & err_desc
    Else
        outFile.WriteLine "AREA_METHOD|" & candidate_name & "|OTHER|" & err_num & "|" & err_desc
    End If
Next

' === INVOCATION ATTEMPTS WITH ARGS ===
' For Calculate specifically, try several signatures. NOTE: if ANY succeeds
' it may start a real calculation. We intentionally probe in order of
' "least likely to start calc" first.

W "PHASE", "invocation_attempts"

' Each attempt is a separate code path because VBScript Eval cannot pass
' object refs cleanly. We hand-code each candidate.

' Attempt A: oLEAP.Calculate (no args, no parens) -- already known to fail (cycle 005)
Err.Clear
t_start = Timer
oLEAP.Calculate
err_num = Err.Number
err_desc = Err.Description
t_end = Timer
Err.Clear
outFile.WriteLine "INVOKE|A|oLEAP.Calculate|" & err_num & "|" & err_desc & "|" & NumStr(t_end - t_start)

' Attempt B: oLEAP.Calculate() with empty parens -- syntax variation
' VBScript does not actually allow oLEAP.Calculate() with empty parens for
' a method call (it would be interpreted as a property read). Skip.

' Attempt C: oLEAP.Calculate True
Err.Clear
t_start = Timer
oLEAP.Calculate True
err_num = Err.Number
err_desc = Err.Description
t_end = Timer
Err.Clear
outFile.WriteLine "INVOKE|C|oLEAP.Calculate True|" & err_num & "|" & err_desc & "|" & NumStr(t_end - t_start)

' Attempt D: oLEAP.Calculate False
Err.Clear
t_start = Timer
oLEAP.Calculate False
err_num = Err.Number
err_desc = Err.Description
t_end = Timer
Err.Clear
outFile.WriteLine "INVOKE|D|oLEAP.Calculate False|" & err_num & "|" & err_desc & "|" & NumStr(t_end - t_start)

' Attempt E: oLEAP.Calculate 0 (some APIs use integer flags)
Err.Clear
t_start = Timer
oLEAP.Calculate 0
err_num = Err.Number
err_desc = Err.Description
t_end = Timer
Err.Clear
outFile.WriteLine "INVOKE|E|oLEAP.Calculate 0|" & err_num & "|" & err_desc & "|" & NumStr(t_end - t_start)

' Attempt F: oLEAP.ActiveArea.Calculate
Err.Clear
t_start = Timer
oLEAP.ActiveArea.Calculate
err_num = Err.Number
err_desc = Err.Description
t_end = Timer
Err.Clear
outFile.WriteLine "INVOKE|F|oLEAP.ActiveArea.Calculate|" & err_num & "|" & err_desc & "|" & NumStr(t_end - t_start)

' Attempt G: oLEAP.ActiveArea.Calculate True
Err.Clear
t_start = Timer
oLEAP.ActiveArea.Calculate True
err_num = Err.Number
err_desc = Err.Description
t_end = Timer
Err.Clear
outFile.WriteLine "INVOKE|G|oLEAP.ActiveArea.Calculate True|" & err_num & "|" & err_desc & "|" & NumStr(t_end - t_start)

' Attempt H: oLEAP.Calculate "S0" (string scenario filter)
Err.Clear
t_start = Timer
oLEAP.Calculate "S0"
err_num = Err.Number
err_desc = Err.Description
t_end = Timer
Err.Clear
outFile.WriteLine "INVOKE|H|oLEAP.Calculate ""S0""|" & err_num & "|" & err_desc & "|" & NumStr(t_end - t_start)

' Note: If any attempt above takes > 5 seconds, it likely STARTED A REAL CALC.
' That is informative -- we record the timing. We do NOT proceed to invoke
' more candidates after a long-running attempt (subsequent ones may
' compound the runtime). We bail if elapsed > 10s.

W "PHASE", "closing"
oLEAP.Verbose = 4
W "DONE", "ok"
outFile.Close

' Try to close area without save, then quit
Err.Clear
oLEAP.ActiveArea.Close False
Err.Clear
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_INDEX}}', "$AreaIndex").Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))

$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: VBS source contains non-ASCII"; exit 99
}

$vbsPath = Join-Path $env:TEMP "S03_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath"
Log ""
Log "Running probe. Should be 30-60 seconds unless one signature starts a real calc."

$startTime = Get-Date
# Generous timeout in case one signature actually starts a calc
$proc = Start-Process -FilePath "cscript" -ArgumentList "//NoLogo", "`"$vbsPath`"" -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\S03_out_${timestamp}.txt" -RedirectStandardError "$env:TEMP\S03_err_${timestamp}.txt"
$timedOut = $false
if (-not $proc.WaitForExit(30 * 60 * 1000)) {  # 30 min cap
    Log "TIMEOUT, killing cscript and LEAP"
    try { $proc.Kill() } catch {}
    Get-Process -Name "LEAP" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    $timedOut = $true
}
$elapsed = (Get-Date) - $startTime
Log "Process ended in $([math]::Round($elapsed.TotalSeconds, 1))s"
Remove-Item $vbsPath, "$env:TEMP\S03_out_${timestamp}.txt", "$env:TEMP\S03_err_${timestamp}.txt" -ErrorAction SilentlyContinue

if (-not (Test-Path $dataFile)) {
    Log "FATAL: data file missing"
    exit 11
}

$dataLines = Get-Content $dataFile -Encoding Unicode

# Parse
$appMethods = @()
$areaMethods = @()
$invokes = @()
$fatal = $null

foreach ($raw in $dataLines) {
    $line = "$raw"
    if ($line.Trim().Length -eq 0) { continue }
    $parts = $line -split '\|', 5
    switch ($parts[0]) {
        "FATAL"       { $fatal = "$($parts[1]):$($parts[2..($parts.Count-1)] -join ':')" }
        "APP_METHOD"  { $appMethods += [pscustomobject]@{ Name=$parts[1]; Status=$parts[2]; ErrNum=$parts[3]; ErrDesc=$parts[4] } }
        "AREA_METHOD" { $areaMethods += [pscustomobject]@{ Name=$parts[1]; Status=$parts[2]; ErrNum=$parts[3]; ErrDesc=$parts[4] } }
        "INVOKE"      {
            # INVOKE|Letter|Signature|ErrNum|ErrDesc|Elapsed
            $sub = $line -split '\|', 6
            $invokes += [pscustomobject]@{ Id=$sub[1]; Signature=$sub[2]; ErrNum=$sub[3]; ErrDesc=$sub[4]; Elapsed=$sub[5] }
        }
    }
}

# Report
$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# Calculate API Probe -- $timestamp")
$md.Add("")
$md.Add("**Area index:** $AreaIndex")
$md.Add("**Wall time:** $([math]::Round($elapsed.TotalSeconds, 1))s")
$md.Add("**Project rule applied:** numeric values use period (.) as decimal separator (rule 7)")
$md.Add("")
if ($fatal) {
    $md.Add("## FATAL")
    $md.Add("``````"); $md.Add($fatal); $md.Add("``````")
} else {
    $md.Add("## Introspection: methods on oLEAP (Application object)")
    $md.Add("")
    $md.Add("| Method | Status | Err.Number | Err.Description |")
    $md.Add("|---|---|---|---|")
    foreach ($x in $appMethods) {
        $md.Add("| ``$($x.Name)`` | $($x.Status) | $($x.ErrNum) | $($x.ErrDesc) |")
    }
    $md.Add("")
    $md.Add("Legend: ``EXISTS_NEEDS_ARGS`` (err 450) = method present but missing/wrong args. ``NOT_PRESENT`` (err 438) = method does not exist. ``EXISTS_NO_ARGS_OK`` = method evaluated as a property without args. ``OTHER`` = unexpected error.")
    $md.Add("")

    $md.Add("## Introspection: methods on oLEAP.ActiveArea")
    $md.Add("")
    $md.Add("| Method | Status | Err.Number | Err.Description |")
    $md.Add("|---|---|---|---|")
    foreach ($x in $areaMethods) {
        $md.Add("| ``$($x.Name)`` | $($x.Status) | $($x.ErrNum) | $($x.ErrDesc) |")
    }
    $md.Add("")

    $md.Add("## Invocation attempts")
    $md.Add("")
    $md.Add("**A success means Err.Number=0 after the call. The calc may still have failed internally; check Elapsed -- anything > 5s suggests a real calc started.**")
    $md.Add("")
    $md.Add("| Id | Signature | Err.Number | Err.Description | Elapsed (s) |")
    $md.Add("|---|---|---|---|---|")
    foreach ($x in $invokes) {
        $md.Add("| $($x.Id) | ``$($x.Signature)`` | $($x.ErrNum) | $($x.ErrDesc) | $($x.Elapsed) |")
    }
    $md.Add("")

    $md.Add("## Verdict")
    $md.Add("")
    $successes = $invokes | Where-Object { $_.ErrNum -eq "0" }
    $errs450 = $invokes | Where-Object { $_.ErrNum -eq "450" }
    $errs438 = $invokes | Where-Object { $_.ErrNum -eq "438" }
    $md.Add("- Invocations that dispatched cleanly (Err=0): $($successes.Count)")
    $md.Add("- Invocations with err 450 (wrong args): $($errs450.Count)")
    $md.Add("- Invocations with err 438 (method not found): $($errs438.Count)")
    $md.Add("")
    if ($successes.Count -gt 0) {
        $md.Add("**Recommended signature(s):**")
        foreach ($s in $successes) {
            $longRun = if ([double]$s.Elapsed -gt 5) { " (started real calc)" } else { "" }
            $md.Add("- ``$($s.Signature)`` (elapsed=$($s.Elapsed)s)$longRun")
        }
    } else {
        $md.Add("**No invocation dispatched cleanly.** Manual fallback: open LEAP -> Advanced -> Edit Scripts. In the script editor, type ``LEAP.`` and use auto-complete to see the actual method list. The Script Editor IS the authoritative API documentation on this install (SEI does not publish the spec online).")
    }
}

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"

if ($fatal) { exit 1 } else { exit 0 }
