# F02_blank_2010_expressions.ps1
# Cycle 010: bulk-blank all expressions in branches that reference pre-BaseYear
# data, replacing them with "0" so calc validation passes.
#
# Approach:
#   - Open KAZ_2024 (index 6) via COM
#   - For each branch in F02_branch_list.txt:
#       Leap.Branches(path).Variable(var).ExpressionRS(kazRegId, scenId) = "0"
#     For scenId in {CA, S0}
#   - Try every plausible save method (we don't know which works on this install)
#   - Log success/error per branch + per save attempt
#
# Project rules applied: ASCII VBS, UTF-16 data output, BranchVariable chained
# (the colleague's pattern, proven to work for writes in hook code at
# beforeCalculation.vbs:79-81).
#
# Reversible: cycle_009_KAZ_2024_current_state.leap is the rollback snapshot.

param(
    [string]$AreaName = "KAZ_2024",
    [string]$BranchList = "scripts\03_fix\F02_branch_list.txt"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\F02_cycle010_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\F02_cycle010_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\F02_cycle010_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== F02: Bulk-blank 2010-era expressions ==="
Log "Target area name: $AreaName"

$branchListPath = Join-Path $projectRoot $BranchList
if (-not (Test-Path $branchListPath)) {
    Log "FATAL: branch list not found at $branchListPath"
    exit 1
}
$branches = Get-Content $branchListPath | Where-Object { $_.Trim().Length -gt 0 }
Log "Branches to process: $($branches.Count)"

# Build a VBS array literal from the branch list
$vbsBranchArrayLines = @()
$vbsBranchArrayLines += "Dim brn(" + ($branches.Count - 1) + ")"
for ($i = 0; $i -lt $branches.Count; $i++) {
    $line = $branches[$i].Trim()
    if ($line.Contains('"')) {
        Log "WARN: skipping line with double-quote: $line"
        continue
    }
    $vbsBranchArrayLines += 'brn(' + $i + ') = "' + $line + '"'
}
$vbsBranchArray = ($vbsBranchArrayLines -join "`r`n")

$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile
Dim a, target_area, current_idx
Dim s, r, kazRegId, caScenId, s0ScenId
Dim i, parts, branchPath, varName
Dim B, V
Dim n_ok_ca, n_err_ca, n_ok_s0, n_err_s0
Dim err_num, err_desc
Dim save_result_a, save_result_b, save_result_c, save_result_d

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
    outFile.Close : WScript.Quit 1
End If
Err.Clear

W "PHASE", "com_created"
WScript.Sleep 4000

' Resolve area by NAME (after cycle 010 bug: index ordering shifts between sessions)
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
    outFile.Close : oLEAP.Quit : WScript.Quit 2
End If

W "RESOLVED_AREA_NAME", target_area.Name

target_area.Open
If Err.Number <> 0 Then
    outFile.WriteLine "FATAL|AREA_OPEN|" & Err.Number & ":" & Err.Description
    outFile.Close : oLEAP.Quit : WScript.Quit 3
End If
Err.Clear

WScript.Sleep 3000
W "PHASE", "area_opened"
W "OPENED_BASE_YEAR", oLEAP.BaseYear
oLEAP.Verbose = 0

' Find Kazakhstan region id and scenario ids
kazRegId = -1
For Each r In oLEAP.Regions
    If r.Name = "Kazakhstan" Then kazRegId = r.Id
Next
W "KAZ_REGION_ID", kazRegId

caScenId = -1
s0ScenId = -1
For Each s In oLEAP.Scenarios
    If s.Name = "Current Accounts" Then caScenId = s.Id
    If s.Name = "S0 Baseline Historical" Then s0ScenId = s.Id
Next
W "CA_SCEN_ID", caScenId
W "S0_SCEN_ID", s0ScenId

If kazRegId = -1 Or caScenId = -1 Or s0ScenId = -1 Then
    outFile.WriteLine "FATAL|IDS_NOT_FOUND|kaz=" & kazRegId & ",ca=" & caScenId & ",s0=" & s0ScenId
    outFile.Close : oLEAP.Quit : WScript.Quit 4
End If

' --- The branch array ---
{{BRANCH_ARRAY}}

W "BRANCH_COUNT", UBound(brn) + 1
W "PHASE", "processing_branches"

n_ok_ca = 0 : n_err_ca = 0 : n_ok_s0 = 0 : n_err_s0 = 0

For i = 0 To UBound(brn)
    parts = Split(brn(i), "|")
    If UBound(parts) < 1 Then
        outFile.WriteLine "BRANCH_PARSE_ERR|" & brn(i)
    Else
        branchPath = parts(0)
        varName    = parts(1)

        ' Check branch + variable existence first
        If Not oLEAP.Branches.Exists(branchPath) Then
            outFile.WriteLine "BRANCH_MISSING|" & branchPath
        Else
            Set B = oLEAP.Branches(branchPath)
            If Not B.VariableExists(varName) Then
                outFile.WriteLine "VAR_MISSING|" & branchPath & "|" & varName
            Else
                ' Write CA expression
                Err.Clear
                B.Variable(varName).ExpressionRS(kazRegId, caScenId) = "0"
                If Err.Number = 0 Then
                    n_ok_ca = n_ok_ca + 1
                Else
                    err_num = Err.Number : err_desc = Err.Description : Err.Clear
                    outFile.WriteLine "WRITE_ERR_CA|" & branchPath & "|" & varName & "|" & err_num & "|" & err_desc
                    n_err_ca = n_err_ca + 1
                End If

                ' Write S0 expression
                Err.Clear
                B.Variable(varName).ExpressionRS(kazRegId, s0ScenId) = "0"
                If Err.Number = 0 Then
                    n_ok_s0 = n_ok_s0 + 1
                Else
                    err_num = Err.Number : err_desc = Err.Description : Err.Clear
                    outFile.WriteLine "WRITE_ERR_S0|" & branchPath & "|" & varName & "|" & err_num & "|" & err_desc
                    n_err_s0 = n_err_s0 + 1
                End If
            End If
        End If
    End If
Next

W "OK_CA", n_ok_ca
W "ERR_CA", n_err_ca
W "OK_S0", n_ok_s0
W "ERR_S0", n_err_s0

' --- Save attempts ---
W "PHASE", "save_attempts"

' Attempt A: oLEAP.SaveArea
Err.Clear
oLEAP.SaveArea
save_result_a = Err.Number & "|" & Err.Description
Err.Clear
outFile.WriteLine "SAVE_A|oLEAP.SaveArea|" & save_result_a

' Attempt B: oLEAP.ActiveArea.Save
Err.Clear
oLEAP.ActiveArea.Save
save_result_b = Err.Number & "|" & Err.Description
Err.Clear
outFile.WriteLine "SAVE_B|oLEAP.ActiveArea.Save|" & save_result_b

' Attempt C: oLEAP.Save
Err.Clear
oLEAP.Save
save_result_c = Err.Number & "|" & Err.Description
Err.Clear
outFile.WriteLine "SAVE_C|oLEAP.Save|" & save_result_c

' Attempt D: oLEAP.ActiveArea.Close True (close-with-save -- nuclear option, do last)
Err.Clear
oLEAP.ActiveArea.Close True
save_result_d = Err.Number & "|" & Err.Description
Err.Clear
outFile.WriteLine "SAVE_D|oLEAP.ActiveArea.Close_True|" & save_result_d

W "PHASE", "closing"
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_NAME}}', $AreaName).Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\')).Replace('{{BRANCH_ARRAY}}', $vbsBranchArray)

# ASCII check
$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: VBS source contains non-ASCII"; exit 99
}

$vbsPath = Join-Path $env:TEMP "F02_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath ($((Get-Item $vbsPath).Length) bytes)"
Log ""
Log "Running. Expected 30-90 seconds for 164 expression writes + save attempts."

$startTime = Get-Date
$proc = Start-Process -FilePath "cscript" -ArgumentList "//NoLogo", "`"$vbsPath`"" -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\F02_out_${timestamp}.txt" -RedirectStandardError "$env:TEMP\F02_err_${timestamp}.txt"
$timedOut = $false
if (-not $proc.WaitForExit(30 * 60 * 1000)) {
    Log "TIMEOUT after 30 min"
    try { $proc.Kill() } catch {}
    Get-Process -Name "LEAP" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    $timedOut = $true
}
$elapsed = (Get-Date) - $startTime
Log "cscript ended in $([math]::Round($elapsed.TotalSeconds, 1))s (timeout=$timedOut)"

$cscriptStdout = Get-Content "$env:TEMP\F02_out_${timestamp}.txt" -ErrorAction SilentlyContinue
$cscriptStderr = Get-Content "$env:TEMP\F02_err_${timestamp}.txt" -ErrorAction SilentlyContinue
Remove-Item "$env:TEMP\F02_out_${timestamp}.txt", "$env:TEMP\F02_err_${timestamp}.txt", $vbsPath -ErrorAction SilentlyContinue
foreach ($l in $cscriptStdout) { if ($l -and $l.Trim().Length -gt 0) { Log "stdout: $l" } }
foreach ($l in $cscriptStderr) { if ($l -and $l.Trim().Length -gt 0) { Log "stderr: $l" } }

if (-not (Test-Path $dataFile)) {
    Log "FATAL: data file missing"
    [System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
    exit 11
}

# Parse
$dataLines = Get-Content $dataFile -Encoding Unicode
Log "Data lines: $($dataLines.Count)"

$meta = @{}
$writeErrCa = @()
$writeErrS0 = @()
$branchMissing = @()
$varMissing = @()
$saveAttempts = @()
$fatal = $null

foreach ($raw in $dataLines) {
    $line = "$raw"
    if ($line.Trim().Length -eq 0) { continue }
    $parts = $line -split '\|'
    switch ($parts[0]) {
        "PHASE"            { $meta["LastPhase"] = $parts[1] }
        "FATAL"            { $fatal = "$($parts[1]):$($parts[2..($parts.Count-1)] -join ':')" }
        "RESOLVED_AREA_NAME" { $meta["ResolvedArea"] = $parts[1] }
        "OPENED_BASE_YEAR" { $meta["BaseYear"] = $parts[1] }
        "KAZ_REGION_ID"    { $meta["KazRegId"] = $parts[1] }
        "CA_SCEN_ID"       { $meta["CaScenId"] = $parts[1] }
        "S0_SCEN_ID"       { $meta["S0ScenId"] = $parts[1] }
        "BRANCH_COUNT"     { $meta["BranchCount"] = $parts[1] }
        "OK_CA"            { $meta["OkCa"] = $parts[1] }
        "ERR_CA"           { $meta["ErrCa"] = $parts[1] }
        "OK_S0"            { $meta["OkS0"] = $parts[1] }
        "ERR_S0"           { $meta["ErrS0"] = $parts[1] }
        "WRITE_ERR_CA"     { $writeErrCa += [pscustomobject]@{ Path=$parts[1]; Var=$parts[2]; ErrNum=$parts[3]; ErrDesc=$parts[4] } }
        "WRITE_ERR_S0"     { $writeErrS0 += [pscustomobject]@{ Path=$parts[1]; Var=$parts[2]; ErrNum=$parts[3]; ErrDesc=$parts[4] } }
        "BRANCH_MISSING"   { $branchMissing += $parts[1] }
        "VAR_MISSING"      { $varMissing += [pscustomobject]@{ Path=$parts[1]; Var=$parts[2] } }
        "SAVE_A"           { $saveAttempts += [pscustomobject]@{ Method=$parts[1]; Sig="oLEAP.SaveArea"; ErrNum=$parts[2]; ErrDesc=$parts[3] } }
        "SAVE_B"           { $saveAttempts += [pscustomobject]@{ Method=$parts[1]; Sig="oLEAP.ActiveArea.Save"; ErrNum=$parts[2]; ErrDesc=$parts[3] } }
        "SAVE_C"           { $saveAttempts += [pscustomobject]@{ Method=$parts[1]; Sig="oLEAP.Save"; ErrNum=$parts[2]; ErrDesc=$parts[3] } }
        "SAVE_D"           { $saveAttempts += [pscustomobject]@{ Method=$parts[1]; Sig="oLEAP.ActiveArea.Close True"; ErrNum=$parts[2]; ErrDesc=$parts[3] } }
    }
}

# Report
$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# F02 cycle 010 -- bulk blank 2010-era expressions -- $timestamp")
$md.Add("")
$md.Add("**Branches in list:** $($meta['BranchCount'])")
$md.Add("**Resolved area:** ``$($meta['ResolvedArea'])`` (BaseYear=$($meta['BaseYear']))")
$md.Add("**Kazakhstan region id:** $($meta['KazRegId'])")
$md.Add("**CA scenario id:** $($meta['CaScenId']), **S0 scenario id:** $($meta['S0ScenId'])")
$md.Add("**Wall time:** $([math]::Round($elapsed.TotalSeconds, 1))s")
$md.Add("")
if ($fatal) {
    $md.Add("## FATAL")
    $md.Add("``````"); $md.Add($fatal); $md.Add("``````")
} else {
    $md.Add("## Write outcomes")
    $md.Add("")
    $md.Add("| Scope | OK | Errors |")
    $md.Add("|---|---|---|")
    $md.Add("| Current Accounts | $($meta['OkCa']) | $($meta['ErrCa']) |")
    $md.Add("| S0 Baseline Historical | $($meta['OkS0']) | $($meta['ErrS0']) |")
    $md.Add("| Branches missing | $($branchMissing.Count) | -- |")
    $md.Add("| Var missing | $($varMissing.Count) | -- |")
    $md.Add("")

    $md.Add("## Save attempts")
    $md.Add("")
    $md.Add("| Method | Err.Number | Err.Description |")
    $md.Add("|---|---|---|")
    foreach ($s in $saveAttempts) {
        $md.Add("| ``$($s.Sig)`` | $($s.ErrNum) | $($s.ErrDesc) |")
    }
    $md.Add("")
    $saveOk = @($saveAttempts | Where-Object { $_.ErrNum -eq "0" })
    if ($saveOk.Count -gt 0) {
        $md.Add("**At least one save method dispatched cleanly:** $($saveOk[0].Sig)")
    } else {
        $md.Add("**No save method returned Err=0.** Changes may not have persisted.")
    }
    $md.Add("")

    if ($writeErrCa.Count -gt 0) {
        $md.Add("## CA write errors (first 20)")
        $md.Add("")
        $md.Add("| Path | Var | Err | Desc |")
        $md.Add("|---|---|---|---|")
        foreach ($e in $writeErrCa | Select-Object -First 20) {
            $md.Add("| ``$($e.Path)`` | $($e.Var) | $($e.ErrNum) | $($e.ErrDesc) |")
        }
        $md.Add("")
    }
    if ($branchMissing.Count -gt 0) {
        $md.Add("## Missing branches (first 20)")
        $md.Add("")
        foreach ($b in $branchMissing | Select-Object -First 20) {
            $md.Add("- ``$b``")
        }
        $md.Add("")
    }
}

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"
if ($fatal -or $timedOut) { exit 1 } else { exit 0 }
