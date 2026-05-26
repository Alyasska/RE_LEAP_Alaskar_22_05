# F03_blank_avg_env_loading.ps1
# Walk Demand tree; for every branch with Avg Environmental Loading variable,
# set its expression to "0" for CA and S0 scenarios. Save the area.
#
# Why: ISSUE-002 ("Invalid numerator unit for emissions factor"). Colleague's
# docs/known_issues.md says this can't be fixed by expression alone (unit
# metadata is validated first). We try it anyway because (a) it's a 5-minute
# experiment, (b) we have a reliable rollback, and (c) the colleague's claim
# may not hold for ALL broken-unit branches -- only the 5 known ones.
#
# If this fails, next try is .Delete on the variable row.

param(
    [string]$AreaName = "KAZ_2024"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\F03_blank_emissions_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\F03_blank_emissions_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\F03_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== F03: Bulk-blank Avg Environmental Loading ==="
Log "Area: $AreaName"

$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile
Dim a, target_area, found_names, B
Dim r, s, kazRegId, caScenId, s0ScenId
Dim n_total, n_with_var, n_ok_ca, n_err_ca, n_ok_s0, n_err_s0
Dim err_num, err_desc

Sub W(key, val)
    outFile.WriteLine(key & "|" & val)
End Sub

On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.OpenTextFile("{{DATA_FILE}}", 2, True, -1)
If Err.Number <> 0 Then WScript.Echo "FATAL_FILE": WScript.Quit 10
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

Set target_area = Nothing
found_names = ""
For Each a In oLEAP.Areas
    found_names = found_names & "," & a.Name
    If a.Name = "{{AREA_NAME}}" Then
        Set target_area = a
        Exit For
    End If
Next

If target_area Is Nothing Then
    outFile.WriteLine "FATAL|NAME_NOT_FOUND|seen=" & found_names
    outFile.Close : oLEAP.Quit : WScript.Quit 2
End If

target_area.Open
If Err.Number <> 0 Then
    outFile.WriteLine "FATAL|AREA_OPEN|" & Err.Number & ":" & Err.Description
    outFile.Close : oLEAP.Quit : WScript.Quit 3
End If
Err.Clear

WScript.Sleep 3000
W "PHASE", "area_opened"
oLEAP.Verbose = 0

kazRegId = -1
For Each r In oLEAP.Regions
    If r.Name = "Kazakhstan" Then kazRegId = r.Id
Next

caScenId = -1
s0ScenId = -1
For Each s In oLEAP.Scenarios
    If s.Name = "Current Accounts" Then caScenId = s.Id
    If s.Name = "S0 Baseline Historical" Then s0ScenId = s.Id
Next

W "KAZ_ID", kazRegId
W "CA_ID", caScenId
W "S0_ID", s0ScenId

If kazRegId = -1 Or caScenId = -1 Or s0ScenId = -1 Then
    outFile.WriteLine "FATAL|IDS|kaz=" & kazRegId & ",ca=" & caScenId & ",s0=" & s0ScenId
    outFile.Close : oLEAP.Quit : WScript.Quit 4
End If

W "PHASE", "iterating_demand"

n_total = 0
n_with_var = 0
n_ok_ca = 0 : n_err_ca = 0 : n_ok_s0 = 0 : n_err_s0 = 0

For Each B In oLEAP.Branches
    If Left(B.FullName, 7) = "Demand\" Then
        n_total = n_total + 1
        Err.Clear
        If B.VariableExists("Avg Environmental Loading") Then
            n_with_var = n_with_var + 1

            ' CA write
            Err.Clear
            B.Variable("Avg Environmental Loading").ExpressionRS(kazRegId, caScenId) = "0"
            If Err.Number = 0 Then
                n_ok_ca = n_ok_ca + 1
            Else
                err_num = Err.Number : err_desc = Err.Description : Err.Clear
                outFile.WriteLine "ERR_CA|" & B.FullName & "|" & err_num & "|" & err_desc
                n_err_ca = n_err_ca + 1
            End If

            ' S0 write
            Err.Clear
            B.Variable("Avg Environmental Loading").ExpressionRS(kazRegId, s0ScenId) = "0"
            If Err.Number = 0 Then
                n_ok_s0 = n_ok_s0 + 1
            Else
                err_num = Err.Number : err_desc = Err.Description : Err.Clear
                outFile.WriteLine "ERR_S0|" & B.FullName & "|" & err_num & "|" & err_desc
                n_err_s0 = n_err_s0 + 1
            End If
        End If
        Err.Clear
    End If
Next

W "TOTAL_DEMAND", n_total
W "WITH_VAR", n_with_var
W "OK_CA", n_ok_ca
W "ERR_CA", n_err_ca
W "OK_S0", n_ok_s0
W "ERR_S0", n_err_s0

W "PHASE", "save"
Err.Clear
oLEAP.SaveArea
If Err.Number = 0 Then
    W "SAVE", "ok"
Else
    outFile.WriteLine "SAVE|err|" & Err.Number & ":" & Err.Description
    Err.Clear
End If

W "PHASE", "closing"
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_NAME}}', $AreaName).Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))

$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: non-ASCII"; exit 99
}

$vbsPath = Join-Path $env:TEMP "F03_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS: $vbsPath"

$startTime = Get-Date
$proc = Start-Process -FilePath "cscript" -ArgumentList "//NoLogo","`"$vbsPath`"" -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\F03_o_${timestamp}.txt" -RedirectStandardError "$env:TEMP\F03_e_${timestamp}.txt"
if (-not $proc.WaitForExit(30 * 60 * 1000)) {
    Log "TIMEOUT"; try { $proc.Kill() } catch {}
    Get-Process -Name "LEAP" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
$elapsed = (Get-Date) - $startTime
Log "Done in $([math]::Round($elapsed.TotalSeconds, 1))s"
Remove-Item $vbsPath, "$env:TEMP\F03_o_${timestamp}.txt", "$env:TEMP\F03_e_${timestamp}.txt" -ErrorAction SilentlyContinue

if (-not (Test-Path $dataFile)) { Log "FATAL: no data file"; exit 11 }

$lines = Get-Content $dataFile -Encoding Unicode
$meta = @{}
$errs = @()
foreach ($raw in $lines) {
    $line = "$raw".Trim()
    if ($line.Length -eq 0) { continue }
    $parts = $line -split '\|'
    switch ($parts[0]) {
        "FATAL"      { $meta["Fatal"] = $line }
        "TOTAL_DEMAND" { $meta["Total"] = $parts[1] }
        "WITH_VAR"   { $meta["WithVar"] = $parts[1] }
        "OK_CA"      { $meta["OkCa"] = $parts[1] }
        "ERR_CA"     { $meta["ErrCa"] = $parts[1] }
        "OK_S0"      { $meta["OkS0"] = $parts[1] }
        "ERR_S0"     { $meta["ErrS0"] = $parts[1] }
        "SAVE"       { $meta["Save"] = $parts[1] }
        "ERR_CA"     { $errs += $line }
        "ERR_S0"     { $errs += $line }
    }
}

$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# F03 cycle 010 -- bulk-blank Avg Environmental Loading -- $timestamp")
$md.Add("")
$md.Add("**Area:** $AreaName")
$md.Add("**Demand branches walked:** $($meta['Total'])")
$md.Add("**Branches with Avg Env Loading:** $($meta['WithVar'])")
$md.Add("**CA writes OK / ERR:** $($meta['OkCa']) / $($meta['ErrCa'])")
$md.Add("**S0 writes OK / ERR:** $($meta['OkS0']) / $($meta['ErrS0'])")
$md.Add("**Save:** $($meta['Save'])")
$md.Add("")
if ($meta["Fatal"]) {
    $md.Add("## FATAL")
    $md.Add("``$($meta['Fatal'])``")
}
[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"
