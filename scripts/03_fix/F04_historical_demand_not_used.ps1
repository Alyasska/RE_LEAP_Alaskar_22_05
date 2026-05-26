# F04_historical_demand_not_used.ps1
# For every Demand branch that has Avg Environmental Loading (= the branches
# LEAP visits during emission factor evaluation), set Historical_Demand to
# "not used". This is the colleague's known-working pattern from
# fix_lubricants_all.vbs / fix_calc_errors_v3.vbs.
#
# Why not "0":
#   F02 set Historical_Demand = "0" on these branches. LEAP still tried to
#   compute emissions = Historical_Demand * Avg Environmental Loading, and
#   hit broken unit metadata during validation. "not used" is a LEAP-special
#   marker that tells calc to skip the branch entirely.

param(
    [string]$AreaName = "KAZ_2024"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\F04_notused_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\F04_notused_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\F04_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== F04: Historical_Demand = not used on branches with Avg Env Loading ==="
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
    If a.Name = "{{AREA_NAME}}" Then Set target_area = a : Exit For
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
W "PHASE", "opened"
oLEAP.Verbose = 0

kazRegId = -1
For Each r In oLEAP.Regions
    If r.Name = "Kazakhstan" Then kazRegId = r.Id
Next
caScenId = -1 : s0ScenId = -1
For Each s In oLEAP.Scenarios
    If s.Name = "Current Accounts" Then caScenId = s.Id
    If s.Name = "S0 Baseline Historical" Then s0ScenId = s.Id
Next
W "KAZ_ID", kazRegId
W "CA_ID", caScenId
W "S0_ID", s0ScenId

W "PHASE", "writing"

n_total = 0 : n_with_var = 0
n_ok_ca = 0 : n_err_ca = 0 : n_ok_s0 = 0 : n_err_s0 = 0

For Each B In oLEAP.Branches
    If Left(B.FullName, 7) = "Demand\" Then
        n_total = n_total + 1
        Err.Clear
        If B.VariableExists("Avg Environmental Loading") And B.VariableExists("Historical_Demand") Then
            n_with_var = n_with_var + 1

            Err.Clear
            B.Variable("Historical_Demand").ExpressionRS(kazRegId, caScenId) = "not used"
            If Err.Number = 0 Then
                n_ok_ca = n_ok_ca + 1
            Else
                err_num = Err.Number : err_desc = Err.Description : Err.Clear
                outFile.WriteLine "ERR_CA|" & B.FullName & "|" & err_num & "|" & err_desc
                n_err_ca = n_err_ca + 1
            End If

            Err.Clear
            B.Variable("Historical_Demand").ExpressionRS(kazRegId, s0ScenId) = "not used"
            If Err.Number = 0 Then
                n_ok_s0 = n_ok_s0 + 1
            Else
                err_num = Err.Number : err_desc = Err.Description : Err.Clear
                outFile.WriteLine "ERR_S0|" & B.FullName & "|" & err_num & "|" & err_desc
                n_err_s0 = n_err_s0 + 1
            End If
        End If
    End If
Next

W "TOTAL", n_total
W "WITH_BOTH_VARS", n_with_var
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
End If

W "PHASE", "closing"
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_NAME}}', $AreaName).Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))
$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) { Log "ABORT non-ASCII"; exit 99 }

$vbsPath = Join-Path $env:TEMP "F04_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS: $vbsPath"

$startTime = Get-Date
$proc = Start-Process -FilePath "cscript" -ArgumentList "//NoLogo","`"$vbsPath`"" -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\F04_o_${timestamp}.txt" -RedirectStandardError "$env:TEMP\F04_e_${timestamp}.txt"
if (-not $proc.WaitForExit(30 * 60 * 1000)) {
    Log "TIMEOUT"; try { $proc.Kill() } catch {}
    Get-Process -Name "LEAP" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
$elapsed = (Get-Date) - $startTime
Log "Done in $([math]::Round($elapsed.TotalSeconds, 1))s"
Remove-Item $vbsPath, "$env:TEMP\F04_o_${timestamp}.txt", "$env:TEMP\F04_e_${timestamp}.txt" -ErrorAction SilentlyContinue

if (-not (Test-Path $dataFile)) { Log "FATAL no data"; exit 11 }

$lines = Get-Content $dataFile -Encoding Unicode
$meta = @{}
foreach ($raw in $lines) {
    $line = "$raw".Trim()
    if ($line.Length -eq 0) { continue }
    $parts = $line -split '\|'
    switch ($parts[0]) {
        "TOTAL"          { $meta["Total"] = $parts[1] }
        "WITH_BOTH_VARS" { $meta["WithBoth"] = $parts[1] }
        "OK_CA"          { $meta["OkCa"] = $parts[1] }
        "ERR_CA"         { $meta["ErrCa"] = $parts[1] }
        "OK_S0"          { $meta["OkS0"] = $parts[1] }
        "ERR_S0"         { $meta["ErrS0"] = $parts[1] }
        "SAVE"           { $meta["Save"] = $parts[1] }
    }
}

$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# F04 -- Historical_Demand = 'not used' for branches with Avg Env Loading -- $timestamp")
$md.Add("")
$md.Add("**Demand branches:** $($meta['Total'])")
$md.Add("**With both Avg Env Loading + Historical_Demand:** $($meta['WithBoth'])")
$md.Add("**CA OK / ERR:** $($meta['OkCa']) / $($meta['ErrCa'])")
$md.Add("**S0 OK / ERR:** $($meta['OkS0']) / $($meta['ErrS0'])")
$md.Add("**Save:** $($meta['Save'])")
[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"
