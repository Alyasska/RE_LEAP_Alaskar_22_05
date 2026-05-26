# S04_find_broken_emission_units.ps1
# Walk the Demand subtree, find every branch with the "Avg Environmental Loading"
# variable, and try to read its unit metadata. Branches where the unit is empty
# or unreadable are candidates for ISSUE-002 (broken unit metadata).
#
# READ-ONLY.
#
# Usage:
#   .\scripts\01_scout\S04_find_broken_emission_units.ps1 -AreaName "KAZ_2024"

param(
    [string]$AreaName = "KAZ_2024"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\S04_find_emissions_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\S04_find_emissions_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\S04_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== S04: Find broken emission units ==="
Log "Area: $AreaName"

$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile
Dim a, target_area, found_names, B, V
Dim has_var, expr, expr_err
Dim probe_unit, probe_num, probe_den, probe_type
Dim n_total_demand, n_with_var, n_err_read

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
W "PHASE", "scanning_demand_branches"
oLEAP.Verbose = 0

n_total_demand = 0
n_with_var = 0
n_err_read = 0

For Each B In oLEAP.Branches
    If Left(B.FullName, 7) = "Demand\" Then
        n_total_demand = n_total_demand + 1
        ' Does it have Avg Environmental Loading?
        Err.Clear
        has_var = B.VariableExists("Avg Environmental Loading")
        If Err.Number <> 0 Then Err.Clear : has_var = False

        If has_var Then
            n_with_var = n_with_var + 1
            Set V = B.Variable("Avg Environmental Loading")
            If Err.Number = 0 And Not (V Is Nothing) Then
                ' Try to read the expression
                Err.Clear
                expr = ""
                expr = V.Expression
                expr_err = Err.Number
                Err.Clear

                ' Try to read various unit-related properties
                Err.Clear : probe_unit = "" : probe_unit = V.Unit
                If Err.Number <> 0 Then probe_unit = "<err:" & Err.Number & ">" : Err.Clear

                Err.Clear : probe_num = "" : probe_num = V.UnitNumerator
                If Err.Number <> 0 Then probe_num = "<err:" & Err.Number & ">" : Err.Clear

                Err.Clear : probe_den = "" : probe_den = V.UnitDenominator
                If Err.Number <> 0 Then probe_den = "<err:" & Err.Number & ">" : Err.Clear

                Err.Clear : probe_type = "" : probe_type = V.TypeName
                If Err.Number <> 0 Then probe_type = "<err:" & Err.Number & ">" : Err.Clear

                If expr_err <> 0 Then n_err_read = n_err_read + 1

                outFile.WriteLine "BRANCH|" & B.FullName & "|expr_err=" & expr_err & "|expr=" & expr & "|unit=" & probe_unit & "|num=" & probe_num & "|den=" & probe_den & "|type=" & probe_type
            Else
                outFile.WriteLine "BRANCH|" & B.FullName & "|VAR_NOTHING"
                Err.Clear
            End If
        End If
    End If
Next

W "TOTAL_DEMAND", n_total_demand
W "WITH_VAR", n_with_var
W "READ_ERRORS", n_err_read

W "PHASE", "closing"
W "DONE", "ok"
outFile.Close
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{AREA_NAME}}', $AreaName).Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))

$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: non-ASCII in VBS"; exit 99
}

$vbsPath = Join-Path $env:TEMP "S04_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS: $vbsPath"
Log ""
Log "Scanning. Expected 60-180s for full Demand tree walk."

$startTime = Get-Date
$proc = Start-Process -FilePath "cscript" -ArgumentList "//NoLogo", "`"$vbsPath`"" -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\S04_o_${timestamp}.txt" -RedirectStandardError "$env:TEMP\S04_e_${timestamp}.txt"
$timedOut = $false
if (-not $proc.WaitForExit(30 * 60 * 1000)) {
    Log "TIMEOUT"
    try { $proc.Kill() } catch {}
    Get-Process -Name "LEAP" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    $timedOut = $true
}
$elapsed = (Get-Date) - $startTime
Log "Done in $([math]::Round($elapsed.TotalSeconds, 1))s"
Remove-Item $vbsPath, "$env:TEMP\S04_o_${timestamp}.txt", "$env:TEMP\S04_e_${timestamp}.txt" -ErrorAction SilentlyContinue

if (-not (Test-Path $dataFile)) { Log "FATAL: no data file"; exit 11 }

$lines = Get-Content $dataFile -Encoding Unicode
$branches = @()
$meta = @{}
$fatal = $null

foreach ($raw in $lines) {
    $line = "$raw"
    if ($line.Trim().Length -eq 0) { continue }
    $parts = $line -split '\|'
    switch ($parts[0]) {
        "FATAL"        { $fatal = $line }
        "TOTAL_DEMAND" { $meta["TotalDemand"] = $parts[1] }
        "WITH_VAR"     { $meta["WithVar"] = $parts[1] }
        "READ_ERRORS"  { $meta["ReadErrors"] = $parts[1] }
        "BRANCH"       {
            $obj = [pscustomobject]@{
                FullName = $parts[1]
                ExprErr  = ($parts[2] -replace '^expr_err=','')
                Expr     = ($parts[3] -replace '^expr=','')
                Unit     = ($parts[4] -replace '^unit=','')
                Num      = ($parts[5] -replace '^num=','')
                Den      = ($parts[6] -replace '^den=','')
                Type     = ($parts[7] -replace '^type=','')
            }
            $branches += $obj
        }
    }
}

# Identify suspect branches: empty unit OR error reading unit/expression
$suspect = $branches | Where-Object {
    ($_.Unit -eq "" -or $_.Unit -like "<err*") -or
    ($_.ExprErr -ne "0") -or
    ($_.Num -eq "" -and $_.Den -ne "")
}

$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# S04 Find broken emission units -- $timestamp")
$md.Add("")
$md.Add("**Area:** $AreaName")
$md.Add("**Total Demand branches walked:** $($meta['TotalDemand'])")
$md.Add("**Branches with 'Avg Environmental Loading':** $($meta['WithVar'])")
$md.Add("**Branches where expression read failed:** $($meta['ReadErrors'])")
$md.Add("**Suspect branches (empty unit or err):** $($suspect.Count)")
$md.Add("")

if ($fatal) {
    $md.Add("## FATAL")
    $md.Add("``$fatal``")
} elseif ($suspect.Count -eq 0) {
    $md.Add("**No suspects identified by Unit/Expression read.** This means either:")
    $md.Add("- The broken-unit metadata is not exposed through Unit/UnitNumerator/UnitDenominator COM properties")
    $md.Add("- OR the broken state is invisible to read calls (only fires at calc time)")
    $md.Add("")
    $md.Add("Sample of first 10 branches with Avg Env Loading (for inspection):")
    $md.Add("")
    $md.Add("| Path | Unit | Num | Den | Expr |")
    $md.Add("|---|---|---|---|---|")
    foreach ($b in $branches | Select-Object -First 10) {
        $md.Add("| ``$($b.FullName)`` | $($b.Unit) | $($b.Num) | $($b.Den) | $($b.Expr) |")
    }
} else {
    $md.Add("## Suspect branches")
    $md.Add("")
    $md.Add("| FullName | ExprErr | Unit | Num | Den | Expr |")
    $md.Add("|---|---|---|---|---|---|")
    foreach ($b in $suspect | Select-Object -First 50) {
        $md.Add("| ``$($b.FullName)`` | $($b.ExprErr) | ``$($b.Unit)`` | ``$($b.Num)`` | ``$($b.Den)`` | ``$($b.Expr)`` |")
    }
    if ($suspect.Count -gt 50) {
        $md.Add("")
        $md.Add("_($($suspect.Count - 50) more not shown)_")
    }
}

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log "Report: $reportPath"
if ($fatal) { exit 1 } else { exit 0 }
