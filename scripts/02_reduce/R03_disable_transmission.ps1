# R03_disable_transmission.ps1
# Disables NEMO transmission simulation on the Electricity Production module.
# This is the v0.1 prototype workaround for ISSUE-001 (nodal distribution sums to 0%).
#
# Usage:
#   .\R03_disable_transmission.ps1 -AreaName "KAZ_2024_v01"
#
# What it does:
#   1. Opens the area in LEAP via COM
#   2. For each region, sets Simulation Type on Electricity Production module to "Standard"
#      (was: "NetworkSimulation(...)" which requires nodal distributions)
#   3. Also sets it for Current Accounts column (scenario_id = 0 / no scenario)
#   4. Saves and reports
#
# Reversible — to re-enable, set Simulation Type back to its original NetworkSimulation expression.
# Get original by extracting .leap first (S01_extract_leap.py) and reading from before image.

param(
    [Parameter(Mandatory=$true)]
    [string]$AreaName,

    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"

# Resolve paths relative to project root
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$logPath = Join-Path $projectRoot $LogDir "R03_disable_transmission_$timestamp.log"
New-Item -ItemType Directory -Force -Path (Split-Path $logPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) { 
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line 
}

Log "=== R03: Disable transmission simulation ==="
Log "Area: $AreaName"
Log ""

# Build the VBScript inline. CRITICAL: write as ASCII no BOM (ISSUE-003).
$vbs = @"
Option Explicit

Dim oLEAP, r, n_changed, original_value

On Error Resume Next

Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    WScript.Echo "FAIL_COM:" & Err.Number & ":" & Err.Description
    WScript.Quit 1
End If
Err.Clear

WScript.Sleep 4000

oLEAP.Areas("$AreaName").Open
If Err.Number <> 0 Then
    WScript.Echo "FAIL_OPEN:" & Err.Number & ":" & Err.Description
    WScript.Quit 2
End If
Err.Clear

WScript.Sleep 2000
WScript.Echo "OPENED:" & oLEAP.ActiveArea.Name

' Confirm the target branch exists
If Not oLEAP.Branches.Exists("Transformation\Electricity Production") Then
    WScript.Echo "FAIL_NOBRANCH"
    oLEAP.Quit
    WScript.Quit 3
End If

' Confirm the Simulation Type variable exists on it
If Not oLEAP.Branches("Transformation\Electricity Production").VariableExists("Simulation Type") Then
    WScript.Echo "FAIL_NOVAR"
    oLEAP.Quit
    WScript.Quit 4
End If

' Disable UI updates during bulk write
oLEAP.Verbose = 0

n_changed = 0

' Get the variable
Dim var
Set var = oLEAP.Branches("Transformation\Electricity Production").Variable("Simulation Type")

' Loop over regions
For Each r In oLEAP.Regions
    WScript.Echo "REGION:" & r.Name & ":id=" & r.Id

    ' Current Accounts (scenario id = 0 = "no scenario" / inheritance root)
    original_value = var.ExpressionRS(r.Id, 0)
    WScript.Echo "  CA_before:" & original_value
    If InStr(original_value, "NetworkSimulation") > 0 Or InStr(original_value, "Network") > 0 Then
        var.ExpressionRS(r.Id, 0) = "Standard"
        n_changed = n_changed + 1
        WScript.Echo "  CA_after:Standard"
    End If

    ' All scenarios — loop and only modify if currently NetworkSimulation
    Dim s
    For Each s In oLEAP.Scenarios
        original_value = var.ExpressionRS(r.Id, s.Id)
        If InStr(original_value, "NetworkSimulation") > 0 Or InStr(original_value, "Network") > 0 Then
            var.ExpressionRS(r.Id, s.Id) = "Standard"
            n_changed = n_changed + 1
            WScript.Echo "  S_" & s.Abbreviation & "_changed"
        End If
    Next
Next

oLEAP.Verbose = 4
oLEAP.Refresh
oLEAP.Save
WScript.Echo "DONE:changes=" & n_changed
oLEAP.Quit
WScript.Quit 0
"@

# Write the VBS file as ASCII no BOM (CRITICAL — ISSUE-003)
$vbsPath = Join-Path $env:TEMP "R03_disable_transmission_$timestamp.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath"

# Execute
Log ""
Log "Running cscript..."
$out = & cscript //NoLogo $vbsPath 2>&1
$out | ForEach-Object { Log "  | $_" }

# Cleanup temp
Remove-Item $vbsPath -ErrorAction SilentlyContinue

# Parse result
$success = ($out | Where-Object { $_ -match "^DONE:changes=" }) -ne $null
$failed  = ($out | Where-Object { $_ -match "^FAIL_" }) -ne $null

Log ""
if ($success) {
    $changeLine = $out | Where-Object { $_ -match "^DONE:changes=" } | Select-Object -First 1
    Log "=== R03 COMPLETE — $changeLine ==="
    Log "Next: open area in LEAP UI, press F9, check whether nodal distribution error is gone."
    $exitCode = 0
} elseif ($failed) {
    $failLine = $out | Where-Object { $_ -match "^FAIL_" } | Select-Object -First 1
    Log "=== R03 FAILED — $failLine ==="
    Log "See log file for details."
    $exitCode = 1
} else {
    Log "=== R03 INCONCLUSIVE — see log ==="
    $exitCode = 2
}

# Persist log
[System.IO.File]::WriteAllText($logPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Write-Host ""
Write-Host "Log: $logPath"

exit $exitCode
