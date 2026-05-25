# F01_neuter_before_calc_hook.ps1
# Day 3, fix B: disable the beforeCalculation.vbs hook in KAZ_2024 by renaming.
#
# Why a filesystem rename (not a COM call):
#   LEAP reads beforeCalculation.vbs from the area folder at calc time.
#   Renaming the file disables the hook on the next calc without any COM
#   interaction. Reversible in 1 command.
#
# What this addresses (ISSUE-001a):
#   The hook tries to populate nodal distribution variables for
#   KAZ_North/West/South. When inputs are wrong (e.g. BaseYear=2024 changed
#   demand totals), it writes 0s, then NEMO calc fails with
#   "nodal distribution shares sum to 0%."
#   Neutralizing the hook lets us see whether calc has any OTHER blockers
#   downstream, independent of the nodal logic.
#
# Usage:
#   .\scripts\03_fix\F01_neuter_before_calc_hook.ps1 -AreaIndex 6
#   .\scripts\03_fix\F01_neuter_before_calc_hook.ps1 -AreaIndex 6 -Restore
#
# Reversible via -Restore flag.
#
# Safe: backs up the original by rename, never deletes. Idempotent on re-run.

param(
    [Parameter(Mandatory=$true)]
    [int]$AreaIndex,
    [switch]$Restore,
    [string]$CycleTag = "cycle005"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runLogPath = Join-Path $projectRoot "logs\F01_${CycleTag}_${timestamp}.log"
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== F01: Neuter beforeCalculation hook ==="
Log "Area index: $AreaIndex"
Log "Mode: $(if ($Restore) { 'RESTORE' } else { 'DISABLE' })"

# ----------------------------------------------------------------------------
# Resolve area directory via COM. Need oLEAP.Areas iterated by index.
# This is a short-lived COM session, just for the directory lookup.
# ----------------------------------------------------------------------------
$dirProbeData = Join-Path $env:TEMP "F01_dir_probe_${timestamp}.txt"

$probeVbs = @"
Option Explicit
Dim oLEAP, fso, outFile, a, current_idx, target_area
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.OpenTextFile("$($dirProbeData.Replace('\','\\'))", 2, True, -1)
Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    outFile.WriteLine "FATAL|COM_CREATE|" & Err.Number & ":" & Err.Description
    outFile.Close
    WScript.Quit 1
End If
WScript.Sleep 4000
current_idx = 0
Set target_area = Nothing
For Each a In oLEAP.Areas
    current_idx = current_idx + 1
    If current_idx = $AreaIndex Then
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
outFile.WriteLine "AREA_NAME|" & target_area.Name
outFile.WriteLine "AREA_DIR|" & target_area.Directory
outFile.Close
oLEAP.Quit
WScript.Quit 0
"@

# ASCII check
$bytes = [System.Text.Encoding]::UTF8.GetBytes($probeVbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: VBS source contains non-ASCII"; exit 99
}

$vbsPath = Join-Path $env:TEMP "F01_probe_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $probeVbs, [System.Text.Encoding]::ASCII)

Log ""
Log "Probing area directory via COM (~5-10 seconds)..."
$startTime = Get-Date
$probeOut = & cscript //NoLogo $vbsPath 2>&1
$elapsed = (Get-Date) - $startTime
Log "Probe done in $([math]::Round($elapsed.TotalSeconds, 1))s"
Remove-Item $vbsPath -ErrorAction SilentlyContinue
foreach ($l in $probeOut) { if ($l -is [string] -and $l.Trim().Length -gt 0) { Log "cscript: $l" } }

if (-not (Test-Path $dirProbeData)) {
    Log "FATAL: probe data file missing"
    exit 11
}

$areaName = $null
$areaDir = $null
$fatal = $null
foreach ($raw in (Get-Content $dirProbeData -Encoding Unicode)) {
    $line = "$raw".Trim()
    if ($line.Length -eq 0) { continue }
    $parts = $line -split '\|', 4
    switch ($parts[0]) {
        "FATAL"     { $fatal = "$($parts[1]):$($parts[2])" }
        "AREA_NAME" { $areaName = $parts[1] }
        "AREA_DIR"  { $areaDir = $parts[1] }
    }
}
Remove-Item $dirProbeData -ErrorAction SilentlyContinue

if ($fatal) {
    Log "FATAL: $fatal"
    exit 1
}

Log "Resolved area: $areaName"
Log "Area directory: $areaDir"

# ----------------------------------------------------------------------------
# Locate the hook file(s) and act
# ----------------------------------------------------------------------------
$candidates = @(
    "$areaDir\beforeCalculation.vbs"
    "$areaDir\beforeCalculation.vbs_Safe"
    "$areaDir\Scripts\beforeCalculation.vbs"
    "$areaDir\Scripts\beforeCalculation.vbs_Safe"
)

$disabledSuffix = ".disabled_${CycleTag}"

if ($Restore) {
    # --- RESTORE MODE: rename disabled-* back to original ---
    Log ""
    Log "Looking for disabled hook files to restore..."
    $restored = 0
    foreach ($c in $candidates) {
        $disabledPath = "$c$disabledSuffix"
        if (Test-Path $disabledPath) {
            if (Test-Path $c) {
                Log "  WARN: both $c and $disabledPath exist. Skipping (cannot determine which is authoritative)."
            } else {
                Move-Item -Path $disabledPath -Destination $c
                Log "  RESTORED: $disabledPath -> $c"
                $restored++
            }
        }
    }
    Log ""
    if ($restored -eq 0) {
        Log "No disabled hook files found to restore. (Already restored, or never disabled by this script.)"
    } else {
        Log "$restored file(s) restored."
    }
} else {
    # --- DISABLE MODE: rename original to disabled-* ---
    Log ""
    Log "Looking for hook files to disable..."
    $disabled = 0
    foreach ($c in $candidates) {
        if (Test-Path $c) {
            $size = (Get-Item $c).Length
            $disabledPath = "$c$disabledSuffix"
            if (Test-Path $disabledPath) {
                Log "  SKIP: $disabledPath already exists (disabled in a previous run). Original: $c ($size bytes). Manually resolve before re-running."
                continue
            }
            Move-Item -Path $c -Destination $disabledPath
            Log "  DISABLED: $c ($size bytes) -> $disabledPath"
            $disabled++
        }
    }
    Log ""
    if ($disabled -eq 0) {
        Log "No hook files found to disable. Either already disabled or never existed."
        Log "Looked in:"
        foreach ($c in $candidates) { Log "  - $c" }
    } else {
        Log "$disabled file(s) disabled. The next Calculate run in LEAP will not execute the hook."
    }
}

[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log ""
Log "Log: $runLogPath"
Log "=== F01 complete ==="

exit 0
