# init_env.ps1 — Bootstrap RE_LEAP_Alaskar_22_05 environment
# Run ONCE after dropping the starter pack into your project folder.
#
# Usage:
#   cd C:\path\to\RE_LEAP_Alaskar_22_05
#   .\init_env.ps1
#
# Idempotent: safe to re-run.

Set-Location $PSScriptRoot

Write-Host "=== RE_LEAP environment setup ===" -ForegroundColor Cyan
Write-Host ""

# 1. Verify directory tree
$dirs = @(
    "docs",
    "docs\ui_procedures",
    "docs\screenshots",
    "scripts\0_setup",
    "scripts\01_scout",
    "scripts\02_reduce",
    "scripts\03_fix",
    "scripts\04_run",
    "scripts\core",
    "data\snapshots",
    "data\audit_reports",
    "data\extracted",
    "data\ground_truth",
    "logs"
)
foreach ($d in $dirs) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Force -Path $d | Out-Null
        Write-Host "  + created $d" -ForegroundColor Green
    } else {
        Write-Host "  . exists  $d"
    }
}
Write-Host ""

# 2. .gitignore (only write if missing)
if (-not (Test-Path ".gitignore")) {
    $gi = @'
# LEAP binaries — too big for git, snapshot manually if needed
*.leap
*.zip

# Decrypted .leap contents — regenerable
data/extracted/*
!data/extracted/.gitkeep

# Snapshots — track names in logs but not files
data/snapshots/*
!data/snapshots/.gitkeep

# Logs and temp
*.tmp
*.log

# Python / Node noise (in case)
__pycache__/
.venv/
node_modules/

# VS Code workspace settings (keep only shared)
.vscode/settings.json
'@
    Set-Content -Path ".gitignore" -Value $gi -Encoding UTF8
    Write-Host "  + .gitignore created" -ForegroundColor Green
} else {
    Write-Host "  . .gitignore exists"
}

# 3. Placeholder .gitkeep files so empty dirs are committed
foreach ($d in "data\snapshots", "data\extracted", "data\audit_reports", "logs") {
    $kp = Join-Path $d ".gitkeep"
    if (-not (Test-Path $kp)) {
        New-Item -ItemType File -Force -Path $kp | Out-Null
    }
}
Write-Host ""

# 4. Quick LEAP COM connection test
Write-Host "=== Testing LEAP COM connection ===" -ForegroundColor Cyan
$vbsTest = @'
On Error Resume Next
Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    WScript.Echo "FAIL: " & Err.Number & " - " & Err.Description
    WScript.Quit 1
End If
WScript.Echo "OK: LEAP COM responsive"
WScript.Echo "Version=" & oLEAP.Version
WScript.Echo "WorkingDirectory=" & oLEAP.WorkingDirectory
oLEAP.Quit
'@
$tmpVbs = Join-Path $env:TEMP "re_leap_conn_test.vbs"
[System.IO.File]::WriteAllText($tmpVbs, $vbsTest, [System.Text.Encoding]::ASCII)

$out = & cscript //NoLogo $tmpVbs 2>&1
$out | ForEach-Object { Write-Host "  $_" }
Remove-Item $tmpVbs -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Cyan
Write-Host "Next steps:"
Write-Host "  1. Copy colleague's .leap file into data\snapshots\cycle_000_colleague_baseline.leap"
Write-Host "  2. git init && git add . && git commit -m 'cycle 000: initial scaffolding'"
Write-Host "  3. Read HANDOFF.md and PROTOTYPE_PLAN.md"
Write-Host "  4. Tomorrow: scripts\01_scout\ -- scout the model"
