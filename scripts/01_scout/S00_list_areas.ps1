# S00_list_areas.ps1
# Probe oLEAP.Areas to enumerate every installed area with disambiguating info.
# READ-ONLY. Does not open any area.
#
# Why this exists (cycle 002 finding):
#   LEAP has two areas resolving under the same key "kaz_workshop exercise":
#     - Parent area, BaseYear=2010 (pre-migration original)
#     - Nested area "kaz_workshop exercise/KAZ_2024", BaseYear=2024 (colleague's)
#   oLEAP.Areas("kaz_workshop exercise").Open non-deterministically picked one
#   or the other across runs. We need to address them by index, not by string.
#
# Usage:
#   .\scripts\01_scout\S00_list_areas.ps1
#
# Output:
#   - data\audit_reports\area_listing_<timestamp>.md
#   - data\audit_reports\area_listing_<timestamp>.data.txt (UTF-16 raw)
#
# Runtime: 30-60 seconds (mostly LEAP cold start).

param(
    [string]$LogDir = "logs"
)

$ErrorActionPreference = "Stop"
$projectRoot = (Get-Item $PSScriptRoot).Parent.Parent.FullName
Set-Location $projectRoot

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$dataFile   = Join-Path $projectRoot "data\audit_reports\area_listing_${timestamp}.data.txt"
$reportPath = Join-Path $projectRoot "data\audit_reports\area_listing_${timestamp}.md"
$runLogPath = Join-Path $projectRoot "logs\S00_list_${timestamp}.log"

New-Item -ItemType Directory -Force -Path (Split-Path $dataFile -Parent) | Out-Null
New-Item -ItemType Directory -Force -Path (Split-Path $runLogPath -Parent) | Out-Null

$log = [System.Collections.Generic.List[string]]::new()
function Log($msg) {
    $line = "$(Get-Date -Format 'HH:mm:ss') $msg"
    $log.Add($line)
    Write-Host $line
}

Log "=== S00: List installed areas (probe) ==="
Log "Data file: $dataFile"

# ----------------------------------------------------------------------------
# VBS: iterate oLEAP.Areas, dump every probable property per entry.
# We do NOT open any area here. Just enumerate.
# ----------------------------------------------------------------------------
$vbsTemplate = @'
Option Explicit

Dim oLEAP, fso, outFile, a, idx, val, propName
Dim propsToTry

On Error Resume Next

Set fso = CreateObject("Scripting.FileSystemObject")
Set outFile = fso.OpenTextFile("{{DATA_FILE}}", 2, True, -1)
If Err.Number <> 0 Then
    WScript.Echo "FATAL_FILE:" & Err.Number & ":" & Err.Description
    WScript.Quit 10
End If
Err.Clear

outFile.WriteLine "PHASE|starting"

Set oLEAP = CreateObject("LEAP.LEAPApplication")
If Err.Number <> 0 Then
    outFile.WriteLine "FATAL|COM_CREATE|" & Err.Number & ":" & Err.Description
    outFile.Close
    WScript.Quit 1
End If
Err.Clear

outFile.WriteLine "PHASE|com_created"
WScript.Sleep 4000

outFile.WriteLine "PHASE|enumerating"

' For each area, try to read a bunch of possible properties.
' We do not know exactly which the LEAP COM exposes, so we try each
' with Err handling and log success/failure separately.
idx = 0
For Each a In oLEAP.Areas
    idx = idx + 1
    outFile.WriteLine "AREA_BEGIN|" & idx

    ' Always-available property
    outFile.WriteLine "AREA_PROP|" & idx & "|Name|" & a.Name

    ' Try a list of properties; log each whether it works or not
    Err.Clear
    val = a.Description
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Description|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Description|" & Err.Number
        Err.Clear
    End If

    val = a.BaseYear
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|BaseYear|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|BaseYear|" & Err.Number
        Err.Clear
    End If

    val = a.FirstScenarioYear
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|FirstScenarioYear|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|FirstScenarioYear|" & Err.Number
        Err.Clear
    End If

    val = a.EndYear
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|EndYear|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|EndYear|" & Err.Number
        Err.Clear
    End If

    val = a.LastModified
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|LastModified|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|LastModified|" & Err.Number
        Err.Clear
    End If

    val = a.Created
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Created|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Created|" & Err.Number
        Err.Clear
    End If

    val = a.Folder
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Folder|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Folder|" & Err.Number
        Err.Clear
    End If

    val = a.FullPath
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|FullPath|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|FullPath|" & Err.Number
        Err.Clear
    End If

    val = a.Directory
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Directory|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Directory|" & Err.Number
        Err.Clear
    End If

    val = a.Author
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Author|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Author|" & Err.Number
        Err.Clear
    End If

    val = a.Comments
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Comments|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Comments|" & Err.Number
        Err.Clear
    End If

    ' Some LEAP versions: Id, Key, ShortName
    val = a.Id
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Id|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Id|" & Err.Number
        Err.Clear
    End If

    val = a.Key
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|Key|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|Key|" & Err.Number
        Err.Clear
    End If

    val = a.ShortName
    If Err.Number = 0 Then
        outFile.WriteLine "AREA_PROP|" & idx & "|ShortName|" & val
    Else
        outFile.WriteLine "AREA_PROP_MISS|" & idx & "|ShortName|" & Err.Number
        Err.Clear
    End If

    outFile.WriteLine "AREA_END|" & idx
Next

outFile.WriteLine "TOTAL|" & idx
outFile.WriteLine "PHASE|done"
outFile.WriteLine "DONE|ok"
outFile.Close

' Try to quit cleanly. If save dialog appears, user must dismiss manually.
oLEAP.Quit
WScript.Quit 0
'@

$vbs = $vbsTemplate.Replace('{{DATA_FILE}}', $dataFile.Replace('\','\\'))

# Verify pure ASCII
$bytes = [System.Text.Encoding]::UTF8.GetBytes($vbs)
if (($bytes | Where-Object { $_ -gt 127 }).Count -gt 0) {
    Log "ABORT: VBS source contains non-ASCII"
    exit 99
}

$vbsPath = Join-Path $env:TEMP "S00_list_${timestamp}.vbs"
[System.IO.File]::WriteAllText($vbsPath, $vbs, [System.Text.Encoding]::ASCII)
Log "VBS written: $vbsPath"
Log ""

$startTime = Get-Date
$cscriptOut = & cscript //NoLogo $vbsPath 2>&1
$elapsed = (Get-Date) - $startTime
Log "cscript finished in $([math]::Round($elapsed.TotalSeconds, 1))s"
Remove-Item $vbsPath -ErrorAction SilentlyContinue

foreach ($l in $cscriptOut) {
    if ($l -is [string] -and $l.Trim().Length -gt 0) { Log "cscript: $l" }
}

if (-not (Test-Path $dataFile)) {
    Log "FATAL: data file missing"
    exit 11
}

# ----------------------------------------------------------------------------
# Parse
# ----------------------------------------------------------------------------
$dataLines = Get-Content $dataFile -Encoding Unicode
Log "Read $($dataLines.Count) data lines"

$areas = @{}
$total = 0
$fatal = $null

foreach ($raw in $dataLines) {
    $line = "$raw".Trim()
    if ($line.Length -eq 0) { continue }
    $parts = $line -split '\|', 4
    switch ($parts[0]) {
        "FATAL"     { $fatal = "$($parts[1]):$($parts[2])" }
        "TOTAL"     { $total = [int]$parts[1] }
        "AREA_BEGIN" {
            $idx = $parts[1]
            if (-not $areas.ContainsKey($idx)) {
                $areas[$idx] = [ordered]@{}
            }
        }
        "AREA_PROP" {
            $idx = $parts[1]
            if (-not $areas.ContainsKey($idx)) { $areas[$idx] = [ordered]@{} }
            $propName = $parts[2]
            $propVal = if ($parts.Count -ge 4) { $parts[3] } else { "" }
            $areas[$idx][$propName] = $propVal
        }
        # AREA_PROP_MISS is silently ignored (we know the property does not exist on this version)
    }
}

# ----------------------------------------------------------------------------
# Markdown report
# ----------------------------------------------------------------------------
$md = [System.Collections.Generic.List[string]]::new()
$md.Add("# Area listing -- $timestamp")
$md.Add("")
$md.Add("**Generated:** $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$md.Add("**Total areas found:** $total")
$md.Add("**Runtime:** $([math]::Round($elapsed.TotalSeconds, 1))s")
$md.Add("")

if ($fatal) {
    $md.Add("## FATAL: $fatal")
} else {
    $md.Add("## Areas (in oLEAP.Areas iteration order)")
    $md.Add("")
    $md.Add("**Use the Index column when calling areas by position.**")
    $md.Add("")

    # Collect union of all property names seen
    $allProps = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($idx in $areas.Keys) {
        foreach ($p in $areas[$idx].Keys) { [void]$allProps.Add($p) }
    }
    # Preferred column order
    $preferred = @("Name", "BaseYear", "FirstScenarioYear", "EndYear", "Folder", "Directory", "FullPath", "LastModified", "Author")
    $cols = @()
    foreach ($p in $preferred) {
        if ($allProps.Contains($p)) { $cols += $p }
    }
    # Append any other props not in preferred list
    foreach ($p in $allProps) {
        if ($p -notin $cols) { $cols += $p }
    }

    $header = "| Index | " + ($cols -join " | ") + " |"
    $separator = "|---|" + (($cols | ForEach-Object { "---" }) -join "|") + "|"
    $md.Add($header)
    $md.Add($separator)

    foreach ($idx in ($areas.Keys | Sort-Object { [int]$_ })) {
        $row = "| $idx | "
        $cells = @()
        foreach ($c in $cols) {
            $v = if ($areas[$idx].Contains($c)) { $areas[$idx][$c] } else { "" }
            $v = "$v".Replace("|", "/").Replace("`r`n", " ").Replace("`n", " ")
            if ($v.Length -gt 80) { $v = $v.Substring(0, 77) + "..." }
            $cells += $v
        }
        $row += ($cells -join " | ") + " |"
        $md.Add($row)
    }
    $md.Add("")

    # Identify the prototype-target area
    $md.Add("## Identifying the prototype target")
    $md.Add("")
    $kazAreas = @()
    foreach ($idx in $areas.Keys) {
        $name = "$($areas[$idx]['Name'])".ToLower()
        $by = "$($areas[$idx]['BaseYear'])"
        if ($name.Contains("kaz") -or $name.Contains("workshop") -or $by -eq "2024") {
            $kazAreas += [pscustomobject]@{
                Index = $idx
                Name = $areas[$idx]['Name']
                BaseYear = $areas[$idx]['BaseYear']
            }
        }
    }

    $md.Add("Candidate areas matching KAZ / workshop / BaseYear=2024:")
    $md.Add("")
    foreach ($k in $kazAreas) {
        $marker = if ("$($k.BaseYear)" -eq "2024") { " <-- PROTOTYPE TARGET" } else { "" }
        $md.Add("- Index **$($k.Index)**: ``$($k.Name)`` (BaseYear=$($k.BaseYear))$marker")
    }
    $md.Add("")

    $target = $kazAreas | Where-Object { "$($_.BaseYear)" -eq "2024" } | Select-Object -First 1
    if ($target) {
        $md.Add("**Recommended next call (S02 v3):**")
        $md.Add('```powershell')
        $md.Add(".\scripts\01_scout\S02_audit_model_v3.ps1 -AreaIndex $($target.Index)")
        $md.Add('```')
        $md.Add("")
        $md.Add("This addresses the area by position (eliminates name-collision ambiguity).")
    }
}

[System.IO.File]::WriteAllText($reportPath, ($md -join "`r`n"), [System.Text.Encoding]::UTF8)
[System.IO.File]::WriteAllText($runLogPath, ($log -join "`r`n"), [System.Text.Encoding]::UTF8)
Log ""
Log "Report: $reportPath"

if ($fatal) { exit 1 } else { exit 0 }
