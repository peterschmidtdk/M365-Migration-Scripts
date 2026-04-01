#Requires -Version 5.1
<#
.SYNOPSIS
    Tracks disk usage of folders (and one level down) under a specified root, with historical trending.

.DESCRIPTION
    Scans each immediate subfolder (L1) and their subfolders (L2) under the given root path.
    Appends results to a CSV history file and displays a delta comparison vs the previous scan.

.PARAMETER RootPath
    The top-level folder to scan (e.g. D:\Shares).

.PARAMETER HistoryFile
    Path to the CSV file used for storing scan history.
    Defaults to 'DiskUsageHistory.csv' in the same folder as the script.

.PARAMETER MinSizeMB
    Only include folders larger than this threshold in MB. Default is 0 (all folders).

.EXAMPLE
    .\Get-FolderDiskUsage.ps1 -RootPath "D:\Shares"

.EXAMPLE
    .\Get-FolderDiskUsage.ps1 -RootPath "D:\Shares" -HistoryFile "C:\Tracking\shares_history.csv"
#>

[CmdletBinding()]
param (
    [Parameter(Position = 0, Mandatory)]
    [ValidateScript({ Test-Path $_ -PathType Container })]
    [string]$RootPath,

    [Parameter()]
    [string]$HistoryFile = (Join-Path $PSScriptRoot "DiskUsageHistory.csv"),

    [Parameter()]
    [double]$MinSizeMB = 0
)

# ── Helpers ──────────────────────────────────────────────────────────────────

function Get-FolderSize {
    param ([string]$Path)
    $bytes = 0; $fileCount = 0
    try {
        $items     = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
        $files     = $items | Where-Object { -not $_.PSIsContainer }
        $measured  = ($files | Measure-Object -Property Length -Sum).Sum
        $bytes     = if ($null -ne $measured) { $measured } else { 0 }
        $fileCount = ($files | Measure-Object).Count
    } catch {
        Write-Warning "Could not fully scan: $Path — $_"
    }
    return [PSCustomObject]@{ Bytes = [long]$bytes; FileCount = $fileCount }
}

function Format-Size {
    param ([long]$Bytes)
    switch ($Bytes) {
        { $_ -ge 1TB } { return "{0:N2} TB" -f ($_ / 1TB) }
        { $_ -ge 1GB } { return "{0:N2} GB" -f ($_ / 1GB) }
        { $_ -ge 1MB } { return "{0:N2} MB" -f ($_ / 1MB) }
        { $_ -ge 1KB } { return "{0:N2} KB" -f ($_ / 1KB) }
        default        { return "$_ B" }
    }
}

function Format-Delta {
    param ([long]$Bytes)
    $sign = if ($Bytes -ge 0) { "+" } else { "" }
    return "$sign$(Format-Size $Bytes)"
}

function Get-DeltaColor {
    param ([long]$Bytes)
    if ($Bytes -gt 0) { return "Red" }
    if ($Bytes -lt 0) { return "Green" }
    return "Gray"
}

# ── Scan ─────────────────────────────────────────────────────────────────────

$scanTime  = Get-Date
$scanStamp = $scanTime.ToString("yyyy-MM-dd HH:mm:ss")

$l1Folders = Get-ChildItem -Path $RootPath -Directory -Force -ErrorAction Stop

if (-not $l1Folders) {
    Write-Warning "No subfolders found in '$RootPath'."
    exit
}

Write-Host "`nScanning '$RootPath' at $scanStamp ...`n" -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$total   = $l1Folders.Count
$i       = 0

foreach ($l1 in $l1Folders) {
    $i++
    Write-Progress -Activity "Scanning" -Status "[$i/$total] $($l1.Name)" `
                   -PercentComplete (($i / $total) * 100)

    # L1
    $l1Size = Get-FolderSize -Path $l1.FullName

    if (($l1Size.Bytes / 1MB) -ge $MinSizeMB) {
        $results.Add([PSCustomObject]@{
            ScanTime  = $scanStamp
            Level     = 1
            L1Folder  = $l1.Name
            L2Folder  = ""
            FullPath  = $l1.FullName
            SizeBytes = $l1Size.Bytes
            SizeMB    = [math]::Round($l1Size.Bytes / 1MB, 4)
        })
    }

    # L2
    $l2Folders = Get-ChildItem -Path $l1.FullName -Directory -Force -ErrorAction SilentlyContinue
    foreach ($l2 in $l2Folders) {
        $l2Size = Get-FolderSize -Path $l2.FullName

        if (($l2Size.Bytes / 1MB) -ge $MinSizeMB) {
            $results.Add([PSCustomObject]@{
                ScanTime  = $scanStamp
                Level     = 2
                L1Folder  = $l1.Name
                L2Folder  = $l2.Name
                FullPath  = $l2.FullName
                SizeBytes = $l2Size.Bytes
                SizeMB    = [math]::Round($l2Size.Bytes / 1MB, 4)
            })
        }
    }
}

Write-Progress -Activity "Scanning" -Completed

# ── Load history & build previous-scan lookup ─────────────────────────────────

$prevScanMap = @{}   # FullPath → SizeBytes from the last scan

if (Test-Path $HistoryFile) {
    $history = Import-Csv -Path $HistoryFile

    $prevStamp = $history |
        Where-Object { $_.ScanTime -ne $scanStamp } |
        Sort-Object ScanTime -Descending |
        Select-Object -ExpandProperty ScanTime -First 1

    if ($prevStamp) {
        $history |
            Where-Object { $_.ScanTime -eq $prevStamp } |
            ForEach-Object { $prevScanMap[$_.FullPath] = [long]$_.SizeBytes }

        Write-Host "Comparing against previous scan: $prevStamp`n" -ForegroundColor DarkGray
    }
}

# ── Append current results to CSV ─────────────────────────────────────────────

$results | Export-Csv -Path $HistoryFile -NoTypeInformation -Encoding UTF8 -Append

if ($prevScanMap.Count -eq 0) {
    Write-Host "History file: $HistoryFile  (first scan — no delta yet)`n" -ForegroundColor DarkGray
} else {
    Write-Host "History appended to: $HistoryFile`n" -ForegroundColor DarkGray
}

# ── Console report ────────────────────────────────────────────────────────────

$l1Results = $results | Where-Object { $_.Level -eq 1 } | Sort-Object SizeBytes -Descending
$colW      = 34
$sep       = "─" * 74

Write-Host $sep -ForegroundColor DarkGray
Write-Host (" {0,-$colW} {1,12}  {2,16}" -f "Folder", "Size", "Delta vs last") -ForegroundColor Yellow
Write-Host $sep -ForegroundColor DarkGray

$prevL1Total = 0

foreach ($l1 in $l1Results) {

    # Delta for L1
    $delta      = if ($prevScanMap.ContainsKey($l1.FullPath)) {
                      $prevL1Total += $prevScanMap[$l1.FullPath]
                      $l1.SizeBytes - $prevScanMap[$l1.FullPath]
                  } else { $null }
    $deltaStr   = if ($null -eq $delta) { "(new)" } else { Format-Delta $delta }
    $deltaColor = if ($null -eq $delta) { "Cyan"  } else { Get-DeltaColor $delta }

    Write-Host (" {0,-$colW} {1,12}" -f $l1.L1Folder, (Format-Size $l1.SizeBytes)) -NoNewline
    Write-Host ("  {0,16}" -f $deltaStr) -ForegroundColor $deltaColor

    # L2 rows indented under L1
    $l2Results = $results |
        Where-Object { $_.Level -eq 2 -and $_.L1Folder -eq $l1.L1Folder } |
        Sort-Object SizeBytes -Descending

    foreach ($l2 in $l2Results) {
        $l2Delta    = if ($prevScanMap.ContainsKey($l2.FullPath)) {
                          $l2.SizeBytes - $prevScanMap[$l2.FullPath]
                      } else { $null }
        $l2DeltaStr = if ($null -eq $l2Delta) { "(new)" } else { Format-Delta $l2Delta }
        $l2Color    = if ($null -eq $l2Delta) { "Cyan"  } else { Get-DeltaColor $l2Delta }
        $l2Label    = "  └─ $($l2.L2Folder)"

        Write-Host ("  {0,-$($colW - 2)} {1,12}" -f $l2Label, (Format-Size $l2.SizeBytes)) -NoNewline
        Write-Host ("  {0,16}" -f $l2DeltaStr) -ForegroundColor $l2Color
    }
}

# ── Totals row ────────────────────────────────────────────────────────────────

$measuredTotal = ($l1Results | Measure-Object -Property SizeBytes -Sum).Sum
$totalBytes    = if ($null -ne $measuredTotal) { $measuredTotal } else { 0 }
$totalDelta = if ($prevScanMap.Count -gt 0) { $totalBytes - $prevL1Total } else { $null }
$totalDeltaStr = if ($null -eq $totalDelta) { "n/a" } else { Format-Delta $totalDelta }
$totalColor    = if ($null -eq $totalDelta) { "Gray" } else { Get-DeltaColor $totalDelta }

Write-Host $sep -ForegroundColor DarkGray
Write-Host (" {0,-$colW} {1,12}" -f "TOTAL ($($l1Results.Count) folders)", (Format-Size $totalBytes)) `
           -ForegroundColor Cyan -NoNewline
Write-Host ("  {0,16}" -f $totalDeltaStr) -ForegroundColor $totalColor
Write-Host $sep -ForegroundColor DarkGray
Write-Host ""
