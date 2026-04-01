# Get-FolderDiskUsage

A PowerShell script for tracking folder disk usage over time. Designed for monitoring shared drives or any folder structure where you want to keep an eye on growth trends.

## Features

- Scans **two levels deep** — top-level folders (L1) and their subfolders (L2)
- **Tracks history** by appending each scan to a CSV file
- **Compares against the previous scan** and shows growth or reduction per folder
- Color-coded delta output — red for growth, green for reduction, cyan for new folders
- Optional **minimum size filter** to cut out noise
- Compatible with **PowerShell 5.1+** (no PS7 required)

## Requirements

- PowerShell 5.1 or higher
- Read access to the folder structure being scanned

## Usage

```powershell
.\Get-FolderDiskUsage.ps1 -RootPath "D:\Shares"
```

```powershell
# Custom history file location
.\Get-FolderDiskUsage.ps1 -RootPath "D:\Shares" -HistoryFile "C:\Tracking\shares_history.csv"
```

```powershell
# Only show folders larger than 500 MB
.\Get-FolderDiskUsage.ps1 -RootPath "D:\Shares" -MinSizeMB 500
```

## Parameters

| Parameter | Required | Default | Description |
|---|---|---|---|
| `-RootPath` | Yes | — | The top-level folder to scan (e.g. `D:\Shares`) |
| `-HistoryFile` | No | `DiskUsageHistory.csv` next to the script | Path to the CSV file used for storing scan history |
| `-MinSizeMB` | No | `0` (all folders) | Skip folders smaller than this size in MB |

## Example Output

On the **first run**, no delta is shown — this becomes the baseline:

```
Scanning 'D:\Shares' at 2026-04-01 08:00:00 ...
History file: D:\DiskUsageHistory.csv  (first scan — no delta yet)

──────────────────────────────────────────────────────────────────────────
 Folder                              Size      Delta vs last
──────────────────────────────────────────────────────────────────────────
 Finance                           45.20 GB             n/a
   └─ Invoices                     30.10 GB             n/a
   └─ Reports                      15.10 GB             n/a
 HR                                12.00 GB             n/a
   └─ Contracts                     8.50 GB             n/a
   └─ Archives                      3.50 GB             n/a
──────────────────────────────────────────────────────────────────────────
 TOTAL (2 folders)                 57.20 GB             n/a
──────────────────────────────────────────────────────────────────────────
```

On **subsequent runs**, each folder is compared against the previous scan:

```
Scanning 'D:\Shares' at 2026-04-08 08:00:00 ...
Comparing against previous scan: 2026-04-01 08:00:00

──────────────────────────────────────────────────────────────────────────
 Folder                              Size      Delta vs last
──────────────────────────────────────────────────────────────────────────
 Finance                           46.50 GB         +1.30 GB
   └─ Invoices                     32.20 GB         +2.10 GB
   └─ Reports                      14.30 GB          -0.80 GB
 HR                                12.00 GB            +0 B
   └─ Contracts                     8.50 GB            +0 B
   └─ Archives                      3.50 GB            +0 B
   └─ NewStarters                   1.20 GB           (new)
──────────────────────────────────────────────────────────────────────────
 TOTAL (2 folders)                 58.50 GB         +1.30 GB
──────────────────────────────────────────────────────────────────────────
```

Delta colors in the terminal:
- 🔴 **Red** — folder has grown since last scan
- 🟢 **Green** — folder has shrunk
- ⚪ **Gray** — no change
- 🔵 **Cyan** — folder is new (not present in previous scan)

## History File (CSV)

Each scan appends rows to the CSV history file with the following columns:

| Column | Description |
|---|---|
| `ScanTime` | Timestamp of the scan (`yyyy-MM-dd HH:mm:ss`) |
| `Level` | `1` for top-level folders, `2` for subfolders |
| `L1Folder` | Name of the top-level folder |
| `L2Folder` | Name of the subfolder (empty for L1 rows) |
| `FullPath` | Full path of the folder |
| `SizeBytes` | Total size in bytes |
| `SizeMB` | Total size in MB (4 decimal places) |

The CSV is well-suited for opening in Excel or Power BI to chart growth over time across multiple scans.

## Scheduling (Task Scheduler)

To run the script automatically on a weekly basis:

1. Open **Task Scheduler** and create a new Basic Task
2. Set the trigger to **Weekly** at your preferred time
3. Set the action to **Start a program**:
   - Program: `powershell.exe`
   - Arguments: `-NonInteractive -File "D:\Get-FolderDiskUsage.ps1" -RootPath "D:\Shares"`

## Author

**Peter Schmidt** — v0.1.0
