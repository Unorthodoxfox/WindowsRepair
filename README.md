# WindowsRepair
This PowerShell script performs common Windows repair and maintenance tasks by running SFC and DISM commands in sequence, monitoring DISM log and CPU activity for possible hangs, saving detailed logs, and checking for available Windows Updates. It is intended to be run as Administrator and does not install updates automatically.
# Windows Repair and Update Check Script

## Overview

`Windows-Repair-And-Update-Check.ps1` is a PowerShell maintenance script that runs common Windows system repair checks and then checks for available Windows Updates.

The script is intended to be run from an elevated PowerShell window as Administrator.

It performs the following steps:

1. Runs `sfc /scannow`
2. Runs `DISM /Online /Cleanup-Image /ScanHealth`
3. Runs `DISM /Online /Cleanup-Image /RestoreHealth`
4. Runs `DISM /Online /Cleanup-Image /StartComponentCleanup`
5. Runs `sfc /scannow` again
6. Checks for available Windows Updates

The script checks for Windows Updates but does not install them.

---

## Requirements

- Windows 10 or Windows 11
- Administrator privileges
- PowerShell
- Internet access is recommended for Windows Update checks and DISM repair operations

---

## Script File Name

Recommended file name:

```powershell
Windows-Repair-And-Update-Check.ps1
```

PowerShell scripts should be saved with the `.ps1` file extension.

---

## How to Run

Open PowerShell as Administrator, browse to the folder containing the script, and run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
.\Windows-Repair-And-Update-Check.ps1
```

The execution policy change only applies to the current PowerShell window. It does not permanently change the computer's execution policy.

To close the PowerShell window automatically after the script completes, run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force; .\Windows-Repair-And-Update-Check.ps1; exit
```

---

## Running from a Network Folder

If the script is stored on a network share, use the full UNC path when possible.

Example:

```powershell
cd "\\server\share\Repair Scripts"
.\Windows-Repair-And-Update-Check.ps1
```

If the folder name contains spaces, wrap the path in quotes:

```powershell
cd "\\server\share\Folder With Spaces"
```

If mapped drives do not appear in an Administrator PowerShell window, use the UNC path instead of the mapped drive letter.

---

## What the Script Does

### SFC

The script runs:

```powershell
sfc /scannow
```

System File Checker scans protected Windows system files and attempts to repair missing or corrupted files.

The script runs SFC twice:

- Once before DISM
- Once after DISM

The second SFC run can sometimes repair files after DISM has repaired the Windows component store.

---

### DISM ScanHealth

The script runs:

```powershell
DISM /Online /Cleanup-Image /ScanHealth
```

This checks whether the Windows component store has corruption.

---

### DISM RestoreHealth

The script runs:

```powershell
DISM /Online /Cleanup-Image /RestoreHealth
```

This attempts to repair corruption in the Windows component store.

---

### DISM StartComponentCleanup

The script runs:

```powershell
DISM /Online /Cleanup-Image /StartComponentCleanup
```

This cleans up superseded Windows component store files.

---

### Windows Update Check

The script checks for available Windows Updates using Windows Update COM objects.

It lists available updates but does not download or install them.

---

## DISM Hang Monitoring

The script includes DISM monitoring logic.

Each DISM command is run with its own log file using `/LogPath`.

While DISM is running, the script checks for:

- Recent DISM log activity
- DISM process CPU activity

If no activity is detected for the configured timeout period, the script writes a warning and displays recent DISM log entries.

By default, the script does not stop DISM automatically.

This is intentional because DISM can appear stuck for a long time while still working.

---

## KillOnHang Option

The script supports a `-KillOnHang` switch in the `Run-DismWithMonitor` function.

Example:

```powershell
Run-DismWithMonitor `
    -Description "Step 3: Running DISM RestoreHealth" `
    -DismArguments @("/Online", "/Cleanup-Image", "/RestoreHealth") `
    -NoActivityTimeoutMinutes 60 `
    -KillOnHang
```

For normal background or unattended use, it is recommended not to use `-KillOnHang`.

Force-stopping DISM could interrupt Windows servicing or repair operations.

Recommended behavior:

- Warn after inactivity
- Continue monitoring
- Review logs afterward if a warning appears

---

## Log Files

The script creates logs in:

```text
C:\WindowsRepairLogs
```

The main transcript log is named similar to:

```text
WindowsRepair_YYYY-MM-DD_HH-mm-ss.log
```

Each DISM step also creates its own log file in the same folder.

Example:

```text
Step_3__Running_DISM_RestoreHealth.log
```

---

## Finding DISM Errors

If the script says DISM exited with a non-zero exit code, check the relevant DISM step log.

List newest logs first:

```powershell
Get-ChildItem "C:\WindowsRepairLogs" -Filter "*.log" |
    Sort-Object LastWriteTime -Descending |
    Select-Object LastWriteTime, Name, FullName
```

Search all logs for errors:

```powershell
Select-String -Path "C:\WindowsRepairLogs\*.log" `
    -Pattern "Error|HRESULT|0x[0-9A-Fa-f]{8}|failed|CBS" `
    -Context 3,3
```

View the last 80 lines of a DISM log:

```powershell
Get-Content "C:\WindowsRepairLogs\Step_3__Running_DISM_RestoreHealth.log" -Tail 80
```

Common DISM exit codes include:

| Exit Code | Meaning |
|---:|---|
| 0 | Success |
| 2 | File not found |
| 50 | Operation not supported |
| 87 | Bad parameter or syntax issue |
| 3010 | Success, reboot required |

---

## Notes

- Run the script as Administrator.
- The script may take a long time to complete.
- DISM may remain at the same percentage for a while and still be working.
- A reboot may be required after repairs or Windows servicing operations.
- The script checks for Windows Updates but does not install them.
- Review `C:\WindowsRepairLogs` after the script finishes.

---

## Troubleshooting

### PowerShell says the script cannot be loaded

Run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
```

Then run the script again.

---

### Network drive is missing in Administrator PowerShell

Use the UNC path instead of the mapped drive letter.

Example:

```powershell
cd "\\server\share\folder"
```

---

### Folder path has spaces

Put the path in quotes.

Example:

```powershell
cd "C:\Folder With Spaces"
```

---

### DISM appears stuck

Check whether the DISM log is still being updated.

The script monitors this automatically and writes a warning if no activity is detected for too long.

Do not force-close DISM unless you are sure it is hung and you are comfortable restarting the repair process afterward.

---

## Recommended Usage

For normal background maintenance, run the script without `-KillOnHang`.

Recommended DISM timeout behavior:

| DISM Step | Suggested No-Activity Warning |
|---|---:|
| ScanHealth | 30 minutes |
| RestoreHealth | 60 minutes |
| StartComponentCleanup | 60 minutes |

This allows the script to continue running while still logging warnings if DISM appears inactive.
