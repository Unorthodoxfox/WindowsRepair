#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Windows repair, component store cleanup, and Windows Update check script.

.DESCRIPTION
    Runs DISM StartComponentCleanup, DISM ScanHealth, conditional DISM RestoreHealth,
    SFC /scannow, and a Windows Update check. It writes a transcript log, individual
    DISM logs, and a CSV summary.

.NOTES
    Run from an elevated Windows PowerShell 5.1 session.
#>

[CmdletBinding()]
param(
    [switch]$SkipDism,
    [switch]$SkipSfc,
    [switch]$SkipUpdates,
    [switch]$InstallUpdates,
    [switch]$CreateRestorePoint,
    [switch]$ForceRestoreHealth,
    [switch]$KillHungDism,
    [switch]$ResetBase,

    [string]$DismSource,
    [switch]$LimitAccess,

    [int]$DismNoActivityTimeoutMinutes = 30,
    [int]$DismCheckIntervalSeconds = 15,
    [int]$MinimumFreeSpaceGB = 10,

    [string]$LogFolder = (Join-Path $env:SystemDrive 'WindowsRepairLogs')
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Continue'

$Script:Results = New-Object System.Collections.Generic.List[object]
$Script:TranscriptStarted = $false

$Timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$LogFile = Join-Path $LogFolder "WindowsRepair_$Timestamp.log"
$SummaryFile = Join-Path $LogFolder "WindowsRepair_$Timestamp.summary.csv"

function Write-Section {
    param([Parameter(Mandatory)][string]$Title)

    Write-Host ''
    Write-Host '========================================'
    Write-Host $Title
    Write-Host '========================================'
}

function Write-Info {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message
}

function Write-Success {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Notice {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Failure {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Add-StepResult {
    param(
        [Parameter(Mandatory)][string]$Step,
        [Parameter(Mandatory)][string]$Command,
        [object]$ExitCode,
        [Parameter(Mandatory)][string]$Status,
        [string]$Log,
        [string]$Message
    )

    $Script:Results.Add([pscustomobject]@{
        Time     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Step     = $Step
        Command  = $Command
        ExitCode = $ExitCode
        Status   = $Status
        Log      = $Log
        Message  = $Message
    }) | Out-Null
}

function Get-ExitCodeHex {
    param([Parameter(Mandatory)][int]$ExitCode)

    $Bytes = [System.BitConverter]::GetBytes([int32]$ExitCode)
    $Unsigned = [System.BitConverter]::ToUInt32($Bytes, 0)
    return ('0x{0:X8}' -f $Unsigned)
}

function Get-ExitCodeDetails {
    param(
        [Parameter(Mandatory)][int]$ExitCode,
        [string]$ToolName = ''
    )

    $HexCode = Get-ExitCodeHex -ExitCode $ExitCode

    $KnownCodes = @{
        '0x00000000' = 'Success.'
        '0x800F081F' = 'DISM could not find required source files. Use -DismSource with matching Windows media, or remove -LimitAccess so Windows Update can be used.'
        '0x800F0906' = 'DISM could not download source files. Check internet, proxy, WSUS policy, or use -DismSource.'
        '0x800F0907' = 'DISM source access may be blocked by policy. Check Windows Update or WSUS policy, or use -DismSource with -LimitAccess.'
        '0x800F0922' = 'Windows servicing failed. Common causes include servicing stack/update issues, partition space issues, VPN/proxy issues, or component store problems.'
        '0x800F0954' = 'The system may be configured to use WSUS and cannot reach the repair source. Check policy or use -DismSource.'
        '0x80073701' = 'A referenced assembly could not be found. Review CBS and DISM logs. Matching Windows media may be needed.'
        '0x80073712' = 'The component store is corrupted or a required manifest/file is missing. Review CBS and DISM logs.'
        '0x80070002' = 'File not found. A required update, package, or repair source file could not be located.'
        '0x80070005' = 'Access denied. Confirm the shell is elevated and security software or policy is not blocking the operation.'
        '0x80070020' = 'A file is in use by another process. Reboot and rerun the repair.'
        '0x80070057' = 'Invalid parameter. Check command arguments, source path, and DISM syntax.'
        '0x80070422' = 'A required service is disabled. Check Windows Update, BITS, Cryptographic Services, and Windows Modules Installer.'
        '0x80070643' = 'Installation failure. Review Windows Update history, CBS log, DISM log, and event logs.'
        '0x80072EE2' = 'Network timeout while contacting update or repair source. Check internet, proxy, firewall, or WSUS.'
        '0x80072EFE' = 'Connection was interrupted while contacting update or repair source. Check network, proxy, firewall, or WSUS.'
        '0x8024001E' = 'Windows Update operation did not complete because the service or system state changed. Reboot and try again.'
        '0x8024401C' = 'Windows Update timed out, often due to proxy, WSUS, firewall, or network issues.'
        '0x8024402C' = 'Windows Update proxy or name resolution issue. Check proxy, DNS, and WSUS settings.'
        '0x80244022' = 'Windows Update server or WSUS service is unavailable or overloaded.'
        '0x8024A105' = 'Windows Update automatic update service error. Restart update services, reboot, and retry.'
        '0x00000001' = 'General failure. Review the command-specific log for details.'
        '0x00000002' = 'Incorrect usage or file not found, depending on the tool. Check command arguments and paths.'
        '0x00000003' = 'Path not found. Check paths used by the command.'
        '0x00000005' = 'Access denied. Confirm elevation and permissions.'
    }

    if ($KnownCodes.ContainsKey($HexCode)) {
        return "$HexCode / $ExitCode - $($KnownCodes[$HexCode])"
    }

    if ($ToolName -match 'sfc') {
        return "$HexCode / $ExitCode - SFC returned an unmapped exit code. Review CBS.log and this transcript."
    }

    if ($ToolName -match 'dism') {
        return "$HexCode / $ExitCode - DISM returned an unmapped exit code. Review the DISM step log and CBS.log."
    }

    return "$HexCode / $ExitCode - Unmapped exit code. Review command output and logs."
}

function Write-ExitSummary {
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][int]$ExitCode,
        [Parameter(Mandatory)][string]$Details,
        [string]$ExtraDetails
    )

    if ($ExitCode -eq 0) {
        Write-Success "$Description completed successfully with exit code: $ExitCode"
        Write-Success "Exit code details: $Details"
        if ($ExtraDetails) { Write-Success $ExtraDetails }
    }
    else {
        Write-Notice "$Description completed with exit code: $ExitCode"
        Write-Notice "Exit code details: $Details"
        if ($ExtraDetails) { Write-Notice $ExtraDetails }
    }
}

function ConvertTo-ArgumentString {
    param([string[]]$Arguments)

    $Quoted = foreach ($Argument in $Arguments) {
        if ($null -eq $Argument) { continue }

        if ($Argument -match '[\s"]') {
            '"' + ($Argument -replace '"', '\"') + '"'
        }
        else {
            $Argument
        }
    }

    return ($Quoted -join ' ')
}

function Initialize-Logging {
    if (-not (Test-Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
    }

    Start-Transcript -Path $LogFile -Append | Out-Null
    $Script:TranscriptStarted = $true

    Write-Section 'Windows Repair and Update Check Script'
    Write-Info "Log file: $LogFile"
    Write-Info "Summary file: $SummaryFile"
}

function Write-SystemInfo {
    Write-Section 'System Information'

    try {
        Get-ComputerInfo |
            Select-Object CsName, WindowsProductName, WindowsVersion, OsBuildNumber, OsArchitecture, BiosFirmwareType |
            Format-List
    }
    catch {
        Write-Notice "Unable to read full system information: $($_.Exception.Message)"
    }
}

function Test-SystemDriveFreeSpace {
    Write-Section 'Free Space Check'

    try {
        $DriveName = $env:SystemDrive.TrimEnd(':')
        $Drive = Get-PSDrive -Name $DriveName -ErrorAction Stop
        $FreeGB = [math]::Round($Drive.Free / 1GB, 2)

        if ($FreeGB -lt $MinimumFreeSpaceGB) {
            $Message = "System drive has only $FreeGB GB free. Recommended minimum is $MinimumFreeSpaceGB GB."
            Write-Notice $Message
            Add-StepResult -Step 'Free Space Check' -Command "Get-PSDrive $DriveName" -ExitCode 0 -Status 'Warning' -Log $LogFile -Message $Message
        }
        else {
            $Message = "System drive has $FreeGB GB free."
            Write-Success $Message
            Add-StepResult -Step 'Free Space Check' -Command "Get-PSDrive $DriveName" -ExitCode 0 -Status 'Success' -Log $LogFile -Message $Message
        }
    }
    catch {
        $Message = "Unable to check free space: $($_.Exception.Message)"
        Write-Notice $Message
        Add-StepResult -Step 'Free Space Check' -Command 'Get-PSDrive' -ExitCode 1 -Status 'Warning' -Log $LogFile -Message $Message
    }
}

function Test-PendingReboot {
    $Reasons = New-Object System.Collections.Generic.List[string]

    $RegistryChecks = @(
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'; Reason = 'Component Based Servicing reboot pending' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'; Reason = 'Windows Update reboot required' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting'; Reason = 'Windows Update post-reboot reporting pending' }
    )

    foreach ($Check in $RegistryChecks) {
        if (Test-Path $Check.Path) {
            $Reasons.Add($Check.Reason) | Out-Null
        }
    }

    try {
        $SessionManagerPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
        $SessionManager = Get-ItemProperty -Path $SessionManagerPath -ErrorAction SilentlyContinue

        if ($null -ne $SessionManager -and ($SessionManager.PSObject.Properties.Name -contains 'PendingFileRenameOperations')) {
            if ($SessionManager.PendingFileRenameOperations) {
                $Reasons.Add('Pending file rename operations') | Out-Null
            }
        }
    }
    catch {
        # Ignore registry read errors for this optional check.
    }

    return [pscustomobject]@{
        IsPending = ($Reasons.Count -gt 0)
        Reasons   = $Reasons.ToArray()
    }
}

function Write-PendingRebootStatus {
    Write-Section 'Pending Reboot Check'

    $Status = Test-PendingReboot

    if ($Status.IsPending) {
        Write-Notice 'A pending reboot was detected. Repairs or updates may not fully complete until the system is restarted.'
        foreach ($Reason in $Status.Reasons) {
            Write-Info "- $Reason"
        }

        Add-StepResult -Step 'Pending Reboot Check' -Command 'Registry reboot checks' -ExitCode 0 -Status 'Warning' -Log $LogFile -Message ($Status.Reasons -join '; ')
    }
    else {
        Write-Success 'No pending reboot indicators were detected.'
        Add-StepResult -Step 'Pending Reboot Check' -Command 'Registry reboot checks' -ExitCode 0 -Status 'Success' -Log $LogFile -Message 'No pending reboot indicators detected.'
    }
}

function New-SystemRestorePointIfRequested {
    if (-not $CreateRestorePoint) {
        Add-StepResult -Step 'Restore Point' -Command 'Checkpoint-Computer' -ExitCode $null -Status 'Skipped' -Log $LogFile -Message 'CreateRestorePoint was not specified.'
        return
    }

    Write-Section 'Creating System Restore Point'

    try {
        Checkpoint-Computer -Description "Before Windows Repair Script $Timestamp" -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop
        Write-Success 'System restore point created successfully.'
        Add-StepResult -Step 'Restore Point' -Command 'Checkpoint-Computer' -ExitCode 0 -Status 'Success' -Log $LogFile -Message 'Restore point created.'
    }
    catch {
        $Message = "Unable to create restore point. System Protection may be disabled, or a recent restore point may already exist. $($_.Exception.Message)"
        Write-Notice $Message
        Add-StepResult -Step 'Restore Point' -Command 'Checkpoint-Computer' -ExitCode 1 -Status 'Warning' -Log $LogFile -Message $Message
    }
}

function Get-DismProgressPercent {
    param([string[]]$Paths)

    foreach ($Path in $Paths) {
        if (-not (Test-Path $Path)) { continue }

        try {
            $RecentLines = Get-Content -Path $Path -Tail 100 -ErrorAction Stop
            $Found = @()

            foreach ($Line in $RecentLines) {
                $Matches = [regex]::Matches($Line, '(?<![0-9])([0-9]{1,3})(?:[.][0-9]+)?%')

                foreach ($Match in $Matches) {
                    $Value = [int]$Match.Groups[1].Value
                    if ($Value -ge 0 -and $Value -le 100) {
                        $Found += $Value
                    }
                }
            }

            if ($Found.Count -gt 0) {
                return ($Found | Select-Object -Last 1)
            }
        }
        catch {
            continue
        }
    }

    return $null
}

function Get-DismHealthStatus {
    param([string[]]$Paths)

    $Lines = @()

    foreach ($Path in $Paths) {
        if (Test-Path $Path) {
            try {
                $Lines += Get-Content -Path $Path -Tail 500 -ErrorAction Stop
            }
            catch {
                # Ignore unreadable files.
            }
        }
    }

    $Text = $Lines -join [Environment]::NewLine

    if ($Text -match 'No component store corruption detected') {
        return [pscustomobject]@{ Status = 'Healthy'; Message = 'No component store corruption detected.' }
    }

    if ($Text -match 'The component store is repairable') {
        return [pscustomobject]@{ Status = 'Repairable'; Message = 'The component store is repairable. RestoreHealth should be run.' }
    }

    if ($Text -match 'The component store cannot be repaired') {
        return [pscustomobject]@{ Status = 'NonRepairable'; Message = 'The component store cannot be repaired by DISM.' }
    }

    if ($Text -match 'The restore operation completed successfully') {
        return [pscustomobject]@{ Status = 'Repaired'; Message = 'The restore operation completed successfully.' }
    }

    return [pscustomobject]@{ Status = 'Unknown'; Message = 'DISM health status could not be determined from captured output or logs.' }
}

function New-DismResult {
    param(
        [int]$ExitCode,
        [string]$HealthStatus,
        [string]$HealthMessage,
        [string]$StepLog,
        [string]$StdOutLog,
        [string]$StdErrLog
    )

    return [pscustomobject]@{
        ExitCode      = $ExitCode
        HealthStatus  = $HealthStatus
        HealthMessage = $HealthMessage
        StepLog       = $StepLog
        StdOutLog     = $StdOutLog
        StdErrLog     = $StdErrLog
    }
}

function Run-Command {
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string]$Command,
        [string[]]$Arguments = @()
    )

    Write-Section $Description
    $CommandLine = "$Command $($Arguments -join ' ')".Trim()
    Write-Info "Running: $CommandLine"

    try {
        $Process = Start-Process -FilePath $Command -ArgumentList $Arguments -Wait -NoNewWindow -PassThru -ErrorAction Stop
        $Details = Get-ExitCodeDetails -ExitCode $Process.ExitCode -ToolName $Command

        Write-Host ''
        Write-ExitSummary -Description $Description -ExitCode $Process.ExitCode -Details $Details

        if ($Process.ExitCode -eq 0) {
            Add-StepResult -Step $Description -Command $CommandLine -ExitCode $Process.ExitCode -Status 'Success' -Log $LogFile -Message $Details
        }
        else {
            Add-StepResult -Step $Description -Command $CommandLine -ExitCode $Process.ExitCode -Status 'Warning' -Log $LogFile -Message $Details
        }

        return $Process.ExitCode
    }
    catch {
        $Message = $_.Exception.Message
        Write-Failure "ERROR: $Message"
        Add-StepResult -Step $Description -Command $CommandLine -ExitCode 1 -Status 'Error' -Log $LogFile -Message $Message
        return 1
    }
}

function Run-DismWithMonitor {
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string[]]$DismArguments,
        [int]$NoActivityTimeoutMinutes = 30,
        [int]$CheckIntervalSeconds = 15,
        [switch]$KillOnHang
    )

    Write-Section $Description

    $SafeName = $Description -replace '[^a-zA-Z0-9_-]', '_'
    $StepLog = Join-Path $LogFolder "$SafeName.dism.log"
    $StdOutLog = Join-Path $LogFolder "$SafeName.stdout.log"
    $StdErrLog = Join-Path $LogFolder "$SafeName.stderr.log"

    $Arguments = @()
    $Arguments += $DismArguments

    if ($DismSource -and ($DismArguments -contains '/RestoreHealth')) {
        $Arguments += "/Source:$DismSource"
    }

    if ($LimitAccess -and ($DismArguments -contains '/RestoreHealth')) {
        $Arguments += '/LimitAccess'
    }

    $Arguments += "/LogPath:$StepLog"
    $Arguments += '/LogLevel:4'

    $ArgumentString = ConvertTo-ArgumentString -Arguments $Arguments
    $CommandLine = "DISM.exe $ArgumentString"

    Write-Info "Running: $CommandLine"
    Write-Info "DISM log: $StepLog"
    Write-Info "DISM output log: $StdOutLog"
    Write-Info "DISM error log: $StdErrLog"

    try {
        $Process = Start-Process -FilePath 'DISM.exe' `
            -ArgumentList $Arguments `
            -PassThru `
            -NoNewWindow `
            -RedirectStandardOutput $StdOutLog `
            -RedirectStandardError $StdErrLog `
            -ErrorAction Stop
    }
    catch {
        $Message = $_.Exception.Message
        Write-Failure "ERROR: Unable to start DISM. $Message"
        Add-StepResult -Step $Description -Command $CommandLine -ExitCode 1 -Status 'Error' -Log $StepLog -Message $Message
        return (New-DismResult -ExitCode 1 -HealthStatus 'Error' -HealthMessage $Message -StepLog $StepLog -StdOutLog $StdOutLog -StdErrLog $StdErrLog)
    }

    $LastActivity = Get-Date
    $LastLogWriteTime = [datetime]::MinValue
    $LastOutWriteTime = [datetime]::MinValue
    $LastErrWriteTime = [datetime]::MinValue
    $LastCpuTime = 0.0
    $LastProgressPercent = $null

    Write-Progress -Activity $Description -Status 'Starting DISM...' -PercentComplete 0

    while (-not $Process.HasExited) {
        Start-Sleep -Seconds $CheckIntervalSeconds
        $Process.Refresh()

        foreach ($PathInfo in @(
            @{ Path = $StepLog; Last = 'LastLogWriteTime' },
            @{ Path = $StdOutLog; Last = 'LastOutWriteTime' },
            @{ Path = $StdErrLog; Last = 'LastErrWriteTime' }
        )) {
            $Item = Get-Item $PathInfo.Path -ErrorAction SilentlyContinue

            if ($null -ne $Item) {
                if ($PathInfo.Last -eq 'LastLogWriteTime' -and $Item.LastWriteTime -gt $LastLogWriteTime) {
                    $LastLogWriteTime = $Item.LastWriteTime
                    $LastActivity = Get-Date
                }
                elseif ($PathInfo.Last -eq 'LastOutWriteTime' -and $Item.LastWriteTime -gt $LastOutWriteTime) {
                    $LastOutWriteTime = $Item.LastWriteTime
                    $LastActivity = Get-Date
                }
                elseif ($PathInfo.Last -eq 'LastErrWriteTime' -and $Item.LastWriteTime -gt $LastErrWriteTime) {
                    $LastErrWriteTime = $Item.LastWriteTime
                    $LastActivity = Get-Date
                }
            }
        }

        try {
            $LiveProcess = Get-Process -Id $Process.Id -ErrorAction Stop
            $CurrentCpuTime = [double]$LiveProcess.CPU

            if ($CurrentCpuTime -gt $LastCpuTime) {
                $LastCpuTime = $CurrentCpuTime
                $LastActivity = Get-Date
            }
        }
        catch {
            # Process may have exited between checks.
        }

        $InactiveFor = New-TimeSpan -Start $LastActivity -End (Get-Date)
        $ProgressPercent = Get-DismProgressPercent -Paths @($StdOutLog, $StepLog)
        $InactiveMinutes = [math]::Round($InactiveFor.TotalMinutes, 1)

        if ($null -ne $ProgressPercent) {
            $LastProgressPercent = $ProgressPercent
            Write-Progress -Activity $Description -Status "DISM running. Last activity: $InactiveMinutes minute(s) ago." -PercentComplete $ProgressPercent
            Write-Info "DISM still running. Progress: $ProgressPercent%. Last detected activity: $InactiveMinutes minute(s) ago."
        }
        elseif ($null -ne $LastProgressPercent) {
            Write-Progress -Activity $Description -Status "DISM running. Last known progress: $LastProgressPercent%. Last activity: $InactiveMinutes minute(s) ago." -PercentComplete $LastProgressPercent
            Write-Info "DISM still running. Last known progress: $LastProgressPercent%. Last detected activity: $InactiveMinutes minute(s) ago."
        }
        else {
            Write-Progress -Activity $Description -Status "DISM running. Progress not reported yet. Last activity: $InactiveMinutes minute(s) ago." -PercentComplete 0
            Write-Info "DISM still running. Progress not reported yet. Last detected activity: $InactiveMinutes minute(s) ago."
        }

        if ($InactiveFor.TotalMinutes -ge $NoActivityTimeoutMinutes) {
            Write-Notice "DISM appears to have had no log, output, error, or CPU activity for $NoActivityTimeoutMinutes minutes."
            Write-Notice 'Recent DISM log entries:'

            if (Test-Path $StepLog) {
                Get-Content -Path $StepLog -Tail 20
            }
            else {
                Write-Notice 'No DISM log file was found yet.'
            }

            if ($KillOnHang) {
                Write-Failure 'Stopping DISM because KillHungDism was specified.'

                try {
                    Stop-Process -Id $Process.Id -Force -ErrorAction Stop
                }
                catch {
                    Write-Notice "Unable to stop DISM process: $($_.Exception.Message)"
                }

                $Message = "DISM was stopped after no activity for $NoActivityTimeoutMinutes minutes."
                Add-StepResult -Step $Description -Command $CommandLine -ExitCode 1 -Status 'Error' -Log $StepLog -Message $Message
                throw $Message
            }

            Write-Notice 'DISM was not stopped. Continuing to monitor.'
            $LastActivity = Get-Date
        }
    }

    Write-Progress -Activity $Description -Completed
    $Process.Refresh()

    $Details = Get-ExitCodeDetails -ExitCode $Process.ExitCode -ToolName 'DISM.exe'
    $Health = Get-DismHealthStatus -Paths @($StdOutLog, $StdErrLog, $StepLog)
    $HealthText = "DISM health result: $($Health.Status) - $($Health.Message)"

    Write-Host ''
    Write-ExitSummary -Description $Description -ExitCode $Process.ExitCode -Details $Details -ExtraDetails $HealthText

    if ($Process.ExitCode -eq 0) {
        Add-StepResult -Step $Description -Command $CommandLine -ExitCode $Process.ExitCode -Status 'Success' -Log $StepLog -Message "$Details $($Health.Message)"
    }
    else {
        Add-StepResult -Step $Description -Command $CommandLine -ExitCode $Process.ExitCode -Status 'Warning' -Log $StepLog -Message "$Details $($Health.Message)"
        Write-Notice "Review the DISM log: $StepLog"
    }

    return (New-DismResult -ExitCode $Process.ExitCode -HealthStatus $Health.Status -HealthMessage $Health.Message -StepLog $StepLog -StdOutLog $StdOutLog -StdErrLog $StdErrLog)
}

function Invoke-DismRepairSequence {
    if ($SkipDism) {
        Write-Section 'DISM Repair Sequence'
        Write-Notice 'Skipped because SkipDism was specified.'
        Add-StepResult -Step 'DISM Repair Sequence' -Command 'DISM' -ExitCode $null -Status 'Skipped' -Log $LogFile -Message 'SkipDism was specified.'
        return
    }

    $CleanupArgs = @('/Online', '/Cleanup-Image', '/StartComponentCleanup')

    if ($ResetBase) {
        Write-Notice 'ResetBase was specified. Installed updates cannot be uninstalled after this cleanup.'
        $CleanupArgs += '/ResetBase'
    }

    Run-DismWithMonitor `
        -Description 'DISM StartComponentCleanup' `
        -DismArguments $CleanupArgs `
        -NoActivityTimeoutMinutes $DismNoActivityTimeoutMinutes `
        -CheckIntervalSeconds $DismCheckIntervalSeconds `
        -KillOnHang:$KillHungDism | Out-Null

    $ScanResult = Run-DismWithMonitor `
        -Description 'DISM ScanHealth' `
        -DismArguments @('/Online', '/Cleanup-Image', '/ScanHealth') `
        -NoActivityTimeoutMinutes 20 `
        -CheckIntervalSeconds $DismCheckIntervalSeconds `
        -KillOnHang:$KillHungDism

    $ShouldRunRestoreHealth = $true

    if ($ForceRestoreHealth) {
        Write-Notice 'ForceRestoreHealth was specified. RestoreHealth will run regardless of ScanHealth result.'
    }
    elseif ($ScanResult.HealthStatus -eq 'Healthy') {
        $ShouldRunRestoreHealth = $false
        Write-Section 'DISM RestoreHealth'
        Write-Success "Skipped because ScanHealth reported: $($ScanResult.HealthMessage)"
        Add-StepResult -Step 'DISM RestoreHealth' -Command 'DISM.exe /Online /Cleanup-Image /RestoreHealth' -ExitCode $null -Status 'Skipped' -Log $LogFile -Message 'Skipped because ScanHealth reported no component store corruption. Use -ForceRestoreHealth to run anyway.'
    }
    elseif ($ScanResult.HealthStatus -eq 'NonRepairable') {
        $ShouldRunRestoreHealth = $false
        Write-Section 'DISM RestoreHealth'
        Write-Notice 'Skipped because ScanHealth reported that the component store cannot be repaired by DISM.'
        Add-StepResult -Step 'DISM RestoreHealth' -Command 'DISM.exe /Online /Cleanup-Image /RestoreHealth' -ExitCode $null -Status 'Skipped' -Log $LogFile -Message $ScanResult.HealthMessage
    }
    elseif ($ScanResult.HealthStatus -eq 'Unknown') {
        Write-Notice 'ScanHealth completed, but the script could not confirm the health result from DISM output/logs. RestoreHealth will run as a conservative fallback.'
    }

    if ($ShouldRunRestoreHealth) {
        Run-DismWithMonitor `
            -Description 'DISM RestoreHealth' `
            -DismArguments @('/Online', '/Cleanup-Image', '/RestoreHealth') `
            -NoActivityTimeoutMinutes $DismNoActivityTimeoutMinutes `
            -CheckIntervalSeconds $DismCheckIntervalSeconds `
            -KillOnHang:$KillHungDism | Out-Null
    }
}

function Invoke-SfcRepair {
    if ($SkipSfc) {
        Write-Section 'System File Checker'
        Write-Notice 'Skipped because SkipSfc was specified.'
        Add-StepResult -Step 'System File Checker' -Command 'sfc.exe /scannow' -ExitCode $null -Status 'Skipped' -Log $LogFile -Message 'SkipSfc was specified.'
        return
    }

    Run-Command -Description 'System File Checker' -Command 'sfc.exe' -Arguments @('/scannow') | Out-Null
}

function Get-WindowsUpdateResultText {
    param([int]$ResultCode)

    switch ($ResultCode) {
        0 { return 'NotStarted' }
        1 { return 'InProgress' }
        2 { return 'Succeeded' }
        3 { return 'SucceededWithErrors' }
        4 { return 'Failed' }
        5 { return 'Aborted' }
        default { return "Unknown($ResultCode)" }
    }
}

function Start-RequiredUpdateServices {
    Write-Section 'Windows Update Service Check'

    $ServiceNames = @('wuauserv', 'bits', 'cryptsvc', 'trustedinstaller')

    foreach ($ServiceName in $ServiceNames) {
        try {
            $Service = Get-Service -Name $ServiceName -ErrorAction Stop
            Write-Info "$ServiceName status: $($Service.Status)"

            if ($Service.Status -ne 'Running') {
                Write-Info "Starting $ServiceName..."
                Start-Service -Name $ServiceName -ErrorAction Stop
                Start-Sleep -Seconds 2
                $Service.Refresh()
                Write-Info "$ServiceName status after start attempt: $($Service.Status)"
            }

            Add-StepResult -Step "Service Check: $ServiceName" -Command "Start-Service $ServiceName" -ExitCode 0 -Status 'Success' -Log $LogFile -Message "Status: $($Service.Status)"
        }
        catch {
            $Message = $_.Exception.Message
            Write-Notice "Unable to verify or start $ServiceName. $Message"
            Add-StepResult -Step "Service Check: $ServiceName" -Command "Start-Service $ServiceName" -ExitCode 1 -Status 'Warning' -Log $LogFile -Message $Message
        }
    }
}

function Invoke-WindowsUpdateCheck {
    if ($SkipUpdates) {
        Write-Section 'Windows Update Check'
        Write-Notice 'Skipped because SkipUpdates was specified.'
        Add-StepResult -Step 'Windows Update Check' -Command 'Microsoft.Update.Session' -ExitCode $null -Status 'Skipped' -Log $LogFile -Message 'SkipUpdates was specified.'
        return
    }

    Write-Section 'Checking for Windows Updates'

    try {
        Start-RequiredUpdateServices

        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

        Write-Info 'Recent Windows Update history, last 10 entries:'
        try {
            $HistoryCount = $UpdateSearcher.GetTotalHistoryCount()
            if ($HistoryCount -gt 0) {
                $HistoryItems = $UpdateSearcher.QueryHistory(0, [Math]::Min(10, $HistoryCount))
                $HistoryItems | Select-Object Date, Title, ResultCode, HResult | Format-Table -AutoSize
            }
            else {
                Write-Info 'No Windows Update history entries found.'
            }
        }
        catch {
            Write-Notice "Unable to read update history: $($_.Exception.Message)"
        }

        $SearchCriteria = "IsInstalled=0 and Type='Software' and IsHidden=0"
        Write-Info "Searching with criteria: $SearchCriteria"
        $SearchResult = $UpdateSearcher.Search($SearchCriteria)

        if ($SearchResult.Updates.Count -eq 0) {
            Write-Success 'No available Windows Updates found.'
            Add-StepResult -Step 'Windows Update Search' -Command $SearchCriteria -ExitCode 0 -Status 'Success' -Log $LogFile -Message 'No available updates found.'
            return
        }

        Write-Success "Available Windows Updates found: $($SearchResult.Updates.Count)"

        for ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
            $Update = $SearchResult.Updates.Item($i)
            Write-Info "$($i + 1). $($Update.Title)"
            Write-Info "   KB Articles: $($Update.KBArticleIDs -join ', ')"
            Write-Info "   Requires Reboot: $($Update.RebootRequired)"
            Write-Info "   Downloaded: $($Update.IsDownloaded)"
            Write-Host ''
        }

        Add-StepResult -Step 'Windows Update Search' -Command $SearchCriteria -ExitCode 0 -Status 'UpdatesAvailable' -Log $LogFile -Message "$($SearchResult.Updates.Count) available update(s) found."

        if (-not $InstallUpdates) {
            Write-Notice 'InstallUpdates was not specified, so updates were listed only.'
            Add-StepResult -Step 'Windows Update Install' -Command 'InstallUpdates switch' -ExitCode $null -Status 'Skipped' -Log $LogFile -Message 'Available updates were listed only.'
            return
        }

        Write-Section 'Installing Windows Updates'
        $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl

        for ($i = 0; $i -lt $SearchResult.Updates.Count; $i++) {
            $Update = $SearchResult.Updates.Item($i)

            if (-not $Update.EulaAccepted) {
                Write-Info "Accepting EULA for: $($Update.Title)"
                $Update.AcceptEula()
            }

            [void]$UpdatesToInstall.Add($Update)
        }

        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $UpdatesToInstall
        $DownloadResult = $Downloader.Download()
        $DownloadText = Get-WindowsUpdateResultText -ResultCode ([int]$DownloadResult.ResultCode)

        if ([int]$DownloadResult.ResultCode -eq 2) {
            Write-Success "Download result: $($DownloadResult.ResultCode) [$DownloadText]"
        }
        else {
            Write-Notice "Download result: $($DownloadResult.ResultCode) [$DownloadText]"
        }

        Add-StepResult -Step 'Windows Update Download' -Command 'CreateUpdateDownloader().Download()' -ExitCode ([int]$DownloadResult.ResultCode) -Status $DownloadText -Log $LogFile -Message "Download result: $DownloadText"

        $Installer = $UpdateSession.CreateUpdateInstaller()
        $Installer.Updates = $UpdatesToInstall
        $InstallResult = $Installer.Install()
        $InstallText = Get-WindowsUpdateResultText -ResultCode ([int]$InstallResult.ResultCode)

        if ([int]$InstallResult.ResultCode -eq 2) {
            Write-Success "Install result: $($InstallResult.ResultCode) [$InstallText]"
        }
        else {
            Write-Notice "Install result: $($InstallResult.ResultCode) [$InstallText]"
        }

        Write-Info "Reboot required: $($InstallResult.RebootRequired)"

        Write-Section 'Per-Update Install Results'
        for ($i = 0; $i -lt $UpdatesToInstall.Count; $i++) {
            $Update = $UpdatesToInstall.Item($i)
            $PerUpdateResult = $InstallResult.GetUpdateResult($i)
            $PerUpdateText = Get-WindowsUpdateResultText -ResultCode ([int]$PerUpdateResult.ResultCode)
            $HResult = [int]$PerUpdateResult.HResult
            $HResultHex = Get-ExitCodeHex -ExitCode $HResult

            Write-Info "$($i + 1). $($Update.Title)"
            Write-Info "   Result: $($PerUpdateResult.ResultCode) [$PerUpdateText]"
            Write-Info "   HResult: $HResultHex / $HResult"

            if ($HResult -ne 0) {
                Write-Notice "   Details: $(Get-ExitCodeDetails -ExitCode $HResult -ToolName 'WindowsUpdate')"
            }
        }

        Add-StepResult -Step 'Windows Update Install' -Command 'CreateUpdateInstaller().Install()' -ExitCode ([int]$InstallResult.ResultCode) -Status $InstallText -Log $LogFile -Message "Reboot required: $($InstallResult.RebootRequired)"
    }
    catch {
        $Message = $_.Exception.Message
        Write-Failure 'ERROR: Failed to check, download, or install Windows Updates.'
        Write-Failure $Message
        Add-StepResult -Step 'Windows Update' -Command 'Microsoft.Update.Session' -ExitCode 1 -Status 'Error' -Log $LogFile -Message $Message
    }
}

function Write-FinalSummary {
    Write-Section 'Summary'

    if ($Script:Results.Count -gt 0) {
        $Script:Results | Format-Table -AutoSize

        try {
            $Script:Results | Export-Csv -Path $SummaryFile -NoTypeInformation -Force
            Write-Success "Summary saved to: $SummaryFile"
        }
        catch {
            Write-Notice "Unable to write CSV summary: $($_.Exception.Message)"
        }
    }
    else {
        Write-Notice 'No step results were recorded.'
    }

    Write-Info "Transcript log saved to: $LogFile"
}

try {
    Initialize-Logging

    Write-SystemInfo
    Test-SystemDriveFreeSpace
    Write-PendingRebootStatus
    New-SystemRestorePointIfRequested

    Invoke-DismRepairSequence
    Invoke-SfcRepair
    Invoke-WindowsUpdateCheck

    Write-Section 'Repair and Update Check Complete'
    Write-Success 'Script completed.'
    Write-Info "Log saved to: $LogFile"
    Write-Info "Summary saved to: $SummaryFile"
}
catch {
    $Message = $_.Exception.Message
    Write-Failure 'ERROR: The repair script encountered a problem.'
    Write-Failure $Message
    Add-StepResult -Step 'Script Error' -Command 'Main' -ExitCode 1 -Status 'Error' -Log $LogFile -Message $Message
}
finally {
    Write-FinalSummary

    if ($Script:TranscriptStarted) {
        try {
            Stop-Transcript | Out-Null
        }
        catch {
            Write-Notice "Unable to stop transcript cleanly: $($_.Exception.Message)"
        }
    }
}
