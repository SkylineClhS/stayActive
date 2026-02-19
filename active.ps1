<#
    Keep-Active-NumLock.ps1 (Enhanced & Robust)
    - Prompts user for how many minutes to run
    - Automatically stops after the specified duration
    - Toggles NumLock twice to simulate activity (no state change)
    - Compatible with Windows PowerShell 5.1 and PowerShell 7+
#>

[CmdletBinding()]
param(
    [int]$IntervalSeconds = 300  # default 5 minutes per cycle
)

# --- Prompt for duration (minutes) ---
$RunMinutesRaw = Read-Host "How many minutes do you want this script to run?"
if (-not ($RunMinutesRaw -as [int])) {
    Write-Host "[ERROR] Invalid number. Please enter an integer (e.g., 15)." -ForegroundColor Red
    exit 1
}
$RunMinutes = [int]$RunMinutesRaw
if ($RunMinutes -le 0) {
    Write-Host "[ERROR] Minutes must be greater than 0." -ForegroundColor Red
    exit 1
}

$StopAt = (Get-Date).AddMinutes($RunMinutes)

# --- Load SendKeys provider ---
$UseComFallback = $false
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
} catch {
    Write-Host "[WARN] System.Windows.Forms not available; falling back to COM (WScript.Shell)." -ForegroundColor Yellow
    $UseComFallback = $true
}

# Create COM object only if needed
$script:WshShell = $null
if ($UseComFallback -and $null -eq $script:WshShell) {
    try {
        $script:WshShell = New-Object -ComObject WScript.Shell
    } catch {
        Write-Host "[ERROR] Failed to initialize WScript.Shell COM object: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

function SendKeys-Activity {
    param([Parameter(Mandatory)][string]$Keys)
    if ($UseComFallback) {
        $script:WshShell.SendKeys($Keys)
    } else {
        [System.Windows.Forms.SendKeys]::SendWait($Keys)
    }
}

Write-Host "----------------------------------------------------------"
Write-Host ("Interval: {0} seconds | Run for: {1} minute(s) | Stop time: {2}" -f $IntervalSeconds, $RunMinutes, $StopAt)
Write-Host ("PowerShell version: {0}" -f $PSVersionTable.PSVersion)
Write-Host "Stop manually anytime with Ctrl+C"
Write-Host "----------------------------------------------------------`n"

try {
    while ($true) {
        # Auto-stop check
        $now = Get-Date
        if ($now -ge $StopAt) {
            Write-Host "`n[INFO] Time limit reached. Script stopped automatically." -ForegroundColor Cyan
            break
        }

        # Remaining time (mm:ss)
        $remainingTs = $StopAt - $now
        $remainingStr = "{0:D2}:{1:D2}" -f [int][math]::Floor($remainingTs.TotalMinutes), $remainingTs.Seconds

        # Heartbeat
        Write-Host ("{0} running ... (Remaining: {1} min:sec)" -f $now.ToString("yyyy-MM-dd HH:mm:ss"), $remainingStr) -ForegroundColor Cyan

        # Simulate activity without changing NumLock state
        SendKeys-Activity '{NUMLOCK}{NUMLOCK}'

        # Sleep per interval (but don't overshoot stop time)
        $sleepSeconds = [math]::Min($IntervalSeconds, [int][math]::Ceiling(($StopAt - (Get-Date)).TotalSeconds))
        if ($sleepSeconds -gt 0) {
            Start-Sleep -Seconds $sleepSeconds
        }
    }
}
catch [System.Management.Automation.StopException] {
    Write-Host "`n[INFO] Stopped by user." -ForegroundColor Cyan
}
catch {
    Write-Host ("`n[ERROR] {0}" -f $_) -ForegroundColor Red
    exit 1
}