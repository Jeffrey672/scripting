<#
.SYNOPSIS
    Migrates users from Microsoft 365 desktop apps to Microsoft 365 web apps.
    - Uninstalls Microsoft 365 desktop apps (non-silent).
    - Creates .URL shortcuts on the Public Desktop for web apps.
    - Logs all actions, errors, and warnings.

.DESCRIPTION
    This script is intended for Windows 10/11 environments.
    Run as Administrator. Users will see the uninstall prompts (not silent).
#>

# --- Setup logging ---
$LogPath = "C:\Temp\MS365_WebMigration.log"
$ErrorLogPath = "C:\Temp\MS365_WebMigration_Errors.log"
New-Item -ItemType Directory -Force -Path (Split-Path $LogPath) | Out-Null

Function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp] [$Level] $Message"

    # Write to console with color
    switch ($Level) {
        "INFO"    { Write-Host $entry -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $entry -ForegroundColor Green }
        "WARNING" { Write-Host $entry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $entry -ForegroundColor Red }
        default   { Write-Host $entry }
    }

    # Write to general log
    Add-Content -Path $LogPath -Value $entry

    # Write errors/warnings separately
    if ($Level -eq "ERROR" -or $Level -eq "WARNING") {
        Add-Content -Path $ErrorLogPath -Value $entry
    }
}

Write-Log "=== Microsoft 365 Web Migration Script Started ===" "INFO"

# --- Prerequisite checks (PowerShell 5+ and Admin) ---
try {
    $psVersion = $PSVersionTable.PSVersion.Major
    if ($psVersion -lt 5) {
        Write-Log "PowerShell $psVersion detected. PowerShell 5 or higher is required." "ERROR"
        throw "PowerShell version not supported"
    }

    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "This script must be run as Administrator." "ERROR"
        throw "Administrator privileges required"
    }
} catch {
    Write-Log "Prerequisite check failed: $($_.Exception.Message)" "ERROR"
    return
}

# --- Step 1: Uninstall Microsoft 365 Desktop Apps ---
Write-Log "Attempting to uninstall Microsoft 365 desktop apps..." "INFO"

Function Get-OfficeUninstallEntries {
    param()

    $uninstallPaths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
    )

    $namePatterns = @(
        'Microsoft 365',
        'Office 365',
        'Microsoft Office',
        'Office'
    )

    $entries = @()
    foreach ($path in $uninstallPaths) {
        try {
            if (Test-Path $path) {
                Get-ChildItem $path | ForEach-Object {
                    try {
                        $displayName = (Get-ItemProperty $_.PSPath -ErrorAction Stop).DisplayName
                        $uninstallString = (Get-ItemProperty $_.PSPath -ErrorAction Stop).UninstallString
                        if ([string]::IsNullOrWhiteSpace($displayName) -or [string]::IsNullOrWhiteSpace($uninstallString)) { return }
                        if ($namePatterns | Where-Object { $displayName -match $_ }) {
                            $entries += [pscustomobject]@{
                                DisplayName     = $displayName
                                UninstallString = $uninstallString
                                PSPath          = $_.PSPath
                            }
                        }
                    } catch {}
                }
            }
        } catch {
            Write-Log "Failed reading uninstall path ${path}: $($_.Exception.Message)" "WARNING"
        }
    }
    return $entries | Sort-Object -Property DisplayName -Unique
}

Function Invoke-UninstallCommand {
    param(
        [Parameter(Mandatory=$true)][string]$UninstallString,
        [Parameter(Mandatory=$true)][string]$DisplayName
    )

    try {
        # Normalize command: some entries are like 'MsiExec.exe /I{GUID}' or have quotes and extra args
        $cmd = $UninstallString.Trim()

        # Convert MSI /I to /X for uninstall when needed
        if ($cmd -match '(?i)msiexec\.exe') {
            $cmd = $cmd -replace '(?i)\s*/I', ' /X'
            if ($cmd -notmatch '(?i)/X|/uninstall') {
                $cmd = "$cmd /X"
            }
        }

        # Split executable and arguments safely
        $exe = $null; $cmdArgs = $null
        if ($cmd.StartsWith('"')) {
            $firstQuoteEnd = $cmd.IndexOf('"',1)
            $exe = $cmd.Substring(1, $firstQuoteEnd-1)
            $cmdArgs = $cmd.Substring($firstQuoteEnd+1).Trim()
        } else {
            $parts = $cmd.Split(' ',2)
            $exe = $parts[0]
            $cmdArgs = if ($parts.Count -gt 1) { $parts[1] } else { '' }
        }

        Write-Log "Launching uninstall for $DisplayName (command: $exe $cmdArgs). This may prompt the user..." "WARNING"
        $process = Start-Process -FilePath $exe -ArgumentList $cmdArgs -PassThru -Wait -WindowStyle Normal
        Write-Log "Uninstall process for $DisplayName exited with code $($process.ExitCode)." "INFO"
    } catch {
        Write-Log "Failed to run uninstall for ${DisplayName}: $($_.Exception.Message)" "ERROR"
    }
}

try {
    $entries = Get-OfficeUninstallEntries
    if ($entries -and $entries.Count -gt 0) {
        foreach ($entry in $entries) {
            Write-Log "Found installed Office app: $($entry.DisplayName)" "INFO"
            Invoke-UninstallCommand -UninstallString $entry.UninstallString -DisplayName $entry.DisplayName
        }
    } else {
        Write-Log "No Microsoft 365/Office apps found to uninstall via registry. Attempting CIM fallback..." "WARNING"
        try {
            $officeApps = Get-CimInstance -ClassName Win32_Product -ErrorAction Stop | Where-Object { $_.Name -match "Microsoft 365|Office 365|Office" }
            if ($officeApps) {
                foreach ($app in $officeApps) {
                    Write-Log "Found installed Office app (CIM): $($app.Name)" "INFO"
                    try {
                        Write-Log "Launching uninstall for $($app.Name). This may prompt the user..." "WARNING"
                        $null = Invoke-CimMethod -InputObject $app -MethodName Uninstall -ErrorAction Stop
                        Write-Log "Uninstall command sent for $($app.Name)." "SUCCESS"
                    } catch {
                        Write-Log "Failed to uninstall $($app.Name): $($_.Exception.Message)" "ERROR"
                    }
                }
            } else {
                Write-Log "No Microsoft 365/Office apps found to uninstall." "WARNING"
            }
        } catch {
            Write-Log "CIM query failed: $($_.Exception.Message)" "ERROR"
        }
    }
} catch {
    Write-Log "Error during uninstall phase: $($_.Exception.Message)" "ERROR"
}

# --- Step 2: Create .URL Shortcuts for Web Apps ---
Write-Log "Creating Microsoft 365 web app shortcuts on the Public Desktop..." "INFO"

$ShortcutPath = "C:\Users\Public\Desktop"
$WebApps = @{
    "Word (Web)"       = "https://www.office.com/launch/word"
    "Excel (Web)"      = "https://www.office.com/launch/excel"
    "PowerPoint (Web)" = "https://www.office.com/launch/powerpoint"
    "Outlook (Web)"    = "https://outlook.office.com/"
    "OneDrive (Web)"   = "https://onedrive.live.com/"
    "Teams (Web)"      = "https://teams.microsoft.com/"
}

foreach ($app in $WebApps.Keys) {
    $filePath = Join-Path $ShortcutPath "$app.url"
    try {
        Write-Log "Creating shortcut: $filePath" "INFO"
        @"
[InternetShortcut]
URL=$($WebApps[$app])
IconFile=%SystemRoot%\system32\SHELL32.dll
IconIndex=0
"@ | Out-File -FilePath $filePath -Encoding ASCII -Force
        Write-Log "Shortcut created: $app" "SUCCESS"
    } catch {
        Write-Log "Failed to create shortcut for ${app}: $($_.Exception.Message)" "ERROR"
    }
}

Write-Log "=== Microsoft 365 Web Migration Script Complete ===" "INFO"
Write-Log "Log file: $LogPath" "INFO"
Write-Log "Error/Warning log: $ErrorLogPath" "INFO"