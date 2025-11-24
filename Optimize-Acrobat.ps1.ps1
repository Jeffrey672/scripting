<#
.SYNOPSIS
  Optimizes Adobe Acrobat / Reader performance on Windows 11
  Author: ChatGPT (GPT-5)
#>

Write-Host "Starting Adobe Acrobat Optimization..." -ForegroundColor Cyan

# --- Step 1: Stop Adobe processes ---
Write-Host "Closing any running Adobe Acrobat processes..."
Stop-Process -Name "AcroRd32","Acrobat","AcroCEF","acrotray" -Force -ErrorAction SilentlyContinue

# --- Step 2: Clear Adobe cache and temp directories ---
$paths = @(
    "$env:LOCALAPPDATA\Adobe\Acrobat",
    "$env:LOCALAPPDATA\Temp\Adobe",
    "$env:APPDATA\Adobe\Acrobat\DC"
)
foreach ($p in $paths) {
    if (Test-Path $p) {
        Write-Host "Clearing cache: $p"
        Remove-Item "$p\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# --- Step 3: Optimize registry for startup and rendering ---
Write-Host "Applying registry optimizations..."

# Disable Adobe splash screen
reg add "HKCU\Software\Adobe\Adobe Acrobat\DC\AVGeneral" /v bDisableSplash /t REG_DWORD /d 1 /f | Out-Null

# Enable hardware acceleration
reg add "HKCU\Software\Adobe\Adobe Acrobat\DC\AVGeneral" /v bUseHWAcceleration /t REG_DWORD /d 1 /f | Out-Null

# Disable startup messages
reg add "HKCU\Software\Adobe\Adobe Acrobat\DC\AVGeneral" /v bShowMsgAtLaunch /t REG_DWORD /d 0 /f | Out-Null

# Disable online storage pane
reg add "HKCU\Software\Adobe\Adobe Acrobat\DC\AVGeneral" /v bToggleAdobeSignPane /t REG_DWORD /d 0 /f | Out-Null

# Disable start page
reg add "HKCU\Software\Adobe\Adobe Acrobat\DC\AVGeneral" /v bShowStartPage /t REG_DWORD /d 0 /f | Out-Null

# --- Step 4: Optional - Move non-essential plug-ins ---
$acroPaths = @(
    "C:\Program Files\Adobe\Acrobat*\Acrobat\plug_ins",
    "C:\Program Files (x86)\Adobe\Acrobat*\Acrobat\plug_ins"
)

foreach ($acroPath in $acroPaths) {
    $resolvedPaths = Get-ChildItem $acroPath -ErrorAction SilentlyContinue
    foreach ($resolved in $resolvedPaths) {
        $optFolder = Join-Path $resolved.FullName "Optional"
        if (!(Test-Path $optFolder)) { New-Item -ItemType Directory -Path $optFolder | Out-Null }

        Write-Host "Moving non-essential plug-ins from $($resolved.FullName)..."
        $nonEssential = @("Accessibility.api","MakeAccessible.api","Weblink.api","Spelling.api","Search5.api")
        foreach ($plugin in $nonEssential) {
            $source = Join-Path $resolved.FullName $plugin
            if (Test-Path $source) {
                Move-Item $source $optFolder -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# --- Step 5: Add Defender Exclusions (optional) ---
$acroDirs = @("C:\Program Files\Adobe","C:\Program Files (x86)\Adobe")
foreach ($dir in $acroDirs) {
    if (Test-Path $dir) {
        Write-Host "Adding Defender exclusion for $dir"
        Add-MpPreference -ExclusionPath $dir -ErrorAction SilentlyContinue
    }
}

Write-Host "`nOptimization Complete!"
Write-Host "Restart your computer or relaunch Adobe Acrobat to apply changes." -ForegroundColor Green
