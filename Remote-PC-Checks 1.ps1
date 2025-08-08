# Remote PC Checks
# Eric Melo
# V1.0

## This script will remotely check computers for:
# Any startup files
# The system model/RAM spec
# How many user profiles exist
# Installed software

## START VARIABLE SELECTION ##
## These variables should be set before first run, but can stay for future runs

# Set where you want to save the report
$outputFolder = "C:\MX\"

# Add any files to ignore within the startup folders
$exclusionListFiles = @(
    "BGinfo.lnk"
    )

# Add any user folders to ignore here when counting how many profiles exist on the computer
# Do not add the Public folder here, it will break the startup check
# Final profile count will subtract 1 user profile from the total count
$exclusionListUsers = @(
    "defaultuser0",
    "MLX-LOCALADMIN"
    )
## END VARIABLE SELECTION ##

# Get the computer name to run this script on
$PC = Read-Host "Enter computer name to get information from"
$outFile = Join-Path $outputFolder "$PC.txt"

# Try enabling WinRM, exit the script if this is not possible
try {
    Get-Service -ComputerName $PC -Name winrm -ErrorAction SilentlyContinue | Start-Service
    Start-Sleep 5 # Needed because it may invoke commands may fail without it
} catch {
    Write-Error "Could not enable WinRM on $PC. Exiting."
    exit
}

# Run remote PowerShell checks and capture output
$results = Invoke-Command -ComputerName $PC -ArgumentList $exclusionListFiles, $exclusionListUsers -ErrorAction SilentlyContinue -ScriptBlock {
    param($exclusionListFiles, $exclusionListUsers)
    # Startup folder check
    $pDataStartup = Get-ChildItem "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup" -File -ErrorAction SilentlyContinue |
                     Where-Object { $_.Name -notin $exclusionListFiles }
    if ($pDataStartup) {
        $pDataStartup | ForEach-Object { "File '$($_.Name)' found in ProgramData startup folder" }
    } else {
        "No files found in ProgramData startup folder (outside exclusions)"
    }

    $userFilesFound = $false
    Get-ChildItem "C:\Users" -Directory | Where-Object Name -notin $exclusionListUsers |
    ForEach-Object {
        $userStartup = "$($_.FullName)\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        if (Test-Path $userStartup) {
            $files = Get-ChildItem $userStartup -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notin $exclusionListFiles }
            if ($files) {
                $userFilesFound = $true
                $files | ForEach-Object { "File '$($_.Name)' found in '$($_.DirectoryName)' for user '$($_.Directory.Parent.Name)'" }
            }
        }
    }
    if ($userFilesFound) {
        # Files were found and already reported above
    } else {
        "No files found in any user startup folders (outside exclusions)"
    }

    # Get computer model & RAM
    "Model: $((gcim Win32_ComputerSystem).Model) | RAM: $([math]::Round((gcim Win32_ComputerSystem).TotalPhysicalMemory / 1GB,2)) GB"

    # Get profile count
    $profiles = Get-ChildItem "C:\Users" -Directory | Where-Object Name -notin $exclusionListUsers
    "Found $($profiles.Count-1) profiles in C:\Users folder"
}

# Save results to file
if ($results) {
    $results | Out-File -FilePath $outFile
Write-Output "Results saved to $outFile"
} else {
    Write-Output "Issue running remote report, try again"
}

# Append a blank line between the above outputs and the WMIC output
Add-Content -Path $outFile -Value ""

# Append WMIC name & product info
Write-Output "Running WMIC report on $($PC)..."
$wmicInfo = wmic /node:$PC product get name,version
$wmicInfo | Add-Content -Path $outFile

# Disable WinRM
Get-Service -ComputerName $PC -Name winrm | Stop-Service
Write-Output "Report on $($PC):","", $results