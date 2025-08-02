# Create a temp if its doesn't exist
if (!(Test-Path "C:\temp")) { New-Item -Path "C:\temp" -ItemType Directory }

# Define output file
$outputFile = "C:\temp\installed_programs.txt"
New-Item -ItemType File -Path $outputFile -Force | Out-Null

# Helper function to append section
function Add-Section {
    param ([string]$title)
    Add-Content -Path $outputFile -Value "`n==== $title ====`n"
}

# 1. Installed Programs (32-bit and 64-bit)
Add-Section "Installed Programs"
$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

foreach ($path in $registryPaths) {
    Get-ItemProperty $path -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -and $_.DisplayVersion } |
        Select-Object DisplayName, DisplayVersion |
        ForEach-Object {
            Add-Content -Path $outputFile -Value "$($_.DisplayName) - Version: $($_.DisplayVersion)"
        }
}

# 2. SCCM Software Info - Requires SCCM PowerShell Module
try {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
    Set-Location "<Your-Site-Code>:" # e.g., XYZ:

    $computerName = $env:COMPUTERNAME
    $sccmSoftware = Get-CMSoftwareInventory -Name $computerName
    $sccmSoftware | Out-File "$exportPath\sccm_installed_programs.txt"
} catch {
    "SCCM Module or command failed: $_" | Out-File "$exportPath\sccm_error.log"
}


# 2. Desktop Shortcuts, Documents, Pictures, Videos
Add-Section "User Files: Shortcuts, Documents, Pictures, Videos"
$userProfile = [Environment]::GetFolderPath("UserProfile")

$folders = @{
    "Desktop Shortcuts" = "$userProfile\Desktop\*.lnk"
    "Documents"         = "$userProfile\Documents\*"
    "Pictures"          = "$userProfile\Pictures\*"
    "Videos"            = "$userProfile\Videos\*"
}

foreach ($key in $folders.Keys) {
    Add-Section $key
    Get-ChildItem -Path $folders[$key] -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            Add-Content -Path $outputFile -Value $_.Name
        }
}

# 3. Sticky Notes (Windows 10/11)
Add-Section "Sticky Notes"
$stickyNotesPath = "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\plum.sqlite"
if (Test-Path $stickyNotesPath) {
    try {
        $connectionString = "Data Source=$stickyNotesPath;Version=3;"
        $query = "SELECT Text FROM Note"
        $conn = New-Object System.Data.SQLite.SQLiteConnection($connectionString)
        $conn.Open()
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = $query
        $reader = $cmd.ExecuteReader()
        while ($reader.Read()) {
            Add-Content -Path $outputFile -Value $reader["Text"]
        }
        $conn.Close()
    } catch {
        Add-Content -Path $outputFile -Value "Failed to read Sticky Notes."
    }
} else {
    Add-Content -Path $outputFile -Value "Sticky Notes file not found."
}

# 4. Chrome Bookmarks and Reading List
$chromePath = "$env:LOCALAPPDATA\Google\Chrome\User Data\Default"

$bookmarksFile = Join-Path $chromePath "Bookmarks"
$passwordsFile = Join-Path $chromePath "Login Data" # This is encrypted and not easily exportable
$backupFolder = "C:\temp\chrome_backup"

if (Test-Path $chromePath) {
    New-Item -ItemType Directory -Path $backupFolder -Force | Out-Null

    Copy-Item $bookmarksFile "$backupFolder\Bookmarks.json" -ErrorAction SilentlyContinue
    Copy-Item $passwordsFile "$backupFolder\LoginData_encrypted.db" -ErrorAction SilentlyContinue
    "Manual Chrome export recommended for passwords via Chrome sync." | Out-File "$backupFolder\info.txt"
} else {
    "Chrome data not found." | Out-File "C:\temp\chrome_backup_error.txt"
}

# Parse Chrome Bookmarks JSON to CSV
if (Test-Path "$backupFolder\Bookmarks.json") {
    $json = Get-Content "$backupFolder\Bookmarks.json" -Raw | ConvertFrom-Json
    $bookmarks = $json.roots.bookmark_bar.children + $json.roots.other.children

    $bookmarks | Select-Object name, url | Export-Csv "C:\temp\chrome_bookmarks.csv" -NoTypeInformation
}

# 5. SCCM Software Info - Requires SCCM PowerShell Module
try {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
    Set-Location "<Your-Site-Code>:" # e.g., XYZ:

    $computerName = $env:COMPUTERNAME
    $sccmSoftware = Get-CMSoftwareInventory -Name $computerName
    $sccmSoftware | Out-File "$exportPath\sccm_installed_programs.txt"
} catch {
    "SCCM Module or command failed: $_" | Out-File "$exportPath\sccm_error.log"
}

# 6. Software Restore (Mockup logic — actual installation logic depends on source availability)
$softwareList = Import-Csv "$exportPath\installed_programs.txt" -Delimiter " " -Header Name, Version
$log = "$exportPath\software_install_log.txt"

foreach ($software in $softwareList) {
    # Simulate software reinstallation from web
    $name = $software.DisplayName
    $version = $software.DisplayVersion
    "$name ($version) would be downloaded and installed from internet." | Out-File -Append $log
    # Actual installation logic would go here: e.g., Chocolatey, Winget, or custom download scripts
} 