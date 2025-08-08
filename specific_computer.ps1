# Output CSV path
$outputPath = "C:\computer_info.csv"

# Get current logon user
$currentUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName

# Get BIOS info
$bios = Get-WmiObject -Class Win32_BIOS
$serial = $bios.SerialNumber
$assetTag = $bios.SMBIOSAssetTag

# Get network info (first active adapter)
$net = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE" | Select-Object -First 1
$ip = $net.IPAddress[0]
$subnet = $net.IPSubnet[0]

# Get computer description from AD (if domain-joined)
$desc = $null
try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $compName = $env:COMPUTERNAME
    $adComp = Get-ADComputer -Identity $compName -Property Description
    $desc = $adComp.Description
} catch {
    $desc = "N/A or not domain-joined"
}

# Create object and export
[PSCustomObject]@{
    ComputerName = $env:COMPUTERNAME
    CurrentUser  = $currentUser
    SerialNumber = $serial
    AssetTag     = $assetTag
    IPAddress    = $ip
    Subnet       = $subnet
    Description  = $desc
} | Export-Csv -Path $outputPath -NoTypeInformation

Write-Host "Export complete: $outputPath"