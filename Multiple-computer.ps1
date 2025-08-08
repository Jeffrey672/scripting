# Requires: ActiveDirectory module, admin rights, and remote WMI access

Import-Module ActiveDirectory

# Output CSV path
$outputPath = "C:\AD_Computer_Inventory.csv"

# Get all computers in the domain
$computers = Get-ADComputer -Filter * -Property Name, Description

$results = foreach ($computer in $computers) {
    $compName = $computer.Name
    $desc = $computer.Description

    # Default values
    $serial = $assetTag = $ip = $subnet = $null

    try {
        # Get BIOS info
        $bios = Get-WmiObject -Class Win32_BIOS -ComputerName $compName -ErrorAction Stop
        $serial = $bios.SerialNumber
        $assetTag = $bios.SMBIOSAssetTag

        # Get network info
        $net = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $compName -Filter "IPEnabled=TRUE" -ErrorAction Stop | Select-Object -First 1
        if ($net) {
            $ip = $net.IPAddress[0]
            $subnet = $net.IPSubnet[0]
        }
    } catch {
        # If the computer is offline or access denied, leave fields blank
    }

    [PSCustomObject]@{
        ComputerName = $compName
        SerialNumber = $serial
        AssetTag     = $assetTag
        IPAddress    = $ip
        Subnet       = $subnet
        Description  = $desc
    }
}

$results | Export-Csv -Path $outputPath -NoTypeInformation
Write-Host "Export complete: $outputPath"