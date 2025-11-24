# **Configuration**: Path to the Excel file (change this to your actual file path)
$excelFilePath = "C:\Path\To\Your\ComputerList.xlsx"

# **1. Start Excel COM Application and Open the Workbook**:contentReference[oaicite:4]{index=4}
# This uses the Excel COM object model. Excel must be installed for this to work.
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false                       # Run Excel in the background (not visible)
$excel.DisplayAlerts = $false                 # Suppress any Excel alerts (e.g., prompts)

# Open the workbook (read-only recommended to avoid locking the file)
$workbook = $excel.Workbooks.Open($excelFilePath, [Type]::Missing, $true) 

# Get the first worksheet in the workbook
$worksheet = $workbook.Worksheets.Item(1)

# **2. Read Computer Names from the first column** (assumes first column has the names)
$computerNames = @()   # Array to store computer names
$row = 1               # Start at the first row. If the first row is a header, we'll skip it below.

# If first cell looks like a header (e.g., "ComputerName"), skip it by starting at row 2
$firstCellValue = $worksheet.Cells.Item(1, 1).Text
if ($firstCellValue -match '^[Cc]omputer\s*Name$' -or $firstCellValue -match '^[Hh]ost(Name)?$') {
    $row = 2
}

# Loop down the first column until an empty cell is encountered
while ($true) {
    $cellValue = $worksheet.Cells.Item($row, 1).Text  # Read the cell value as text:contentReference[oaicite:5]{index=5}
    if ([string]::IsNullOrWhiteSpace($cellValue)) { 
        break  # Stop when we reach an empty cell (end of list)
    }
    $computerNames += $cellValue
    $row++
}

# **Close the Excel workbook and quit Excel** (no longer needed now that we have the data)
$workbook.Close($false)     # Close without saving changes
$excel.Quit()               # Quit the Excel application:contentReference[oaicite:6]{index=6}
# Release COM objects to fully quit Excel process
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)  | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)     | Out-Null

# **3. Ping Each Computer and collect results**
$results = @()  # will hold objects with ComputerName, Status, IP

foreach ($name in $computerNames) {
    if ([string]::IsNullOrWhiteSpace($name)) { continue }  # skip empty or whitespace names if any

    # Initialize status and IP variables for this computer
    $status = "Offline"
    $ipAddress = ""

    # Use Test-Connection to send a single ping (ICMP echo):contentReference[oaicite:7]{index=7}:contentReference[oaicite:8]{index=8}
    # -Count 1 sends one ping. -ErrorAction SilentlyContinue prevents errors from unresolvable names.
    $pingResult = Test-Connection -ComputerName $name -Count 1 -ErrorAction SilentlyContinue

    if ($pingResult) {
        # If $pingResult is not $null, we got a response
        $status = "Online"
        # Extract the IPv4 address from the ping result object
        # The ping result has an IPv4Address property (an IP object); convert it to string:contentReference[oaicite:9]{index=9}:contentReference[oaicite:10]{index=10}
        try {
            $ipAddress = ($pingResult.IPV4Address).IPAddressToString
        } catch {
            $ipAddress = ""
        }
    } else {
        # No ping response (Offline). 
        # Optionally attempt DNS resolution to get IP even if host is offline:
        try {
            $resolvedIP = [System.Net.Dns]::GetHostAddresses($name) | 
                          Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | 
                          Select-Object -First 1
            if ($resolvedIP) {
                $ipAddress = $resolvedIP.ToString()
            }
        } catch {
            # If name cannot be resolved (e.g., DNS failure), leave $ipAddress blank.
            $ipAddress = ""
        }
    }

    # Create an object with the results for this computer
    $results += [PSCustomObject]@{
        ComputerName = $name
        Status       = $status    # "Online" or "Offline"
        IP           = $ipAddress # IP address if available, otherwise empty
    }
}

# **4. Export results to a CSV file** in the same folder as the Excel file
$excelFolder = Split-Path -Path $excelFilePath -Parent
# Construct output CSV path with same name + "_PingResults.csv"
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($excelFilePath)
$outputCsvPath = Join-Path -Path $excelFolder -ChildPath "${baseName}_PingResults.csv"

# Use Export-Csv to save the results. This will include the headers "ComputerName,Status,IP".
$results | Export-Csv -Path $outputCsvPath -NoTypeInformation

# Optionally, print a message to console indicating where the file was saved
Write-Host "Ping results saved to $outputCsvPath"