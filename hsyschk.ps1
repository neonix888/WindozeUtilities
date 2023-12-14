# Function to display the last 20 entries of specified event logs
# REQUIREMENT: PowerShell - Administrator mode else you will get unpleasant messages
function Show-LastEventLogEntries {
    Write-Host "Displaying the last 20 entries of the System log..."
    Get-EventLog -LogName System -Newest 20 | Format-Table -AutoSize

    Write-Host "`nDisplaying the last 20 entries of the Security log..."
    Get-EventLog -LogName Security -Newest 20 | Format-Table -AutoSize

    Write-Host "`nDisplaying the last 20 entries of the Application log..."
    Get-EventLog -LogName Application -Newest 20 | Format-Table -AutoSize
}

# Display Event Logs
Show-LastEventLogEntries

# Function to backup the registry
function Backup-Registry {
    $backupFolder = "C:\tmp"
    if (-not (Test-Path -Path $backupFolder)) {
        $backupFolder = "C:\temp"
        New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null
    }

    $date = Get-Date -Format "yyyyMMdd"
    $backupFileHKLM = "$backupFolder\regbackup_HKLM_$date.reg"
    $backupFileHKCU = "$backupFolder\regbackup_HKCU_$date.reg"

    reg export HKLM $backupFileHKLM /y
    reg export HKCU $backupFileHKCU /y

    Write-Host "Registry backup created at $backupFileHKLM and $backupFileHKCU"
}

# Check Installed Printers
Write-Host "Checking installed printers..."
$printers = Get-Printer
if ($printers) {
    $printers | Format-Table Name, PrinterStatus
} else {
    Write-Host "No printers are installed."
}

# Check Print Spooler Service
Write-Host "`nChecking Print Spooler service..."
$spooler = Get-Service -Name Spooler
Write-Host "Print Spooler Status: $($spooler.Status)"

# Check Network Connectivity for Network Printers
Write-Host "`nChecking network connectivity for network printers..."
$networkPrinters = $printers | Where-Object { $_.PortName -like "IP_*" }
foreach ($printer in $networkPrinters) {
    $ip = $printer.PortName -replace "IP_", ""
    $ping = Test-Connection -ComputerName $ip -Count 1 -Quiet
    if ($ping) {
        Write-Host "Printer $($printer.Name) at $ip is reachable."
    } else {
        Write-Host "Printer $($printer.Name) at $ip is NOT reachable."
    }
}

# Network Configuration Check
Write-Host "`nChecking current network configuration..."
$netConfig = Get-NetIPConfiguration
$netConfig | Format-Table InterfaceAlias, IPv4Address, IPv4DefaultGateway, IPv4SubnetMask

# Test Print to Default Printer
$defaultPrinter = Get-WmiObject -Query " SELECT * FROM Win32_Printer WHERE Default=$true"
if ($defaultPrinter) {
    Write-Host "`nDefault printer is set to: $($defaultPrinter.Name)"
    $testPage = Read-Host "Do you want to print a test page to the default printer? (Y/N)"
    if ($testPage -eq 'Y') {
        $null = $defaultPrinter.PrintTestPage()
        Write-Host "Sent a test page to $($defaultPrinter.Name)..."
    }
} else {
    Write-Host "No default printer is set."
}

# Backup the registry
Write-Host "`nCreating registry backup..."
Backup-Registry
