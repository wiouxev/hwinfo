# Remote Windows System Information Query Script
# Usage: .\Get-RemoteSystemInfo.ps1 -ComputerName "TARGET_IP_OR_NAME" [-Credential $cred]

param(
    [Parameter(Mandatory=$true)]
    [string]$ComputerName,
    [PSCredential]$Credential = $null,
    [string]$ExportPath = $null,
    [ValidateSet('TXT', 'CSV', 'JSON', 'XML', 'HTML')]
    [string]$ExportFormat = 'TXT'
)

# Function to format bytes to human readable
function Format-Bytes {
    param([long]$Size)
    if ($Size -gt 1TB) { return "{0:N2} TB" -f ($Size / 1TB) }
    elseif ($Size -gt 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
    elseif ($Size -gt 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
    else { return "{0:N2} KB" -f ($Size / 1KB) }
}

# Initialize output collection
$outputData = @{}
$outputText = @()

Write-Host "Querying system information for: $ComputerName" -ForegroundColor Green
Write-Host "=" * 60

# Helper function to add to both console and export
function Add-Output {
    param($Text, $Category = "General", $Key = $null, $Value = $null)
    Write-Host $Text
    $script:outputText += $Text
    
    if ($Key -and $Value) {
        if (!$script:outputData[$Category]) { $script:outputData[$Category] = @{} }
        $script:outputData[$Category][$Key] = $Value
    }
}

try {
    # Test connection first
    if (!(Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
        throw "Unable to reach $ComputerName"
    }

    # Set up session parameters
    $sessionParams = @{ ComputerName = $ComputerName }
    if ($Credential) { $sessionParams.Credential = $Credential }

    # === BASIC SYSTEM INFO ===
    Write-Host "`n[SYSTEM INFORMATION]" -ForegroundColor Yellow
    
    $systemInfo = Get-WmiObject -Class Win32_ComputerSystem @sessionParams
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem @sessionParams
    $biosInfo = Get-WmiObject -Class Win32_BIOS @sessionParams
    
    Write-Host "Computer Name    : $($systemInfo.Name)"
    Write-Host "Domain/Workgroup : $($systemInfo.Domain)"
    Write-Host "Manufacturer     : $($systemInfo.Manufacturer)"
    Write-Host "Model            : $($systemInfo.Model)"
    Write-Host "Serial Number    : $($biosInfo.SerialNumber)"
    Write-Host "OS Version       : $($osInfo.Caption) $($osInfo.Version)"
    Write-Host "Architecture     : $($osInfo.OSArchitecture)"
    Write-Host "Last Boot Time   : $($osInfo.ConvertToDateTime($osInfo.LastBootUpTime))"
    Write-Host "Uptime           : $([math]::Round((New-TimeSpan -Start $osInfo.ConvertToDateTime($osInfo.LastBootUpTime)).TotalDays, 1)) days"

    # === NETWORK INFORMATION ===
    Write-Host "`n[NETWORK INFORMATION]" -ForegroundColor Yellow
    
    $networkAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration @sessionParams | 
        Where-Object { $_.IPEnabled -eq $true }
    
    foreach ($adapter in $networkAdapters) {
        Write-Host "Adapter          : $($adapter.Description)"
        Write-Host "IP Address       : $($adapter.IPAddress -join ', ')"
        Write-Host "Subnet Mask      : $($adapter.IPSubnet -join ', ')"
        Write-Host "Default Gateway  : $($adapter.DefaultIPGateway -join ', ')"
        Write-Host "DNS Servers      : $($adapter.DNSServerSearchOrder -join ', ')"
        Write-Host "DHCP Enabled     : $($adapter.DHCPEnabled)"
        if ($adapter.DHCPEnabled) {
            Write-Host "DHCP Server      : $($adapter.DHCPServer)"
        }
        Write-Host ""
    }

    # === CPU INFORMATION ===
    Write-Host "[PROCESSOR INFORMATION]" -ForegroundColor Yellow
    
    $cpuInfo = Get-WmiObject -Class Win32_Processor @sessionParams
    foreach ($cpu in $cpuInfo) {
        Write-Host "Processor        : $($cpu.Name.Trim())"
        Write-Host "Cores            : $($cpu.NumberOfCores)"
        Write-Host "Logical Procs    : $($cpu.NumberOfLogicalProcessors)"
        Write-Host "Max Clock Speed  : $($cpu.MaxClockSpeed) MHz"
        Write-Host "Current Load     : $($cpu.LoadPercentage)%"
        Write-Host ""
    }

    # === MEMORY INFORMATION ===
    Write-Host "[MEMORY INFORMATION]" -ForegroundColor Yellow
    
    $totalRAM = [math]::Round($osInfo.TotalVisibleMemorySize * 1024, 0)
    $freeRAM = [math]::Round($osInfo.FreePhysicalMemory * 1024, 0)
    $usedRAM = $totalRAM - $freeRAM
    
    Write-Host "Total RAM        : $(Format-Bytes $totalRAM)"
    Write-Host "Used RAM         : $(Format-Bytes $usedRAM) ($([math]::Round(($usedRAM / $totalRAM) * 100, 1))%)"
    Write-Host "Free RAM         : $(Format-Bytes $freeRAM) ($([math]::Round(($freeRAM / $totalRAM) * 100, 1))%)"
    
    # Physical memory modules
    $memoryModules = Get-WmiObject -Class Win32_PhysicalMemory @sessionParams
    Write-Host "`nMemory Modules   : $($memoryModules.Count)"
    foreach ($module in $memoryModules) {
        $capacity = Format-Bytes $module.Capacity
        Write-Host "  - $capacity @ $($module.Speed) MHz ($($module.Manufacturer))"
    }

    # === DISK INFORMATION ===
    Write-Host "`n[DISK INFORMATION]" -ForegroundColor Yellow
    
    # Logical disks
    $disks = Get-WmiObject -Class Win32_LogicalDisk @sessionParams | 
        Where-Object { $_.DriveType -eq 3 }  # Fixed disks only
    
    Write-Host "Logical Drives:"
    foreach ($disk in $disks) {
        $totalSize = $disk.Size
        $freeSpace = $disk.FreeSpace
        $usedSpace = $totalSize - $freeSpace
        $percentFree = [math]::Round(($freeSpace / $totalSize) * 100, 1)
        
        Write-Host "  Drive $($disk.DeviceID)"
        Write-Host "    Total Size   : $(Format-Bytes $totalSize)"
        Write-Host "    Used Space   : $(Format-Bytes $usedSpace) ($([math]::Round(($usedSpace / $totalSize) * 100, 1))%)"
        Write-Host "    Free Space   : $(Format-Bytes $freeSpace) ($percentFree%)"
        Write-Host "    File System  : $($disk.FileSystem)"
        Write-Host ""
    }
    
    # Physical disks
    $physicalDisks = Get-WmiObject -Class Win32_DiskDrive @sessionParams
    Write-Host "Physical Disks:"
    foreach ($pDisk in $physicalDisks) {
        Write-Host "  $($pDisk.Model.Trim())"
        Write-Host "    Size         : $(Format-Bytes $pDisk.Size)"
        Write-Host "    Interface    : $($pDisk.InterfaceType)"
        Write-Host ""
    }

    # === GRAPHICS INFORMATION ===
    Write-Host "[GRAPHICS INFORMATION]" -ForegroundColor Yellow
    
    $videoCards = Get-WmiObject -Class Win32_VideoController @sessionParams | 
        Where-Object { $_.Name -notlike "*Remote*" -and $_.Name -notlike "*Virtual*" }
    
    foreach ($gpu in $videoCards) {
        Write-Host "Graphics Card    : $($gpu.Name)"
        if ($gpu.AdapterRAM -gt 0) {
            Write-Host "Video Memory     : $(Format-Bytes $gpu.AdapterRAM)"
        }
        Write-Host "Driver Version   : $($gpu.DriverVersion)"
        Write-Host "Current Mode     : $($gpu.VideoModeDescription)"
        Write-Host ""
    }

    # === SOFTWARE INFORMATION ===
    Write-Host "[SOFTWARE INFORMATION]" -ForegroundColor Yellow
    
    # Windows version details
    $windowsVersion = Get-WmiObject -Class Win32_OperatingSystem @sessionParams
    Write-Host "Windows Edition  : $($windowsVersion.Caption)"
    Write-Host "Build Number     : $($windowsVersion.BuildNumber)"
    Write-Host "Install Date     : $($windowsVersion.ConvertToDateTime($windowsVersion.InstallDate))"
    Write-Host "Last Update      : $($windowsVersion.ConvertToDateTime($windowsVersion.LastBootUpTime))"
    
    # PowerShell version (requires PS remoting)
    try {
        $psVersion = Invoke-Command @sessionParams -ScriptBlock { $PSVersionTable.PSVersion } -ErrorAction SilentlyContinue
        if ($psVersion) {
            Write-Host "PowerShell Ver   : $($psVersion.ToString())"
        }
    } catch {
        Write-Host "PowerShell Ver   : Unable to determine (PS Remoting may be disabled)"
    }

    # === SERVICES STATUS ===
    Write-Host "`n[CRITICAL SERVICES STATUS]" -ForegroundColor Yellow
    
    $criticalServices = @('Spooler', 'BITS', 'Winmgmt', 'EventLog', 'Themes', 'AudioSrv')
    $services = Get-WmiObject -Class Win32_Service @sessionParams | 
        Where-Object { $_.Name -in $criticalServices }
    
    foreach ($service in $services) {
        $status = if ($service.State -eq 'Running') { 'Running' } else { $service.State }
        Write-Host "$($service.DisplayName.PadRight(25)) : $status"
    }

    Write-Host "`n" + "=" * 60
    Write-Host "System information query completed successfully!" -ForegroundColor Green

} catch {
    Write-Host "`nError occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure the target computer is reachable"
    Write-Host "2. Verify you have administrative privileges"
    Write-Host "3. Check if Windows Firewall is blocking WMI"
    Write-Host "4. Ensure WMI service is running on target machine"
    Write-Host "5. Use -Credential parameter if different authentication is needed"
}

# Example usage:
# .\Get-RemoteSystemInfo.ps1 -ComputerName "192.168.1.100"
# .\Get-RemoteSystemInfo.ps1 -ComputerName "COMPUTER-NAME" -Credential (Get-Credential)