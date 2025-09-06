# Remote Windows System Information Query Script
# Usage: .\Get-RemoteSystemInfo.ps1 -ComputerName "TARGET_IP_OR_NAME" [-Credential $cred]

param(
    [Parameter(Mandatory=$true)]
    [string]$ComputerName,
    [PSCredential]$Credential = $null
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

# Progress tracking
$sections = @('Network Test', 'System Info', 'Network', 'CPU', 'Memory', 'Disk', 'Graphics', 'Software', 'Services')
$currentSection = 0

try {
    # Improved connectivity test
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    
    Write-Host "Testing connectivity to $ComputerName..." -ForegroundColor Cyan
    
    if (!(Test-NetConnection -ComputerName $ComputerName -Port 135 -InformationLevel Quiet)) {
        throw "Unable to reach $ComputerName on WMI port (135). Check firewall and WMI service."
    }

    # Set up session parameters
    $sessionParams = @{ 
        ComputerName = $ComputerName
        ErrorAction = 'Stop'
    }
    if ($Credential) { $sessionParams.Credential = $Credential }

    # === BASIC SYSTEM INFO ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "`n[SYSTEM INFORMATION]" -ForegroundColor Yellow
    
    $systemInfo = Get-WmiObject -Class Win32_ComputerSystem @sessionParams
    $osInfo = Get-WmiObject -Class Win32_OperatingSystem @sessionParams
    $biosInfo = Get-WmiObject -Class Win32_BIOS @sessionParams
    
    Add-Output "Computer Name    : $($systemInfo.Name)" "System" "ComputerName" $systemInfo.Name
    Add-Output "Domain/Workgroup : $($systemInfo.Domain)" "System" "Domain" $systemInfo.Domain
    Add-Output "Manufacturer     : $($systemInfo.Manufacturer)" "System" "Manufacturer" $systemInfo.Manufacturer
    Add-Output "Model            : $($systemInfo.Model)" "System" "Model" $systemInfo.Model
    Add-Output "Serial Number    : $($biosInfo.SerialNumber)" "System" "SerialNumber" $biosInfo.SerialNumber
    Add-Output "OS Version       : $($osInfo.Caption) $($osInfo.Version)" "System" "OSVersion" "$($osInfo.Caption) $($osInfo.Version)"
    Add-Output "Architecture     : $($osInfo.OSArchitecture)" "System" "Architecture" $osInfo.OSArchitecture
    
    $lastBootTime = $osInfo.ConvertToDateTime($osInfo.LastBootUpTime)
    $uptime = [math]::Round((New-TimeSpan -Start $lastBootTime).TotalDays, 1)
    Add-Output "Last Boot Time   : $lastBootTime" "System" "LastBootTime" $lastBootTime
    Add-Output "Uptime           : $uptime days" "System" "Uptime" "$uptime days"

    # === NETWORK INFORMATION ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "`n[NETWORK INFORMATION]" -ForegroundColor Yellow
    
    $networkAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration @sessionParams | 
        Where-Object { $_.IPEnabled -eq $true }
    
    $adapterCount = 0
    foreach ($adapter in $networkAdapters) {
        $adapterCount++
        Add-Output "Adapter          : $($adapter.Description)" "Network" "Adapter$adapterCount" $adapter.Description
        Add-Output "IP Address       : $($adapter.IPAddress -join ', ')" "Network" "IPAddress$adapterCount" ($adapter.IPAddress -join ', ')
        Add-Output "Subnet Mask      : $($adapter.IPSubnet -join ', ')" "Network" "SubnetMask$adapterCount" ($adapter.IPSubnet -join ', ')
        Add-Output "Default Gateway  : $($adapter.DefaultIPGateway -join ', ')" "Network" "Gateway$adapterCount" ($adapter.DefaultIPGateway -join ', ')
        Add-Output "DNS Servers      : $($adapter.DNSServerSearchOrder -join ', ')" "Network" "DNS$adapterCount" ($adapter.DNSServerSearchOrder -join ', ')
        Add-Output "DHCP Enabled     : $($adapter.DHCPEnabled)" "Network" "DHCP$adapterCount" $adapter.DHCPEnabled
        if ($adapter.DHCPEnabled) {
            Add-Output "DHCP Server      : $($adapter.DHCPServer)" "Network" "DHCPServer$adapterCount" $adapter.DHCPServer
        }
        Add-Output ""
    }

    # === CPU INFORMATION ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "[PROCESSOR INFORMATION]" -ForegroundColor Yellow
    
    $cpuInfo = Get-WmiObject -Class Win32_Processor @sessionParams
    $cpuCount = 0
    foreach ($cpu in $cpuInfo) {
        $cpuCount++
        Add-Output "Processor        : $($cpu.Name.Trim())" "CPU" "Processor$cpuCount" $cpu.Name.Trim()
        Add-Output "Cores            : $($cpu.NumberOfCores)" "CPU" "Cores$cpuCount" $cpu.NumberOfCores
        Add-Output "Logical Procs    : $($cpu.NumberOfLogicalProcessors)" "CPU" "LogicalProcs$cpuCount" $cpu.NumberOfLogicalProcessors
        Add-Output "Max Clock Speed  : $($cpu.MaxClockSpeed) MHz" "CPU" "MaxClock$cpuCount" "$($cpu.MaxClockSpeed) MHz"
        Add-Output "Current Load     : $($cpu.LoadPercentage)%" "CPU" "Load$cpuCount" "$($cpu.LoadPercentage)%"
        Add-Output ""
    }

    # === MEMORY INFORMATION ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "[MEMORY INFORMATION]" -ForegroundColor Yellow
    
    $totalRAM = [math]::Round($osInfo.TotalVisibleMemorySize * 1024, 0)
    $freeRAM = [math]::Round($osInfo.FreePhysicalMemory * 1024, 0)
    $usedRAM = $totalRAM - $freeRAM
    
    Add-Output "Total RAM        : $(Format-Bytes $totalRAM)" "Memory" "TotalRAM" (Format-Bytes $totalRAM)
    Add-Output "Used RAM         : $(Format-Bytes $usedRAM) ($([math]::Round(($usedRAM / $totalRAM) * 100, 1))%)" "Memory" "UsedRAM" "$(Format-Bytes $usedRAM) ($([math]::Round(($usedRAM / $totalRAM) * 100, 1))%)"
    Add-Output "Free RAM         : $(Format-Bytes $freeRAM) ($([math]::Round(($freeRAM / $totalRAM) * 100, 1))%)" "Memory" "FreeRAM" "$(Format-Bytes $freeRAM) ($([math]::Round(($freeRAM / $totalRAM) * 100, 1))%)"
    
    # Physical memory modules
    $memoryModules = Get-WmiObject -Class Win32_PhysicalMemory @sessionParams
    Add-Output "`nMemory Modules   : $($memoryModules.Count)" "Memory" "ModuleCount" $memoryModules.Count
    $moduleCount = 0
    foreach ($module in $memoryModules) {
        $moduleCount++
        $capacity = Format-Bytes $module.Capacity
        Add-Output "  - $capacity @ $($module.Speed) MHz ($($module.Manufacturer))" "Memory" "Module$moduleCount" "$capacity @ $($module.Speed) MHz ($($module.Manufacturer))"
    }

    # === DISK INFORMATION ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "`n[DISK INFORMATION]" -ForegroundColor Yellow
    
    # Logical disks
    $disks = Get-WmiObject -Class Win32_LogicalDisk @sessionParams | 
        Where-Object { $_.DriveType -eq 3 }  # Fixed disks only
    
    Add-Output "Logical Drives:"
    foreach ($disk in $disks) {
        $totalSize = $disk.Size
        $freeSpace = $disk.FreeSpace
        $usedSpace = $totalSize - $freeSpace
        $percentFree = [math]::Round(($freeSpace / $totalSize) * 100, 1)
        
        Add-Output "  Drive $($disk.DeviceID)" "Disk" "Drive$($disk.DeviceID)_Label" $disk.DeviceID
        Add-Output "    Total Size   : $(Format-Bytes $totalSize)" "Disk" "Drive$($disk.DeviceID)_Total" (Format-Bytes $totalSize)
        Add-Output "    Used Space   : $(Format-Bytes $usedSpace) ($([math]::Round(($usedSpace / $totalSize) * 100, 1))%)" "Disk" "Drive$($disk.DeviceID)_Used" "$(Format-Bytes $usedSpace) ($([math]::Round(($usedSpace / $totalSize) * 100, 1))%)"
        Add-Output "    Free Space   : $(Format-Bytes $freeSpace) ($percentFree%)" "Disk" "Drive$($disk.DeviceID)_Free" "$(Format-Bytes $freeSpace) ($percentFree%)"
        Add-Output "    File System  : $($disk.FileSystem)" "Disk" "Drive$($disk.DeviceID)_FileSystem" $disk.FileSystem
        Add-Output ""
    }
    
    # Physical disks
    $physicalDisks = Get-WmiObject -Class Win32_DiskDrive @sessionParams
    Add-Output "Physical Disks:"
    $physicalCount = 0
    foreach ($pDisk in $physicalDisks) {
        $physicalCount++
        Add-Output "  $($pDisk.Model.Trim())" "Disk" "PhysicalDisk$physicalCount" $pDisk.Model.Trim()
        Add-Output "    Size         : $(Format-Bytes $pDisk.Size)" "Disk" "PhysicalDisk$($physicalCount)_Size" (Format-Bytes $pDisk.Size)
        Add-Output "    Interface    : $($pDisk.InterfaceType)" "Disk" "PhysicalDisk$($physicalCount)_Interface" $pDisk.InterfaceType
        Add-Output ""
    }

    # === GRAPHICS INFORMATION ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "[GRAPHICS INFORMATION]" -ForegroundColor Yellow
    
    $videoCards = Get-WmiObject -Class Win32_VideoController @sessionParams | 
        Where-Object { $_.Name -notlike "*Remote*" -and $_.Name -notlike "*Virtual*" }
    
    $gpuCount = 0
    foreach ($gpu in $videoCards) {
        $gpuCount++
        Add-Output "Graphics Card    : $($gpu.Name)" "Graphics" "GPU$gpuCount" $gpu.Name
        if ($gpu.AdapterRAM -gt 0) {
            Add-Output "Video Memory     : $(Format-Bytes $gpu.AdapterRAM)" "Graphics" "GPU$($gpuCount)_Memory" (Format-Bytes $gpu.AdapterRAM)
        }
        Add-Output "Driver Version   : $($gpu.DriverVersion)" "Graphics" "GPU$($gpuCount)_Driver" $gpu.DriverVersion
        Add-Output "Current Mode     : $($gpu.VideoModeDescription)" "Graphics" "GPU$($gpuCount)_Mode" $gpu.VideoModeDescription
        Add-Output ""
    }

    # === SOFTWARE INFORMATION ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "[SOFTWARE INFORMATION]" -ForegroundColor Yellow
    
    # Windows version details
    Add-Output "Windows Edition  : $($osInfo.Caption)" "Software" "WindowsEdition" $osInfo.Caption
    Add-Output "Build Number     : $($osInfo.BuildNumber)" "Software" "BuildNumber" $osInfo.BuildNumber
    
    $installDate = $osInfo.ConvertToDateTime($osInfo.InstallDate)
    Add-Output "Install Date     : $installDate" "Software" "InstallDate" $installDate
    
    # PowerShell version (requires PS remoting)
    try {
        $psVersion = Invoke-Command @sessionParams -ScriptBlock { $PSVersionTable.PSVersion } -ErrorAction SilentlyContinue -TimeoutSec 10
        if ($psVersion) {
            Add-Output "PowerShell Ver   : $($psVersion.ToString())" "Software" "PowerShellVersion" $psVersion.ToString()
        }
    } catch {
        Add-Output "PowerShell Ver   : Unable to determine (PS Remoting may be disabled)" "Software" "PowerShellVersion" "Unable to determine"
    }

    # === SERVICES STATUS ===
    $currentSection++
    Write-Progress -Activity "Gathering System Information" -Status $sections[$currentSection-1] -PercentComplete (($currentSection / $sections.Count) * 100)
    Write-Host "`n[CRITICAL SERVICES STATUS]" -ForegroundColor Yellow
    
    try {
        $criticalServices = @('Spooler', 'BITS', 'Winmgmt', 'EventLog', 'Themes', 'AudioSrv')
        
        # Add timeout to prevent hanging
        $services = Get-WmiObject -Class Win32_Service @sessionParams | 
            Where-Object { $_.Name -in $criticalServices }
        
        foreach ($service in $services) {
            $status = if ($service.State -eq 'Running') { 'Running' } else { $service.State }
            Add-Output "$($service.DisplayName.PadRight(25)) : $status" "Services" $service.Name $status
        }
    } catch {
        Add-Output "Unable to retrieve service status: $($_.Exception.Message)" "Services" "Error" "Unable to retrieve"
        Write-Host "Warning: Could not retrieve service status" -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host ("=" * 60)
    Write-Host "System information query completed successfully!" -ForegroundColor Green

} catch {
    Write-Host "`nError occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure the target computer is reachable"
    Write-Host "2. Verify you have administrative privileges"
    Write-Host "3. Check if Windows Firewall is blocking WMI"
    Write-Host "4. Ensure WMI service is running on target machine"
    Write-Host "5. Use -Credential parameter if different authentication is needed"
} finally {
    Write-Progress -Completed -Activity "Gathering System Information"
}

# Simple export functionality
Write-Host ""
Write-Host "Export Options" -ForegroundColor Cyan
$exportChoice = Read-Host "Export report? (Y/N)"

if ($exportChoice -eq 'Y' -or $exportChoice -eq 'y') {
    Write-Host "1. TXT file"
    Write-Host "2. CSV file" 
    Write-Host "3. JSON file"
    $format = Read-Host "Choose format (1-3)"
    
    # Get file extension based on choice
    $extension = ""
    if ($format -eq '1') { $extension = "txt" }
    elseif ($format -eq '2') { $extension = "csv" }
    elseif ($format -eq '3') { $extension = "json" }
    else { 
        Write-Host "Invalid choice. Exiting export." -ForegroundColor Red
        return
    }
    
    # Suggest default path
    $defaultFileName = "$ComputerName-Report-$(Get-Date -Format 'yyyy-MM-dd-HHmm').$extension"
    $suggestedPath = "C:\temp\$defaultFileName"
    
    Write-Host ""
    Write-Host "Suggested path: $suggestedPath" -ForegroundColor Gray
    Write-Host "You can change the path and/or filename as needed." -ForegroundColor Gray
    $exportPath = Read-Host "Enter full path and filename"
    
    # If user just pressed Enter, use the suggested path
    if ([string]::IsNullOrWhiteSpace($exportPath)) {
        $exportPath = $suggestedPath
        # Create C:\temp if it doesn't exist
        $tempDir = Split-Path $exportPath -Parent
        if (!(Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            Write-Host "Created directory: $tempDir" -ForegroundColor Yellow
        }
    }
    
    try {
        Write-Host "Exporting..." -ForegroundColor Cyan
        
        if ($format -eq '1') {
            $outputText | Out-File -FilePath $exportPath -Encoding UTF8
        } elseif ($format -eq '2') {
            $csvData = @()
            foreach ($category in $outputData.Keys) {
                foreach ($key in $outputData[$category].Keys) {
                    $csvData += [PSCustomObject]@{
                        Category = $category
                        Property = $key
                        Value = $outputData[$category][$key]
                    }
                }
            }
            $csvData | Export-Csv -Path $exportPath -NoTypeInformation
        } elseif ($format -eq '3') {
            $outputData | ConvertTo-Json -Depth 3 | Out-File -FilePath $exportPath -Encoding UTF8
        }
        
        Write-Host "Successfully exported to: $exportPath" -ForegroundColor Green
        Write-Host "File size: $([math]::Round((Get-Item $exportPath).Length / 1KB, 1)) KB" -ForegroundColor Gray
        
    } catch {
        Write-Host "Export failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Make sure the directory exists and you have write permissions." -ForegroundColor Yellow
    }
}

Write-Host "Script completed!" -ForegroundColor Green