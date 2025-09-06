### REMOTE SYSTEM INFO ###

# Display only (no export)
.\Get-RemoteSystemInfo.ps1 -ComputerName "192.168.1.100"

# Export to text file
.\Get-RemoteSystemInfo.ps1 -ComputerName "192.168.1.100" -ExportPath "C:\Reports" -ExportFormat "TXT"

# Export to CSV for Excel
.\Get-RemoteSystemInfo.ps1 -ComputerName "192.168.1.100" -ExportPath "C:\Reports" -ExportFormat "CSV"

# Export to HTML for easy viewing/sharing
.\Get-RemoteSystemInfo.ps1 -ComputerName "192.168.1.100" -ExportPath "C:\Reports" -ExportFormat "HTML"

# Export to JSON for automation
.\Get-RemoteSystemInfo.ps1 -ComputerName "192.168.1.100" -ExportPath "C:\Reports" -ExportFormat "JSON"


### REMOTE SYSTEM INFO 2 ###

fixed export functionality and made it guided
missing xml/html export
