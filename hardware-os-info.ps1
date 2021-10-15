###########################################
#  _   ___        __  ___        __       
# | | | \ \      / / |_ _|_ __  / _| ___  
# | |_| |\ \ /\ / /   | || '_ \| |_ / _ \ 
# |  _  | \ V  V /    | || | | |  _| (_) |
# |_| |_|  \_/\_/    |___|_| |_|_|  \___/ 
#     _    ____  _   _ 
#    / \  |  _ \| | | |
#   / _ \ | | | | | | |
#  / ___ \| |_| | |_| |
# /_/   \_\____/ \___/                                          
#
##########################################

# Shows hardware and OS details
# A.DUPOUY 19/09/2021

# Nettoyage de la console
Clear-Host

#Recuperation des infos dans les variables
$computerSystem = get-wmiobject Win32_ComputerSystem
$computerBIOS = get-wmiobject Win32_BIOS
$computerOS = get-wmiobject Win32_OperatingSystem
$computerCPU = get-wmiobject Win32_Processor
$computerHDD = Get-WmiObject Win32_LogicalDisk -Filter drivetype=3
$computerScreen = Get-WmiObject -Class Win32_VideoController


#Output
"System Information for: " + $computerSystem.Name
""
### GENERALITES PC ###
"== Generalites PC =="
"Manufacturer: " + $computerSystem.Manufacturer
"Model: " + $computerSystem.Model
"Serial Number: " + $computerBIOS.SerialNumber
"BIOS Version: " + $computerBIOS.SMBIOSBIOSVersion

### CPU ###
""
"== CPU =="
"CPU: " + $computerCPU.Name
"Number of CPUs: " + $computerSystem.NumberOfProcessors
"Nbr of logical processors: " + $computerSystem.NumberOfLogicalProcessors

### DISKS ####
""
"== Disks =="
foreach ($disk in $computerHDD) {
    "Drive Letter: " + $disk.DeviceID
    "Drive Capacity: "  + "{0:N2}" -f ($disk.Size/1GB) + "GB"
    "Drive Free Space: " + "{0:P2}" -f ($disk.FreeSpace/$disk.Size) + " Free (" + "{0:N2}" -f ($disk.FreeSpace/1GB) + "GB)"
    "--"       
}

### RAM ###
""
"== RAM =="
"RAM Installed: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"

## Current User ##
""
"== Current session =="
"Domain member: " + $computerSystem.PartOfDomain
"Domain: " + $computerSystem.Domain
"Current user: " + $computerSystem.UserName

### OS ###
""
"== OS =="
"Operating System: " + $computerOS.caption + " " + $computerOS.OSArchitecture
"Windows Build number: " + $computerOS.BuildNumber
"OS Install date: " + $computerOS.ConvertToDateTime($computerOS.InstallDate)
"Last Reboot: " + $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)

### Screen resolution ###
""
"== Screen resolution =="
Add-Type -AssemblyName System.Windows.Forms
$screen_cnt  = [System.Windows.Forms.Screen]::AllScreens.Count
$col_screens = [system.windows.forms.screen]::AllScreens

$info_screens = ($col_screens | ForEach-Object {
if ("$($_.Primary)" -eq "True") {$monitor_type = "Primary Monitor    "} else {$monitor_type = "Secondary Monitor  "}
if ("$($_.Bounds.Width)" -gt "$($_.Bounds.Height)") {$monitor_orientation = "Landscape"} else {$monitor_orientation = "Portrait"}
$monitor_type + "(Bounds)                          " + "$($_.Bounds)"
$monitor_type + "(Primary)                         " + "$($_.Primary)"
$monitor_type + "(Device Name)                     " + "$($_.DeviceName)"
$monitor_type + "(Bounds Width x Bounds Height)    " + "$($_.Bounds.Width) x $($_.Bounds.Height) ($monitor_orientation)"
"--"
}
)
Write-Host "TOTAL SCREEN COUNT: $screen_cnt"
$info_screens

### Report time&date ###
""
"-------------------------------------------------------"
"Report generated: " + (Get-Date)
"-------------------------------------------------------"

#DEBUG#
#Get the full list of properties
#Get-WmiObject -Class "Win32_operatingsystem" | Format-List *