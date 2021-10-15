###########################################
#  ______        __  ___        __       
# / ___\ \      / / |_ _|_ __  / _| ___  
# \___ \\ \ /\ / /   | || '_ \| |_ / _ \ 
#  ___) |\ V  V /    | || | | |  _| (_) |
# |____/  \_/\_/    |___|_| |_|_|  \___/                                    
#     _    ____  _   _ 
#    / \  |  _ \| | | |
#   / _ \ | | | | | | |
#  / ___ \| |_| | |_| |
# /_/   \_\____/ \___/ 
#                      
###########################################

# Shows SOFTWARE
# A.DUPOUY 19/09/2021

# Nettoyage de la console
Clear-Host

#Recuperation des infos dans les variables
$office = New-Object -ComObject Excel.Application
$AntivirusProduct = Get-WmiObject -Namespace "root\SecurityCenter2" -Query "SELECT * FROM AntiVirusProduct"  @psboundparameters      
$allsoftsWOW64 = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, Publisher | Format-Table -AutoSize
$allsoftsWOW = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, Publisher | Format-Table -AutoSize

#Output
""
"== Microsoft Office =="
"Version: " + $office.version
"Path: " + $office.path
""
"--- Licence Office ---"
cd $office.path
cd ..\..
cd Office*
cscript ospp.vbs /dstatus

""
"== Antivirus =="
foreach($avproduct in $AntivirusProduct){
"Product name: "+ $avproduct.displayName
"Product executable: " + $avproduct.pathToSignedProductExe
"Product reporting exe: " + $avproduct.pathToSignedReportingExe
"--"
}

""
"== All softwares 64bit Node =="
$allsoftsWOW64

""
"== All softwares 32bit Node =="
$allsoftsWOW

### Report time&date ###
""
"Report generated: " + (Get-Date)
"-------------------------------------------------------"