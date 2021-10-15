$filename = $env:COMPUTERNAME

#Create folder with name of PC
New-Item -Path .\ -Name $filename -ItemType "directory"

#Launch hardware script
.\hardware-os-info.ps1 > .\$filename\$filename-hw-os-info.txt

#Launch software script
.\software-info.ps1 > .\$filename\$filename-sw-info.txt