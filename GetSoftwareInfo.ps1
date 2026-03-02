powershell -NoExit {
$Computers = Get-Content "C:/Users/A28033/Documents/PC_LIST.TXT" 
ForEach ($comp in $computers)
{
Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $comp | Select -Property Name,Model,USername | ft -autosize -wrap 
Get-CimInstance Win32_BIOS -ComputerName $comp | Select -Property SerialNumber | ft -auto
Get-CimInstance -ClassName Win32_Printer -ComputerName $comp | Select-Object Name,DriverName,Portname | Where-Object -FilterScript { ($_.Name -notlike 'WebEx*' -and $_.Name -notlike 'Send*' -and $_.Name -notlike 'Microsoft*' -and $_.Name -notlike 'Cortex*' -and $_.Name -notlike 'Fax' ) } | ft -autosize -wrap 
Get-CimInstance -ClassName Win32_Product -ComputerName $comp | Select -Property Name,Vendor | Where-Object -FilterScript { ($_.Vendor -notlike 'Microsoft*' -and $_.Name -notlike 'Microsoft*' -and $_.Vendor -notlike 'Citrix*' -and $_.Vendor -notlike 'PrinterLogic' -and $_.Vendor -notlike 'McAfee*' -and $_.Vendor -notlike 'Advanced Micro Devices, Inc.*' -and $_.Vendor -notlike 'Intel*' -and $_.Vendor -notlike 'Hewlett-Packard*' -and $_.Vendor -notlike 'Sun*' -and $_.Vendor -notlike 'Cisco*' -and $_.Vendor -notlike 'Husdawg*' -and $_.Vendor -notlike 'Oracle*' -and $_.Vendor -notlike 'HP*' -and $_.Vendor -notlike 'HP Inc.*' -and $_.Name -notlike 'Google Update Helper*' -and $_.Name -notlike 'Adobe Acrobat Reader DC*' -and $_.Name -notlike 'Adobe Refresh Manager*' -and $_.Name -notlike 'Google Toolbar*' -and $_.Name -notlike 'Adobe Reader X*' -and $_.Name -notlike 'McAfee Drive Encryption' -and $_.Name -notlike 'Adobe Flash Player 21 ActiveX' -and $_.Name -notlike 'Adobe Refresh Manager' -and $_.Name -notlike 'Adobe Acrobat Reader DC' -and $_.Name -notlike 'LWS*' -and $_.Name -notlike 'CameraHelperMsi*' -and $_.Name -notlike 'Alcor*' -and $_.Name -notlike 'Amtel*'-and $_.Name -notlike 'Visual C*') } | ft -autosize -wrap  
Get-ADComputer -Identity $Comp -Properties Description | ft -a DistinguishedName,Description
} 
}