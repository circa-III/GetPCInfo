$Version = "GetPCInfo | Version 26.02"
# Script made by Brad Linder - blinder@ecommunity.com
$continueSearching = $true

while ($continueSearching) {

    
Write-Host "========================="
    Write-Host $Version -ForegroundColor Cyan
    Write-Host "========================="

    # Prompt the user to enter computer names using a text box popup
    $Computers = Read-Host -Prompt "Enter computer names (separated by commas)"

    # Split the input into an array of computer names
    $Computers = $Computers -split ','

    ForEach ($comp in $Computers) {
        # Check if the computer is reachable
        if (Test-Connection -ComputerName $comp -Count 1 -Quiet) {
            try {
                Write-Host "Gathering information for $comp..."
               
                # System Information
                Write-Host "-===[ System Info ]===-"
                Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $comp | Select-Object -Property Name,Model,Username | Format-Table -AutoSize
                Get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp | Select-Object -Property Description | Format-Table -AutoSize
                Get-CimInstance Win32_BIOS -ComputerName $comp | Select-Object -Property SerialNumber | Format-Table -AutoSize
                Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $comp -Filter 'IpEnabled=True' | Select-Object -Property IPAddress, MACAddress | Format-Table -AutoSize

                # Define script block within the loop
                $scriptBlock = {
                    $softwareKeys = @(
                        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
                        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                        'HKCU:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                    )

                    foreach ($key in $softwareKeys) {
                        if (Test-Path $key) {
                            Get-ItemProperty -Path $key |
                            Where-Object {
                                $_.DisplayName -and (
                                    ($_.Publisher -notmatch 'Microsoft*' -or
                                    $_.DisplayName -match '365' -or
                                    $_.DisplayName -match 'Teams') -and
                                    $_.DisplayName -notlike 'BloxOne*' -and
                                    $_.DisplayName -notlike 'Cortex*' -and
                                    $_.DisplayName -notlike 'Bomgar Button*' -and
                                    $_.DisplayName -notlike '*Acrobat Reader*' -and
                                    $_.DisplayName -notlike '*Refresh Manager*' -and
                                    $_.DisplayName -notlike '*Genuine Service*' -and
                                    $_.DisplayName -notlike '*Webex*' -and
                                    $_.DisplayName -notlike '*Update Helper*' -and
                                    $_.Publisher -notlike 'PrinterLogic*' -and
                                    $_.DisplayName -notlike '*Windows Driver*' -and
                                    $_.DisplayName -notlike '*Realtek*' -and
                                    $_.Publisher -notlike '*Cisco*' -and
                                    $_.DisplayName -notlike 'Microsoft Edge*' -and
                                    $_.Publisher -notlike 'Midmark Diagnostics*' -and
                                    $_.Publisher -notlike 'Intel*' -and
                                    $_.Publisher -notlike 'Epic*' -and
                                    $_.DisplayName -notlike '*Vulkan*' -and
                                    $_.Publisher -notlike 'Conexant*' -and
                                    $_.Publisher -notlike 'HP*' -and
                                    $_.DisplayName -notlike 'Update for*' -and
					$_.DisplayName -notlike '*Support Button*' -and
					$_.DisplayName -notlike 'PaperCut*' -and
					$_.Publisher -notlike 'ANCILE*' -and
                                    $_.Publisher -notlike '*Citrix*'
                                )
                            } |
                            Select-Object DisplayName, Publisher
                        }
                    }
                }

                # Execute the script block remotely
                $softwareList = Invoke-Command -ComputerName $comp -ScriptBlock $scriptBlock | Sort-Object DisplayName -Unique | Format-Table -AutoSize
                Write-Host "-===[ Installed Software ]===-"
                $softwareList

                # Printers
                Write-Host "-===[ Installed Printers ]===-"
                Get-CimInstance -ClassName Win32_Printer -ComputerName $comp | 
                Where-Object { $_.Name -notlike 'WebEx*' -and $_.Name -notlike 'OneNote*' -and $_.Name -notlike 'Send*' -and $_.Name -notlike 'Microsoft*' -and $_.Name -notlike 'Fax' } | 
                Select-Object Name, DriverName, Portname | 
                Format-Table -AutoSize

                # Active Directory Computer Information
                Write-Host "-==[ Active Directory ]==-"
                Get-ADComputer -Identity $comp -Properties Description | 
                Select-Object DistinguishedName, Description | 
                Format-Table -AutoSize

            } catch {
                Write-Host "Error accessing $comp - $_"
            }
        } else {
            Write-Host "Computer $comp is offline or unreachable."
        }
    }
}