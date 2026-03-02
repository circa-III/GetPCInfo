<#
.SYNOPSIS
  Get system, software, printers, monitors, OneDrive status, and AD info for one or more PCs.
  Script made by Brad Linder - blinder@ecommunity.com

.EXAMPLES
  .\GetPCInfo.ps1 -Computers 'NR-TCPC1,PC02' /t
  .\GetPCInfo.ps1 -Computers 'NR-TCPC1','PC02' -Transcript
  .\GetPCInfo.ps1
    (Interactive mode: enter hostnames, add /t in the prompt to save transcripts)

.NOTES
  Built for Windows PowerShell 5.1; PS7+ also works.
#>

[CmdletBinding()]
param(
    # Accepts a comma-separated string or an array: -Computers 'PC1,PC2' or -Computers 'PC1','PC2'
    [Parameter(Position=0)]
    [string[]]$Computers,

    # Normal PowerShell switch for transcript
    [Alias('t')]
    [switch]$Transcript,

    # Default transcript folder (auto-created on first use)
    [string]$TranscriptFolder = 'C:\Computer Reports'
)

$Version = "GetPCInfo | Version 26.03"

# Detect /t as a standalone token (PowerShell passes unbound tokens in $args)
$Transcript = $Transcript -or ($args | Where-Object { $_ -ieq '/t' } | ForEach-Object { $true } | Select-Object -First 1)

# ---------------- Formatting helpers ----------------
function Write-Section {
    param([Parameter(Mandatory)][string]$Title)
    Write-Host ""  # leading blank line between sections
    Write-Host ("-===[ {0} ]===-" -f $Title) -ForegroundColor Yellow
}

function Write-ComputerHeader {
    param([Parameter(Mandatory)][string]$CompName)
    Write-Host ""
    Write-Host ("================  {0}  ================" -f $CompName) -ForegroundColor Cyan
}

function Write-AfterTable {
    Write-Host ""  # uniform trailing blank line
}

# ---------------- Utility helpers ----------------
function Convert-UInt16ArrayToString {
    param([uint16[]]$arr)
    if (-not $arr) { return $null }
    ($arr | ForEach-Object { [char]$_ }) -join '' -replace '\u0000',''
}

function Get-VideoOutputTechName {
    param([uint32]$code)
    switch ($code) {
        4294967294 { 'Uninitialized' }                 # -2
        4294967295 { 'Other' }                         # -1
        0 { 'VGA/HD15' }
        1 { 'S-Video' }
        2 { 'Composite' }
        3 { 'Component' }
        4 { 'DVI' }
        5 { 'HDMI' }
        6 { 'LVDS (Internal Panel)' }
        7 { 'UDI' }
        8 { 'D-JPN' }
        9 { 'SDTV Dongle' }
        10 { 'DisplayPort' }
        11 { 'DisplayPort (Embedded)' }
        12 { 'UDI (Embedded)' }
        13 { 'UDI (External)' }
        14 { 'SDI' }
        15 { 'Virtual' }
        16 { 'DisplayPort over USB-C' }
        2147483648 { 'Internal (Embedded - vendor specific)' }  # 0x80000000
        default { "Code $code" }
    }
}

function Get-InternalPanelTag {
    param(
        [string] $Manufacturer,
        [string] $ConnName,
        [uint32] $ConnCode,
        [bool]   $IsPortableChassis = $false,
        [string] $ModelName = $null,
        [string] $Serial = $null
    )

    # Strong signals for internal panel
    $isEmbedded = ($ConnCode -eq 6) -or ($ConnCode -eq 11) -or ($ConnCode -eq 0x80000000) -or ($ConnName -match 'LVDS|Embedded')
    if ($isEmbedded) { return ' (Internal Panel)' }

    # Conservative fallback (on laptops only, weak EDID, not obvious external link)
    $isNameBlank = [string]::IsNullOrWhiteSpace($ModelName)
    $isSerialMissing = [string]::IsNullOrWhiteSpace($Serial) -or $Serial -eq '0' -or $Serial -eq '(n/a)'
    $likelyExternalConn = $ConnName -match 'HDMI|DisplayPort(?!.*Embedded)|DVI|USB-C'  # excludes Embedded DP
    if ($IsPortableChassis -and ($isNameBlank -or $isSerialMissing) -and -not $likelyExternalConn) {
        return ' (Internal Panel)'
    }
    return ''
}

function Convert-ODSyncProgressState {
    param([int64]$code)
    switch ($code) {
        0 { 'Up-to-date' }
        16777216 { 'Up-to-date' }
        65536 { 'Paused' }
        8194 { 'Not syncing' }
        1854 { 'Syncing problems' }
        default { "Unknown ($code)" }
    }
}

# ---------------- Core routine (runs once for a set of computers) ----------------
function Invoke-GetPCInfo {
    param(
        [string[]]$ComputerList,
        [switch]$ExportTranscript,
        [string]$OutFolder
    )

    # Normalize and clean computer names
    $ComputerList = $ComputerList | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if (-not $ComputerList -or $ComputerList.Count -eq 0) {
        Write-Host "No computer names provided."
        return
    }

    # Ensure transcript folder exists if needed
    if ($ExportTranscript) {
        try {
            if (-not (Test-Path -LiteralPath $OutFolder)) {
                New-Item -ItemType Directory -Path $OutFolder -Force | Out-Null
            }
        } catch {
            Write-Host "(!) Could not create transcript folder '$OutFolder': $($_.Exception.Message)"
            Write-Host "    Falling back to current directory."
            $OutFolder = (Get-Location).Path
        }
    }

    foreach ($comp in $ComputerList) {

        # Optional per-computer transcript
        $transcriptStarted = $false
        $transcriptPath = $null
        if ($ExportTranscript) {
            $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $transcriptPath = Join-Path $OutFolder ("GetPCInfo_{0}_{1}.txt" -f $comp, $stamp)
            try {
                Start-Transcript -Path $transcriptPath -Append | Out-Null
                $transcriptStarted = $true
            } catch {
                Write-Host "(!) Could not start transcript: $($_.Exception.Message)"
            }
        }

        try {
    Write-ComputerHeader -CompName $comp

    # --- Existence & reachability checks ---
    $adObj = $null
    $adQueryError = $null
    try {
        # Use Name filter (robust for regular hostnames)
        $adObj = Get-ADComputer -Filter "Name -eq '$comp'" -Properties Description, DistinguishedName -ErrorAction Stop
    } catch {
        $adQueryError = $_
    }

    $dnsRecord = $null
    try {
        $dnsRecord = Resolve-DnsName -Name $comp -ErrorAction Stop | Select-Object -First 1
    } catch { }

    $online = $false
    try {
        $online = Test-Connection -ComputerName $comp -Count 1 -Quiet
    } catch {
        $online = $false
    }

    # --- Report state & decide next action ---
    if (-not $adObj) {
        Write-Host "Result: Computer account '$comp' not found in Active Directory." -ForegroundColor Yellow
        if ($adQueryError) {
            Write-Host ("Note: AD query error: {0}" -f $adQueryError.Exception.Message) -ForegroundColor DarkYellow
        }
        if ($dnsRecord) {
            Write-Host ("DNS: Resolved to {0}" -f $dnsRecord.IPAddress) -ForegroundColor DarkYellow
        } else {
            Write-Host "Could not ping this computer name." -ForegroundColor DarkYellow
        }
        Write-Host ""
        continue
    }

    if (-not $dnsRecord) {
        Write-Host "Result: AD account exists, but the hostname does not resolve in DNS." -ForegroundColor Yellow
        Write-Host ("AD DN: {0}" -f $adObj.DistinguishedName) -ForegroundColor DarkYellow
        Write-Host ""
        continue
    }

    if (-not $online) {
        Write-Host "Result: AD and DNS OK, but the host is offline or unreachable (powered off, network issue, or firewall)." -ForegroundColor Yellow
        Write-Host ("AD DN: {0}" -f $adObj.DistinguishedName) -ForegroundColor DarkYellow
        Write-Host ("DNS:   {0}" -f $dnsRecord.IPAddress) -ForegroundColor DarkYellow
        Write-Host ""
        continue
    }

    # If we got here, it's online and we can collect normally
    Write-Host "Result: Online and reachable. Gathering information..." -ForegroundColor Green

    try {
        # -------- System Info --------
        Write-Section "System Info"
        Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $comp |
            Select-Object Name, Model, Username | Format-Table -AutoSize
        Get-WmiObject -Class Win32_OperatingSystem -ComputerName $comp |
            Select-Object Description | Format-Table -AutoSize
        Get-CimInstance Win32_BIOS -ComputerName $comp |
            Select-Object SerialNumber | Format-Table -AutoSize
        Get-WmiObject -Class Win32_NetworkAdapterConfiguration -ComputerName $comp -Filter 'IpEnabled=True' |
            Select-Object IPAddress, MACAddress | Format-Table -AutoSize
        Write-AfterTable

        # -------- Installed Software --------
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
        $softwareList = Invoke-Command -ComputerName $comp -ScriptBlock $scriptBlock |
                        Sort-Object DisplayName -Unique

        Write-Section "Installed Software"
        $softwareList | Format-Table DisplayName, Publisher -AutoSize
        Write-AfterTable

        # -------- Installed Printers --------
        Write-Section "Installed Printers"
        Get-CimInstance -ClassName Win32_Printer -ComputerName $comp |
            Where-Object { $_.Name -notlike 'WebEx*' -and $_.Name -notlike 'OneNote*' -and $_.Name -notlike 'Send*' -and $_.Name -notlike 'Microsoft*' -and $_.Name -notlike 'Fax' } |
            Select-Object Name, DriverName, Portname |
            Format-Table -AutoSize
        Write-AfterTable

        # -------- Monitors --------
        Write-Section "Monitors"
        try {
            $chassis = Get-CimInstance -Class Win32_SystemEnclosure -ComputerName $comp -ErrorAction SilentlyContinue
            $portableChassisTypes = 8,9,10,14
            $isPortable = $false
            if ($chassis -and $chassis.ChassisTypes) {
                $isPortable = ($chassis.ChassisTypes | ForEach-Object { [int]$_ }) |
                              Where-Object { $portableChassisTypes -contains $_ } |
                              ForEach-Object { $true } | Select-Object -First 1
                if (-not $isPortable) { $isPortable = $false }
            }

            $monId = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ComputerName $comp -ErrorAction Stop |
                     Where-Object { $_.Active -eq $true }
            $monConn = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorConnectionParams -ComputerName $comp -ErrorAction SilentlyContinue

            $results = foreach ($m in $monId) {
                $conn = $monConn | Where-Object { $_.InstanceName -eq $m.InstanceName } | Select-Object -First 1

                $name   = (Convert-UInt16ArrayToString $m.UserFriendlyName)
                $manu   = (Convert-UInt16ArrayToString $m.ManufacturerName)
                $serial = (Convert-UInt16ArrayToString $m.SerialNumberID)
                $year   = $m.YearOfManufacture

                if ([string]::IsNullOrWhiteSpace($name)) { $name = '(Unknown Model)' }
                if ([string]::IsNullOrWhiteSpace($serial) -or $serial -eq '0') { $serial = '(n/a)' }

                $connCode = if ($conn) { [uint32]$conn.VideoOutputTechnology } else { [uint32]0xFFFFFFFF }
                $connName = if ($conn) { Get-VideoOutputTechName -code $connCode } else { 'Unknown' }

                $tag = Get-InternalPanelTag -Manufacturer $manu -ConnName $connName -ConnCode $connCode -IsPortableChassis $isPortable -ModelName $name -Serial $serial
                $displayName = $name + $tag

                [PSCustomObject]@{
                    Name         = $displayName
                    Manufacturer = $manu
                    Serial       = $serial
                    Year         = $year
                    Connection   = $connName
                    Instance     = $m.InstanceName
                }
            }

            $activeCount = ($monId | Measure-Object).Count
            Write-Host ("Active Monitors: {0}" -f $activeCount)

            if ($results) {
                $results | Select-Object Name, Manufacturer, Serial, Year, Connection | Format-Table -AutoSize
            }
            else {
                Write-Host "No active monitors detected (this can happen over RDP). Trying fallback..."

                $pnpmons = Get-CimInstance -ClassName Win32_PnPEntity -ComputerName $comp -Filter "PNPClass = 'Monitor'" |
                           Where-Object { $_.ConfigManagerErrorCode -eq 0 }

                $fallbackCount = ($pnpmons | Measure-Object).Count
                Write-Host ("Present Monitors (PnP): {0}" -f $fallbackCount)

                if ($pnpmons) {
                    $pnpmons | Select-Object Name, Manufacturer, PNPDeviceID | Format-Table -AutoSize
                } else {
                    Write-Host "No monitors reported by PnP either."
                }
            }
        }
        catch {
            Write-Host "Unable to query monitor info: $_"
        }
        Write-AfterTable

       # -------- OneDrive Status --------
                    Write-Section "OneDrive Status"
                    try {
                        # 1) Interactive user + SID + profile
                        $sessionInfo = Invoke-Command -ComputerName $comp -ScriptBlock {
                            $cs = Get-CimInstance -Class Win32_ComputerSystem
                            $user = $cs.UserName

                            if (-not $user) {
                                $explorer = Get-Process -Name explorer -ErrorAction SilentlyContinue | Select-Object -First 1
                                if ($explorer) {
                                    $owner = (Get-CimInstance Win32_Process -Filter "ProcessId=$($explorer.Id)").GetOwner()
                                    if ($owner -and $owner.Domain -and $owner.User) { $user = "$($owner.Domain)\$($owner.User)" }
                                }
                            }

                            $sid = $null
                            if ($user) {
                                try {
                                    $nt = New-Object System.Security.Principal.NTAccount($user)
                                    $sid = ($nt.Translate([System.Security.Principal.SecurityIdentifier])).Value
                                } catch {}
                            }

                            $profile = $null
                            if ($sid) {
                                $profile = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue).ProfileImagePath
                            }

                            [PSCustomObject]@{ User=$user; Sid=$sid; Profile=$profile }
                        }

                        $hasInteractiveUser = $sessionInfo -and $sessionInfo.User -and $sessionInfo.Sid -and $sessionInfo.Profile
                        if (-not $hasInteractiveUser) {
                            Write-Host "No interactive user session detected (skipping OneDrive checks for this computer)."
                        }
                        else {
                            $user = $sessionInfo.User
                            $sid  = $sessionInfo.Sid
                            $prof = $sessionInfo.Profile

                            Write-Host ("Interactive User: {0}" -f $user)

                            # 2) HKU:\<SID>\Software\Microsoft\OneDrive\Accounts\Business1
                            $odKeyPath = "Registry::HKEY_USERS\$sid\Software\Microsoft\OneDrive\Accounts\Business1"
                            $odAccount = Invoke-Command -ComputerName $comp -ScriptBlock {
                                param($k)
                                if (Test-Path $k) {
                                    Get-ItemProperty -Path $k | Select-Object UserEmail, UserFolder, TenantName, DisplayName, LastSignInTime, LastCandidateUpdateTime
                                }
                            } -ArgumentList $odKeyPath

                            if (-not $odAccount) {
                                Write-Host "OneDrive for Business: Not configured for this user (no Business1 key)."
                            }
                            else {
                                $odEmail  = $odAccount.UserEmail
                                $odFolder = $odAccount.UserFolder
                                $tenant   = $odAccount.TenantName

                                $odEmailOut = if ([string]::IsNullOrWhiteSpace($odEmail)) { '(unknown)'} else { $odEmail }
                                $tenantOut  = if ([string]::IsNullOrWhiteSpace($tenant))  { '(unknown)'} else { $tenant }
                                $folderOut  = if ([string]::IsNullOrWhiteSpace($odFolder)){ '(unknown)'} else { $odFolder }

                                Write-Host ("Signed-in: {0} (Tenant: {1})" -f $odEmailOut, $tenantOut)
                                Write-Host ("Local Folder: {0}" -f $folderOut)

                                # 3) Logs: SyncDiagnostics.log (state + recent time)
                                $logInfo = Invoke-Command -ComputerName $comp -ScriptBlock {
                                    param($profilePath)
                                    $logFile = Join-Path $profilePath "AppData\Local\Microsoft\OneDrive\logs\Business1\SyncDiagnostics.log"
                                    if (Test-Path $logFile) {
                                        $lastWrite = (Get-Item $logFile).LastWriteTimeUtc
                                        $content = Get-Content -Path $logFile -Tail 400 -ErrorAction SilentlyContinue

                                        $stateLine = $content | Select-String -Pattern 'SyncProgressState' | Select-Object -Last 1
                                        $utcLine   = $content | Select-String -Pattern 'UtcNow:' | Select-Object -Last 1

                                        $stateCode = $null
                                        if ($stateLine) {
                                            $m = [regex]::Match($stateLine.Line, 'SyncProgressState\s*:\s*(\d+)')
                                            if ($m.Success) { $stateCode = [int64]$m.Groups[1].Value }
                                        }

                                        $utc = $null
                                        if ($utcLine) {
                                            $m2 = [regex]::Match($utcLine.Line, 'UtcNow:\s*([0-9T:\-\.Z]+)')
                                            if ($m2.Success) {
                                                [datetime]::Parse($m2.Groups[1].Value, $null, [System.Globalization.DateTimeStyles]::AssumeUniversal)
                                            }
                                        }

                                        [PSCustomObject]@{
                                            LogPath    = $logFile
                                            LastWrite  = $lastWrite
                                            UtcNow     = $utc
                                            StateCode  = $stateCode
                                        }
                                    }
                                } -ArgumentList $prof

                                if ($logInfo) {
                                    $stateCode = if ($null -ne $logInfo.StateCode) { [int64]$logInfo.StateCode } else { 0 }
                                    $state     = Convert-ODSyncProgressState -code $stateCode
                                    $when      = if ($null -ne $logInfo.UtcNow) { $logInfo.UtcNow } else { $logInfo.LastWrite }

                                    Write-Host ("Sync Status: {0}" -f $state)
                                    if ($when) {
                                        $local = [System.TimeZoneInfo]::ConvertTimeFromUtc([datetime]$when, [System.TimeZoneInfo]::Local)
                                        Write-Host ("Last Sync Activity: {0}" -f $local.ToString("yyyy-MM-dd HH:mm:ss"))
                                    } else {
                                        Write-Host "Last Sync Activity: (unknown)"
                                    }
                                }
                                else {
                                    Write-Host "No SyncDiagnostics.log found for this user. OneDrive may be idle, not started, or logs are cleared."
                                }
                            }
                        }
                    }
                    catch {
                        Write-Host "Unable to query OneDrive status: $_"
                    }
                    Write-AfterTable

        Write-AfterTable

        # -------- Active Directory --------
        Write-Section "Active Directory"
        try {
            # Reuse $adObj from the availability checks
            $dn = $adObj.DistinguishedName

            # Split DN safely (handles escaped commas)
            $parts = $dn -split ',(?=(?:[^\\]|\\.)*$)'

            # Extract only OUs → remove "OU=" and unescape '\,' → ','
            $ous = $parts |
                Where-Object { $_ -like 'OU=*' } |
                ForEach-Object { ($_ -replace '^OU=', '') -replace '\\,', ',' }

            # Take bottom-most 4 (leaf OU + 3 parents)
            $bottom4 = $ous | Select-Object -First 4

            $ouPath = if ($bottom4) { $bottom4 -join ' / ' } else { '<no OUs>' }

            [PSCustomObject]@{
                Computer              = $adObj.Name
                Description           = $adObj.Description
                'OUPath (bottom->up)' = $ouPath
            } | Format-Table -AutoSize
        }
        catch {
            Write-Host "Get-ADComputer not available or failed: $($_.Exception.Message)"
        }
        Write-AfterTable

        # -------- End-of-computer separator --------
        Write-Host ""
        Write-Host ("-" * 70) -ForegroundColor DarkGray
        Write-Host ""

    } catch {
        Write-Host "Error accessing $comp - $_"
    }
}
        finally {
            if ($ExportTranscript -and $transcriptStarted) {
                try {
                    Stop-Transcript | Out-Null
                    Write-Host "Transcript saved to: $transcriptPath"
                } catch {
                    Write-Host "(!) Could not stop transcript: $($_.Exception.Message)"
                }
            }
        }
    }
}

# ---------------- Entry behavior ----------------
if ($Computers -and $Computers.Count -gt 0) {
    # Non-interactive mode (parameters supplied)
    Invoke-GetPCInfo -ComputerList $Computers -ExportTranscript:$Transcript -OutFolder $TranscriptFolder
}
else {
    # Interactive mode (prompt once per run; empty input exits)
    while ($true) {
        Write-Host "========================="
        Write-Host $Version -ForegroundColor Cyan
        Write-Host "========================="
        Write-Host "Type computer names (comma-separated) and press ENTER."
        Write-Host "Add /t after the PC name to save a transcript text file (i.e. NR-ISPC1 /t)"
        Write-Host ""

        $raw = Read-Host -Prompt "Enter computer names"
        if ([string]::IsNullOrWhiteSpace($raw)) { break }

        # Tokenize on whitespace; detect /t token
        $tokens = $raw -split '\s+' | Where-Object { $_ -ne '' }
        $wantTranscript = $tokens | Where-Object { $_ -ieq '/t' } | ForEach-Object { $true } | Select-Object -First 1
        $remaining = $tokens | Where-Object { $_ -ine '/t' }
        $namesPart = ($remaining -join ' ')
        $list = ($namesPart -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }

        if (-not $list) {
            Write-Host "No computer names provided. Try again or press ENTER to exit."
            continue
        }

        Invoke-GetPCInfo -ComputerList $list -ExportTranscript:$wantTranscript -OutFolder $TranscriptFolder
    }
}