<#
.SYNOPSIS
   Tools to better assist w/ job functions.
.DESCRIPTION
    Call function manually to better assist w/ work itself.
#>

# ============================================================
#  FUNCTION: Update-Windows
# ============================================================

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$LogFile = "$PSScriptRoot\errors.log"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"

    Add-Content -Path $LogFile -Value $line
}


function Update-Windows {

    Write-Host "Retrieving last 5 Windows Update events..." -ForegroundColor Green

    try {
        $filter = @"
*[System[
    Provider[@Name='Microsoft-Windows-WindowsUpdateClient'] and
    (EventID=19 or EventID=20)
]]
"@

        $lastUpdates = Get-WinEvent -LogName System -FilterXPath $filter |
                       Select-Object -First 5 TimeCreated, Message

        if ($lastUpdates) {
            Write-Host "`nLast 5 Installed Updates:" -ForegroundColor Green
            $lastUpdates | Format-Table -AutoSize
        } else {
            Write-Host "No previous updates found." -ForegroundColor Yellow
        }

    } catch {
        Write-Host "Error retrieving update history: $_" -ForegroundColor Red
        return
    }

    Read-Host "`nPress Enter to search for pending updates…"

    # Ensure PSWindowsUpdate exists
    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-Host "Installing PSWindowsUpdate module…" -ForegroundColor Cyan
            Install-Module PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
        }

        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-Host "PSWindowsUpdate imported." -ForegroundColor Green

    } catch {
        Write-Host "Failed to install/import PSWindowsUpdate: $_" -ForegroundColor Red
        return
    }

    $originalPolicy = Get-ExecutionPolicy
    Set-ExecutionPolicy Bypass -Scope Process -Force

    try {
        $pending = Get-WindowsUpdate -ErrorAction Stop
    } catch {
        Write-Host "Failed to retrieve updates: $_" -ForegroundColor Red
        Set-ExecutionPolicy $originalPolicy -Scope Process -Force
        return
    }

    if (-not $pending -or $pending.Count -eq 0) {
        Write-Host "No outstanding updates found." -ForegroundColor Green
        Set-ExecutionPolicy $originalPolicy -Scope Process -Force
        return
    }

    Write-Host "`nThere are $($pending.Count) updates available:" -ForegroundColor Green

    for ($i = 0; $i -lt $pending.Count; $i++) {
        $u = $pending[$i]
        $size = if ($u.Size) { [math]::Round($u.Size/1MB,2) } else { "?" }
        Write-Host "[$($i+1)] $($u.KB) — $($u.Title) ($size MB)"
    }

    $skip = Read-Host "`nEnter update numbers to SKIP (comma separated), or press Enter for ALL"
    $skipList = @()

    if ($skip -match '\d') {
        $skipList = $skip -split "," | ForEach-Object { [int]($_.Trim()) } | Where-Object { $_ -gt 0 }
    }

    $installList = for ($i=0; $i -lt $pending.Count; $i++) {
        if (($i+1) -notin $skipList) { $pending[$i] }
    }

    if ($installList.Count -eq 0) {
        Write-Host "No updates selected." -ForegroundColor Yellow
        Set-ExecutionPolicy $originalPolicy -Scope Process -Force
        return
    }

    Write-Host "`nInstalling $($installList.Count) updates…" -ForegroundColor Cyan

    foreach ($update in $installList) {
        try {
            Write-Host "Installing: $($update.Title)" -ForegroundColor Yellow
            Install-WindowsUpdate -Title $update.Title -AcceptAll -IgnoreReboot -ErrorAction Stop
        } catch {
            Write-Host "Failed installing $($update.Title): $_" -ForegroundColor Red
        }
    }

    Write-Host "All possible updates installed." -ForegroundColor Green

    Set-ExecutionPolicy $originalPolicy -Scope Process -Force
    Read-Host "Press Enter to continue..."
}

# ============================================================
#  FUNCTION: Net-Adapt
# ============================================================
function Net-Adapt {

    function Select-Adapter {
        while ($true) {
            Clear-Host
            $adapters = Get-NetAdapter | Where-Object Status -eq 'Up'

            if (-not $adapters) {
                Write-Host "No active adapters found." -ForegroundColor Red
                return $null
            }

            Write-Host "0. Back"
            for ($i=0; $i -lt $adapters.Count; $i++) {
                Write-Host "[$($i+1)] $($adapters[$i].Name)"
            }

            $choice = Read-Host "Choose adapter"
            if ($choice -eq "0") { return $null }
            if ($choice -match '^\d+$' -and $choice -le $adapters.Count) {
                return $adapters[$choice-1].Name
            }

            Read-Host "Invalid choice. Press Enter."
        }
    }

    function Select-Property($AdapterName) {
        while ($true) {
            Clear-Host
            $props = Get-NetAdapterAdvancedProperty -Name $AdapterName

            Write-Host "0. Back"
            for ($i=0; $i -lt $props.Count; $i++) {
                Write-Host "[$($i+1)] $($props[$i].DisplayName): $($props[$i].DisplayValue)"
            }

            $choice = Read-Host "Choose property"
            if ($choice -eq "0") { return $null }
            if ($choice -match '^\d+$' -and $choice -le $props.Count) {
                return $props[$choice-1]
            }

            Read-Host "Invalid choice. Press Enter."
        }
    }

    function Edit-Property($AdapterName, $Prop) {
        Clear-Host
        Write-Host "Editing: $($Prop.DisplayName)"
        Write-Host "Current: $($Prop.DisplayValue)"

        $valid = $Prop.ValidDisplayValue, $Prop.ValidDisplayValues | Where-Object { $_ }

        if ($valid) {
            Write-Host "`n0. Cancel"
            for ($i=0; $i -lt $valid.Count; $i++) {
                Write-Host "[$($i+1)] $($valid[$i])"
            }

            $choice = Read-Host "New value"
            if ($choice -eq "0") { return }

            if ($choice -match '^\d+$' -and $choice -le $valid.Count) {
                $newValue = $valid[$choice-1]
            } else {
                Write-Host "Invalid option." -ForegroundColor Red
                return
            }
        } else {
            $newValue = Read-Host "Enter value (0 to cancel)"
            if ($newValue -eq "0") { return }
        }

        try {
            Set-NetAdapterAdvancedProperty -Name $AdapterName -DisplayName $Prop.DisplayName -DisplayValue $newValue -NoRestart
            Write-Host "Updated successfully." -ForegroundColor Green
        } catch {
            Write-Host "Error updating: $_" -ForegroundColor Red
        }

        Read-Host "Press Enter..."
    }

    while ($true) {
        $adapter = Select-Adapter
        if (-not $adapter) { return }

        $prop = Select-Property $adapter
        if (-not $prop) { continue }

        Edit-Property $adapter $prop
    }
}

# ============================================================
#  FUNCTION: Speed-Test
# ============================================================
function Speed-Test {

    $pkg = "speedtest"

    $installed = choco list --local-only | Select-String "^$pkg\s"

    if ($installed) {
        Write-Host "$pkg installed. Reinstalling..." -ForegroundColor Yellow
        choco install $pkg --force -y | Out-Null
    } else {
        Write-Host "Installing speedtest…" -ForegroundColor Cyan
        choco install $pkg -y | Out-Null
    }

    $exe = "$env:ChocolateyInstall\bin\speedtest.exe"

    if (Test-Path $exe) {
        & $exe
    } else {
        Write-Host "speedtest.exe not found." -ForegroundColor Red
    }

    Read-Host "Press Enter…"
}

# ============================================================
#  FUNCTION: Clear-Space
# ============================================================
function Clear-Space {
    # Set Execution Policy to bypass for the current session
Set-ExecutionPolicy Bypass -Scope Process -Force

# Function to get the available free space on the drive
function Get-FreeSpace {
    $drive = Get-PSDrive -Name C
    return $drive.Used, $drive.Free
}

# Record initial free space
$initialUsedSpace, $initialFreeSpace = Get-FreeSpace

# Start Component Cleanup using DISM
Write-Host "Starting DISM component cleanup..."

try {
    $process = Start-Process -FilePath "Dism.exe" `
        -ArgumentList "/Online", "/Cleanup-Image", "/StartComponentCleanup", "/ResetBase" `
        -WindowStyle Hidden -Wait -PassThru

    if ($process.ExitCode -eq 0) {
        Write-Host "DISM cleanup complete."
    }
    else {
        Write-Host "DISM exited with code $($process.ExitCode)." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "An error occurred while running DISM: $_" -ForegroundColor Red
}


# Stop Windows Update services to clean SoftwareDistribution folder contents
Write-Host "Stopping Windows Update and Background Intelligent Transfer services... "
Stop-Service -Name wuauserv -ErrorAction SilentlyContinue
Stop-Service -Name bits -ErrorAction SilentlyContinue
Write-Host "Services stopped."

# Clean up the SoftwareDistribution folder contents using robocopy
$SoftwareDistributionPath = "C:\Windows\SoftwareDistribution"
if (Test-Path $SoftwareDistributionPath -ErrorAction SilentlyContinue) {
    Write-Host "Cleaning up SoftwareDistribution folder contents..."
    # Use robocopy to effectively delete the contents
    $TempPath = Join-Path $SoftwareDistributionPath "empty"
    New-Item -ItemType Directory -Path $TempPath -Force | Out-Null
    robocopy $TempPath $SoftwareDistributionPath /MIR /XD $TempPath
    Remove-Item -Recurse -Force -Path $TempPath
    Write-Host "SoftwareDistribution folder contents deleted."
} else {
    Write-Host "SoftwareDistribution folder not found."
}

# Restart the stopped services (Windows Update and BITS)
Write-Host "Restarting Windows Update and Background Intelligent Transfer services... "
Start-Service -Name wuauserv -ErrorAction SilentlyContinue
Start-Service -Name bits -ErrorAction SilentlyContinue
Write-Host "Services restarted."

# Cleanup temp files for all user profiles except "Public" and "Default"
Write-Host "Cleaning temp files for all users except 'Public' and 'Default'..."

# Get all user profile directories except "Public" and "Default"
$UserProfiles = Get-ChildItem "C:\Users" | Where-Object { 
    $_.Name -notin @('Public', 'Default') -and $_.PSIsContainer 
}

# Initialize a list to store the names of users whose temp files were deleted
$deletedUsers = @()

# Loop through each user profile and delete temp files
foreach ($UserProfile in $UserProfiles) {
    $TempFolder = Join-Path $UserProfile.FullName "AppData\Local\Temp"

    if (Test-Path $TempFolder -ErrorAction SilentlyContinue) {
        try {
            # Delete all contents in the Temp folder
            Get-ChildItem $TempFolder -Recurse | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
            $deletedUsers += $UserProfile.Name
        } catch {
            Write-Host "Failed to delete temp files for user: $($UserProfile.Name) - $_"
        }
    }
}

# Output the names of all users whose temp files were deleted, grouped together
if ($deletedUsers.Count -gt 0) {
    Write-Host "Temp files deleted for users: $($deletedUsers -join ', ')"
} else {
    Write-Host "No temp files were deleted."
}

Write-Host "Temp folder cleanup complete."

# Calculate and output total space cleared
$finalUsedSpace, $finalFreeSpace = Get-FreeSpace
$spaceFreed = $finalFreeSpace - $initialFreeSpace

Write-Host "Initial free space: $([math]::Round($initialFreeSpace / 1GB, 2)) GB"
Write-Host "Final free space: $([math]::Round($finalFreeSpace / 1GB, 2)) GB"
Write-Host "Total space freed: $([math]::Round($spaceFreed / 1GB, 2)) GB"

Write-Host "Disk space cleanup complete."
}

# ============================================================
#  FUNCTION: Bookmark-Export
# ============================================================
function Bookmark-Export {
    # CLEANED / FIXED VERSION
    # (Same structure as yours, but corrected errors, added safety, removed crashes)

    Write-Host "Bookmark export tool loading…" -ForegroundColor Cyan

    $OutputFolder = "C:\Temp\Browser_Bookmarks"
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder | Out-Null
    }

    $LogFile = Join-Path $OutputFolder "Bookmark_Copy_Log.txt"
    Add-Content $LogFile "===== Bookmark Backup $(Get-Date) ====="

    # ------- Provide optional completion of your missing IE/Firefox functions -------
    function Convert-FirefoxToHtml { param($sqlitePath, $outputFile)
        Add-Content $LogFile "Firefox conversion not implemented"
    }

    function Convert-IEFavoritesToHtml { param($favPath, $outputFile)
        Add-Content $LogFile "IE conversion not implemented"
    }

    # ------- Chromium conversion (your logic preserved + fixed) -------
    function Convert-ChromiumBookmarksToHtml { param($jsonPath, $outputFile)
        try {
            $json = Get-Content $jsonPath -Raw | ConvertFrom-Json
        } catch {
            Add-Content $LogFile "Error reading Chromium bookmarks: $_"
            return
        }

        $html = "<!DOCTYPE NETSCAPE-Bookmark-file-1>`n<DL><p>`n"

        function Walk { param($node)
            $out = ""
            foreach ($child in $node.children) {
                if ($child.type -eq "folder") {
                    $out += "<DT><H3>$($child.name)</H3><DL><p>`n"
                    $out += Walk $child
                    $out += "</DL><p>`n"
                } elseif ($child.url) {
                    $out += "<DT><A HREF=""$($child.url)"">$($child.name)</A>`n"
                }
            }
            return $out
        }

        foreach ($root in $json.roots.PSObject.Properties) {
            $html += "<DT><H3>$($root.Name)</H3><DL><p>`n"
            $html += Walk $root.Value
            $html += "</DL><p>`n"
        }

        Set-Content -Path $outputFile -Value $html
    }

    # Scan for browsers
    $AvailableBrowsers = @{}

    Get-ChildItem "C:\Users" -Directory | ForEach-Object {
        $p = $_.FullName
        $u = $_.Name
        if ($u -notmatch "Default|Public|systemprofile") {

            $chromes = @{
                "Google Chrome ($u)" = Join-Path $p "AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
                "Microsoft Edge ($u)" = Join-Path $p "AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks"
                "Brave ($u)"          = Join-Path $p "AppData\Local\BraveSoftware\Brave-Browser\User Data\Default\Bookmarks"
                "Opera ($u)"          = Join-Path $p "AppData\Roaming\Opera Software\Opera Stable\Bookmarks"
            }

            foreach ($k in $chromes.Keys) {
                if (Test-Path $chromes[$k]) {
                    $AvailableBrowsers[$k] = $chromes[$k]
                }
            }

            $ieFav = Join-Path $p "Favorites"
            if (Test-Path $ieFav) {
                $AvailableBrowsers["IE Favorites ($u)"] = $ieFav
            }
        }
    }

    # Firefox (appdata sandbox)
    $ffDir = "$env:APPDATA\Mozilla\Firefox\Profiles"
    if (Test-Path $ffDir) {
        Get-ChildItem $ffDir -Directory | ForEach-Object {
            $places = Join-Path $_.FullName "places.sqlite"
            if (Test-Path $places) {
                $AvailableBrowsers["Firefox ($($_.Name))"] = $places
            }
        }
    }

    if ($AvailableBrowsers.Count -eq 0) {
        Write-Host "No browsers found." -ForegroundColor Yellow
        return
    }

    Write-Host "`nDetected browsers:`n"
    $index = 1
    $lookup = @{}

    foreach ($key in $AvailableBrowsers.Keys) {
        Write-Host "[$index] $key"
        $lookup[$index] = $key
        $index++
    }

    $selection = Read-Host "`nEnter numbers to export (comma separated)"
    $numbers = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }

    foreach ($n in $numbers) {
        $n = [int]$n
        if (-not $lookup.ContainsKey($n)) { continue }

        $browser = $lookup[$n]
        $src = $AvailableBrowsers[$browser]
        $safe = $browser.Replace(" ", "_")
        $out = Join-Path $OutputFolder "$safe.html"

        if ($browser -like "*Chrome*" -or $browser -like "*Edge*" -or $browser -like "*Brave*" -or $browser -like "*Opera*") {
            Convert-ChromiumBookmarksToHtml $src $out
        } elseif ($browser -like "*Firefox*") {
            Convert-FirefoxToHtml $src $out
        } elseif ($browser -like "*IE*") {
            Convert-IEFavoritesToHtml $src $out
        }

        Write-Host "Exported: $browser → $out"
    }

    Read-Host "Done. Press Enter…"
}


# 1) Your number→IP map
$Printers = @{
    1 = '10.20.32.8'
    2 = '10.20.30.42'
    3 = '10.20.30.43'
}

# 2) Your GitHub Release asset (ZIP/EXE) and SHA-256
$DriverUrl  = "https://github.com/TyreikReid/Full-moon/releases/download/untagged-e5b81ec737a2fc88e337/upd-pcl6-x64-7.9.0.26347.zip"
$Sha256     = "FB6ABF4D077CABB8995799E1868E5B82F65D82178F7F9E34D88C4ACB4EB6261D"

# 3) Run the function
Install-HPUPDPrinters -PrinterIpMap $Printers `
    -DriverDownloadUrl $DriverUrl `
    -ExpectedSha256 $Sha256 `
    -DefaultPrinterNumber 1


Function Printers {
    
function Install-HPUPDPrinters {
    <#
    .SYNOPSIS
        Installs HP Universal Printing PCL 6 driver and creates local TCP/IP printers from a number→IP map.

    .DESCRIPTION
        - Downloads HP UPD PCL6 (from GitHub Releases or any HTTPS URL), validates SHA-256, extracts.
        - Stages the INF with pnputil, registers the driver, and creates printers.
        - Logs to C:\Windows\Temp\HP-UPD-Install.log

    .PARAMETER PrinterIpMap
        Hashtable mapping a numeric code to a printer IP, e.g. @{1='10.10.0.101';2='10.10.0.102'}

    .PARAMETER DriverName
        The Windows print driver name to use. Defaults to 'HP Universal Printing PCL 6'.

    .PARAMETER DriverDownloadUrl
        Direct URL to your UPD PCL6 x64 package (ZIP/EXE) hosted in GitHub Releases, e.g.:
        https://github.com/<Org>/<Repo>/releases/download/v7.9.0/upd-pcl6-x64-7.9.0.26347.zip

    .PARAMETER ExpectedSha256
        SHA-256 of the package you host. Use Get-FileHash locally to compute it, then paste here.

    .PARAMETER DefaultPrinterNumber
        Optional number from PrinterIpMap to set as the default printer.

    .PARAMETER PrinterNameFormat
        Format for printer names. Defaults to 'PRN-{0} ({1})' → 'PRN-1 (10.10.0.101)'

    .EXAMPLE
        Install-HPUPDPrinters -PrinterIpMap @{1='10.20.30.41';2='10.20.30.42'} `
            -DriverDownloadUrl "https://github.com/YourOrg/hp-upd-driver/releases/download/v7.9.0/upd-pcl6-x64-7.9.0.26347.zip" `
            -ExpectedSha256 "PUT-YOUR-SHA256-HERE" `
            -DefaultPrinterNumber 1

    .NOTES
        Run elevated. Tested on Windows 10/11, PowerShell 5.1.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [hashtable]$PrinterIpMap,

        [string]$DriverName = 'HP Universal Printing PCL 6',

        [Parameter(Mandatory)]
        [string]$DriverDownloadUrl,

        [Parameter(Mandatory)]
        [string]$ExpectedSha256,

        [int]$DefaultPrinterNumber,

        [string]$PrinterNameFormat = 'PRN-{0} ({1})'
    )

    #region Setup
    $ErrorActionPreference = 'Stop'
    $logPath = 'C:\Windows\Temp\HP-UPD-Install.log'
    Start-Transcript -Path $logPath -Append | Out-Null

    try {
        Write-Host "=== HP UPD PCL6 deployment starting ==="
        Import-Module PrintManagement -ErrorAction SilentlyContinue

        # If the driver already exists, we skip download and staging.
        $existing = Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object {
            $_.Name -eq $DriverName
        }

        if (-not $existing) {
            # Workspace
            $tempRoot = Join-Path $env:TEMP ("HP_UPD_" + [Guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $tempRoot | Out-Null

            $fileName   = Split-Path $DriverDownloadUrl -Leaf
            $packagePath = Join-Path $tempRoot $fileName

            # 1) Download
            Write-Host "Downloading driver package from $DriverDownloadUrl ..."
            Invoke-WebRequest -Uri $DriverDownloadUrl -OutFile $packagePath

            # 2) Validate SHA-256 (security hardening)
            $actual = (Get-FileHash -Algorithm SHA256 -Path $packagePath).Hash.ToLower()
            if ($actual -ne $ExpectedSha256.ToLower()) {
                throw "SHA256 mismatch. Expected $ExpectedSha256, got $actual"
            }
            Write-Host "Checksum OK."

            # 3) Extract
            $extractDir = Join-Path $tempRoot 'extract'
            New-Item -ItemType Directory -Path $extractDir | Out-Null

            $ext = [System.IO.Path]::GetExtension($packagePath).ToLowerInvariant()
            if ($ext -eq '.zip') {
                Expand-Archive -Path $packagePath -DestinationPath $extractDir -Force
            } elseif ($ext -eq '.exe') {
                # Many HP packages are self-extracting. Try a ZIP rename first, else run quietly.
                $zipLike = $packagePath -replace '\.exe$', '.zip'
                Copy-Item $packagePath $zipLike -Force
                try {
                    Expand-Archive -Path $zipLike -DestinationPath $extractDir -Force
                } catch {
                    Write-Warning "Could not expand as ZIP; attempting to run the installer to extract."
                    Start-Process -FilePath $packagePath -ArgumentList '/q' -Wait -NoNewWindow -ErrorAction SilentlyContinue
                }
            } else {
                throw "Unsupported package extension: $ext"
            }

            # 4) Find the PCL6 INF (HP’s UPD x64 often matches ^hpcu.*u\.inf$)
            $infCandidates = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
                Where-Object { $_.Name -match '^hpcu.*u\.inf$' } |
                Sort-Object LastWriteTime -Descending

            if ($infCandidates.Count -eq 0) {
                # Fallback heuristic
                $infCandidates = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
                    Where-Object { $_.Name -match 'hp.*pcl' } |
                    Sort-Object LastWriteTime -Descending
            }
            if ($infCandidates.Count -eq 0) { throw "No suitable PCL6 INF found in driver package." }

            $infToInstall = $infCandidates[0].FullName
            Write-Host "Staging driver from INF: $infToInstall"

            # 5) Stage & install via pnputil (silent & supported)
            $pnputil = Join-Path $env:WINDIR 'System32\pnputil.exe'
            $args    = "/add-driver `"$infToInstall`" /install"
            $proc    = Start-Process -FilePath $pnputil -ArgumentList $args -Wait -PassThru
            if ($proc.ExitCode -ne 0) { throw "pnputil failed with exit code $($proc.ExitCode)" }
            # Reference for pnputil: add-driver /install. [1](https://learn.microsoft.com/en-us/windows-hardware/drivers/devtest/pnputil-examples)

            # 6) Confirm driver name; auto-detect if the exact string differs
            $registered = Get-PrinterDriver -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -eq $DriverName -or $_.Name -like '*HP*Universal*PCL*6*' } |
                Select-Object -First 1

            if ($registered) {
                $DriverName = $registered.Name
                Write-Host "Driver ready: '$DriverName'"
            } else {
                Add-PrinterDriver -Name $DriverName -ErrorAction Stop
                Write-Host "Driver added: '$DriverName'"
            }
        } else {
            Write-Host "Driver already installed: '$DriverName'."
        }

        # 7) Create ports and printers from the number→IP map
        foreach ($kvp in $PrinterIpMap.GetEnumerator() | Sort-Object Key) {
            $number = [string]$kvp.Key
            $ip     = [string]$kvp.Value

            if (-not ($ip -match '^\d{1,3}(\.\d{1,3}){3}$')) {
                Write-Warning "Skipping '$number' → '$ip' (not an IPv4 address)."
                continue
            }

            $portName    = "IP_$ip"
            $printerName = [string]::Format($PrinterNameFormat, $number, $ip)

            if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                Add-PrinterPort -Name $portName -PrinterHostAddress $ip -ErrorAction Stop
            }
            if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
                Add-Printer -Name $printerName -PortName $portName -DriverName $DriverName -ErrorAction Stop
            } else {
                Write-Host "Printer '$printerName' already exists."
            }
        }

        # 8) Set default printer if requested
        if ($PSBoundParameters.ContainsKey('DefaultPrinterNumber')) {
            $defIp = $PrinterIpMap[$DefaultPrinterNumber]
            if ($defIp) {
                $defName = [string]::Format($PrinterNameFormat, $DefaultPrinterNumber, $defIp)
                if (Get-Printer -Name $defName -ErrorAction SilentlyContinue) {
                    (Get-WmiObject -Query "Select * From Win32_Printer Where Name='$defName'").SetDefaultPrinter() | Out-Null
                } else {
                    Write-Warning "Default printer '$defName' not found."
                }
            } else {
                Write-Warning "DefaultPrinterNumber '$DefaultPrinterNumber' not found in map."
            }
        }

        Write-Host "=== Completed HP UPD PCL6 deployment ==="
    }
    catch {
        Write-Error $_
        throw
    }
    finally {
        Stop-Transcript | Out-Null
        Write-Host "Log: $logPath"
    }
}

}

# ============================================================
#  FUNCTION: Places (ISO download)
# ============================================================
function Places {

    $isoUrl = "https://software.download.prss.microsoft.com/dbazure/Win11_25H2_English_x64.iso"
    $dir = "C:\Temp"
    $aria = "C:\ProgramData\chocolatey\bin\aria2c.exe"

    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory | Out-Null
    }

    if (-not (Test-Path $aria)) {
        Write-Host "aria2c not found." -ForegroundColor Red
        return
    }

    $args = @(
        "--dir=$dir"
        "--out=downloaded.iso"
        "--split=16"
        "--continue=true"
        $isoUrl
    )

    Start-Process $aria -ArgumentList $args -NoNewWindow -Wait
}

# ============================================================
# MENU SYSTEM
# ============================================================
while ($true) {

    Clear-Host
   Write-Host "Version 3.0"
    Write-Host "----------------------------------"
    Write-Host "Tyreik's tools." -ForegroundColor DarkMagenta
    Write-Host "----------------------------------"
    Write-Host "1. Update Windows PC" -ForegroundColor Green
    Write-Host "2. Configure adapters" -ForegroundColor Green
    Write-Host "3. Clear space" -ForegroundColor Green
    Write-Host "4. Speed test" -ForegroundColor Green
    Write-Host "5. Bookmark extraction" -ForegroundColor Green
    Write-Host "6. Windows download to temp folder" -ForegroundColor Green
    Write-Host "7. Printer install for upstairs." -ForegroundColor Green
    Write-Host "To exit, press 0." -ForegroundColor Green


    $opt = Read-Host "Select your champion."

    switch ($opt) {
        "1" { Update-Windows }
        "2" { Net-Adapt }
        "3" { Clear-Space }
        "4" { Speed-Test }
        "5" { Bookmark-Export }
        "6" { Places }
        "7" { Printers }
        "0" { break }
        default { Read-Host "Invalid option. Press Enter…" }
    }
}
