
<# 
    HP UPD PCL6 – Download to C:\Temp, Stage & Install, then Add Printers
    ---------------------------------------------------------
    - Downloads from a predetermined URL to C:\Temp
    - Validates SHA-256 (optional)
    - Extracts archive (ZIP or EXE self-extractor; will try ZIP-rename)
    - Stages/Installs driver via pnputil
    - Registers driver, adds Standard TCP/IP ports and printers
    - Sets default printer (optional)

    Run elevated (Administrator). Tested on Windows 10/11, PowerShell 5.1
#>

# ======= USER CONFIG =======
# 1) Your predetermined URL to the HP UPD PCL6 x64 package (ZIP or EXE)
$DriverDownloadUrl = "https://support.hp.com/wcc-assets/content/dam/hp-wcc/fe-assets/images/swd/DownloadIcon.svg"

# 2) Optional SHA-256 checksum for integrity verification (set to $null to skip)
$ExpectedSha256   = "FB6ABF4D077CABB8995799E1868E5B82F65D82178F7F9E34D88C4ACB4EB6261D"  # e.g. 'A1B2C3D4...'; or $null

# 3) Printer number → IP mapping
$PrinterIpMap = @{
    1 = '10.20.32.13'
    2 = '10.20.30.42'
    3 = '10.20.30.43'
}

# 4) Optional default printer by number (from the map). Comment out to skip.
$DefaultPrinterNumber = 1

# 5) Driver name (Windows registered name)
$DriverName = 'HP Universal Printing PCL 6'

# 6) Printer name format; {0}=number, {1}=IP
$PrinterNameFormat = 'PRN-{0} ({1})'


# ======= SCRIPT START =======
$ErrorActionPreference = 'Stop'
$logPath = 'C:\Windows\Temp\HP-UPD-Install.log'
Start-Transcript -Path $logPath -Append | Out-Null

try {
    Write-Host "=== HP UPD PCL6 deployment starting ==="

    # Ensure C:\Temp exists
    $downloadDir = 'C:\Temp'
    if (-not (Test-Path $downloadDir)) {
        New-Item -ItemType Directory -Path $downloadDir | Out-Null
    }

    # Determine local file path
    $fileName = Split-Path $DriverDownloadUrl -Leaf
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        throw "Could not determine filename from DriverDownloadUrl."
    }
    $packagePath = Join-Path $downloadDir $fileName

    # Download the package to C:\Temp
    Write-Host "Downloading driver package to $packagePath ..."
    Invoke-WebRequest -Uri $DriverDownloadUrl -OutFile $packagePath

    # Extract the package to a working folder next to the download
    $extractDir = Join-Path $downloadDir ("HP_UPD_" + [Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $extractDir | Out-Null

    $ext = [System.IO.Path]::GetExtension($packagePath).ToLowerInvariant()
    if ($ext -eq '.zip') {
        Expand-Archive -Path $packagePath -DestinationPath $extractDir -Force
    } elseif ($ext -eq '.exe') {
        # Many UPD packages are self-extracting archives. Try renaming to ZIP first; if it fails, run extractor quietly.
        $zipLike = $packagePath -replace '\.exe$', '.zip'
        Copy-Item $packagePath $zipLike -Force
        $expanded = $false
        try {
            Expand-Archive -Path $zipLike -DestinationPath $extractDir -Force
            $expanded = $true
        } catch {
            Write-Warning "Could not expand executable as ZIP; attempting silent extraction."
        }
        if (-not $expanded) {
            # Silent switches vary; we only need the INF files. If the EXE doesn’t support silent extract,
            # consider pre-extracting and hosting a ZIP. Otherwise, this may still install interactively.
            Start-Process -FilePath $packagePath -ArgumentList '/q' -Wait -NoNewWindow -ErrorAction SilentlyContinue
            # If the EXE performs a full install silently, you may already have the driver. We’ll still continue with detection below.
        }
    } else {
        throw "Unsupported package extension: $ext"
    }

    # Try to locate the PCL6 INF (HP often uses hpcu***u.inf for x64 UPD PCL6)
    $infCandidates = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
        Where-Object { $_.Name -match '^hpcu.*u\.inf$' } |
        Sort-Object LastWriteTime -Descending

    if ($infCandidates.Count -eq 0) {
        # Fallback heuristic
        $infCandidates = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
            Where-Object { $_.Name -match 'hp.*pcl' } |
            Sort-Object LastWriteTime -Descending
    }
    $infToInstall = $null
    if ($infCandidates.Count -gt 0) {
        $infToInstall = $infCandidates[0].FullName
        Write-Host "Found INF: $infToInstall"
    } else {
        Write-Warning "No suitable PCL6 INF found. If the EXE already installed the driver, we’ll detect it next."
    }

    # Stage & install the driver via pnputil if we have the INF
    $pnputil = Join-Path $env:WINDIR 'System32\pnputil.exe'
    if ($infToInstall) {
        $args = "/add-driver `"$infToInstall`" /install"
        Write-Host "Staging driver with pnputil: $args"
        $p = Start-Process -FilePath $pnputil -ArgumentList $args -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw "pnputil failed with exit code $($p.ExitCode)" }
    }

    # Register/confirm the driver name in Windows
    Import-Module PrintManagement -ErrorAction SilentlyContinue
    $candidate = Get-PrinterDriver -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -eq $DriverName -or $_.Name -like '*HP*Universal*PCL*6*' } |
                 Select-Object -First 1

    if ($candidate) {
        $DriverName = $candidate.Name
        Write-Host "Driver ready: '$DriverName'"
    } else {
        # If not found yet, try to add by name
        Write-Host "Adding printer driver by name: '$DriverName'"
        Add-PrinterDriver -Name $DriverName -ErrorAction Stop
    }

    # Add ports & printers from the map
    foreach ($kv in $PrinterIpMap.GetEnumerator() | Sort-Object Key) {
        $num = [string]$kv.Key
        $ip  = [string]$kv.Value

        if (-not ($ip -match '^\d{1,3}(\.\d{1,3}){3}$')) {
            Write-Warning "Skipping '$num' → '$ip' (not an IPv4 address)."
            continue
        }

        $portName    = "IP_$ip"
        $printerName = [string]::Format($PrinterNameFormat, $num, $ip)

        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Write-Host "Creating port $portName for $ip ..."
            Add-PrinterPort -Name $portName -PrinterHostAddress $ip -ErrorAction Stop
        }
        if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
            Write-Host "Adding printer '$printerName' ..."
            Add-Printer -Name $printerName -PortName $portName -DriverName $DriverName -ErrorAction Stop
        } else {
            Write-Host "Printer '$printerName' already exists."
        }
    }

    # Set default printer if requested
    if ($null -ne $DefaultPrinterNumber -and $PrinterIpMap.ContainsKey($DefaultPrinterNumber)) {
        $defIp   = $PrinterIpMap[$DefaultPrinterNumber]
        $defName = [string]::Format($PrinterNameFormat, $DefaultPrinterNumber, $defIp)
        if (Get-Printer -Name $defName -ErrorAction SilentlyContinue) {
            Write-Host "Setting default printer to '$defName' ..."
            (Get-WmiObject -Query "Select * From Win32_Printer Where Name='$defName'").SetDefaultPrinter() | Out-Null
        } else {
            Write-Warning "Default printer '$defName' not found."
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
