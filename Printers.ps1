
<# 
    HP UPD PCL6 Deployment Script
    -----------------------------------------
    - Downloads driver from a pre-determined URL to C:\Temp
    - Validates SHA-256 (optional)
    - Rejects non-driver URLs (.svg, etc.)
    - Extracts ZIP or EXE (tries ZIP rename for EXE)
    - Stages driver via pnputil, registers with Add-PrinterDriver
    - Adds printers from a number→IP map
    - Sets default printer (optional)
#>

# ======= USER CONFIG =======
$DriverDownloadUrl = "https://github.com/YourOrg/hp-upd-driver/releases/download/v7.9.0/upd-pcl6-x64-7.9.0.26347.zip"
$ExpectedSha256    = "PASTE-YOUR-SHA256-HERE"   # Set to $null to skip validation
$PrinterIpMap      = @{
    1 = '10.20.30.41'
    2 = '10.20.30.42'
    3 = '10.20.30.43'
}
$DefaultPrinterNumber = 1
$DriverName           = 'HP Universal Printing PCL 6'
$PrinterNameFormat    = 'PRN-{0} ({1})'

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

    # Validate URL extension
    $ext = [System.IO.Path]::GetExtension($DriverDownloadUrl).ToLowerInvariant()
    if ($ext -notin @('.zip', '.exe')) {
        throw "Unsupported package extension: $ext. The URL likely points to an icon (.svg) or placeholder, not the driver asset."
    }

    # Download driver
    $fileName    = Split-Path $DriverDownloadUrl -Leaf
    $packagePath = Join-Path $downloadDir $fileName
    Write-Host "Downloading driver package to $packagePath ..."
    Invoke-WebRequest -Uri $DriverDownloadUrl -OutFile $packagePath

    # Validate SHA-256 if provided
    if ($ExpectedSha256 -and $ExpectedSha256 -ne "") {
        $actualHash = (Get-FileHash -Algorithm SHA256 -Path $packagePath).Hash.ToLower()
        if ($actualHash -ne $ExpectedSha256.ToLower()) {
            throw "SHA256 mismatch. Expected $ExpectedSha256, got $actualHash"
        }
        Write-Host "Checksum OK."
    } else {
        Write-Host "Checksum validation skipped."
    }

    # Extract driver
    $extractDir = Join-Path $downloadDir ("HP_UPD_" + [Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $extractDir | Out-Null

    if ($ext -eq '.zip') {
        Expand-Archive -Path $packagePath -DestinationPath $extractDir -Force
    } elseif ($ext -eq '.exe') {
        $zipLike = $packagePath -replace '\.exe$', '.zip'
        Copy-Item $packagePath $zipLike -Force
        try {
            Expand-Archive -Path $zipLike -DestinationPath $extractDir -Force
        } catch {
            Write-Warning "Could not expand EXE as ZIP; attempting silent extraction."
            Start-Process -FilePath $packagePath -ArgumentList '/q' -Wait -NoNewWindow
        }
    }

    # Locate INF for PCL6
    $inf = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
           Where-Object { $_.Name -match '^hpcu.*u\.inf$' } |
           Select-Object -First 1
    if (-not $inf) {
        $inf = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
               Where-Object { $_.Name -match 'hp.*pcl' } |
               Select-Object -First 1
    }
    if (-not $inf) { throw "No suitable PCL6 INF found." }

    # Stage driver via pnputil
    $pnputil = "$env:WINDIR\System32\pnputil.exe"
    Start-Process -FilePath $pnputil -ArgumentList "/add-driver `"$($inf.FullName)`" /install" -Wait

    # Register driver
    Import-Module PrintManagement
    $candidate = Get-PrinterDriver -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -eq $DriverName -or $_.Name -like '*HP*Universal*PCL*6*' } |
                 Select-Object -First 1
    if ($candidate) { $DriverName = $candidate.Name } else { Add-PrinterDriver -Name $DriverName }

    # Add printers
    foreach ($kv in $PrinterIpMap.GetEnumerator() | Sort-Object Key) {
        $num = $kv.Key; $ip = $kv.Value
        $portName    = "IP_$ip"
        $printerName = [string]::Format($PrinterNameFormat, $num, $ip)
        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Add-PrinterPort -Name $portName -PrinterHostAddress $ip
        }
        if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
            Add-Printer -Name $printerName -PortName $portName -DriverName $DriverName
        }
    }

    # Set default printer
    if ($DefaultPrinterNumber -and $PrinterIpMap.ContainsKey($DefaultPrinterNumber)) {
        $defIp   = $PrinterIpMap[$DefaultPrinterNumber]
        $defName = [string]::Format($PrinterNameFormat, $DefaultPrinterNumber, $defIp)
        if (Get-Printer -Name $defName -ErrorAction SilentlyContinue) {
            (Get-WmiObject -Query "Select * From Win32_Printer Where Name='$defName'").SetDefaultPrinter() | Out-Null
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
