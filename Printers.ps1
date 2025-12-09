
<# 
    HP UPD PCL6 – Hardened Download + Install + Add Printers
    ---------------------------------------------------------
    - Validates source URL via HEAD (detects 404/HTML/icon)
    - If source is a page, extracts first .zip/.exe asset link automatically
    - Downloads to C:\Temp, optional SHA-256, extracts ZIP/EXE
    - Stages driver via pnputil, registers driver, adds TCP/IP printers
    - Clear step-by-step diagnostics
#>

# ======= USER CONFIG =======
$DriverSourceUrl      = "https://github.com/YourOrg/hp-upd-driver/releases/download/v7.9.0/upd-pcl6-x64-7.9.0.26347.zip"  # asset or page URL
$ExpectedSha256       = "PASTE-YOUR-SHA256-HERE"   # or $null to skip
$PrinterIpMap         = @{
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

# Prefer TLS 1.2+
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

function Resolve-DriverAssetUrl {
    param([Parameter(Mandatory)][string]$Url)

    Write-Host "STEP A: Validating URL (HEAD) → $Url"
    try {
        # HEAD request to see what we get (status + content-type + disposition)
        $head = Invoke-WebRequest -Uri $Url -Method Head -ErrorAction Stop
        $ct   = $head.ContentType
        $disp = $head.Headers['Content-Disposition']
        $sc   = $head.StatusCode

        Write-Host ("HEAD OK: Status={0}  Content-Type={1}  Disposition={2}" -f $sc, $ct, $disp)

        # If it already looks like a file (zip/exe)
        if ($Url.ToLower().EndsWith(".zip") -or $Url.ToLower().EndsWith(".exe")) {
            return $Url
        }

        # If content type looks like a file (octet-stream/zip), trust this URL anyway
        if ($ct -and ($ct -match 'octet|zip|application/x-msdownload')) {
            return $Url
        }

        # If it’s HTML (typical HP/GitHub page), fetch and find the first .zip/.exe link
        if ($ct -match 'html' -or $Url -match '/drivers/' -or $Url -match '/releases/') {
            Write-Host "STEP B: Source is an HTML page—extracting asset links (.zip/.exe)"
            $page = Invoke-WebRequest -Uri $Url -UseBasicParsing
            $links = @()

            if ($page.Links) { $links += $page.Links | ForEach-Object { $_.href } }

            # Regex scrape in case Links collection misses script-built anchors
            $links += ([regex]::Matches($page.Content, 'href="([^"]+)"')).Groups |
                      Where-Object { $_.Value -like 'http*' } |
                      ForEach-Object { $_.Value }

            $asset = $links | Where-Object { $_ -match '\.(zip|exe)(\?.*)?$' } | Select-Object -First 1
            if ($asset) {
                Write-Host "Found asset: $asset"
                return $asset
            }

            throw "No .zip/.exe asset link found on page: $Url"
        }

        # Otherwise, treat as unsupported (e.g., SVG icon)
        throw "Unsupported URL or content-type ($ct). Supply a direct .zip/.exe or a page that lists the driver asset."
    }
    catch {
        $ex = $_.Exception
        $status = $null; $respUri = $null
        if ($ex -is [System.Net.WebException] -and $ex.Response) {
            $status  = [int]$ex.Response.StatusCode
            $respUri = $ex.Response.ResponseUri
        }
        $msg = "HEAD/resolve failed: $($ex.Message)  Status=$status  ResponseUri=$respUri  URL=$Url"
        throw [System.Exception]::new($msg, $ex)
    }
}

try {
    Write-Host "=== HP UPD PCL6 deployment starting ==="

    # STEP C: Ensure C:\Temp exists
    $downloadDir = 'C:\Temp'
    if (-not (Test-Path $downloadDir)) {
        New-Item -ItemType Directory -Path $downloadDir | Out-Null
    }

    # STEP D: Resolve to an actual ZIP/EXE asset
    $DriverDownloadUrl = Resolve-DriverAssetUrl -Url $DriverSourceUrl

    # STEP E: Download asset
    $fileName    = Split-Path $DriverDownloadUrl -Leaf
    $packagePath = Join-Path $downloadDir $fileName
    Write-Host "Downloading asset → $DriverDownloadUrl"
    try {
        $resp = Invoke-WebRequest -Uri $DriverDownloadUrl -OutFile $packagePath -ErrorAction Stop
        Write-Host "Download OK. Status=$($resp.StatusCode)  Saved=$packagePath"
    } catch {
        $ex = $_.Exception
        $status = $null; $respUri = $null
        if ($ex -is [System.Net.WebException] -and $ex.Response) {
            $status  = [int]$ex.Response.StatusCode
            $respUri = $ex.Response.ResponseUri
        }
        $msg = "Download failed: $($ex.Message)  Status=$status  ResponseUri=$respUri  URL=$DriverDownloadUrl"
        throw [System.Exception]::new($msg, $ex)
    }

    # STEP F: Optional SHA-256 validation
    if ($ExpectedSha256 -and $ExpectedSha256 -ne "") {
        $actualHash = (Get-FileHash -Algorithm SHA256 -Path $packagePath).Hash.ToLower()
        if ($actualHash -ne $ExpectedSha256.ToLower()) {
            throw [System.Security.SecurityException]::new("SHA256 mismatch. Expected $ExpectedSha256, got $actualHash")
        }
        Write-Host "Checksum OK."
    } else {
        Write-Host "Checksum validation skipped."
    }

    # STEP G: Extract ZIP/EXE
    $extractDir = Join-Path $downloadDir ("HP_UPD_" + [Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $extractDir | Out-Null

    $ext = [System.IO.Path]::GetExtension($packagePath).ToLowerInvariant()
    if ($ext -eq '.zip') {
        Write-Host "Expanding ZIP → $extractDir"
        Expand-Archive -Path $packagePath -DestinationPath $extractDir -Force
    } elseif ($ext -eq '.exe') {
        Write-Host "Attempting EXE-as-ZIP expansion …"
        $zipLike = $packagePath -replace '\.exe$', '.zip'
        Copy-Item $packagePath $zipLike -Force
        $expanded = $false
        try {
            Expand-Archive -Path $zipLike -DestinationPath $extractDir -Force
            $expanded = $true
        } catch {
            Write-Warning "Could not expand EXE as ZIP; trying silent extraction '/q'."
        }
        if (-not $expanded) {
            Start-Process -FilePath $packagePath -ArgumentList '/q' -Wait -NoNewWindow
            # If the EXE performs a silent full install, we’ll still detect the driver later.
        }
    } else {
        throw [System.IO.FileLoadException]::new("Unsupported asset extension: $ext")
    }

    # STEP H: Locate PCL6 INF
    Write-Host "Locating PCL6 INF under $extractDir …"
    $inf = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
           Where-Object { $_.Name -match '^hpcu.*u\.inf$' } |
           Select-Object -First 1
    if (-not $inf) {
        $inf = Get-ChildItem -Path $extractDir -Recurse -Filter *.inf |
               Where-Object { $_.Name -match 'hp.*pcl' } |
               Select-Object -First 1
    }
    if (-not $inf) {
        throw [System.IO.FileNotFoundException]::new("No suitable PCL6 INF found in $extractDir. Ensure the package is UPD PCL6 x64.")
    }
    Write-Host "INF: $($inf.FullName)"

    # STEP I: Stage driver via pnputil
    $pnputil = "$env:WINDIR\System32\pnputil.exe"
    $args    = "/add-driver `"$($inf.FullName)`" /install"
    Write-Host "Staging with pnputil $args …"
    $p = Start-Process -FilePath $pnputil -ArgumentList $args -Wait -PassThru
    if ($p.ExitCode -ne 0) {
        throw [System.Exception]::new("pnputil failed with exit code $($p.ExitCode)")
    }

    # STEP J: Register / confirm driver
    Import-Module PrintManagement -ErrorAction SilentlyContinue
    $candidate = Get-PrinterDriver -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -eq $DriverName -or $_.Name -like '*HP*Universal*PCL*6*' } |
                 Select-Object -First 1
    if ($candidate) {
        $DriverName = $candidate.Name
        Write-Host "Driver present: '$DriverName'"
    } else {
        Write-Host "Adding printer driver by name: '$DriverName'"
        Add-PrinterDriver -Name $DriverName -ErrorAction Stop
    }

    # STEP K: Add ports & printers
    foreach ($kv in $PrinterIpMap.GetEnumerator() | Sort-Object Key) {
        $num = [string]$kv.Key
        $ip  = [string]$kv.Value

        if (-not ($ip -match '^\d{1,3}(\.\d{1,3}){3}$')) {
            Write-Warning "Skipping '$num' → '$ip' (invalid IPv4)."
            continue
        }

        $portName    = "IP_$ip"
        $printerName = [string]::Format($PrinterNameFormat, $num, $ip)

        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Write-Host "Creating port $portName ($ip)"
            Add-PrinterPort -Name $portName -PrinterHostAddress $ip -ErrorAction Stop
        }
        if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
            Write-Host "Adding printer '$printerName'"
            Add-Printer -Name $printerName -PortName $portName -DriverName $DriverName -ErrorAction Stop
        } else {
            Write-Host "Printer '$printerName' already exists."
        }
    }

    # STEP L: Default printer
    if ($DefaultPrinterNumber -and $PrinterIpMap.ContainsKey($DefaultPrinterNumber)) {
        $defIp   = $PrinterIpMap[$DefaultPrinterNumber]
        $defName = [string]::Format($PrinterNameFormat, $DefaultPrinterNumber, $defIp)
        if (Get-Printer -Name $defName -ErrorAction SilentlyContinue) {
            Write-Host "Setting default printer to '$defName'"
            (Get-WmiObject -Query "Select * From Win32_Printer Where Name='$defName'").SetDefaultPrinter() | Out-Null
        } else {
            Write-Warning "Default printer '$defName' not found."
        }
    }

    Write-Host "=== Completed HP UPD PCL6 deployment ==="
}
catch {
    $ex = $_.Exception
    $detail = if ($ex.InnerException) { $ex.InnerException.Message } else { $null }
    Write-Error ("ERROR: {0}`nDETAIL: {1}`nSTACK: {2}" -f $ex.Message, $detail, $ex.StackTrace)
    throw
}
finally {
    Stop-Transcript | Out-Null
    Write-Host "Log written to: $logPath"
