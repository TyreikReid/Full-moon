
# =========================
# HP UPD PCL6 – Download → Validate → Install → Add Printers
# =========================

# ======= USER CONFIG =======
$DriverSourceUrl      = "https://support.hp.com/wcc-assets/content/dam/hp-wcc/fe-assets/images/swd/DownloadIcon.svg"  # asset or page URL
$ExpectedSha256       = "FB6ABF4D077CABB8995799E1868E5B82F65D82178F7F9E34D88C4ACB4EB6261D"   # or $null to skip
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

# Optional: quick elevation guard
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

Start-Transcript -Path $logPath -Append | Out-Null

# Prefer TLS 1.2+
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

function Resolve-DriverAssetUrl {
    param([Parameter(Mandatory)][string]$Url)

    Write-Host "STEP A: Validating URL (HEAD) -> $Url"
    try {
        # HEAD request to see status + content-type + disposition
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

        # If it's HTML (typical HP/GitHub page), fetch and find the first .zip/.exe link
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
    Write-Host "Resolved asset URL: $DriverDownloadUrl"

    # STEP E: Download asset
    $fileName    = Split-Path $DriverDownloadUrl -Leaf
    $packagePath = Join-Path $downloadDir $fileName
    Write-Host "Downloading asset -> $DriverDownloadUrl"
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

    # Guard: refuse obvious non-assets early
    if ($ext -notin @('.zip', '.exe')) {
        throw [System.IO.FileLoadException]::new("Unsupported asset extension: $ext. Expected .zip or .exe.")
    }

    if ($ext -eq '.zip') {
        Write-Host "Expanding ZIP -> $extractDir"
        Expand-Archive -Path $packagePath -DestinationPath $extractDir -Force
    } else {
        # EXE case: try zip rename first, fall back to silent extraction
        Write-Host "Attempting EXE-as-ZIP expansion ..."
        $zipLike = $packagePath -replace '\.exe$', '.zip'
        Copy-Item $packagePath $zipLike -Force
        $expanded = $false
        try {
            Expand-Archive -Path $zipLike -DestinationPath $extractDir -Force
            $expanded = $true
            Write-Host "EXE expanded as ZIP successfully."
        } catch {
            Write-Warning "Could not expand EXE as ZIP; trying silent extraction '/q'."
        }
        if (-not $expanded) {
            Start-Process -FilePath $packagePath -ArgumentList '/q' -Wait -NoNewWindow
            Write-Host "EXE executed with '/q'. If it performed a silent full install, driver detection will follow."
        }
    }

    # STEP H: Locate PCL6 INF
    Write-Host "Locating PCL6 INF under $extractDir ..."
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
    $pnputil = Join-Path $env:WINDIR 'System32\pnputil.exe'
    $infPath = $inf.FullName
    Write-Host "Staging with pnputil: /add-driver `"$infPath`" /install ..."
    & $pnputil /add-driver "$infPath" /install
    if ($LASTEXITCODE -ne 0) {
        throw [System.Exception]::new("pnputil failed with exit code $LASTEXITCODE while staging '$infPath'")
    }

    # STEP J: Register / confirm driver
    Import-Module PrintManagement -ErrorAction SilentlyContinue
    $candidate = Get-PrinterDriver -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -eq $DriverName -or $_.Name -like '*HP*Universal*PCL*6*' } |
                 Select-Object -First 1

    if ($candidate) {
        $DriverName = $candidate.Name
        Write-Host "Driver present and will be used as: '$DriverName'"
    } else {
        Write-Host "Adding printer driver by name: '$DriverName'"
        try {
            Add-PrinterDriver -Name $DriverName -ErrorAction Stop
            Write-Host "Driver added: '$DriverName'"
        } catch {
            throw [System.Exception]::new("Add-PrinterDriver failed. Name tried: '$DriverName'. Consider inspecting INF to confirm exact display name.", $_.Exception)
        }
    }

    # STEP K: Add ports & printers
    foreach ($kv in $PrinterIpMap.GetEnumerator() | Sort-Object Key) {
        $num = [string]$kv.Key
        $ip  = [string]$kv.Value

        if (-not ($ip -match '^\d{1,3}(\.\d{1,3}){3}$')) {
            Write-Warning "Skipping '$num' -> '$ip' (invalid IPv4)."
            continue
        }

        $portName    = "IP_$ip"
        $printerName = [string]::Format($PrinterNameFormat, $num, $ip)

        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
            Write-Host "Creating port $portName ($ip)"
            Add-PrinterPort -Name $portName -PrinterHostAddress $ip -ErrorAction Stop
        } else {
            Write-Host "Port '$portName' already exists."
        }

        if (-not (Get-Printer -Name $printerName -ErrorAction SilentlyContinue)) {
            Write-Host "Adding printer '$printerName' with driver '$DriverName'"
            Add-Printer -Name $printerName -PortName $portName -DriverName $DriverName -ErrorAction Stop
        } else {
            Write-Host "Printer '$printerName' already exists."
        }
    }

    # STEP L: Default printer
    if ($DefaultPrinterNumber -and $PrinterIpMap.ContainsKey($DefaultPrinterNumber)) {
        $defIp   = $PrinterIpMap[$DefaultPrinterNumber]
        $defName = [string]::Format($PrinterNameFormat, $DefaultPrinterNumber, $defIp)

        $defPrinter = Get-Printer -Name $defName -ErrorAction SilentlyContinue
        if ($defPrinter) {
            Write-Host "Setting default printer to '$defName'"
            $null = (Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Name='$defName'").SetDefaultPrinter()
        } else {
            Write-Warning "Default printer '$defName' not found; skipping."
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
}
