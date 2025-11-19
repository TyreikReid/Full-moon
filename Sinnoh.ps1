<#
.SYNOPSIS
   Tools to better assist w/ job functions.
.DESCRIPTION
    Call function manually to better assist w/ work itself.
#>

# ============================================================
#  FUNCTION: Update-Windows
# ============================================================
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
    Set-ExecutionPolicy Bypass -Scope Process -Force
    Write-Host "TODO: Add your cleanup logic here." -ForegroundColor Yellow
    Read-Host "Press Enter…"
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
    Write-Host "System Tools 2.0"
    Write-Host "-----------------------------------------"
    Write-Host "1. Update Windows"
    Write-Host "2. Configure Adapters"
    Write-Host "3. Clear Space"
    Write-Host "4. Speed Test"
    Write-Host "5. Bookmark Extraction"
    Write-Host "6. Download Windows ISO"
    Write-Host "0. Exit"
    Write-Host "-----------------------------------------"

    $opt = Read-Host "Choose your fighter"

    switch ($opt) {
        "1" { Update-Windows }
        "2" { Net-Adapt }
        "3" { Clear-Space }
        "4" { Speed-Test }
        "5" { Bookmark-Export }
        "6" { Places }
        "0" { break }
        default { Read-Host "Invalid option. Press Enter…" }
    }
}

