<#
.SYNOPSIS
   Tools to better assist w/ job functions.

.DESCRIPTION
    Call function manually to better assist w/ work itself.
#>

function Update-Windows {
    # Part 1: Show last 5 installed updates
    Write-Host "Retrieving the last 5 Windows Update events..." -ForegroundColor Green
    try {
        $lastUpdates = Get-WinEvent -LogName System -FilterXPath '
            *[System[Provider[@Name="Microsoft-Windows-WindowsUpdateClient"] and
            (EventID=19 or EventID=20)]]' | 
            Select-Object -First 5 -Property TimeCreated, Message

        if ($lastUpdates) {
            Write-Host "`nLast 5 Installed Updates:" -ForegroundColor Green
            $lastUpdates | Format-Table TimeCreated, Message -AutoSize
        } else {
            Write-Host "No previous updates found in the system logs." -ForegroundColor Red
        }
    } catch {
        Write-Host "Error: Failed to retrieve update history." -ForegroundColor Red
        exit
    }

    # Wait for user input before proceeding
    Write-Host "`nPress Enter to search for pending updates..." -ForegroundColor Green
    Read-Host

    # Part 2: Search for and handle pending updates
    Write-Host "Checking for pending updates..." -ForegroundColor Green

    # Ensure PSWindowsUpdate is installed
    try {
        if (!(Get-Module -ListAvailable PSWindowsUpdate)) {
            Install-Module PSWindowsUpdate -Force -Scope CurrentUser -ErrorAction Stop
            Write-Host "PSWindowsUpdate module installed successfully." -ForegroundColor Green
        }
        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-Host "PSWindowsUpdate module imported successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error: Failed to install or import PSWindowsUpdate module." -ForegroundColor Red
        exit
    }

    # Temporarily set execution policy to Bypass for this session
    $originalPolicy = Get-ExecutionPolicy
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

    # Get outstanding updates
    try {
        $outstanding = Get-WindowsUpdate -ErrorAction Stop

        if ($outstanding.Count -eq 0) {
            Write-Host "No outstanding updates found." -ForegroundColor Green
        } else {
            Write-Host "`nThere are $($outstanding.Count) outstanding updates available." -ForegroundColor Green

            # Display numbered list
            Write-Host "`nAvailable Updates:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $outstanding.Count; $i++) {
                $u = $outstanding[$i]
                $num = $i + 1
                Write-Host "[$num] $($u.KB) - $($u.Title) ($([math]::Round($u.Size/1MB,2)) MB)"
            }

            # Ask which to skip
            Write-Host "`nEnter update numbers to SKIP (comma-separated), or press Enter to install all:" -ForegroundColor Cyan
            $skipInput = Read-Host
            $toInstall = if ([string]::IsNullOrWhiteSpace($skipInput)) {
                $outstanding
            } else {
                $skipNums = $skipInput -split '\s*,\s*' |
                            Where-Object { $_ -match '^[0-9]+$' } |
                            ForEach-Object { [int]$_ }
                $outstanding | Where-Object { ($outstanding.IndexOf($_) + 1) -notin $skipNums }
            }

            if ($toInstall.Count -eq 0) {
                Write-Host "No updates selected for installation." -ForegroundColor Yellow
            } else {
                Write-Host "`nInstalling $($toInstall.Count) updates..." -ForegroundColor Green
                try {
                    foreach ($update in $toInstall) {
                        try {
                            Write-Host "Installing: $($update.Title)" -ForegroundColor Cyan
                            Install-WindowsUpdate -Title $update.Title -AcceptAll -IgnoreReboot -ErrorAction Stop
                        } catch {
                            Write-Host "Error installing $($update.Title): $_" -ForegroundColor Red
                        }
                    }
                    Write-Host "Updates installed successfully." -ForegroundColor Green
                } catch {
                    Write-Host "Error: Failed to install updates. $_" -ForegroundColor Red
                }
            }
        }
    } catch {
        Write-Host "Error: Failed to retrieve updates. $_" -ForegroundColor Red
    }

    # Reset execution policy to the original setting
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy $originalPolicy -Force
    Write-Host "Execution policy reset to original settings." -ForegroundColor Green
    Read-Host "Press enter to continue."
}

function Net-Adapt {

    function Show-Adapters {
        do {
            Clear-Host
            $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
            if (-not $adapters) {
                Write-Error "No active network adapters found."
                return
            }

            Write-Host "`nAvailable Network Adapters:" -ForegroundColor Cyan
            Write-Host "0. Exit"
            $adapters | ForEach-Object -Begin { $i = 1 } -Process {
                Write-Host "$i. $($_.Name)"
                $i++
            }

            $adapterChoice = Read-Host "`nSelect adapter number"
            if ($adapterChoice -eq '0') { return }

            $selectedAdapter = $adapters[$adapterChoice - 1]
            if ($selectedAdapter) {
                Show-Properties -AdapterName $selectedAdapter.Name
            } else {
                Write-Host "❌ Invalid selection. Press Enter to try again." -ForegroundColor Red
                Read-Host
            }

        } while ($true)
    }

    function Show-Properties {
        param ([string]$AdapterName)
        do {
            Clear-Host
            $props = Get-NetAdapterAdvancedProperty -Name $AdapterName
            if (-not $props) {
                Write-Error "No advanced properties found for $AdapterName"
                return
            }

            Write-Host "`nAdvanced Properties for adapter: $AdapterName" -ForegroundColor Cyan
            Write-Host "0. Back"
            $props | ForEach-Object -Begin { $j = 1 } -Process {
                Write-Host "$j. $($_.DisplayName): $($_.DisplayValue)"
                $j++
            }

            $propChoice = Read-Host "`nSelect property number to modify"
            if ($propChoice -eq '0') { return }

            $selectedProp = $props[$propChoice - 1]
            if ($selectedProp) {
                Show-PropertyEditor -AdapterName $AdapterName -Prop $selectedProp
            } else {
                Write-Host "❌ Invalid selection. Press Enter to try again." -ForegroundColor Red
                Read-Host
            }

        } while ($true)
    }

    function Show-PropertyEditor {
        param (
            [string]$AdapterName,
            $Prop
        )

        do {
            Clear-Host
            Write-Host "`nModify Property: $($Prop.DisplayName)" -ForegroundColor Yellow
            Write-Host "Current Value: $($Prop.DisplayValue)"
            Write-Host ""

            $validValues = @()
            if ($Prop.PSObject.Properties.Name -contains 'ValidDisplayValue') {
                $validValues = $Prop.ValidDisplayValue
            } elseif ($Prop.PSObject.Properties.Name -contains 'ValidDisplayValues') {
                $validValues = $Prop.ValidDisplayValues
            }

            if ($validValues.Count -gt 0) {
                Write-Host "0. Back"
                $validValues | ForEach-Object -Begin { $k = 1 } -Process {
                    Write-Host "$k. $_"
                    $k++
                }

                $valChoice = Read-Host "`nSelect new value number"
                if ($valChoice -eq '0') { return }
                $newValue = $validValues[$valChoice - 1]
            } else {
                $newValue = Read-Host "`nNo predefined values found. Enter new value manually (or 0 to cancel)"
                if ($newValue -eq '0') { return }
            }

            try {
                Set-NetAdapterAdvancedProperty -Name $AdapterName `
                    -DisplayName $Prop.DisplayName `
                    -DisplayValue $newValue -NoRestart
                Write-Host "`n✅ Property updated successfully!" -ForegroundColor Green
            } catch {
                Write-Error "❌ Failed to update property: $_"
            }

            Read-Host "`nPress Enter to return"
            return

        } while ($true)
    }

    Show-Adapters
}

function Speed-Test {
    $PackageName = "speedtest"

    $PackageCheck = choco list --local-only | Select-String "^$PackageName\s"

    if ($PackageCheck) {
        Write-Host "✔ $PackageName is already installed. Reinstalling..."
        choco install $PackageName --force -y > $null 2>&1
    } else {
        Write-Host "ℹ $PackageName not found. Installing..."
        choco install $PackageName -y > $null 2>&1
    }

    $ExePath = "$env:ChocolateyInstall\bin\speedtest.exe"

    if (Test-Path $ExePath) {
        Write-Host "`n=== 25 down, 3 up ==="
        & $ExePath
        Write-Host "`nPress Enter to continue..."
        Read-Host
    } else {
        Write-Host "ERROR: speedtest.exe not found at $ExePath"
    }
}

function Clear-Space {
    # Set Execution Policy to bypass for the current session
    Set-ExecutionPolicy Bypass -Scope Process -Force
    ...
}

function Bookmark-Export {
     # --------------------
    # Menu-driven browser bookmark exporter for SYSTEM perspective
    # --------------------

    # Output setup
# Output setup (replace the previous $OutputFolder definition)
   # Output setup
    $OutputFolder = "C:\Temp\Browser_Bookmarks"
    if (-not (Test-Path $OutputFolder)) { 
        New-Item -ItemType Directory -Path $OutputFolder | Out-Null 
    }

    $LogFile = Join-Path $OutputFolder "Bookmark_Copy_Log.txt"
    "===== Bookmark Backup Run: $(Get-Date -Format u) =====" | 
        Out-File -FilePath $LogFile -Append


    # --------------------
    # HTML header
    function Start-BookmarkHtml { param($title = "Bookmarks")
        return @"
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!-- Automatically generated -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>$title</TITLE>
<H1>$title</H1>
<DL><p>
"@
    }

    function Convert-ChromeTimestampToUnix { param([string]$chromeTs)
        if (-not $chromeTs) { return "" }
        try {
            $n = [int64]$chromeTs
            if ($n -le 0) { return "" }
            $unixSeconds = [math]::Floor(($n / 1000000) - 11644473600)
            return [string]$unixSeconds
        } catch { return "" }
    }

    function HtmlEncode([string]$s) { if ($null -eq $s) { return "" } ; return [System.Net.WebUtility]::HtmlEncode($s) }

    function Make-SafeFileName([string]$s) {
        if ($null -eq $s) { return "bookmarks" }
        $invalid = [System.IO.Path]::GetInvalidFileNameChars() + [System.IO.Path]::GetInvalidPathChars()
        $out = $s
        foreach ($c in $invalid) { $out = $out -replace [regex]::Escape($c), "_" }
        $out = $out -replace "_+", "_"
        return $out.Trim("_")
    }

    # --------------------
    # Chromium JSON -> HTML
    function Convert-ChromiumBookmarksToHtml {
        param([string]$jsonPath, [string]$outputFile)
        if (-not (Test-Path $jsonPath)) { throw "Bookmarks JSON not found: $jsonPath" }

        try {
            $tmp = Join-Path $env:TEMP ("Bookmarks_copy_" + ([System.Guid]::NewGuid().ToString()) + ".json")
            Copy-Item -Path $jsonPath -Destination $tmp -Force -ErrorAction Stop
        } catch {
            Add-Content -Path $LogFile -Value "Warning: could not copy Bookmarks file. Reading in place. $($_.Exception.Message)"
            $tmp = $jsonPath
        }

        $json = Get-Content $tmp -Raw | ConvertFrom-Json -ErrorAction Stop
        $html = Start-BookmarkHtml "Bookmarks"

        function Build-BookmarksHtml {
            param($nodes, $level)
            $out = ""
            if (-not $nodes) { return $out }
            foreach ($n in $nodes) {
                if ($n.type -and ($n.type.ToString().ToLower() -eq "folder")) {
                    $name = if ($n.name) { $n.name } else { "Folder" }
                    $add = Convert-ChromeTimestampToUnix $n.date_added
                    $mod = Convert-ChromeTimestampToUnix $n.date_modified
                    $attrs = @()
                    if ($add) { $attrs += "ADD_DATE=`"$add`"" }
                    if ($mod) { $attrs += "LAST_MODIFIED=`"$mod`"" }
                    $attrString = if ($attrs.Count -gt 0) { " " + ($attrs -join " ") } else { "" }
                    $out += ("`t" * $level) + "<DT><H3$attrString>$([System.Net.WebUtility]::HtmlEncode($name))</H3>`n"
                    $out += ("`t" * $level) + "<DL><p>`n"
                    if ($n.children) { $out += Build-BookmarksHtml $n.children ($level + 1) }
                    $out += ("`t" * $level) + "</DL><p>`n"
                }
                elseif ($n.url -or ($n.type -and ($n.type.ToString().ToLower() -eq "url"))) {
                    $title = if ($n.name) { $n.name } else { $n.url }
                    $url = $n.url
                    $add = Convert-ChromeTimestampToUnix $n.date_added
                    $itemAttrs = @()
                    if ($add) { $itemAttrs += "ADD_DATE=`"$add`"" }
                    if ($n.icon) { $itemAttrs += "ICON=`"$([System.Net.WebUtility]::HtmlEncode($n.icon))`"" }
                    $itemAttrString = if ($itemAttrs.Count -gt 0) { " " + ($itemAttrs -join " ") } else { "" }
                    $out += ("`t" * $level) + "<DT><A HREF=`"" + [System.Net.WebUtility]::HtmlEncode($url) + "`"$itemAttrString>$([System.Net.WebUtility]::HtmlEncode($title))</A>`n"
                }
                elseif ($n.children) { $out += Build-BookmarksHtml $n.children $level }
            }
            return $out
        }

        foreach ($rootProp in $json.roots.PSObject.Properties) {
            $rootName = $rootProp.Name
            $root = $json.roots.$rootName
            if (-not $root) { continue }
            $isToolbar = ($rootName -eq "bookmark_bar")
            $rootTitle = if ($root.name) { $root.name } else { $rootName }
            $rootAdd = Convert-ChromeTimestampToUnix $root.date_added
            $rootMod = Convert-ChromeTimestampToUnix $root.date_modified
            $attrs = @()
            if ($rootAdd) { $attrs += "ADD_DATE=`"$rootAdd`"" }
            if ($rootMod) { $attrs += "LAST_MODIFIED=`"$rootMod`"" }
            if ($isToolbar) { $attrs += "PERSONAL_TOOLBAR_FOLDER=`"true`"" }
            $attrString = if ($attrs.Count -gt 0) { " " + ($attrs -join " ") } else { "" }
            $html += "    <DT><H3$attrString>$([System.Net.WebUtility]::HtmlEncode($rootTitle))</H3>`n"
            $html += "    <DL><p>`n"
            $html += Build-BookmarksHtml $root.children 2
            $html += "    </DL><p>`n"
        }

        $html += "</DL><p>`n"
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
        [System.IO.File]::WriteAllBytes($outputFile, $bytes)
        if ($tmp -ne $jsonPath -and (Test-Path $tmp)) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    }

    # --------------------
    # IE / Firefox functions remain the same
    # Convert-FirefoxToHtml {...}
    # Convert-IEFavoritesToHtml {...}
    # (Omitted for brevity; reuse your existing functions)

    # --------------------
    # SYSTEM perspective: scan user profiles for Chromium-based browsers
    $AvailableBrowsers = @{}

    Get-ChildItem "C:\Users" -Directory | ForEach-Object {
        $profile = $_.FullName
        $userName = $_.Name
        if ($userName -notmatch 'Default|Public|All Users|Default User|systemprofile') {
            # Chrome
            $chromePath = Join-Path $profile "AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
            if (Test-Path $chromePath) { $AvailableBrowsers["Google Chrome ($userName)"] = $chromePath }
            # Edge
            $edgePath = Join-Path $profile "AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks"
            if (Test-Path $edgePath) { $AvailableBrowsers["Microsoft Edge ($userName)"] = $edgePath }
            # Brave
            $bravePath = Join-Path $profile "AppData\Local\BraveSoftware\Brave-Browser\User Data\Default\Bookmarks"
            if (Test-Path $bravePath) { $AvailableBrowsers["Brave ($userName)"] = $bravePath }
            # Opera
            $operaPath = Join-Path $profile "AppData\Roaming\Opera Software\Opera Stable\Bookmarks"
            if (Test-Path $operaPath) { $AvailableBrowsers["Opera ($userName)"] = $operaPath }
        }
    }

    # Firefox (still per profile)
    $ffBase = "$env:APPDATA\Mozilla\Firefox\Profiles"
    if (Test-Path $ffBase) {
        Get-ChildItem $ffBase -Directory | ForEach-Object {
            $places = Join-Path $_.FullName "places.sqlite"
            if (Test-Path $places) { $AvailableBrowsers["Mozilla Firefox ($($_.Name))"] = $places }
        }
    }

    # IE Favorites (system perspective)
    $userProfiles = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notmatch 'Default|Public|All Users|Default User|systemprofile' }
    foreach ($u in $userProfiles) {
        $favPath = Join-Path $u.FullName "Favorites"
        if (Test-Path $favPath) { $AvailableBrowsers["Internet Explorer ($($u.Name))"] = $favPath }
    }

    if ($AvailableBrowsers.Count -eq 0) { Write-Host "No supported browsers found."; exit }

    # Menu display
    Write-Host "`nDetected browsers with bookmark/favorite files:`n"
    $i = 1; $BrowserIndex = @{}
    foreach ($browser in $AvailableBrowsers.Keys) {
        Write-Host "[$i] $browser"
        $BrowserIndex[$i] = $browser
        $i++
    }

    $selection = Read-Host "`nEnter number(s) of browser(s) to export (comma separated)"
    $selectedIndexes = $selection -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' }

    # Export selected entries
    foreach ($index in $selectedIndexes) {
        if ($BrowserIndex.ContainsKey([int]$index)) {
            $browserName = $BrowserIndex[[int]$index]
            $sourcePath = $AvailableBrowsers[$browserName]
            $safeName = Make-SafeFileName($browserName.Replace(" ", "_"))
            $destFile = Join-Path $OutputFolder ($safeName + "_Bookmarks.html")
            try {
                if ($browserName -like "Google Chrome*" -or $browserName -like "Microsoft Edge*" -or $browserName -like "Brave*" -or $browserName -like "Opera*") {
                    Convert-ChromiumBookmarksToHtml -jsonPath $sourcePath -outputFile $destFile
                } elseif ($browserName -like "Mozilla Firefox*") {
                    Convert-FirefoxToHtml -sqlitePath $sourcePath -outputFile $destFile
                } elseif ($browserName -like "Internet Explorer*") {
                    Convert-IEFavoritesToHtml -favPath $sourcePath -outputFile $destFile
                } else {
                    Add-Content -Path $LogFile -Value "No conversion rule for $browserName"; Write-Host "No conversion rule for $browserName"
                }
                Write-Host "Exported $browserName bookmarks to $destFile"
                Add-Content -Path $LogFile -Value "Exported $browserName -> $destFile"
            } catch {
                Write-Host ("Failed to export {0}: {1}" -f $browserName, $_.Exception.Message) -ForegroundColor Red
                Add-Content -Path $LogFile -Value ("Failed to export {0}: {1}" -f $browserName, $_.Exception.Message)
            }
        }
    }

    Write-Host "`nAll done! Bookmarks exported as HTML in: $OutputFolder"
    Add-Content -Path $LogFile -Value "Completed at $(Get-Date -Format u)"
    Read-Host
}

function Places {
    # Define variables
    $isoUrl = "https://software.download.prss.microsoft.com/dbazure/Win11_25H2_English_x64.iso?t=a792f8f4-d7cf-44cb-b967-84c4ff20e6f0&P1=1760647882&P2=601&P3=2&P4=MMlrDkw9GKanorw5unGDZESwBsRPpIJb4iN5k76K9nAjHHu1N0YBgBKCKKa5B555Hliohf0ITVa%2bDUW3n1mppjeh212wSgGEwpOakUeOb7EVFzWCVwqMaVto9KrScRFvXM9x1dcYDjbQ%2fEA77UsqIwhsu2JfxGi5axO6Jkl8d1tx42SzAMiZ48FGfHI7haxuyfqePw71Csqb1Bf2XP8zzznCKQg%2fIkMZhzoQBnt56v2bU0jJHys42JlLZ0ZVa0%2fZePEO0P%2bWhuFnsY7JjQnUSIteOlK05Sq6zEZDerS7cuzhL5OpGCCHgh0FCaLJd7poIi0AojA7CD%2bOQ6T7EIJ2mA%3d%3d"  # Replace with actual ISO URL
    $downloadDir = "C:\Temp"
    $aria2Path = "C:\ProgramData\chocolatey\bin\aria2c.exe"  # Adjust if needed

    # Create download directory if it doesn't exist
    if (-not (Test-Path $downloadDir)) {
        New-Item -Path $downloadDir -ItemType Directory | Out-Null
        Write-Host "Created directory: $downloadDir"
    }

    # Build argument string
    $aria2Args = @(
        "--dir=$downloadDir"
        "--out=downloaded.iso"
        "--max-connection-per-server=16"
        "--split=16"
        "--min-split-size=1M"
        "--enable-http-pipelining=true"
        "--continue=true"
        "--check-certificate=false"
        "$isoUrl"
    )

    # Launch aria2c with arguments
    Start-Process -FilePath $aria2Path -ArgumentList $aria2Args -NoNewWindow -Wait
}


# === MENU ===
$MenuVar = 999
while ($MenuVar -ne 0) {
    Clear-Host
    Write-Host "Version 2.0"
    Write-Host "----------------------------------"
    Write-Host "Pick your poison." -ForegroundColor Red
    Write-Host "----------------------------------"
    Write-Host "1. Update Windows PC" -ForegroundColor Blue
    Write-Host "2. Configure adapters" -ForegroundColor Blue
    Write-Host "3. Clear space" -ForegroundColor Blue
    Write-Host "4. Speed test" -ForegroundColor Blue
    Write-Host "5. Bookmark extraction" -ForegroundColor Blue
    Write-Host "6. Windows download to temp folder" -ForegroundColor Blue
    Write-Host "To exit, press 0." -ForegroundColor Blue

    $MenuVar = Read-Host "Choose your fighter"

    switch ($MenuVar) {
        1 { Update-Windows }
        2 { Net-Adapt }
        3 { Clear-Space }
        4 { Speed-Test }
        5 { Bookmark-Export }
        6 { Places }
    }
}
