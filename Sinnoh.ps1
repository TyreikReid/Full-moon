function Clear-Space {

    # ----------------------------
    # FUNCTION: Get free disk space
    # ----------------------------
    function Get-FreeSpace {
        $drive = Get-PSDrive -Name C
        return $drive.Used, $drive.Free
    }

    # Set execution policy for session
    Set-ExecutionPolicy Bypass -Scope Process -Force

    # Record starting space
    $initialUsedSpace, $initialFreeSpace = Get-FreeSpace

    Write-Host "Starting DISM component cleanup..."

    try {
        $process = Start-Process -FilePath "Dism.exe" `
            -ArgumentList "/Online", "/Cleanup-Image", "/StartComponentCleanup", "/ResetBase" `
            -WindowStyle Hidden -Wait -PassThru

        if ($process.ExitCode -eq 0) {
            Write-Host "DISM cleanup complete."
        } else {
            Write-Host "DISM exited with code $($process.ExitCode)." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "DISM failed: $_" -ForegroundColor Red
    }

    # Stop services
    Write-Host "Stopping WU and BITS..."
    Stop-Service wuauserv -ErrorAction SilentlyContinue
    Stop-Service bits -ErrorAction SilentlyContinue

    # Clean SoftwareDistribution
    $SoftwareDistributionPath = "C:\Windows\SoftwareDistribution"
    if (Test-Path $SoftwareDistributionPath) {
        Write-Host "Cleaning SoftwareDistribution folder..."

        $TempPath = Join-Path $SoftwareDistributionPath "empty"
        New-Item -ItemType Directory -Path $TempPath -Force | Out-Null

        robocopy $TempPath $SoftwareDistributionPath /MIR /XD $TempPath | Out-Null
        
        Remove-Item -Recurse -Force -Path $TempPath
        Write-Host "SoftwareDistribution cleaned."
    }

    # Restart services
    Write-Host "Restarting WU and BITS..."
    Start-Service wuauserv -ErrorAction SilentlyContinue
    Start-Service bits -ErrorAction SilentlyContinue

    # Clean user temp folders
    Write-Host "Cleaning all user TEMP folders..."

    $UserProfiles = Get-ChildItem "C:\Users" | Where-Object {
        $_.Name -notin @('Public', 'Default') -and $_.PSIsContainer
    }

    $deletedUsers = @()

    foreach ($UserProfile in $UserProfiles) {
        $TempFolder = Join-Path $UserProfile.FullName "AppData\Local\Temp"

        if (Test-Path $TempFolder) {
            try {
                Get-ChildItem $TempFolder -Recurse -ErrorAction SilentlyContinue |
                Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
                $deletedUsers += $UserProfile.Name
            }
            catch {
                Write-Host "Failed to clean: $($UserProfile.Name) - $_"
            }
        }
    }

    if ($deletedUsers.Count) {
        Write-Host "Temp cleaned for: $($deletedUsers -join ', ')"
    } else {
        Write-Host "No temp deletions performed."
    }

    # Free space results
    $finalUsedSpace, $finalFreeSpace = Get-FreeSpace
    $spaceFreed = $finalFreeSpace - $initialFreeSpace

    Write-Host "`nInitial free space: $([math]::Round($initialFreeSpace / 1GB, 2)) GB"
    Write-Host "Final free space:   $([math]::Round($finalFreeSpace / 1GB, 2)) GB"
    Write-Host "Total space freed:  $([math]::Round($spaceFreed / 1GB, 2)) GB"

    Write-Host "`nCleanup complete."
}
