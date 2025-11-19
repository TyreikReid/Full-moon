function DISM {
    Write-Host "Starting the DISM process now, please standby."
    $outputFile = DISMlog.txt
    $OutputFolder = C:/temp
DISM /Online /Cleanup-Image /RestoreHealth; sfc /scannow }
