# Load MDT Task Sequence Environment and Logs
$TSenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
$logPath = $tsenv.Value("LogPath")
$logFile = "$logPath\BIOS_Update.log"
  
# Start the logging 
  
Write-Output "Logging to $logFile." > $logFile
  
# Collect data
Write-Output "Collecting Data" >> $logFile
$ScriptRoot = (Get-location).Path
$Model = $TSenv.Value("Model")
$CompBiosVersion = (Get-WmiObject WIN32_BIOS).SMBIOSBIOSVersion
$CurrentBiosVersion = Get-Content "$ScriptRoot\$Model\BIOS.txt"
$Installer = "UpgradeBIOS.cmd"
 
try {
    Test-Path $CurrentBiosVersion -ErrorAction Stop
}
catch {
    Write-Output "BIOS.txt does not exist!" >> $logFile
}
 
Write-Output "Copying $ScriptRoot\$Model to C:\Temp\$Model" >> $logFile
Copy-Item "$ScriptRoot\$Model" "C:\Temp\$Model" -Force -Recurse
  
# Checking for BIOS update
if($CompBiosVersion.replace(' ' , '') -eq $CurrentBiosVersion.replace(' ' , '')) {
    Write-Output "BIOS is up to date." >> $logFile
    Exit
}
else {
    Write-Output "Updating BIOS $CompBiosVersion to $CurrentBiosVersion." >> $logFile
    Start-Process "cmd.exe" "/c C:\Temp\$Model\$Installer" -Wait
    $tsenv.Value("NeedReboot") = "YES"
    Write-Output "Update has been completed successfully." >> $logFile
    Exit
}
