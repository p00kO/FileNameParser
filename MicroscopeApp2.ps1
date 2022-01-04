Set-ExecutionPolicy unrestricted
Add-Type -AssemblyName System.Windows.Forms

#CHANGE THESE 3 LINES FOR EACH INSTALL:
# -- Variables to set at install: Microscope Application name, Path to executable, Path to the Watcher's directory -- 

$micAppName = "amscope"
$micAppPath = "C:\Program Files\AmScope\AmScope\x64\"
$localWatcherDir = "C:\Users\P00ko\source\repos\FileNameParser\"

# -------------------------------------------------------

$watchercmd = $localWatcherDir + "WindowsFormsApp1\bin\Release\WindowsFormsApp1.exe"
$batPath = $localWatcherDir + "Test1.bat"
$micAppExe = $micAppName + ".exe"
$micAppCmdPath = $micAppPath + $micAppExe


#1) Check that AmScope isn't running:
$procTest = Get-Process -Name $micAppName -ErrorAction SilentlyContinue
if($procTest){
    Write-Output Fizzle snaps! 
    [System.Windows.Forms.MessageBox]::Show('Hmmmm... Looks like you already have a microscope application already running. Please close it and try again', 
    'TurretMicroscopeWatcherApp',
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information)
    exit
}

#2) Start the microscope program:
# Set the location to the registry
Set-Location -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'

# Remove the key
Remove-Item -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\$micAppExe" -Force -Verbose

# Start Microscope app process
$micApp= Start-Process $micAppCmdPath -PassThru

# Restore the RegKey
Get-Item -Path 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options' | New-Item -Name $micAppExe -Force
New-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\$micAppExe" -Name 'Debugger' -Value $batPath -PropertyType String -Force

#Start process/turret watcher application:
$ID=$micApp.Id
$client = Start-Process $watchercmd -ArgumentList "$ID" -PassThru

#Kill watcher when user quits microscope application
$micApp.WaitForExit()
Stop-Process($client.Id)