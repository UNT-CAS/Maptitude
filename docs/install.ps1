# I made this file to showcase the install process outside of PSADT's fluff.
# It's essentially the same code as what is in the `Deploy-Application.ps1`.
# This particular file will run without any dependencies, just need the following:
# - settings.json file as shown in `/` of this repo.
# - Installer media from Maptitude as shown in `/Files/` of this repo.
# - AppData.zip which you will have to get your own from installing Maptitude on a test machine.

$Settings = Get-Content "$PSScriptRoot/settings.json" | Out-String | ConvertFrom-Json
[IO.FileInfo] $issMaptitudeInstallation = "$PSScriptRoot\Files\Maptitude 2020 Build 4720\MaptitudeInstallation.iss"
[IO.FileInfo] $issDataInstaller = "$PSScriptRoot\Files\USA Country Package 2020\DataInstaller.iss"

# Licenses are good for one computer.
# The setting.json defines which computer gets which license.
if ($Settings.licenses.PSObject.Properties.Name -notcontains $env:COMPUTERNAME) {
    Throw [System.Management.Automation.ItemNotFoundException] "This computer cannot be licensed. Confirm $env:COMPUTERNAME is in the ``settings.json`` licenses list."
}

# Create Setup.iss file for silent installation.
$registration = @{
    license = $Settings.licenses.($env:COMPUTERNAME)
    user = $Settings.registration.user
    company = $Settings.registration.company
    email = $Settings.registration.email
}

$issSetupContent = Get-Content $issMaptitudeInstallation
$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%LICENSE%', ('$1={0}' -f $registration.license))
$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%USER%', ('$1={0}' -f $registration.user))
$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%COMPANY%', ('$1={0}' -f $registration.company))
$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%EMAIL%', ('$1={0}' -f $registration.email))

[IO.FileInfo] $issSetup = "$PSScriptRoot\Files\Maptitude 2020 Build 4720\Setup.iss"
$issSetupContent | Out-File -FilePath $issSetup.FullName -Force
$issSetup.Refresh()

# Install Maptitude
$installMaptitudeInstallation = @{
    FilePath = "$PSScriptRoot\Files\Maptitude 2020 Build 4720\MaptitudeInstallation.exe"
    ArgumentList = @(
        '-s',
        '/BypassActivation',
        "-f1""$issSetup"""
    )
    Wait = $true
    PassThru = $true
}
$result = Start-Process @installMaptitudeInstallation

if ($result.ExitCode) {
    Throw "Process exited with unexpected code: $($result.ExitCode)"
}

# Install Data
$installDataInstaller = @{
    FilePath = "$PSScriptRoot\Files\USA Country Package 2020\DataInstaller.exe"
    ArgumentList = @(
        '-s',
        "-f1""$issDataInstaller"""
    )
    Wait = $true
    PassThru = $true
}
$result = Start-Process @installDataInstaller

if ($result.ExitCode) {
    Throw "Process exited with unexpected code: $($result.ExitCode)"
}

# The `/BypassActivation` prevents Serial Number prompt in ISS.
# The answer file won't work for the licensing.
# Adding to registry.
$regKeys = @(
    'HKLM:\SOFTWARE\Caliper Corporation\Maptitude\2020\',
    'HKLM:\SOFTWARE\WOW6432Node\Caliper Corporation\Maptitude\2020\'
)

foreach ($regKey in $regKeys) {
    Set-ItemProperty $regKey -Name 'Serial Number' -Type String -Value $registration.license -Force
}

$installDir = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\').'Installed In'

# Run the silent license activation
#   Seems to always Exits 0, so just run now and confirm later.
$instantActivatorApp = @{
    FilePath = "$installDir\ActivateLicense\InstantActivatorApp.exe"
    Wait = $true
    PassThru = $true
}
$result = Start-Process @instantActivatorApp

if ($result.ExitCode) {
    Throw "Process exited with unexpected code: $($result.ExitCode)"
}

# Validate licensing worked.
Add-Type -Path "$installDir\ActivateLicense\InstantActivator.dll"
[com.caliper.softwarekey.InstantActivator] $InstantActivator = New-Object -TypeName 'com.caliper.softwarekey.InstantActivator'

$InstantActivator.Start() | Out-Null
if (($InstantActivator.KeyStatus -ne 'SSCP_ACTIVATED') -or (-not $InstantActivator.KeyHasSerialNumber)) {
    Throw "Licensing Failed: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber)"
}

# Setup Default Profile
#   Turn off Registration Prompt
#   Turn off Software Updates
[IO.DirectoryInfo] $publicAppData = '{0}\AppData\Roaming' -f (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList').Default
Expand-Archive "$PSScriptRoot\Files\AppData.zip" -DestinationPath ('{0}\Caliper\Maptitude 2020' -f $publicAppData) -Force

# Delete Icon From AllUser's Desktop
[IO.DirectoryInfo] $publicDesktop = [System.Environment]::GetFolderPath('CommonDesktopDirectory')
Remove-Item ('{0}\Maptitude *.lnk' -f $publicDesktop) -Force
