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
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\',
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Caliper Corporation\Maptitude\2020\'
)

foreach ($regKey in $regKeys) {
    Set-ItemProperty $regKey -Name 'Serial Number' -Type String -Value $registration.license -Force
}

$installDir = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\').'Installed In'

# Run the silent license activation
$instantActivatorApp = @{
    FilePath = "$installDir\ActivateLicense\InstantActivatorApp.exe"
    Wait = $true
    PassThru = $true
}
$result = Start-Process @instantActivatorApp

if ($result.ExitCode -eq 1) {
    Throw "Error with activation; Serial Number may already be in use."
} elseif ($result.ExitCode) {
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
[IO.DirectoryInfo] $defaultProfile = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList').Default
[IO.DirectoryInfo] $defaultAppData = '{0}\AppData\Roaming' -f $defaultProfile.FullName
Expand-Archive "$PSScriptRoot\Files\AppData.zip" -DestinationPath ('{0}\Caliper\Maptitude 2020' -f $defaultAppData.FullName) -Force

# Setup each User Profiles
$userProfiles = Get-ChildItem 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
foreach ($userProfile in $userProfiles) {
    $userSid = $userProfile.PSChildName   
    if (-not $userSid.StartsWith('S-1-5-21-')) {
        # Skip System Users and Services
        continue
    }

    $userDomain, $userName = (New-Object System.Security.Principal.SecurityIdentifier($userSid)).Translate([System.Security.Principal.NTAccount]).Value.Split('\')

    try {
        $userShellFolders = Get-Item ('Registry::HKEY_USERS\{0}\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -f $userSid) -ErrorAction Stop
        $userEnvironment = Get-Item ('Registry::HKEY_USERS\{0}\Environment' -f $userSid) -ErrorAction Stop
    } catch [System.Management.Automation.ItemNotFoundException] {
        # Not a full/real User Profile
        # Likely was cerated using RunAs
        continue
    }

    $userAppData = $userShellFolders.GetValue('AppData', '', 'DoNotExpandEnvironmentNames')

    # Build the user environment based on the target User and System Environments
    [hashtable] $userEnv = @{}

    foreach ($var in (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment').PSObject.Properties) {
        if (@('PSPath','PSParentPath','PSChildName','PSProvider') -notcontains $var.Name) {
            $userEnv.Set_Item(('%{0}%' -f $var.Name), $userEnvironment.GetValue($var.Name, '', 'DoNotExpandEnvironmentNames'))
        }
    }

    foreach ($var in $userEnvironment.PSObject.Properties) {
        if (@('PSPath','PSParentPath','PSChildName','PSProvider') -contains $var.Name) {
            $userEnv.Set_Item(('%{0}%' -f $var.Name), $var.Value)
        }
    }

    $userEnv.Set_Item('%USERDOMAIN%', $userDomain)
    $userEnv.Set_Item('%USERNAME%', $userName)
    $userEnv.Set_Item('%USERPROFILE%', (Get-ItemProperty $userProfile.PSPath).ProfileImagePath)

    # Expand Env Strings
    $regexes = $userEnv.Keys | ForEach-Object {[System.Text.RegularExpressions.Regex]::Escape($_)}
    $regex = [regex]('(?i)' + ($regexes -join '|')) # (?i) makes it case-insensitive
    $userAppData = $regex.Replace($userAppData, { $userEnv[$args[0].Value] })

    # Expand Archive
    Expand-Archive "$PSScriptRoot\Files\AppData.zip" -DestinationPath ('{0}\Caliper\Maptitude 2020' -f $userAppData) -Force
}

# Delete Icon From AllUser's Desktop
[IO.DirectoryInfo] $publicDesktop = [System.Environment]::GetFolderPath('CommonDesktopDirectory')
Remove-Item ('{0}\Maptitude *.lnk' -f $publicDesktop) -Force
