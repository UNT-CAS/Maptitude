$displayNameMaptitude = 'Maptitude 2020 (64-bit)'
$displayNameData = 'Maptitude Data for USA (HERE) - 2019 Quarter 4'
[IO.FileInfo] $issMaptitudeInstallation = "$PSScriptRoot\Files\Maptitude 2020 Build 4720\MaptitudeInstallation.x.iss"
[IO.FileInfo] $issDataInstaller = "$PSScriptRoot\Files\USA Country Package 2020\DataInstaller.x.iss"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$installDir = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\').'Installed In'

# Run License Deactivation
$instantActivatorApp = @{
    FilePath = "$installDir\ActivateLicense\InstantActivatorApp.exe"
    ArgumentList = @(
        '/D'
    )
    WindowStyle = 'Hidden'
    Wait = $true
    PassThru = $true
}
$result = Start-Process @instantActivatorApp

if ($result.ExitCode -eq 2) {
    Write-Information "ExitCode 2: likely that no active license found"
} elseif ($result.ExitCode) {
    Throw "Process exited with unexpected code: $($result.ExitCode)"
}

# Confirm License Deactivation
Add-Type -Path "$installDir\ActivateLicense\InstantActivator.dll"
[com.caliper.softwarekey.InstantActivator] $InstantActivator = New-Object -TypeName 'com.caliper.softwarekey.InstantActivator'
$InstantActivator.Start() | Out-Null
if (($InstantActivator.KeyStatus -eq 'SSCP_ACTIVATED')) {
    Throw "Licensing Failed: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber)"
}

$regKeys = @(
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\',
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\'
)

$uninstallStringMaptitude = $null
$uninstallStringData = $null

:Uninstalls foreach ($regUninstallKey in $regKeys) {
    foreach ($regSubKey in (Get-ChildItem $regUninstallKey)) {
        $regValues = Get-ItemProperty $regSubKey.PSPath
        switch ($regValues.DisplayName) {
            $displayNameMaptitude {
                $uninstallStringMaptitude = $regValues.UninstallString
            }
            
            $displayNameData {
                $uninstallStringData = $regValues.UninstallString
            }
        }

        if (($uninstallStringMaptitude -as [bool]) -band ($uninstallStringData -as [bool])) {
            break Uninstalls
        }
    }
}

# Sample UninstallString
#   "C:\Program Files (x86)\InstallShield Installation Information\{1AC9AF81-4426-11D7-BD59-0002B34B98FF}\setup.exe" -runfromtemp -l0x0409 AddRemove -removeonly
$regexUninstallString = '(?(?=\")("([^"]+)")|([^ ]+))\s+(.+)'

if ($uninstallStringData) { # If nothing found, likely not installed.
    if ($uninstallStringData -match $regexUninstallString) {
        $uninstallData = @{
            FilePath = if ([string]::IsNullOrEmpty($Matches[1])) { $Matches[2] } else { $Matches[1]}
            ArgumentList = '{0} -s -f1"{1}"' -f $Matches[3], $issDataInstaller.FullName
            Wait = $true
            PassThru = $true
        }
        Write-Information "Uninstall Data: $($uninstallData | ConvertTo-Json)"
        $result = Start-Process @uninstallData

        if ($result.ExitCode) {
            Write-Information "Unexpected Exit Code: $($result.ExitCode)"
        }
    } else {
        Throw "Uninstall String matcher failed: $uninstallStringData"
    }
}

if ($uninstallStringMaptitude) { # If nothing found, likely not installed.
    if ($uninstallStringMaptitude -match $regexUninstallString) {
        $uninstallMaptitude = @{
            FilePath = if ([string]::IsNullOrEmpty($Matches[1])) { $Matches[2] } else { $Matches[1]}
            ArgumentList = '{0} -s -f1"{1}"' -f $Matches[3], $issMaptitudeInstallation.FullName
            Wait = $true
            PassThru = $true
        }
        Write-Information "Uninstall Maptitude: $($uninstallMaptitude | ConvertTo-Json)"
        $result = Start-Process @uninstallMaptitude

        if ($result.ExitCode) {
            Write-Information "Unexpected Exit Code: $($result.ExitCode)"
        }
    } else {
        Throw "Uninstall String matcher failed: $uninstallStringData"
    }
}



# Remove AppData from Default Profile

[IO.DirectoryInfo] $defaultProfile = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList').Default
[IO.DirectoryInfo] $defaultAppData = '{0}\AppData\Roaming' -f $defaultProfile.FullName
[IO.DirectoryInfo] $defaultMaptitudeAppData = '{0}\Caliper\Maptitude 2020' -f $defaultAppData.FullName
Remove-Item $defaultMaptitudeAppData.FullName -Recurse -Force -ErrorAction SilentlyContinue
if (-not (Get-ChildItem $defaultMaptitudeAppData.Paren -ErrorAction SilentlyContinue)) {
    Remove-Item $defaultMaptitudeAppData.Parent.FullName -Recurse -Force -ErrorAction SilentlyContinue
}

# Remove AppData from each User Profile
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
    $regexes = $userEnv.Keys | ForEach-Object { [regex]::Escape($_) }
    $regex = [regex]('(?i)' + ($regexes -join '|')) # (?i) makes it case-insensitive
    $userAppData = $regex.Replace($userAppData, { $userEnv[$args[0].Value] })

    # Remove AppData
    [IO.DirectoryInfo] $userMaptitudeAppData = '{0}\Caliper\Maptitude 2020' -f $userAppData
    Remove-Item $userMaptitudeAppData.FullName -Recurse -Force -ErrorAction SilentlyContinue
    if (-not (Get-ChildItem $userMaptitudeAppData.Parent -ErrorAction SilentlyContinue)) {
        Remove-Item $userMaptitudeAppData.Parent.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Remove lingering artifacts
Remove-Item $installDir -Recurse -Force -ErrorAction SilentlyContinue
