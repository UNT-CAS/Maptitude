$ErrorActionPreference = 'Stop'
Trap {
    # For successful "Application Not Installed" detection, need NO STDOUT and EXIT 0
    #    https://docs.microsoft.com/en-us/mem/configmgr/apps/deploy-use/create-applications#about-custom-script-detection-methods
    Exit 0
}

$displayNameMaptitude = 'Maptitude 2020 (64-bit)'
[version] $versionMaptitude = '2020.0.4720'
$displayNameData = 'Maptitude Data for USA (HERE) - 2019 Quarter 4'

# Fastest thing we can do to check if it's installed.
$installDir = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\' -ErrorAction Ignore).'Installed In'
if (-not $installDir) {
    Throw [System.Management.Automation.ItemNotFoundException] "Registry key for '${displayNameMaptitude}' install dir was not found."
}

# Look for Maptitude
$regUninstallKeys = @(
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
    'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall'
)

$foundMaptitude = $false
$foundData = $false

:Uninstalls foreach ($regUninstallKey in $regUninstallKeys) {
    foreach ($regSubKey in (Get-ChildItem $regUninstallKey)) {
        $regValues = Get-ItemProperty $regSubKey.PSPath
        switch ($regValues.DisplayName) {
            $displayNameMaptitude {
                [IO.FileInfo] $mapt = '{0}\mapt.exe' -f $regValues.InstallLocation
                if (([version] $mapt.VersionInfo.ProductVersion) -eq $versionMaptitude) {
                    $foundMaptitude = $true
                }
            }
            
            $displayNameData {
                $foundData = $true
            }
        }
                
        if ($foundMaptitude -band $foundData) {
            break Uninstalls
        }
    }
}

# Is Maptitude installed? Is it the correct version?
if (-not $foundMaptitude) {
    Throw [System.Management.Automation.ItemNotFoundException] "Either '${displayNameMaptitude}' or the correct version '${versionMaptitude}' was not found."
}

# Is Data installed?
if (-not $foundData) {
    Throw [System.Management.Automation.ItemNotFoundException] "${displayNameData} was not found."
}

# Is Access DB Engine installed?
if (-not (Get-CimInstance Win32_Product -Filter 'IdentifyingNumber = "{90160000-00D1-0409-1000-0000000FF1CE}"')) {
    Throw [System.Management.Automation.ItemNotFoundException] "Microsoft Access database engine 2016 (English) v16.0.4519.1000 was not found."
}

# Is the License valid?
Add-Type -Path "$installDir\ActivateLicense\InstantActivator.dll"
[com.caliper.softwarekey.InstantActivator] $InstantActivator = New-Object -TypeName 'com.caliper.softwarekey.InstantActivator'

$InstantActivator.Start() | Out-Null
if (($InstantActivator.KeyStatus -ne 'SSCP_ACTIVATED') -or (-not $InstantActivator.KeyHasSerialNumber)) {
    Throw "Licensing Failed: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber)"
}

# Does Default Users's AppData have some files?
[IO.DirectoryInfo] $defaultProfile = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList').Default
[IO.DirectoryInfo] $defaultAppData = '{0}\AppData\Roaming' -f $defaultProfile.FullName
if (-not (Get-ChildItem ('{0}\Caliper\Maptitude 2020' -f $defaultAppData.FullName))) {
    Throw [System.Management.Automation.ItemNotFoundException] "Default Users's AppData files were not found."
}

# Choosing to not run detection on each User Profile.

# Is the shortcut removed from the All User's Desktop?
[IO.DirectoryInfo] $publicDesktop = [System.Environment]::GetFolderPath('CommonDesktopDirectory')
if (Get-ChildItem ('{0}\Maptitude *.lnk' -f $publicDesktop)) {
    # This is like a double negative ItemNotFound ... :D
    Throw [System.Management.Automation.ItemNotFoundException] "Maptitude Desktop shortcut should not have been found, but it was."
}

# For successful "Application Installed" detection, need STDOUT and EXIT 0
#    https://docs.microsoft.com/en-us/mem/configmgr/apps/deploy-use/create-applications#about-custom-script-detection-methods
Write-Output 'Success'
Exit 0
