$ErrorActionPreference = 'Stop'
Trap {
    # For successful "Application Not Installed" detection, need NO STDOUT and EXIT 0
    #    https://docs.microsoft.com/en-us/mem/configmgr/apps/deploy-use/create-applications#about-custom-script-detection-methods
    Exit 0
}

$displayNameMaptitude = 'Maptitude 2020 (64-bit)'
[version] $versionMaptitude = '2020.0.4720'
$displayNameData = 'Maptitude Data for USA (HERE) - 2019 Quarter 4'

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
                    
                    if ($foundMaptitude -band $foundData) {
                        break Uninstalls
                    }
                }
            }
            
            $displayNameData {
                $foundData = $true
                
                if ($foundMaptitude -band $foundData) {
                    break Uninstalls
                }
            }
        }
    }
}

if (-not $foundMaptitude) {
    Throw [System.Management.Automation.ItemNotFoundException] "Either '${displayNameMaptitude}' or the correct version '${versionMaptitude}' was not found."
}

if (-not $foundData) {
    Throw [System.Management.Automation.ItemNotFoundException] "${displayNameData} was not found."
}

# Validate licensing worked.
$installDir = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\').'Installed In'
Add-Type -Path "$installDir\ActivateLicense\InstantActivator.dll"
[com.caliper.softwarekey.InstantActivator] $InstantActivator = New-Object -TypeName 'com.caliper.softwarekey.InstantActivator'

$InstantActivator.Start() | Out-Null
if (($InstantActivator.KeyStatus -ne 'SSCP_ACTIVATED') -or (-not $InstantActivator.KeyHasSerialNumber)) {
    Throw "Licensing Failed: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber)"
}

# For successful "Application Installed" detection, need STDOUT and EXIT 0
#    https://docs.microsoft.com/en-us/mem/configmgr/apps/deploy-use/create-applications#about-custom-script-detection-methods
Write-Output 'Success'
Exit 0
