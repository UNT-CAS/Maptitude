<#
.SYNOPSIS
	This script performs the installation or uninstallation of an application(s).
	# LICENSE #
	PowerShell App Deployment Toolkit - Provides a set of functions to perform common application deployment tasks on Windows.
	Copyright (C) 2017 - Sean Lillis, Dan Cunningham, Muhammad Mashwani, Aman Motazedian.
	This program is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or any later version. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
	You should have received a copy of the GNU Lesser General Public License along with this program. If not, see <http://www.gnu.org/licenses/>.
.DESCRIPTION
	The script is provided as a template to perform an install or uninstall of an application(s).
	The script either performs an "Install" deployment type or an "Uninstall" deployment type.
	The install deployment type is broken down into 3 main sections/phases: Pre-Install, Install, and Post-Install.
	The script dot-sources the AppDeployToolkitMain.ps1 script which contains the logic and functions required to install or uninstall an application.
.PARAMETER DeploymentType
	The type of deployment to perform. Default is: Install.
.PARAMETER DeployMode
	Specifies whether the installation should be run in Interactive, Silent, or NonInteractive mode. Default is: Interactive. Options: Interactive = Shows dialogs, Silent = No dialogs, NonInteractive = Very silent, i.e. no blocking apps. NonInteractive mode is automatically set if it is detected that the process is not user interactive.
.PARAMETER AllowRebootPassThru
	Allows the 3010 return code (requires restart) to be passed back to the parent process (e.g. SCCM) if detected from an installation. If 3010 is passed back to SCCM, a reboot prompt will be triggered.
.PARAMETER TerminalServerMode
	Changes to "user install mode" and back to "user execute mode" for installing/uninstalling applications for Remote Destkop Session Hosts/Citrix servers.
.PARAMETER DisableLogging
	Disables logging to file for the script. Default is: $false.
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeployMode 'Silent'; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -AllowRebootPassThru; Exit $LastExitCode }"
.EXAMPLE
    powershell.exe -Command "& { & '.\Deploy-Application.ps1' -DeploymentType 'Uninstall'; Exit $LastExitCode }"
.EXAMPLE
    Deploy-Application.exe -DeploymentType "Install" -DeployMode "Silent"
.NOTES
	Toolkit Exit Code Ranges:
	60000 - 68999: Reserved for built-in exit codes in Deploy-Application.ps1, Deploy-Application.exe, and AppDeployToolkitMain.ps1
	69000 - 69999: Recommended for user customized exit codes in Deploy-Application.ps1
	70000 - 79999: Recommended for user customized exit codes in AppDeployToolkitExtensions.ps1
.LINK
	http://psappdeploytoolkit.com
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false)]
	[ValidateSet('Install','Uninstall','Repair')]
	[string]$DeploymentType = 'Install',
	[Parameter(Mandatory=$false)]
	[ValidateSet('Interactive','Silent','NonInteractive')]
	[string]$DeployMode = 'Interactive',
	[Parameter(Mandatory=$false)]
	[switch]$AllowRebootPassThru = $false,
	[Parameter(Mandatory=$false)]
	[switch]$TerminalServerMode = $false,
	[Parameter(Mandatory=$false)]
	[switch]$DisableLogging = $false
)

Try {
	## Set the script execution policy for this process
	Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force -ErrorAction 'Stop' } Catch {}

	##*===============================================
	##* VARIABLE DECLARATION
	##*===============================================
	## Variables: Application
	[string]$appVendor = 'Caliper'
	[string]$appName = 'Maptitude'
	[string]$appVersion = '2020'
	[string]$appArch = 'x64'
	[string]$appLang = 'EN'
	[string]$appRevision = '01'
	[string]$appScriptVersion = '1.0.0'
	[string]$appScriptDate = '09/19/2020'
	[string]$appScriptAuthor = 'Raymond Piller'
	##*===============================================
	## Variables: Install Titles (Only set here to override defaults set by the toolkit)
	[string]$installName = '{0} {1} {2} ({3})' -f $appVendor, $appName, $appVersion, $appScriptVersion
	[string]$installTitle = '{0} {1} {2} ({3})' -f $appVendor, $appName, $appVersion, $appScriptVersion

	##* Do not modify section below
	#region DoNotModify

	## Variables: Exit Code
	[int32]$mainExitCode = 0

	## Variables: Script
	[string]$deployAppScriptFriendlyName = 'Deploy Application'
	[version]$deployAppScriptVersion = [version]'3.8.2'
	[string]$deployAppScriptDate = '08/05/2020'
	[hashtable]$deployAppScriptParameters = $psBoundParameters

	## Variables: Environment
	If (Test-Path -LiteralPath 'variable:HostInvocation') { $InvocationInfo = $HostInvocation } Else { $InvocationInfo = $MyInvocation }
	[string]$scriptDirectory = Split-Path -Path $InvocationInfo.MyCommand.Definition -Parent

	## Dot source the required App Deploy Toolkit Functions
	Try {
		[string]$moduleAppDeployToolkitMain = "$scriptDirectory\AppDeployToolkit\AppDeployToolkitMain.ps1"
		If (-not (Test-Path -LiteralPath $moduleAppDeployToolkitMain -PathType 'Leaf')) { Throw "Module does not exist at the specified location [$moduleAppDeployToolkitMain]." }
		If ($DisableLogging) { . $moduleAppDeployToolkitMain -DisableLogging } Else { . $moduleAppDeployToolkitMain }
	}
	Catch {
		If ($mainExitCode -eq 0){ [int32]$mainExitCode = 60008 }
		Write-Error -Message "Module [$moduleAppDeployToolkitMain] failed to load: `n$($_.Exception.Message)`n `n$($_.InvocationInfo.PositionMessage)" -ErrorAction 'Continue'
		## Exit the script, returning the exit code to SCCM
		If (Test-Path -LiteralPath 'variable:HostInvocation') { $script:ExitCode = $mainExitCode; Exit } Else { Exit $mainExitCode }
	}

	#endregion
	##* Do not modify section above
	##*===============================================
	##* END VARIABLE DECLARATION
	##*===============================================

	if (Get-Module PsIni -ListAvailable) {
		Install-Module PsIni -Scope CurrentUser -Force
	}

	Import-Module PsIni

	$Settings = Get-Content "$PSScriptRoot/settings.json" | Out-String | ConvertFrom-Json
	$issMaptitudeInstallation = "$PSScriptRoot/MaptitudeInstallation.iss"
	$issDataInstaller = "$PSScriptRoot/DataInstaller.iss"

	If ($deploymentType -ine 'Uninstall' -and $deploymentType -ine 'Repair') {
		##*===============================================
		##* PRE-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Installation'

		## Show Welcome Message, close Internet Explorer if required, allow up to 3 deferrals, verify there is enough disk space to complete the install, and persist the prompt
		# Show-InstallationWelcome -CloseApps 'iexplore' -AllowDefer -DeferTimes 3 -CheckDiskSpace -PersistPrompt

		## Show Progress Message (with the default message)
		# Show-InstallationProgress

		## <Perform Pre-Installation tasks here>
		
		if ($Settings.licenses.PSObject.Properties.Name -notcontains $env:COMPUTERNAME) {
			Throw [System.Management.Automation.ItemNotFoundException] "This computer cannot be licensed. Confirm $env:COMPUTERNAME is in the ``settings.json`` licenses list."
		}

		[IO.FileInfo] $issMaptitudeInstallation = "$PSScriptRoot\Files\Maptitude 2020 Build 4720\MaptitudeInstallation.iss"
		[IO.FileInfo] $issDataInstaller = "$PSScriptRoot\Files\USA Country Package 2020\DataInstaller.iss"

		$registration = @{
			license = $Settings.licenses.($env:COMPUTERNAME)
			user = $Settings.registration.user
			company = $Settings.registration.company
			email = $Settings.registration.email
		}

		##*===============================================
		##* INSTALLATION
		##*===============================================
		[string]$installPhase = 'Installation'

		# Setup Installer ISS file with license and registration.
		$issSetupContent = Get-Content $issMaptitudeInstallation
		$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%LICENSE%', ('$1={0}' -f $registration.license))
		$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%USER%', ('$1={0}' -f $registration.user))
		$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%COMPANY%', ('$1={0}' -f $registration.company))
		$issSetupContent = $issSetupContent -replace ('(szEdit\d)=%EMAIL%', ('$1={0}' -f $registration.email))

		[IO.FileInfo] $issSetup = "$PSScriptRoot\Files\Maptitude 2020 Build 4720\Setup.iss"
		$issSetupContent | Out-File -FilePath $issSetup.FullName -Force
		$issSetup.Refresh()

		# Install Main
		$InstallParameters = @(
			'-s',
			'/BypassActivation',
			"-f1""$issSetup"""
		)
		Execute-Process -Path "$PSScriptRoot\Files\Maptitude 2020 Build 4720\MaptitudeInstallation.exe" -Parameters $InstallParameters -WindowStyle Hidden

		# Install Data
		$InstallParameters = @(
			'-s',
			"-f1""$issDataInstaller"""
		)
		Execute-Process -Path "$PSScriptRoot\Files\USA Country Package 2020\DataInstaller.exe" -Parameters $InstallParameters -WindowStyle Hidden

		# The `/BypassActivation` prevents Serial Number prompt in ISS.
		# The answer file won't work for the licensing.
		# Adding to registry.
		$regKeys = @(
			'HKLM:\SOFTWARE\Caliper Corporation\Maptitude\2020\',
			'HKLM:\SOFTWARE\WOW6432Node\Caliper Corporation\Maptitude\2020\'
		)

		foreach ($regKey in $regKeys) {
			Write-Log "RegKey Before: $regKey $(Get-ItemProperty $regKey | ConvertTo-Json)"
			Set-ItemProperty $regKey -Name 'Serial Number' -Type String -Value $registration.license -Force
			Write-Log "RegKey After: $regKey $(Get-ItemProperty $regKey | ConvertTo-Json)"
		}

		$installDir = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\').'Installed In'

		# Run License Activation
		Execute-Process -Path "$installDir\ActivateLicense\InstantActivatorApp.exe" -WindowStyle Hidden
		
		# Confirm License Activation
		Add-Type -Path "$installDir\ActivateLicense\InstantActivator.dll"
		[com.caliper.softwarekey.InstantActivator] $InstantActivator = New-Object -TypeName 'com.caliper.softwarekey.InstantActivator'
		$InstantActivator.Start() | Out-Null
		Write-Log "License Status: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber) $($InstantActivator | ConvertTo-Json)"
		if (($InstantActivator.KeyStatus -ne 'SSCP_ACTIVATED') -or (-not $InstantActivator.KeyHasSerialNumber)) {
			Throw "Licensing Failed: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber)"
		}

		# Setup Default Profile
		#   Turn off Registration Prompt (doesn't appear to work)
		#   Turn off Software Upates
		[IO.DirectoryInfo] $publicDesktop = [System.Environment]::GetFolderPath('CommonDesktopDirectory')
		[IO.DirectoryInfo] $publicAppData = '{0}\AppData\Roaming' -f $publicDesktop.Parent.FullName
		Write-Log "Public Desktop: $publicDesktop $($publicDesktop | Select-Object * | Out-String)"
		Write-Log "Public AppData: $publicAppData $($publicAppData | Select-Object * | Out-String)"
		
		$extracted = Expand-Archive "$PSScriptRoot\Files\AppData.zip" -DestinationPath ('{0}\Caliper\Maptitude 2020' -f $publicAppData) -Force -Verbose 4>&1
		Write-Log "Extracted: Default Profile`n`n$($extracted | Out-String)"

		# Delete Icon From AllUser's Desktop
		Remove-Item ('{0}\Maptitude *.lnk' -f $publicDesktop) -Force

		##*===============================================
		##* POST-INSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Installation'

		## Display a message at the end of the install
		# If (-not $useDefaultMsi) { Show-InstallationPrompt -Message 'You can customize text to appear at the end of an install or remove it completely for unattended installations.' -ButtonRightText 'OK' -Icon Information -NoWait }
	}
	ElseIf ($deploymentType -ieq 'Uninstall')
	{
		##*===============================================
		##* PRE-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Pre-Uninstallation'

		## Show Welcome Message, close Internet Explorer with a 60 second countdown before automatically closing
		# Show-InstallationWelcome -CloseApps 'iexplore' -CloseAppsCountdown 60

		## Show Progress Message (with the default message)
		# Show-InstallationProgress

		$installDir = (Get-ItemProperty 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Caliper Corporation\Maptitude\2020\').'Installed In'

		# Run License Activation
		$InstallParameters = @(
			'/D'
		)
		Execute-Process -Path "$installDir\ActivateLicense\InstantActivatorApp.exe" -Parameters $InstallParameters -WindowStyle Hidden

		# Confirm License Activation
		Add-Type -Path "$installDir\ActivateLicense\InstantActivator.dll"
		[com.caliper.softwarekey.InstantActivator] $InstantActivator = New-Object -TypeName 'com.caliper.softwarekey.InstantActivator'
		$InstantActivator.Start() | Out-Null
		Write-Log "License Status: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber) $($InstantActivator | ConvertTo-Json)"
		if (($InstantActivator.KeyStatus -eq 'SSCP_ACTIVATED') -or (-not $InstantActivator.KeyHasSerialNumber)) {
			Throw "Licensing Failed: KeyStatus $($InstantActivator.KeyStatus); KeyHasSerialNumber $($InstantActivator.KeyHasSerialNumber)"
		}

		##*===============================================
		##* UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Uninstallation'

		$InstallParameters = @(
			'-s',
			'-uninstal'
		)

		Execute-Process -Path "Maptitude 2020 Build 4720\MaptitudeInstallation.exe" -Parameters $InstallParameters -WindowStyle Hidden

		$InstallParameters = @(
			'-s',
			'-uninstal'
		)

		Execute-Process -Path "USA Country Package 2020\DataInstallation.exe" -Parameters $InstallParameters -WindowStyle Hidden


		##*===============================================
		##* POST-UNINSTALLATION
		##*===============================================
		[string]$installPhase = 'Post-Uninstallation'

		Remove-Item 'C:\Users\Default\AppData\Roaming\Caliper\Maptitude 2020' -Recurse -Force
		
		if (-not (Get-ChildItem 'C:\Users\Default\AppData\Roaming\Caliper')) {
			Remove-Item 'C:\Users\Default\AppData\Roaming\Caliper' -Recurse -Force
		}

	}
	ElseIf ($deploymentType -ieq 'Repair')
	{
		##*===============================================
		##* PRE-REPAIR
		##*===============================================
		[string]$installPhase = 'Pre-Repair'

		## Show Progress Message (with the default message)
		Show-InstallationProgress

		## <Perform Pre-Repair tasks here>

		##*===============================================
		##* REPAIR
		##*===============================================
		[string]$installPhase = 'Repair'

		## Handle Zero-Config MSI Repairs
		If ($useDefaultMsi) {
			[hashtable]$ExecuteDefaultMSISplat =  @{ Action = 'Repair'; Path = $defaultMsiFile; }; If ($defaultMstFile) { $ExecuteDefaultMSISplat.Add('Transform', $defaultMstFile) }
		Execute-MSI @ExecuteDefaultMSISplat
		}
		# <Perform Repair tasks here>

		##*===============================================
		##* POST-REPAIR
		##*===============================================
		[string]$installPhase = 'Post-Repair'

		## <Perform Post-Repair tasks here>


    }
	##*===============================================
	##* END SCRIPT BODY
	##*===============================================

	## Call the Exit-Script function to perform final cleanup operations
	Exit-Script -ExitCode $mainExitCode
}
Catch {
	[int32]$mainExitCode = 60001
	[string]$mainErrorMessage = "$(Resolve-Error)"
	Write-Log -Message $mainErrorMessage -Severity 3 -Source $deployAppScriptFriendlyName
	Show-DialogBox -Text $mainErrorMessage -Icon 'Stop'
	Exit-Script -ExitCode $mainExitCode
}
