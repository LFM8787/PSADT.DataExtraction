<#
.SYNOPSIS
	Data Extraction Extension script file, must be dot-sourced by the AppDeployToolkitExtension.ps1 script.
.DESCRIPTION
	Contains various functions used for extracting data from files.
.NOTES
	Extension Exit Codes:
	70301: Get-StringFromFile - Unable to get string ID from file path.
	70302: Get-StringFromFile - Unable to locate file.
	70303: Get-FileVersionInfo - Unable to locate file.
	70304: Get-IconFromFile - The file does not exist.
	70305: Get-IconFromFile - Unable to save icon into image file using Win32API method.
	70306: Get-IconFromFile - Unable to locate the saved image file.
	70307: Get-IconFromFile - The icon index is out of range.
	70308: Get-IconFromFile - There is no icon available for the icon index.
	70309: Get-IconFromFile - Unable to get display scale factor.

	Author:  Leonardo Franco Maragna
	Version: 1.0
	Date:    2023/02/08
#>
[CmdletBinding()]
Param (
)

##*===============================================
##* VARIABLE DECLARATION
##*===============================================
#region VariableDeclaration

## Variables: Extension Info
$DataExtractionExtName = "DataExtractionExtension"
$DataExtractionExtScriptFriendlyName = "Data Extraction Extension"
$DataExtractionExtScriptVersion = "1.0"
$DataExtractionExtScriptDate = "2023/02/08"
$DataExtractionExtSubfolder = "PSADT.DataExtraction"
$DataExtractionExtCustomTypesName = "DataExtractionExtension.cs"
$DataExtractionExtIconExtractorCustomTypesName = "TsudaKageyuIconExtractor.cs"
$DataExtractionExtConfigFileName = "DataExtractionConfig.xml"

## Variables: Data Extraction Script Dependency Files
[IO.FileInfo]$dirDataExtractionExtFiles = Join-Path -Path $scriptRoot -ChildPath $DataExtractionExtSubfolder
[IO.FileInfo]$DataExtractionCustomTypesSourceCode = Join-Path -Path $dirDataExtractionExtFiles -ChildPath $DataExtractionExtCustomTypesName
[IO.FileInfo]$DataExtractionIconExtractorCustomTypesSourceCode = Join-Path -Path $dirDataExtractionExtFiles -ChildPath $DataExtractionExtIconExtractorCustomTypesName
[IO.FileInfo]$DataExtractionConfigFile = Join-Path -Path $dirDataExtractionExtFiles -ChildPath $DataExtractionExtConfigFileName
if (-not $DataExtractionCustomTypesSourceCode.Exists) { throw "$($DataExtractionExtScriptFriendlyName) custom types source code file [$DataExtractionCustomTypesSourceCode] not found." }
if (-not $DataExtractionIconExtractorCustomTypesSourceCode.Exists) { throw "$($DataExtractionExtScriptFriendlyName) custom types source code file [$DataExtractionIconExtractorCustomTypesSourceCode] not found." }
if (-not $DataExtractionConfigFile.Exists) { throw "$($DataExtractionExtScriptFriendlyName) XML configuration file [$DataExtractionConfigFile] not found." }

## Import variables from XML configuration file
[Xml.XmlDocument]$xmlDataExtractionConfigFile = Get-Content -LiteralPath $DataExtractionConfigFile -Encoding UTF8
[Xml.XmlElement]$xmlDataExtractionConfig = $xmlDataExtractionConfigFile.DataExtraction_Config

#  Get Config File Details
[Xml.XmlElement]$configDataExtractionConfigDetails = $xmlDataExtractionConfig.Config_File

#  Check compatibility version
$configDataExtractionConfigVersion = [string]$configDataExtractionConfigDetails.Config_Version
#$configDataExtractionConfigDate = [string]$configDataExtractionConfigDetails.Config_Date

try {
	if ([version]$DataExtractionExtScriptVersion -ne [version]$configDataExtractionConfigVersion) {
		Write-Log -Message "The $($DataExtractionExtScriptFriendlyName) version [$([version]$DataExtractionExtScriptVersion)] is not the same as the $($DataExtractionExtConfigFileName) version [$([version]$configDataExtractionConfigVersion)]. Problems may occurs." -Severity 2 -Source ${CmdletName}
	}
}
catch {}

#  Get Data Extraction General Options
[Xml.XmlElement]$xmlDataExtractionOptions = $xmlDataExtractionConfig.DataExtraction_Options
$configDataExtractionGeneralOptions = [PSCustomObject]@{
	ExitScriptOnError = Invoke-Expression -Command 'try { [boolean]::Parse([string]($xmlDataExtractionOptions.ExitScriptOnError)) } catch { $false }'
}

#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings

#region Function Get-StringFromFile
Function Get-StringFromFile {
	<#
	.SYNOPSIS
		Gets a string contained in an executbale or system file.
	.DESCRIPTION
		Gets a string contained in an executbale or system file.
	.PARAMETER Path
		Fully qualified path name of the executable file.
	.PARAMETER StringID
		String number used as id by the api function.
	.PARAMETER ContinueOnError
		Continue if an error is encountered. Default is: $true.
	.PARAMETER DisableFunctionLogging
		If specified disables logging messages to the script log file.
	.INPUTS
		None
		You cannot pipe objects to this function.
	.OUTPUTS
		None
		The string id contained in the file.
	.EXAMPLE
		Get-StringFromFile -Path 'c:\windows\system32\user32.dll' -StringID 800
	.NOTES
		Author: Leonardo Franco Maragna
		Part of Data Extraction Extension
	.LINK
		https://github.com/LFM8787/PSADT.DataExtraction
		http://psappdeploytoolkit.com
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true)]
		[IO.FileInfo]$Path,
		[Parameter(Mandatory = $true)]
		[ValidateScript({ [int32]$_ -ge 0 })]
		[int32]$StringID,
		[Parameter(Mandatory = $false)]
		[boolean]$ContinueOnError = $true,
		[switch]$DisableFunctionLogging
	)

	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {
		if ($Path.Exists) {
			try {
				$String = [PSADT.DataExtraction]::ExtractStringFromFile($Path.FullName, $StringID)

				if ($null -eq $String) {
					if (-not ($DisableFunctionLogging)) { Write-Log -Message "The returned string with string ID [$StringID] from file path [$Path] appers to be null." -Severity 2 -Source ${CmdletName} }
				}

				return $String
			}
			catch {
				if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to get string ID [$StringID] from file path [$Path].`r`n$(Resolve-Error)" -Severity 3 -Source ${CmdletName} }
				if (-not $ContinueOnError) {
					if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70301 }
					throw "Unable to get string ID [$StringID] from file path [$Path]: $($_.Exception.Message)"
				}
				return
			}
		}
		else {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to locate file [$Path]." -Severity 3 -Source ${CmdletName} }
			if (-not $ContinueOnError) {
				if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70302 }
				throw "Unable to locate file [$Path]."
			}
			return
		}
	}
	End {
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion


#region Function Get-ApplicationMuiCache
Function Get-ApplicationMuiCache {
	<#
	.SYNOPSIS
		Gets information about an application name and company from the MuiCache in registry.
	.DESCRIPTION
		Gets information about an application name and company from the MuiCache in registry.
	.PARAMETER ProcessName
		Name of the process.
	.INPUTS
		None
		You cannot pipe objects to this function.
	.OUTPUTS
		None
		Returns a hashtable containing the process name, the mui cache name and company.
	.EXAMPLE
		Get-ApplicationMuiCache -ProcessName 'notepad'
	.NOTES
		Author: Leonardo Franco Maragna
		Part of Data Extraction Extension
	.LINK
		https://github.com/LFM8787/PSADT.DataExtraction
		http://psappdeploytoolkit.com
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true)]
		[string]$ProcessName
	)

	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {
		## Extract extension if given
		$ProcessName = [IO.Path]::GetFileNameWithoutExtension($ProcessName)

		$RegistryPaths = @()
		$RegistryPaths += Convert-RegistryPath -Key "Registry::HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\MuiCache"
		$RegistryPaths += "Registry::HKEY_CLASSES_ROOT\Local Settings\Software\Microsoft\Windows\Shell\MuiCache"

		foreach ($RegistryPath in $RegistryPaths) {
			## Get MuiCache items in the current RegistryPath
			$MuiCacheItems = Get-Item -Path $RegistryPath -ErrorAction SilentlyContinue

			if ($MuiCacheItems) {
				$ProcessMuiCacheItem = $MuiCacheItems.Property -replace ".FriendlyAppName", "" -replace ".ApplicationCompany", "" | Sort-Object -Unique | Where-Object { $_ -like "*$($ProcessName).exe" } | Select-Object -First 1
	
				if ($ProcessMuiCacheItem) {
					$FriendlyAppName = ""
					$ApplicationCompany = ""
					$FriendlyAppName = Get-ItemPropertyValue -Path $RegistryPath -Name "$($ProcessMuiCacheItem).FriendlyAppName" -ErrorAction SilentlyContinue
					$ApplicationCompany = Get-ItemPropertyValue -Path $RegistryPath -Name "$($ProcessMuiCacheItem).ApplicationCompany" -ErrorAction SilentlyContinue

					return [PSCustomObject]@{
						ProcessName        = $ProcessName
						FriendlyAppName    = $FriendlyAppName
						ApplicationCompany = $ApplicationCompany
					}
				}
			}
		}

		## Return empty object if not found
		return [PSCustomObject]@{
			ProcessName        = $ProcessName
			FriendlyAppName    = ""
			ApplicationCompany = ""
		}
	}
	End {
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion


#region Function Get-FileVersionInfo
Function Get-FileVersionInfo {
	<#
	.SYNOPSIS
		Gets the file version information from an executable file, like FileDescription and CompanyName properties.
	.DESCRIPTION
		Gets the file version information from an executable file, like FileDescription and CompanyName properties.
	.PARAMETER Path
		Fully qualified path name of the executable file.
	.PARAMETER ContinueOnError
		Continue if an error is encountered. Default is: $true.
	.PARAMETER DisableFunctionLogging
		If specified disables logging messages to the script log file.
	.INPUTS
		None
		You cannot pipe objects to this function.
	.OUTPUTS
		None
		Returns a hashtable with the FileVersionInfo from the file.
	.EXAMPLE
		Get-FileVersionInfo -Path 'C:\Windows\explorer.exe'
	.NOTES
		Author: Leonardo Franco Maragna
		Part of Data Extraction Extension
	.LINK
		https://github.com/LFM8787/PSADT.DataExtraction
		http://psappdeploytoolkit.com
	#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[IO.FileInfo]$Path,
		[Parameter(Mandatory = $false)]
		[boolean]$ContinueOnError = $true,
		[switch]$DisableFunctionLogging
	)

	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
	}
	Process {
		if ($Path.Exists) {
			try {
				return [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
			}
			catch {}
		}
		else {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to locate file [$Path]." -Severity 3 -Source ${CmdletName} }
			if (-not $ContinueOnError) {
				if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70303 }
				throw "Unable to locate file [$Path]."
			}
			return
		}
	}
	End {
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion


#region Function Get-IconFromFile
Function Get-IconFromFile {
	<#
	.SYNOPSIS
		Extracts an icon contained in an executable or library.
	.DESCRIPTION
		Extracts an icon contained in an executable or library.
		Uses TsudaKageyu IconExtractor class and native methods.
	.PARAMETER SystemIcon
		Used to extract extension defined system icons.
	.PARAMETER Path
		Path to the executable or library containing the icons.
	.PARAMETER IconIndex
		Icon index number.
	.PARAMETER SavePath
		Path where to save the image file.
	.PARAMETER TargetSize
		Dimension of the icon to search and extract, if it is not found, use the one immediately above.
	.PARAMETER ScaleToDPI
		Scale the TargetSize using the current dpi scale factor.
	.PARAMETER SaveFormat
		One of the system defined image format in [System.Drawing.Imaging.ImageFormat].
	.PARAMETER ContinueOnError
		Continue if an error is encountered. Default is: $true.
	.PARAMETER DisableFunctionLogging
		If specified disables logging messages to the script log file.
	.INPUTS
		None
		You cannot pipe objects to this function.
	.OUTPUTS
		None
		This function does not generate any output.
	.EXAMPLE
		Get-IconFromFile -Path 'C:\Windows\explorer.exe' -IconIndex 2 -SavePath 'C:\icon.png' -SaveFormat Png -TargetSize 48 -ScaleToDPI
	.NOTES
		Author: Leonardo Franco Maragna
		Part of Data Extraction Extension
	.LINK
		https://github.com/LFM8787/PSADT.DataExtraction
		http://psappdeploytoolkit.com
	#>
	[CmdletBinding(DefaultParameterSetName = "Path")]
	Param (
		[Parameter(Mandatory = $true,
			ParameterSetName = "SystemIcon")]
		[ValidateSet("Application", "Asterisk", "Error", "Exclamation", "Hand", "Info", "Information", "Lock", "MultipleWindows", "Question", "Shield", "Stop", "Warning", "WinLogo")]
		[string]$SystemIcon,
		[Parameter(Mandatory = $true,
			ParameterSetName = "Path")]
		[ValidateNotNullorEmpty()]
		[IO.FileInfo]$Path,
		[Parameter(Mandatory = $false,
			ParameterSetName = "Path")]
		[ValidateScript({ [int32]$_ -ge 0 })]
		[int32]$IconIndex = 0,
		[Parameter(Mandatory = $true)]
		[ValidateNotNullorEmpty()]
		[IO.FileInfo]$SavePath,
		[Parameter(Mandatory = $false)]
		[ValidateSet("Bmp", "Emf", "Wmf", "Gif", "Jpeg", "Png", "Tiff", "Exif", "Icon")]
		[System.Drawing.Imaging.ImageFormat]$SaveFormat = "Png",
		[Parameter(Mandatory = $false)]
		[ValidateScript({ [int32]$_ -ge 0 })]
		[int32]$TargetSize = 48,
		[Parameter(Mandatory = $false)]
		[boolean]$ScaleToDPI = $true,
		[Parameter(Mandatory = $false)]
		[boolean]$ContinueOnError = $true,
		[switch]$DisableFunctionLogging
	)

	Begin {
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header

		## Force function logging if debugging
		if ($configToolkitLogDebugMessage) { $DisableFunctionLogging = $false }
	}
	Process {
		## System icons defined in native library
		if ($SystemIcon) {
			$Path = $envImageresLibraryPath
			$IconIndex = switch -regex ($SystemIcon) {
				"Application" { 260 }
				"(Asterisk|Info|Information)" { 76 }
				"(Error|Hand|Stop)" { 93 }
				"(Exclamation|Warning)" { 79 }
				"Lock" { 54 }
				"MultipleWindows" { 261 }
				"Question" { 94 }
				"Shield" { 73 }
				"WinLogo" { 1 }
			}
		}

		## Check if path exists
		if (-not $Path.Exists) {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "The file [$Path] does not exist." -Severity 3 -Source ${CmdletName} }
			if (-not $ContinueOnError) {
				if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70304 }
				throw "The file [$Path] does not exist."
			}
			return
		}

		## Native methods to retrieve a system large icon
		[scriptblock]$UseWin32APIMethod = {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "Selected icon from index [$IconIndex] of [$Path] with max size [32x32] will be saved into image file [$SavePath]." -Source ${CmdletName} }
			try {
				if ($IconIndex -gt 0) {
					#  Extract the icon located in the index parameter
					$null = [PSADT.ExtractData]::ExtractIcon($Path, $IconIndex, $true).ToBitmap().Save($SavePath, $SaveFormat)
				}
				else {
					#  Extract the first icon
					$null = [System.Drawing.Icon]::ExtractAssociatedIcon($Path).ToBitmap().Save($SavePath, $SaveFormat)
				}
			}
			catch {
				if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to save icon from index [$IconIndex] of [$Path] into image file [$SavePath] using Win32API method.`r`n$(Resolve-Error)" -Severity 3 -Source ${CmdletName} }
				if (-not $ContinueOnError) {
					if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70305 }
					throw "Unable to save icon from index [$IconIndex] of [$Path] into image file [$SavePath] using Win32API method: $($_.Exception.Message)"
				}
			}
		}

		## Test if the saved image exists
		[scriptblock]$TestSaveImage = {
			$SavePath.Refresh()
			if ($SavePath.Exists) {
				if (-not ($DisableFunctionLogging)) { Write-Log -Message "Destination image file [$SavePath] created successfully." -Source ${CmdletName} }
				return $SavePath
			}
			else {
				if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to locate the saved image file [$SavePath]." -Severity 3 -Source ${CmdletName} }
				if (-not $ContinueOnError) {
					if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70306 }
					throw "Unable to locate the saved image file [$SavePath]."
				}
				return
			}
		}
		
		## Get icons from file
		try {
			$IconExtractor = [TsudaKageyu.IconExtractor]($Path.FullName)
		}
		catch {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to retrieve available icons from file [$Path], native Win32API method will be used.`r`n$(Resolve-Error)" -Severity 2 -Source ${CmdletName} }
			Invoke-Command -ScriptBlock $UseWin32APIMethod -NoNewScope
			Invoke-Command -ScriptBlock $TestSaveImage -NoNewScope
			return
		}

		## Check if index is in range
		if ($IconIndex -gt $IconExtractor.Count) {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "The icon index [$IconIndex] is out of range, max index of [$Path] is [$($IconExtractor.Count)]." -Severity 3 -Source ${CmdletName} }
			if (-not $ContinueOnError) {
				if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70307 }
				throw "The icon index [$IconIndex] is out of range, max index of [$Path] is [$($IconExtractor.Count)]."
			}
			return
		}
		else {
			#  Get IconDirectory object from specified index
			try {
				$IconDirectory = $IconExtractor.GetIcon($IconIndex)
			}
			catch {
				if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to retrieve available icons in index [$IconIndex] from file [$Path], native Win32API method will be used.`r`n$(Resolve-Error)" -Severity 2 -Source ${CmdletName} }
				Invoke-Command -ScriptBlock $UseWin32APIMethod -NoNewScope
				Invoke-Command -ScriptBlock $TestSaveImage -NoNewScope
				return
			}
		}

		## Get available icons in the index
		$IconsAvailable = @()
		try {
			[TsudaKageyu.IconUtil]::Split($IconDirectory) | ForEach-Object {
				$IconsAvailable += [PSCustomObject]@{
					Icon     = $_
					Width    = $_.Width
					Height   = $_.Height
					BitDepth = [TsudaKageyu.IconUtil]::GetBitCount($_)
				}
			}
		}
		catch {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to retrieve available icons in index [$IconIndex] from file [$Path], native Win32API method will be used.`r`n$(Resolve-Error)" -Severity 2 -Source ${CmdletName} }
			Invoke-Command -ScriptBlock $UseWin32APIMethod -NoNewScope
			Invoke-Command -ScriptBlock $TestSaveImage -NoNewScope
			return
		}

		if ($IconsAvailable.Count -eq 0) {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "There is no icon available for the icon index [$IconIndex] of [$Path]." -Severity 3 -Source ${CmdletName} }
			if (-not $ContinueOnError) {
				if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70308 }
				throw "There is no icon available for the icon index [$IconIndex] of [$Path]."
			}
			return
		}
		else {
			$IconsAvailable = $IconsAvailable | Sort-Object -Property @{Expression = "Width"; Descending = $true }, @{Expression = "Height"; Descending = $true }, @{Expression = "BitDepth"; Descending = $false }
		}

		## Scale target size icon to current scale factor
		if ($ScaleToDPI) {
			if ($dpiScale -isnot [int32]) {
				if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to get display scale factor." -Severity 2 -Source ${CmdletName} }
				if (-not $ContinueOnError) {
					if ($configDataExtractionGeneralOptions.ExitScriptOnError) { Exit-Script -ExitCode 70309 }
					throw "Unable to get display scale factor."
				}
			}
			elseif ($dpiScale -gt 100) {
				if (-not ($DisableFunctionLogging)) { Write-Log -Message "Original target size [$TargetSize] will be scaled to a [$($dpiScale)%] factor." -Source ${CmdletName} }
				$TargetSize = [System.Math]::Round($TargetSize * ($dpiScale / 100))
			}
		}

		## Get the better icon available
		$SelectedIcon = @()

		#  Select the equal size or next with the highest quality
		$SelectedIcon = $IconsAvailable | Where-Object { $_.Width -ge $TargetSize -or $_.Height -ge $TargetSize } | Select-Object -Last 1

		if ($SelectedIcon.Count -eq 0) {
			#  Select the biggest one with the highest quality
			$BestBitDepth = ($IconsAvailable.BitDepth | Measure-Object -Maximum).Maximum
			$SelectedIcon = $IconsAvailable | Group-Object -Property BitDepth | Where-Object { $_.Name -eq $BestBitDepth } | Select-Object -ExpandProperty Group | Sort-Object -Property Width, Height | Select-Object -Last 1
		}

		## Save the selected icon
		if (-not ($DisableFunctionLogging)) { Write-Log -Message "Selected icon from index [$IconIndex] of [$Path] with size [$($SelectedIcon.Width)x$($SelectedIcon.Height)] and bit depth [$($SelectedIcon.BitDepth)] will be saved into image file [$SavePath]." -Source ${CmdletName} }
		try {
			$null = [TsudaKageyu.IconUtil]::ToBitmap($SelectedIcon.Icon).Save($SavePath, $SaveFormat)
			if ($?) {
				Invoke-Command -ScriptBlock $TestSaveImage -NoNewScope
				return
			}
		}
		catch {
			if (-not ($DisableFunctionLogging)) { Write-Log -Message "Unable to save icon from index [$IconIndex] of [$Path] into image file [$SavePath], native Win32API method will be used.`r`n$(Resolve-Error)" -Severity 3 -Source ${CmdletName} }
			Invoke-Command -ScriptBlock $UseWin32APIMethod -NoNewScope
			Invoke-Command -ScriptBlock $TestSaveImage -NoNewScope
			return
		}
	}
	End {
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
}
#endregion

#endregion
##*===============================================
##* END FUNCTION LISTINGS
##*===============================================

##*===============================================
##* SCRIPT BODY
##*===============================================
#region ScriptBody

if ($scriptParentPath) {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] dot-source invoked by [$(((Get-Variable -Name MyInvocation).Value).ScriptName)]" -Source $DataExtractionExtName
}
else {
	Write-Log -Message "Script [$($MyInvocation.MyCommand.Definition)] invoked directly" -Source $DataExtractionExtName
}

## Add the custom types required for the toolkit
if (-not ([Management.Automation.PSTypeName]"PSADT.DataExtraction").Type) {
	[string[]]$ReferencedAssemblies = "System.Drawing"
	Add-Type -Path $DataExtractionCustomTypesSourceCode -ReferencedAssemblies $ReferencedAssemblies -IgnoreWarnings -ErrorAction Stop
}

if (-not ([Management.Automation.PSTypeName]"TsudaKageyu.IconExtractor").Type) {
	[string[]]$ReferencedAssemblies = "System.Drawing"
	Add-Type -Path $DataExtractionIconExtractorCustomTypesSourceCode -ReferencedAssemblies $ReferencedAssemblies -IgnoreWarnings -ErrorAction Stop
}

#endregion
##*===============================================
##* END SCRIPT BODY
##*===============================================