# PSADT.DataExtraction
Extension for PowerShell App Deployment Toolkit to get strings and icons contained in executables and library as well as cache MUI information from registry.

## Features
- Gets a string inside any file.
- MUI format buttons and messages using system multilanguage support.
- Get MUI information of any process previously cached.
- Get information from files.
- Extract any icon in any size from any executable or library.
- Scale the extracted icon based in scale factor.
- Native methods used to perform the tasks.
- *ContinueOnError* and *ExitScriptOnError* support.

## Disclaimer
```diff
- Test the functions before production.
- Make a backup before applying.
- Check the config file options description.
- Run AppDeployToolkitHelp.ps1 for more help and parameter descriptions.
```

## Functions
* **Get-StringFromFile** - Gets a string contained in an executbale or system file.
* **Get-ApplicationMuiCache** - Gets information about an application name and company from the MuiCache in registry.
* **Get-FileVersionInfo** - Gets the file version information from an executable file, like FileDescription and CompanyName properties.
* **Get-IconFromFile** - Extracts an icon contained in an executable or library.

## Usage
```PowerShell
# Gets a string contained in an executbale or system file.
Get-StringFromFile -Path 'c:\windows\system32\user32.dll' -StringID 800
  OK

# Gets information about an application name and company from the MuiCache in registry.
Get-ApplicationMuiCache -ProcessName 'notepad'
  ProcessName = notepad
  FirendlyAppName = Windows Notepad
  ApplicationCompany = Microsoft Corporation

# Gets the file version information from an executable file, like FileDescription and CompanyName properties.
Get-FileVersionInfo -Path 'C:\Windows\explorer.exe'
  OriginalFilename  : EXPLORER.EXE.MUI
  FileDescription   : Windows Explorer
  ProductName       : Microsoft® Windows® Operating System
  Comments          : 
  CompanyName       : Microsoft Corporation
  FileName          : C:\windows\explorer.exe
  FileVersion       : 10.0.22000.184 (WinBuild.160101.0800)
  ProductVersion    : 10.0.22000.184
  IsDebug           : False
  IsPatched         : False
  IsPreRelease      : False
  IsPrivateBuild    : False
  IsSpecialBuild    : False
  Language          : English (United States)
  LegalCopyright    : © Microsoft Corporation. All rights reserved.
  LegalTrademarks   : 
  PrivateBuild      : 
  SpecialBuild      :
  FileVersionRaw    : 10.0.22000.1455
  ProductVersionRaw : 10.0.22000.1455
 
 # Extracts an icon contained in an executable or library.
 Get-IconFromFile -Path 'C:\Windows\explorer.exe' -IconIndex 2 -SavePath 'C:\icon.png' -SaveFormat Png -TargetSize 48 -ScaleToDPI
```

## Extension Exit Codes
|Exit Code|Function|Exit Code Detail|
|:----------:|:--------------------|:-|
|70301|Get-StringFromFile|Unable to get string ID from file path.|
|70302|Get-StringFromFile|Unable to locate file.|
|70303|Get-FileVersionInfo|Unable to locate file.|
|70304|Get-IconFromFile|The file does not exist.|
|70305|Get-IconFromFile|Unable to save icon into image file using Win32API method.|
|70306|Get-IconFromFile|Unable to locate the saved image file.|
|70307|Get-IconFromFile|The icon index is out of range.|
|70308|Get-IconFromFile|There is no icon available for the icon index.|
|70309|Get-IconFromFile|Unable to get display scale factor.|

## How to Install
#### 1. Download and extract into Toolkit folder.
#### 2. Edit *AppDeployToolkitExtensions.ps1* file and add the following lines.
#### 3. Create an empty array (only once if multiple extensions):
```PowerShell
## Variables: Extensions to load
$ExtensionToLoad = @()
```
#### 4. Add Extension Path and Script filename (repeat for multiple extensions):
```PowerShell
$ExtensionToLoad += [PSCustomObject]@{
	Path   = "PSADT.DataExtraction"
	Script = "DataExtractionExtension.ps1"
}
```
#### 5. Complete with the remaining code to load the extension (only once if multiple extensions):
```PowerShell
## Loading extensions
foreach ($Extension in $ExtensionToLoad) {
	$ExtensionPath = $null
	if ($Extension.Path) {
		[IO.FileInfo]$ExtensionPath = Join-Path -Path $scriptRoot -ChildPath $Extension.Path | Join-Path -ChildPath $Extension.Script
	}
	else {
		[IO.FileInfo]$ExtensionPath = Join-Path -Path $scriptRoot -ChildPath $Extension.Script
	}
	if ($ExtensionPath.Exists) {
		try {
			. $ExtensionPath
		}
		catch {
			Write-Log -Message "An error occurred while trying to load the extension file [$($ExtensionPath)].`r`n$(Resolve-Error)" -Severity 3 -Source $appDeployToolkitExtName
		}
	}
	else {
		Write-Log -Message "Unable to locate the extension file [$($ExtensionPath)]." -Severity 2 -Source $appDeployToolkitExtName
	}
}
```

## Requirements
* Powershell 5.1+
* PSAppDeployToolkit 3.8.4+

## External Links
* [PowerShell App Deployment Toolkit](https://psappdeploytoolkit.com/)
* [TsudaKageyu/IconExtractor: Icon Extractor Library for .NET](https://github.com/TsudaKageyu/IconExtractor)
