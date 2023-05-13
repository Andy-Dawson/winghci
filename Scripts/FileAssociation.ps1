################################################
#
# This script associates .hs and .lhs files with
# WinGHCi
#
# Author: Andy Dawson
# Script Version: 1.0.0: Initial version
#
################################################

# The variables below control whether messages will be displayed or logged, and the location of the log file
$global:LogOutput = $True
$global:WriteStdOutput = $True
$global:LogFileLoc = "C:\Support\HaskellInstall.txt"

# Function to notify the user what's going on
function NotifyUser {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$Message
	)

    # If we're writing std output...
    if ($global:WriteStdOutput) {
        Write-Host $Message
    }

    # If we're writing log file output...
    if ($global:LogOutput) {
        # Find current date and time to prepend the message
        $date = (Get-Date -Format "dd/MM/yyyy HH:mm:ss").ToString()
        $Message = "[" + $date + "]: " + $Message 
        $Message | Out-File -FilePath $global:LogFileLoc -Append
    }
}

# Functions required for registering file extensions
function CreateHKCRKey {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$PathToCreate
	)
    
    # Assemble the full path to the key to be created
    $RegKeyToCreate = "Registry::HKEY_CLASSES_ROOT" + $PathToCreate
    # Test if the key already exists
    If (-NOT (Test-Path $RegKeyToCreate)) {
        # If it doesnt, create it
        NotifyUser -Message "Creating registry key at $($RegKeyToCreate)"
        # Write-Host "Creating registry key at $RegKeyToCreate"
        New-Item -Path $RegKeyToCreate -Force | Out-Null
    } else {
        # If it does, let the user know that it already exists
        NotifyUser -Message "Registry key at $($RegKeyToCreate) already exists"
        # Write-Host "Registry key at $RegKeyToCreate already exists"
    }
}

function CreateHKCRValue {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory)]
		[string]$RegistryPath,

        [Parameter(Mandatory)]
		[string]$RegistryName,

		[Parameter(Mandatory)]
		[string]$ValueToCreate
	)

    # Assemble the full path to the key to be created
    $RegKey = "Registry::HKEY_CLASSES_ROOT" + $RegistryPath
    NotifyUser -Message "Creating registry value $($RegistryName) at key $($RegKey)"
    # Write-Host "Creating registry value $($RegistryName) at key $($RegKey)"

    # Test that the key exists in preparation for creating the value
    If (Test-Path $RegKey) {
        # Check whether the registry value already exists
        $RegValueProps = Get-ItemProperty -Path $RegKey 
        if(-not $RegValueProps) {
            # The value doesnt exist, create it
            if ($RegistryName -ieq "Default") {
                # Special case; use this command to set the value
                Set-Item -Path $RegKey -Value $ValueToCreate
            } else {
                # It's a regular registry value, create it
                New-ItemProperty -Path $RegKey -Name $RegistryName -Value $ValueToCreate -PropertyType REG_SZ
            }
        } else {
            # The value exists already, update it
            NotifyUser -Message "  o The registry value already exists, it will be updated"
            # Write-Host "  o The registry value already exists, it will be updated"
            if ($RegistryName -ieq "Default") {
                # As we're updating an existing value, write out the current value
                $ExistingValue = (Get-ItemProperty -Path $regkey -Name '(default)').'(default)'
                NotifyUser -Message "  o Original value: $($ExistingValue)"
                # Write-Host "  o Original value: $($ExistingValue)" -ForegroundColor Yellow
                # Special case; use this command to set the value
                Set-Item -Path $RegKey -Value $ValueToCreate
            } else {
                # As we're updating an existing value, write out the current value
                $ExistingValue = (Get-ItemProperty -Path $regkey -Name $RegsitryName).$RegistryName
                NotifyUser -Message "  o Original value: $($ExistingValue)"
                # Write-Host "  o Original value: $($ExistingValue)" -ForegroundColor Yellow
                # It's a regular registry value, create it
                Set-ItemProperty -Path $RegKey -Name $RegistryName -Value $ValueToCreate -PropertyType REG_SZ
            }
        }
    }
}

# Create the file associations
NotifyUser -Message "Creating Haskell file associations"

# Assemble the required paths using $PSScriptRoot to determine where the script is being run from
$WinGHCiPath = $PSScriptRoot + "\WinGHCi.exe"
$WinGHCiIconPath = $PSScriptRoot + "\winghciFile.ico"

if (Test-Path -Path $($WinGHCiPath)) {
    # The application we're targeting exists in the bin folder
    if (Test-Path -Path $($WinGHCiIconPath)) {
        # The required icon exists in the same location
        NotifyUser -Message "WinGHCi.exe and winghciFile.ico both exist in the bin folder"

        # Create the base structure in the registry, if it doesnt exist already
        NotifyUser -Message "Creating base registry structure for file associations"
        CreateHKCRKey -PathToCreate "\HaskellScript"
        CreateHKCRKey -PathToCreate "\HaskellScript\DefaultIcon"
        CreateHKCRKey -PathToCreate "\HaskellScript\Shell"
        CreateHKCRKey -PathToCreate "\HaskellScript\Shell\Open"
        CreateHKCRKey -PathToCreate "\HaskellScript\Shell\Open\Command"
        # Create the registry keys for the file extensions
        CreateHKCRKey -PathToCreate "\.hs"
        CreateHKCRKey -PathToCreate "\.lhs"
    
        # Create the required registry values
        NotifyUser -Message "Creating registry values for file associations"
        CreateHKCRValue -RegistryPath "\HaskellScript\DefaultIcon" -RegistryName "Default" -ValueToCreate $WinGHCiIconPath
        CreateHKCRValue -RegistryPath "\HaskellScript\Shell\Open\Command" -RegistryName "Default" -ValueToCreate "`"$WinGHCiPath`" `"%1`""
        # Create the required registry values
        CreateHKCRValue -RegistryPath "\.hs" -RegistryName "Default" -ValueToCreate "HaskellScript"
        CreateHKCRValue -RegistryPath "\.lhs" -RegistryName "Default" -ValueToCreate "HaskellScript"

    } else { # $GHCBinFolder\winghciFile.ico does not exist
        NotifyUser -Message "winghciFile.ico NOT found in the bin folder"
        Throw "winghciFile.ico NOT found in the bin folder."
    }

} else { # $GHCBinFolder\WinGHCi.exe does not exist
    NotifyUser -Message "WinGHCi.exe NOT found in the bin folder"
    Throw "WinGHCi.exe NOT found in the bin folder."
}

# Installation of Haskell completed
NotifyUser -Message "Association of Haskell file types complete"
