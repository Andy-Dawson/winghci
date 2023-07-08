################################################
#
# This script installs or upgrades Haskell using Chocolatey
#
# Author: Andy Dawson
# Script Version: 1.0.0: Initial version
#                 1.0.1: Using Start-Process to wait for choco command completion
#                 1.0.2: Added file extension registration PowerShell script call
#                 1.0.3: Merged file extension registration script with main script
#                 1.0.4: Removed '-localonly' switch from choco commands [As of Choco 2.0.0
#                        this paramter, along with others, has been removed]
#
################################################

# The variables below control whether messages will be displayed or logged, and the location of the log file
$global:LogOutput = $True
$global:WriteStdOutput = $True
$global:LogFileLoc = "C:\Support\HaskellInstall.txt"

# This is how the file should be called from SCCM:
# Powershell.exe -ExecutionPolicy ByPass -File "InstallOrUpgradeHaskell.ps1"

# If running this script manually you need to have run
# Set-ExecutionPolicy Unrestricted
# Before this script can be run

# Packages to install for Haskell:
#  GHC
#  cabal (installed as part of the GHC installation)
#  haskell-stack

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


# script starts here
NotifyUser -Message "Starting Haskell install"

# Check for Chocolatey, and that it is up-to-date and correct this, if this is not the case:
try {
    if($env:Path -match "chocolatey") {
        $vers = ([regex]::Matches((choco upgrade chocolatey --yes --force --whatif), "(?<=(v|\s))(0|[1-9]\d*)(\.(0|[1-9]\d*)){1,3}")).Value # gives the versions that are installed and available as two (or more) values (assuming Chocolatey is installed) and allows a comparison and upgrade, or installation
        if (@($vers).Count -ge 2) { # We have an installed version and an available version
            # Note that depending on exactly what gets returned, we may have pending reboot warnings etc. that complicate things
            # We may get two values to compare, or more than two, however we typically want the first and last values to compare
            # Check the versions to ensure that they are the same
            $FinalArrayIndex = ($vers.count - 1)
            if ($vers[0] -eq $vers[$FinalArrayIndex]) {
                NotifyUser -Message "Chocolatey is installed and up-to-date"
            } else {
                # Chocolatey is installed, but out of date and needs upgrading
                NotifyUser -Message "Upgrading Chocolatey"
                Start-Process -FilePath "choco" -ArgumentList "upgrade chocolatey --yes" -Wait -NoNewWindow
                # start /wait choco upgrade chocolatey --yes
				#Start-Process <path to exe> -NoNewWindow -Wait
            }
        }
    } else {
        # Chocolatey is not installed and should be
        NotifyUser -Message "Installing Chocolatey"
        iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }
} catch { }

# Get the version of Choco that we have on the system - We'll want this later
$ChocoVer = & choco -v
if ([System.Version]$ChocoVer -ge [System.Version]"2.0.0") {
    Import-Module $env:ChocolateyInstall\helpers\chocolateyProfile.psm1
}

# Install or upgrade GHC
# Check to see whether GHC is already installed
try {
    $ghcvers = ([regex]::Matches((choco list), "(ghc\s){1}(\d+|\.)+")).Value # Returns ghc x.y.z
    if ($ghcvers.Count -eq 1) {
        # It appears we have GHC installed, check whether we can upgrade it
        NotifyUser -Message "GHC installed - checking for upgrade"
        # Get the current version from the above test
        $ghcver = ($ghcvers -split " ")[1]
        # Now get the available version
        $ghcavailablever = ([regex]::Matches((choco upgrade ghc --yes --force --whatif), "(?<=(ghc v|\s))(\d+|\.)+(?=\s)")).Value
        if ($ghcver -eq $ghcavailablever) {
            NotifyUser -Message "GHC is installed and up-to-date"
        } else {
            NotifyUser -Message "Upgrading GHC"
            Start-Process -FilePath "choco" -ArgumentList "upgrade ghc --yes" -Wait -NoNewWindow
            # Now remove any old versions that are hanging around. Choco doesn't list them as installed once the new version is in place
            # so we can remove the old folder and any child items remaining within it.
            $GHCInstalledVers = Get-ChildItem -Path "C:\Tools" -Directory -Filter "ghc-*" -Depth 0 | Sort-Object -Property FullName -Descending
            $NumGHCInstalledVers = $GHCInstalledVers.Count
            if ($NumGHCInstalledVers -gt 1) {
                # We have multiple versions of GHC installed and need to deal with this
                NotifyUser -Message "Multiple versions of GHC found on disk, tidying up..."
                for ($i = 1; $i -le ($NumGHCInstalledVers - 1); $i++) {
                    NotifyUser -Message "Removing GHC installation at $($GHCInstalledVers[$i].FullName)"
                    Remove-Item -Path $GHCInstalledVers[$i].FullName -Recurse -Force
                }
            } elseif ($NumGHCInstalledVers -eq 1) {
                # We have exactly one version installed
                NotifyUser -Message "A single version of GHC has been located on disk"
            } else {
                # We appear to have zero GHC installations, which will cause an issue
                NotifyUser -Message "No installations of GHC have been located during the GHC post-upgrade operations - This will cause an issue"
            }
        }
    } else {
        # GHC is not installed, install it
        NotifyUser -Message "Installing GHC"
        Start-Process -FilePath "choco" -ArgumentList "install GHC -y --package-parameters=`"'/globalinstall:True'`"" -Wait -NoNewWindow
    }
} catch { }

# Refresh the Chocolatey environment so we don't have to close and re-open the PowerShell/cmd window
NotifyUser -Message "Refreshing the Chocolatey environment"
if ([System.Version]$ChocoVer -ge [System.Version]"2.0.0") {
    # This seems to work, whereas the command within the 'else' no longer does with Choco >= 2.0.0
    # & refreshenv
    Update-SessionEnvironment
} else {
    # This seems to work with Choco < 2.0.0
    Start-Process -FilePath "refreshenv" -Wait -NoNewWindow
}

# During the GHC install, there's a step to rename a folder inside C:\Tools that seems to fail (anti-virus scanning the dropped files perhaps?)
# Need to ensure that this folder has been renamed before going any further
NotifyUser -Message "Checking to see whether the GHC folder needs renaming"
$ghcver = (([regex]::Matches((choco list), "(ghc\s){1}(\d+|\.)+")).Value -split " ")[1]
if ($ghcver -ne $null) {
	#Check whether the 'C:\Tools\ghc-<version>' folder exists
	$folder = "C:\Tools\ghc-" + $ghcver
    NotifyUser -Message "Folder we want: $($folder)"
    $folderToRename = "C:\Tools\ghc-" + $ghcver + "-x86_64-unknown-mingw32"
    NotifyUser -Message "Folder that we may have instead: $($folderToRename)"
    if (Test-Path -PathType Container $folder) {
        NotifyUser -Message "The correct folder exists, nothing to do"
	} elseif (Test-Path -PathType Container $folderToRename) {
        NotifyUser -Message "Incorrect path found"
        # Now try to rename the folder
		$numtries = 0
		do {
			$numtries++
			try {
                NotifyUser -Message "Trying to rename the path - (attempt $($numtries))"
				Rename-Item -Path $folderToRename -NewName $folder
				return
			} catch {
                NotifyUser -Message "Caught an exception $($_.Exception.InnerException.Message)"
				Start-Sleep 30
			}
		} while ($numtries -lt 10)
		# Throw an exception after 10 retries
        NotifyUser -Message "Tried to rename the path 10 times, giving up"
		Throw 'Maximum of 10 attempts to rename the GHC folder reached, giving up'
	} else {
		# The folder is not the one we're expecting and we need to stop here...
        NotifyUser -Message "Did NOT find the expected folder to rename $($folderToRename)"
		Throw 'Did NOT find the expected folder to rename for C:\Tools\GHC-X.Y.Z'
	}
} else {
    NotifyUser -Message "Did NOT find a version of Haskell listed as installed, quitting"
    Throw 'Did NOT find a version of Haskell listed as installed, quitting'
}

# Now install/upgrade Haskell-Stack
NotifyUser -Message "Starting Haskell-Stack installation"

# Check to see whether Haskell-Stack is already installed
try {
    $hsvers = ([regex]::Matches((choco list), "(haskell-stack\s){1}(\d+|\.)+")).Value # Returns ghc x.y.z
    if ($hsvers.Count -eq 1) {
        # It appears we have Haskell-Stack installed, check whether we can upgrade it
		NotifyUser -Message "Haskell-Stack appears to be installed, checking for possible upgrade"
        # Get the current version from the above test
        $hscver = ($hscvers -split " ")[1]
        # Now get the available version
        $hscavailablever = ([regex]::Matches((choco upgrade haskell-stack --yes --force --whatif), "(?<=(haskell-stack v|\s))(\d+|\.)+(?=\s)")).Value
        if ($hscver -eq $hscavailablever) {
		    NotifyUser -Message "Haskell-Stack is installed and up-to-date"
        } else {
		    NotifyUser -Message "Upgrading Haskell-Stack"
            Start-Process -FilePath "choco" -ArgumentList "upgrade haskell-stack --yes" -Wait -NoNewWindow
        }
    } else {
        # Haskell-Stack is not installed, install it
		NotifyUser -Message "Installing Haskell-Stack"
        Start-Process -FilePath "choco" -ArgumentList "install haskell-stack -y" -Wait -NoNewWindow
    }
} catch { }

# Refresh the Chocolatey environment so we don't have to close and re-open the PowerShell/cmd window
NotifyUser -Message "Refreshing the Chocolatey environment (again)"
if ([System.Version]$ChocoVer -ge [System.Version]"2.0.0") {
    # This seems to work, whereas the command within the 'else' no longer does with Choco >= 2.0.0
    # & refreshenv
    Update-SessionEnvironment
} else {
    # This seems to work with Choco < 2.0.0
    Start-Process -FilePath "refreshenv" -Wait -NoNewWindow
}

# Now copy the WinGHCi items into the correct location
NotifyUser -Message "Copying WinGHCi files to GHC bin folder"
$GHCBinFolder = $folder + "\bin"
Copy-Item -Path ".\WinGHCi\*" -Destination $GHCBinFolder -Force

# Finally, create a link (or two) in the global start menu for WinGHCi (+anything else we need to point people at)
NotifyUser -Message "Creating Start Menu links"

# Create some variables we're going to need
$StartMenuLoc = $env:ProgramData + "\Microsoft\Windows\Start menu\Programs"
$StartMenuHaskellFolderLoc = $StartMenuLoc + "\Haskell"

# Create a Haskell folder in the start menu for all users
New-Item -Path $StartMenuHaskellFolderLoc -ItemType Directory

# Create a shortcut to the WinGHCi executable
$WinGHCiLnkLoc = $StartMenuHaskellFolderLoc + "\WinGHCi.lnk"
$WinGHCiExeLoc = $GHCBinFolder + "\WinGHCi.exe"

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($WinGHCiLnkLoc)
$Shortcut.TargetPath = $WinGHCiExeLoc
$Shortcut.Save()

# Create a shortcut to the GHCi executable
$GHCiLnkLoc = $StartMenuHaskellFolderLoc + "\GHCi.lnk"
$GHCiExeLoc = $GHCBinFolder + "\GHCi.exe"

$WshShell = New-Object -comObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut($GHCiLnkLoc)
$Shortcut.TargetPath = $GHCiExeLoc
$Shortcut.Save()

# Create the file associations
NotifyUser -Message "Creating Haskell file associations"

# Assemble the required paths
$WinGHCiPath = $GHCBinFolder + "\WinGHCi.exe"
$WinGHCiIconPath = $GHCBinFolder + "\winghciFile.ico"

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
NotifyUser -Message "Finished Haskell installation"
