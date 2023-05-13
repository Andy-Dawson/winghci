################################################
#
# This script performs MECM detection for GHC
# and WinGHCi.
# This script must be run from the detection node
# of the MECM application.
#
# Author: Andy Dawson
# Script Version: 1.0.0: Initial version
#
################################################

# Check to see whether Chocolatey is installed - if not, then GHC cannot have been installed
if (Test-Path -Path "C:\ProgramData\chocolatey\choco.exe" -PathType Leaf) {
	# Check to see whether GHC has been installed, and in which case, which version
	# the following should return the installed version number (e.g. '9.6.1')
	$ghcver = (([regex]::Matches((C:\ProgramData\chocolatey\choco.exe list -localonly), "(ghc\s){1}(\d+|\.)+")).Value -split " ")[1]

	# The following is the target version number - the installed version must be greater than or equal
	# to the following version for installation to be deemed acceptable.
	# Change the version number below when superceding the current version of the GHC installation.
	$ghvtargetver = "9.6.1"

	# Did we get an installed version number?
	if ($ghcver -ne $null) {
		# We got a version number
		if ([System.Version]$ghcver -ge [System.Version]$ghvtargetver) {
			# We have a suitable target version installed
			#Check whether the 'C:\Tools\ghc-<version>' folder exists
			$GHCFolder = "C:\Tools\ghc-" + $ghcver
			if (Test-Path -Path $GHCFolder) {
				# The folder exists, there's a good chance that GHC is installed...
				# Now test that the bin folder exists
				$GHCBinFolder = $GHCFolder + "\bin"
				if (Test-Path -Path $GHCBinFolder) {
					# The bin folder exists - Now look for the exe files that we want
					$WinGHCi = $GHCBinFolder + "\WinGHCi.exe"
					$GHCi = $GHCBinFolder + "\GHCi.exe"
					$WinGHCiExists  = $false
					$GHCiExists = $false
					if (Test-Path -Path $WinGHCi -PathType Leaf) {
						$WinGHCiExists = $true
					}
					if (Test-Path -Path $GHCi -PathType Leaf) {
						$GHCiExists = $true
					}
					if ($WinGHCiExists -And $GHCiExists) {
						# Both GHC and WinGHCi are in place, this should signify that everything is installed
						Write-Host "GHC installed"
					} else { # Both files do not exist
					}
				} else { # The bin folder doesnt exist
				}
			} else { # The GHC folder doesnt exist
			}
		} else { # The version returned is < target version
		}
	} else { # No version was returned by the test
	}
} else { # choco.exe doesnt exist on the machine
}
