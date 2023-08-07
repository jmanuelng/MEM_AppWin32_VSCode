<#
.DESCRIPTION
    Automate the process of installing software using the Windows Package Manager (Winget).
    Structure can be adjusted to support any Winget application by changing the WingetAppID variable.

.EXAMPLE
    $WingetAppID = "Your.ApplicationID"
    .\ScriptName.ps1
    Installs the application associated with "Your.ApplicationID" using Winget, 
    assuming Winget is available on the device.

.NOTES
    Inspired by John Bryntze; Twitter: @JohnBryntze

.DISCLAIMER
    This script is delivered as-is without any guarantees or warranties. Always ensure 
    you have backups and take necessary precautions when executing scripts, particularly 
    in production environments.

.LAST MODIFIED
    August 7th, 2023

#>

#region Functions

function Find-WingetPath {
    <#
    .SYNOPSIS
        Locates the winget.exe executable within a system.

    .DESCRIPTION
        Finds the path of the `winget.exe` executable on a Windows system. 
        Aimed at finding Winget when main script is executed as SYSTEM, but will also work under USER
        
        Windows Package Manager (`winget`) is a command-line tool that facilitates the 
        installation, upgrade, configuration, and removal of software packages. Identifying the 
        exact path of `winget.exe` allows for execution (installations) under SYSTEM context.

        METHOD
        1. Defining Potential Paths:
        - Specifies potential locations of `winget.exe`, considering:
            - Standard Program Files directory (64-bit systems).
            - 32-bit Program Files directory (32-bit applications on 64-bit systems).
            - Local application data directory.
            - Current user's local application data directory.
        - Paths may utilize wildcards (*) for flexible directory naming, e.g., version-specific folder names.

        2. Iterating Through Paths:
        - Iterates over each potential location.
        - Resolves paths containing wildcards to their actual path using `Resolve-Path`.
        - For each valid location, uses `Get-ChildItem` to search for `winget.exe`.

        3. Returning Results:
        - If `winget.exe` is located, returns the full path to the executable.
        - If not found in any location, outputs an error message and returns `$null`.

    .EXAMPLE
        $wingetLocation = Find-WingetPath
        if ($wingetLocation) {
            Write-Output "Winget found at: $wingetLocation"
        } else {
            Write-Error "Winget was not found on this system."
        }

    .NOTES
        While this function is designed for robustness, it relies on current naming conventions and
        structures used by the Windows Package Manager's installation. Future software updates may
        necessitate adjustments to this function.

    .DISCLAIMER
        This function and script is provided as-is with no warranties or guarantees of any kind. 
        Always test scripts and tools in a controlled environment before deploying them in a production setting.
        
        This function's design and robustness were enhanced with the assistance of ChatGPT, it's important to recognize that 
        its guidance, like all automated tools, should be reviewed and tested within the specific context it's being 
        applied. 

    #>
    # Define possible locations for winget.exe
    $possibleLocations = @(
        "${env:ProgramFiles}\WindowsApps\Microsoft.DesktopAppInstaller*_x64__8wekyb3d8bbwe\winget.exe", 
        "${env:ProgramFiles(x86)}\WindowsApps\Microsoft.DesktopAppInstaller*_8wekyb3d8bbwe\winget.exe",
        "${env:LOCALAPPDATA}\Microsoft\WindowsApps\winget.exe",
        "${env:USERPROFILE}\AppData\Local\Microsoft\WindowsApps\winget.exe"
    )

    # Iterate through the potential locations and return the path if found
    foreach ($location in $possibleLocations) {
        try {
            # Resolve path if it contains a wildcard
            if ($location -like '*`**') {
                $resolvedPaths = Resolve-Path $location -ErrorAction SilentlyContinue
                # If the path is resolved, update the location for Get-ChildItem
                if ($resolvedPaths) {
                    $location = $resolvedPaths.Path
                }
                else {
                    # If path couldn't be resolved, skip to the next iteration
                    Write-Warning "Couldn't resolve path for: $location"
                    continue
                }
            }
            
            # Try to find winget.exe using Get-ChildItem
            $items = Get-ChildItem -Path $location -ErrorAction Stop
            if ($items) {
                Write-Host "Found Winget at: $items"
                return $items[0].FullName
                break
            }
        }
        catch {
            Write-Warning "Couldn't search for winget.exe at: $location"
        }
    }

    Write-Error "Winget wasn't located in any of the specified locations."
    return $null
}


#endregion Functions

#region Main

#region Initialization
$wingetPath = ""                            # Path to Winget executable
$detectSummary = ""                         # Script execution summary
$result = 0                                 # Exit result (default to 0)
$WingetAppID = "Microsoft.VisualStudioCode" # Winget Application ID
$processResult = $null                      # Winget process result
$exitCode = $null                           # Software installation exit code
$installInfo                                # Information about the Winget installation process
#endregion Initialization

# Make the log easier to read
Write-Host `n`n

# Check if Winget is available in current context, if not, find it
$wingetPath = (Get-Command -Name winget -ErrorAction SilentlyContinue).Source

if (-not $wingetPath) {
    Write-Host "Winget not detected, attempting to locate in system..."
    $wingetPath = Find-WingetPath
}

if (-not $wingetPath) {
    Write-Host "Winget (Windows Package Manager) is absent on this device." 
    $detectSummary += "Winget NOT detected. "
    $result = 5
} else {
    $detectSummary += "Winget located at $wingetPath. "
}

# Use Winget to install the desired software
if ($result -eq 0) {
    try {
        $tempFile = New-TemporaryFile
        Write-Host "Initiating App $WingetAppID Installation"
        $processResult = Start-Process -FilePath "$wingetPath" -ArgumentList "install -e --id ""$WingetAppID"" --scope=machine --silent --accept-package-agreements --accept-source-agreements --force" -NoNewWindow -Wait -RedirectStandardOutput $tempFile.FullName -PassThru

        $exitCode = $processResult.ExitCode
        $installInfo = Get-Content $tempFile.FullName
        Remove-Item $tempFile.FullName

        Write-Host "Winget install exit code: $exitCode"
        #Write-Host "Winget installation output: $installInfo"          #Remove comment to troubleshoot.
        
        if ($exitCode -eq 0) {
            Write-Host "Winget successfully installed application."
            $detectSummary += "Installed $WingetAppID via Winget. "
            $result = 0
        } else {
            $detectSummary += "Error during installation, exit code: $exitCode. "
            $result = 1
        }
    }
    catch {
        Write-Host "Encountered an error during installation: $_"
        $detectSummary += "Installation failed with exit code $($processResult.ExitCode). "
        $result = 1
    }
}

# Simplify reading in the AgentExecutor Log
Write-Host `n`n

# Output the final results
if ($result -eq 0) {
    Write-Host "OK $([datetime]::Now) : $detectSummary"
    Exit 0
} elseif ($result -eq 1) {
    Write-Host "FAIL $([datetime]::Now) : $detectSummary"
    Exit 1
} else {
    Write-Host "NOTE $([datetime]::Now) : $detectSummary"
    Exit 0
}

#endregion Main
