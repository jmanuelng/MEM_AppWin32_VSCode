# MEM_AppWin32_VSCode

## Overview

The `MEM_AppWin32_VSCode` repository provides a solution for deploying the Visual Studio Code (VSCode) editor using Microsoft Intune as a Win32 app (as of August 2023, can't be deployed via "New" Microsoft Store). By leveraging the power of `winget`, Microsoft's package manager for Windows, this method ensures that devices always receive the most recent version of VSCode upon deployment. 

## Scripts in this Repository

- **[Detect_VSCode.ps1](https://github.com/jmanuelng/MEM_AppWin32_VSCode/blob/main/Detect_VSCode.ps1):** Checks if VSCode is already installed on the device. 

- **[Install_VSCode.ps1](https://github.com/jmanuelng/MEM_AppWin32_VSCode/blob/main/Install_VSCode.ps1):** Script uses `winget` to install the latest version of VSCode. It includes a function to ensure it runs in a 64-bit PowerShell environment, even if deployed as a "SYSTEM" in Intune, which defaults to 32-bit.

- **[Uninstall_VSCode.ps1](https://github.com/jmanuelng/MEM_AppWin32_VSCode/blob/main/Uninstall_VSCode.ps1):** Facilitates the removal of VSCode from the device.

## Usage

1. **Preparation**: Before deploying the scripts, ensure you have the IntuneWinAppUtil tool. This utility is essential for packaging the scripts into a format suitable for Intune.

2. **Packaging for Intune**:
   - Navigate to the directory containing the scripts and the IntuneWinAppUtil tool.
   - Run the IntuneWinAppUtil tool.
   - When prompted, provide the source folder, the setup file (script), and the output folder.
   - The tool will generate an `.intunewin` file, which is suitable for uploading to Intune.

3. **Uploading to Intune**:
   - Go to the Microsoft Endpoint Manager admin center.
   - Navigate to Apps > All apps > Add.
   - Select `Windows app (Win32)` from the list.
   - Upload the `.intunewin` file generated in the previous step.
   - Configure the app information, settings, and assignments as needed.
       - Install command: `powershell.exe -executionpolicy ByPass -file .\Install_VScode.ps1`
       - Uninstall command: `powershell.exe -executionpolicy ByPass -file .\Uninstall_VScode.ps1`
       - Install behavior: System
       - Detection rules: Use custom script `Detect_VSCode.ps1`
   - Save and assign the app to desired group(s).

4. **Setting Detection Rules**:
   - For detection, utilize the `Detect_VSCode.ps1` script. This script will verify if VSCode is already installed on the device, ensuring that the installation process is only initiated when necessary.

5. **Deployment**:
   - Once the app is assigned, devices/users in the target group(s) will receive the latest version of VSCode upon their next check-in with Intune.

## Credits

A special acknowledgment to [John Bryntze](https://twitter.com/JohnBryntze) for the inspiration, based on the tutorials below for packaging applications for Intune using `winget`:
- [Packaging Zoom with Winget and Intune - Part 1](https://www.youtube.com/watch?v=0Ov4AcRM4jI)
- [Packaging Zoom with Winget and Intune - Part 2](https://www.youtube.com/watch?v=MnFL2FQLjp4)
  
Also, huge thanks to [Robert Milner](https://twitter.com/robm82) and [Ruby Ooms](https://twitter.com/Mister_MDM), creators of the referenced articles and scripts that guided me to understand how to ensure Win32 packaged scripts run in a 64-bit environment when deployed as "SYSTEM" in Intune:
- [Running 64-bit PowerShell Scripts using Intune Win32 App Install](https://italik.co.uk/running-64-bit-powershell-scripts-using-intune-win32-app-install/)
- [IntunePS-x64 Script Sample](https://gist.github.com/robm82/8946aa0460a1feb00c434768b4ed1329#file-intuneps-x64-ps1)
- [The Sysnative Witch Project](https://call4cloud.nl/2021/05/the-sysnative-witch-project/)

