# IEModeExpiryFix

[![image](https://user-images.githubusercontent.com/79026235/152910441-59ba653c-5607-4f59-90c0-bc2851bf2688.png)Download the zip file](https://github.com/LesFerch/IEModeExpiryFix/archive/refs/heads/main.zip)

[**Version: 1.1.0**](https://github.com/LesFerch/IEModeExpiryFix/blob/main/VersionHistory.md)\
**Last updated: 2023-02-27**

## Set IE Mode pages to expire far in the future

This script sets your Edge **IE Mode** pages to a "Date added" in 2099 (can be edited to any date), which causes the expiry to also be in 2099. It also provides features for clearing all existing IE Mode pages, adding pages via the script, removing individual pages via the script, and searching and replacing strings in the Edge Preferences file. See the comments in the script for more details.

This package now includes both VBScript and PowerShell versions. Use whichever one you like.

The VBScript version is faster and can be double-clicked to run. It will work for anyone, unless VBScript is blocked by security polices on your computer.

The PowerShell script is slower and you must right-click and "Run with Powershell" (or run it via the command line). You may have to also set the initial PowerShell execution policy if you've never run PowerShell scripts before.

For those who care about the technical details, the VBScript version makes all the edits to the Edge JSON Preferences file by direct string search and replace. The PowerShell script converts the JSON data to an object, modifies the object, and then converts the object back to JSON data.

The script requires no modifications to set the expiry on URLs that are already in Edge's IE Mode list.

To use it, first ensure that you are NOT logged into Edge (the script will not work with synced profiles). Then close Edge (the script will forcefully close Edge if it's still running). Then, just download and unzip the file and either double-click IEModeExpiryFix.vbs or right-click IEModeExpiryFix.ps1 and select "Run with PowerShell".

See the comments in the script to see how to modify it to add sites to the IE Mode list at the same time as setting their expiry.

**Note**: The script only works with 100% local Edge profiles. It will not work if you are logged into Edge (i.e. synced profiles).

**Note**: The script will create a backup of the Edge Preferences file whenever it makes a change, but please test carefully (especially if deploying to multiple users). 
As with any script, use at your own risk.

**Note**: If you got the script from this repository, it's 100% clean, but you may find that some security software will detect it as potentially unwanted or potentially malicious. That's the nature of such software. It will err on the side of caution. If you encounter that situation, you will need to disable or, at least, dial-back the protection settings of your security software.

**Note for system administrators**: If your computers are in Active Directory, please consider using the Enterprise Mode Site List instead of this script:
[Enterprise Mode and the Enterprise Mode Site List](https://docs.microsoft.com/en-us/internet-explorer/ie11-deploy-guide/what-is-enterprise-mode)


[![image](https://user-images.githubusercontent.com/79026235/153264696-8ec747dd-37ec-4fc1-89a1-3d6ea3259a95.png)](https://github.com/LesFerch/IEModeExpiryFix)
