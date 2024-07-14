# IEModeExpiryFix

[![image](https://github.com/LesFerch/WinSetView/assets/79026235/0188480f-ca53-45d5-b9ff-daafff32869e)Download the zip file](https://github.com/LesFerch/IEModeExpiryFix/archive/refs/heads/main.zip)

**Note**: Some antivirus software may falsely detect the download as a virus. This can happen any time you download a new executable and may require extra steps to whitelist the file.


[**Version: 1.2.0**](./VersionHistory.md)\
**Last updated: 2024-07-14**

## Set IE Mode pages to expire far in the future

This script sets your Edge **IE Mode** pages to a "Date added" in 2099 (can be edited to any date), which causes the expiry to also be in 2099. It also provides features for clearing all existing IE Mode pages, adding pages via the script, removing individual pages via the script, and searching and replacing strings in the Edge Preferences file. See the comments in the script for more details.

This package includes both VBScript and PowerShell versions. Use whichever one you like.

The VBScript version is faster and can be double-clicked to run. It will work for anyone, unless VBScript is blocked by security polices on your computer.

The PowerShell script is slower and you must right-click and "Run with Powershell" (or run it via the command line). You may have to also set the initial PowerShell execution policy if you've never run PowerShell scripts before.

For those who care about the technical details, the VBScript version makes all the edits to the Edge JSON Preferences file by direct string search and replace. The PowerShell script converts the JSON data to an object, modifies the object, and then converts the object back to JSON data.

The script requires no modifications to set the expiry on URLs that are already in Edge's IE Mode list.

To use it, first ensure that you are NOT logged into Edge (the script will not work with synced profiles). Then close Edge (the script will forcefully close Edge if it's still running). Then, just download and unzip the file and either double-click **IEModeExpiryFix.vbs** or right-click **IEModeExpiryFix.ps1** and select "Run with PowerShell".

See the comments in the script to see how to modify it to add sites to the IE Mode list at the same time as setting their expiry.

The script settings can also be configured using an INI file instead of directly modifying the variables in the script. Provide the INI file name as a command line argument. See `Example.ini` for format.

You may update all user profiles at the same time by setting the `AllUsers` variable to `True`. You will have to ensure that the script is run via an account that can update all user profiles, such as SYSTEM, in order for that feature to work.

**Note**: The script only works with 100% local Edge profiles. It will not work if you are logged into Edge (i.e. synced profiles).

**Note**: The script will not work if your Edge profile is centrally managed (i.e. controlled by your IT department).

**Note**: The script will create a backup of the Edge Preferences file whenever it makes a change, but please test carefully (especially if deploying to multiple users). 
As with any script, use at your own risk.

**Note**: If you got the script from this repository, it's 100% clean, but you may find that some security software will falsely detect it as potentially unwanted or potentially malicious. That's the nature of such software. It will err on the side of caution. If you encounter that situation, you will need to disable or, at least, dial-back the protection settings of your security software.

**Note for system administrators**: If your computers are in Active Directory, please consider using the Enterprise Mode Site List instead of this script:
[Enterprise Mode and the Enterprise Mode Site List](https://docs.microsoft.com/en-us/internet-explorer/ie11-deploy-guide/what-is-enterprise-mode)

## Alternative Solution

You can also load web pages directly with Internet Explorer using **[LaunchIE](https://lesferch.github.io/LaunchIE/)** or by using one of the included launcher scripts (LaunchIE.vbs or LaunchIE.js).

There is no more risk using a web page via the launcher as there is accessing the page using Edge IE Mode. Edge IE Mode actually runs IE in the background, so it's really all the same, other than the annoying expiry date. What you should NOT do is use IE as a general purpose browser. IE (or Edge IE Mode) should only be used for specific pages, that you trust, that only work in IE.

\
[![image](https://github.com/LesFerch/WinSetView/assets/79026235/63b7acbc-36ef-4578-b96a-d0b7ea0cba3a)](https://github.com/LesFerch/IEModeExpiryFix)
