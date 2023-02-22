# IEModeExpiryFix
## Set IE Mode pages to expire far in the future

[![image](https://user-images.githubusercontent.com/79026235/152910441-59ba653c-5607-4f59-90c0-bc2851bf2688.png)Download the zip file](https://github.com/LesFerch/IEModeExpiryFix/archive/refs/heads/main.zip)

This VBS script sets your Edge **IE Mode** pages to a "Date added" in 2099 (can be edited to any date), which causes the expiry to also be in 2099. It also provides features for clearing all existing IE Mode pages, adding pages via the script, and searching and replacing strings in the Edge Preferences file. See the comments in the script for more details.

You can download the script in **ZIP** format via the link above or as a **VBS** file by clicking on the [script link](https://github.com/LesFerch/IEModeExpiryFix/blob/main/IEModeExpiryFix.vbs) and then either **Alt-Click** the **Raw** button (disabled in some browsers) or **Right-Click** the **Raw** button and select **Save link as...**.

The script requires no modifications to set the expiry on URLs that are already in Edge's IE Mode list. See the comments in the script to see how to use it to add sites to the IE Mode list at the same time as setting their expiry.

**Note**: The script only works with 100% local Edge profiles. It will not work if you are logged into Edge (i.e. synced profiles).

**Note**: The script will create a backup of the Edge Preferences file whenever it makes a change, but please test carefully (especially if deploying to multiple users). 
As with any script, use at your own risk.

**Note**: If you got the script from this repository, it's 100% clean, but you may find that some security software will detect it as potentially unwanted or potentially malicious. That's the nature of such software. It will err on the side of caution. If you encounter that situation, you will need to disable or, at least, dial-back the protection settings of your security software.

**Note for system administrators**: If your computers are in Active Directory, please consider using the Enterprise Mode Site List instead of this script:
[Enterprise Mode and the Enterprise Mode Site List](https://docs.microsoft.com/en-us/internet-explorer/ie11-deploy-guide/what-is-enterprise-mode)
