#IEModeExpiryFix.ps1
#https://lesferch.github.io/IEModeExpiryFix/

#Sets the date added for all Edge IE Mode pages to any date you specify below.
#This causes the expiry dates to be the specified date plus 30 days.
#The default date added is a date in 2099, making the expiry a long way in the future.

#This script only works with completely local Edge profiles. It will not work if Edge is signed-in.

#Note for system administrators: If your computers are in Active Directory, please consider
#using the Enterprise Mode Site List instead of this script. See this link:
#https://docs.microsoft.com/en-us/internet-explorer/ie11-deploy-guide/what-is-enterprise-mode

#How to use:

#1. Add your IE Mode pages in Microsoft Edge
#   (or add them to the $AddPages variable below or add them via an INI file)
#2. Close Microsoft Edge (if open)
#3. Run this script

#Repeat the above steps to add more IE Mode pages

#Variables are read from an INI file if the INI file name is provided as a parameter on the command line.
#INI file settings override settings supplied directly in the script.

Param (
  $INIFile = ''
)

$Version = '1.2.0'

$invoc = $MyInvocation.InvocationName

$Silent = $False #Change to True for no prompts and no report.

Function ExitScript {
  If (-Not $Silent) {
    Write-Host $MyLog
    #Pause if launched via right-click
    If ($invoc -eq '&') {
      Read-Host "Press Enter to close window"
    }
  }
  Exit
}

Function LogMsg($Msg) {
  $Script:MyLog = "$MyLog$Msg`r`n`r`n"
}

if ($PSVersionTable.PSEdition -eq "Core") {
  LogMsg "This script is not compatible with PowerShell Core. Please run it with PowerShell."
  ExitScript
}

If ($ExecutionContext.SessionState.LanguageMode -eq "ConstrainedLanguage") {
  LogMsg "This script cannot be run in constrained language mode."
  ExitScript
}

If (-Not $Silent) {
  try { $host.ui.RawUI.WindowTitle = "IEModeExpiryFix" }
  catch { }
}

$AllUsers = $False #Set to True to process all user folders (typically used with SYSTEM account and $Silent = $True).
$RemoveAll = $False #Set to True to remove all existing IE Mode pages.
$Backup = $True #Set to False for no backup.
$ForceLowercase = $True #Force domain part of URL to be lowercase.
$currentThread = [System.Threading.Thread]::CurrentThread
$currentThread.CurrentCulture = "en-US" #Do NOT change unless you change the date format for "DateAdded" below.
$DateAdded = '10/28/2099 10:00:00 PM' #Specify the date here (ensure format is consistent with "Set-Culture").

$RemovePages = '' #Don't touch this!
$AddPages = '' #Don't touch this!
$FindReplace = '' #Don't touch this!

#To remove pages, copy, uncomment and edit the $RemovePages line below. Separate each page entry with a |.
#Entries must match the existing URL EXACTLY, including case and trailing slash (if any).

#$RemovePages = 'http://www.example.com/|http://www.example.com/page.asp'

#To add pages, copy, uncomment and edit the $AddPages line below. Separate each page entry with a |.
#Entries must end with a slash unless the URL ends with a file such as .html, .aspx, etc.
#The domain part of the entry must be all lowercase.
#Edge IE Mode will not accept URL parameters, so the script will trim URLs at the first "?" character.

#$AddPages = 'http://www.fiat.it/|http://www.ferrari.it/|http://www.some.com/page.asp'

#To find and replace any string in the Preferences file, copy, uncomment and edit the $FindReplace line below.
#Separate find and replace strings with a comma and separate each find/replace pair with a |.
#Be very careful that you specify exact, unique, case-sensitive text! Use at your own risk!
#Be careful not to generate duplicates by making an incorrect entry the same as an existing entry!
#This feature operates on the whole file, not just the IE Mode pages section. You have been warned!
#It's safer to use $RemovePages and $AddPages to correct an error! Use FindReplace as a last resort!

#$FindReplace = 'www.oops.com,www.correct.com|www.oops.net,www.correct.net'

#Used to Read $INIfile contents
Function Get-IniContent ($FilePath) {
  $ini = @{}
  Switch -regex -file $FilePath {
    '^\[(.+)\]' # Section
    {
      $section = $matches[1]
      $ini[$section] = @{}
    }
    '(.+?)\s*=(.*)' # Key
    {
      $name,$value = $matches[1..2]
      $value = $value.Trim()
      $v = $value.ToLower()
      If (($v -eq '0') -Or ($v -eq 'false') -Or ($v -eq '$false')) {$value = $False}
      If (($v -eq '1') -Or ($v -eq 'true') -Or ($v -eq '$true')) {$value = $True}
      $ini[$section][$name] = $value
    }
  }
  Return $ini
}

#Read variables from $INIFile (if provided on command line)
If ($INIFile -ne '') {
  Set-Location $PSScriptRoot
  If (-Not(Test-Path -Path $INIFile)) {
    LogMsg "File not found: $INIFile"
    ExitScript
  }
  $iniContent = Get-INIContent $INIFile
  $Silent = $iniContent['Options']['Silent']
  $AllUsers = $iniContent['Options']['AllUsers']
  $RemoveAll = $iniContent['Options']['RemoveAll']
  $Backup = $iniContent['Options']['Backup']
  $DateAdded = $iniContent['Content']['DateAdded']
  $RemovePages = $iniContent['Content']['RemovePages']
  $AddPages = $iniContent['Content']['AddPages']
  $FindReplace = $iniContent['Content']['FindReplace']
  If ($AllUsers -eq $Null) {$AllUsers = $False}
  If ($RemoveAll -eq $Null) {$RemoveAll = $False}
  If ($Backup -eq $Null) {$Backup = $True}
  If ($Silent -eq $Null) {$Silent = $False}
  If ($DateAdded -eq $Null) {$DateAdded = '10/28/2099 10:00:00 PM'}
  If ($RemovePages -eq $Null) {$RemovePages = ''}
  If ($AddPages -eq $Null) {$AddPages = ''}
  If ($FindReplace -eq $Null) {$FindReplace = ''}
}

#Convert variable lists to arrays
#Do NOT edit any of the variables below!
#Edits are done using the commented variables above!
$aRemovePages = @()
$aAddPages = @()
$aFindReplace = @()
$aRemovePages = $RemovePages.Split("|")
$aAddPages = $AddPages.Split("|")
$aFindReplace = $FindReplace.Split("|")

#Convert the date
$EdgeDateAdded = (Get-Date $DateAdded).ToFileTime()
$EdgeDateAdded = "$EdgeDateAdded".Substring(0,17)

#Erase one or more lines from the console window
Function Clear-Lines ($Count) {
  try {
    $CurrentLine  = $Host.UI.RawUI.CursorPosition.Y
    $ConsoleWidth = $Host.UI.RawUI.BufferSize.Width
    $i = 1
    For ($i; $i -le $Count; $i++) {
      [Console]::SetCursorPosition(0,($CurrentLine - $i))
      [Console]::Write("{0,-$ConsoleWidth}" -f " ")
    }
    [Console]::SetCursorPosition(0,($CurrentLine - $Count))
  }
  catch { }
}

#Display list of actions to be performed
If (-Not $Silent) {
  $MC = "Remove ALL IE Mode pages.`r`n`r`n"
  $MD = "Remove these IE Mode pages:`r`n$RemovePages`r`n`r`n"
  $MA = "Add these IE Mode pages:`r`n$AddPages`r`n`r`n"
  $MF = "Find and replace these strings:`r`n$FindReplace`r`n`r`n"
  $MX = "Set all IE Mode pages to expire 30 days after: $DateAdded`r`n`r`n"
  $MB = "Create a Preferences backup file if any changes are made.`r`n`r`n"
  $MS = "Skip any synced profiles (because Edge will reject edits).`r`n`r`n"
  $MK = "Kill any running MSEdge.exe tasks.`r`n`r`n"
  If ($RemoveAll) {$Msg = "$Msg$MC"}
  If (($RemovePages -ne '') -And (-Not $RemoveAll)) {$Msg = "$Msg$MD"}
  If ($AddPages -ne '') {$Msg = "$Msg$MA"}
  If ($FindReplace -ne '') {$Msg = "$Msg$MF"}
  If ((-Not $RemoveAll) -Or ($AddPages -ne '')) {$Msg = "$Msg$MX"}
  If ($Backup) {$Msg = "$Msg$MB"}
  $Msg = "$Msg$MS$MK"
  Cls
  Write-Host "The following actions will be performed:`r`n`r`n$Msg"
  Read-Host `n'Press Enter to continue or Ctrl-C to exit'
}

#Clear the prompt line
If (-Not $Silent) {Clear-Lines 2}

#Edge must be closed to modify the Preferences file
Stop-Process -Force -ErrorAction SilentlyContinue -ProcessName MSEdge

#Functions defined below. Scroll down to see remaining main code.

#Check $URL for too many slashes in prefix
Function BadURL($URL) {
  $URL = $URL.ToLower()
  Return (($URL.IndexOf('http:///') -ge 0) -Or ($URL.IndexOf('https:///') -ge 0) -Or ($URL.IndexOf('file:////') -ge 0))
}

#Ensure $URL starts with something legit, is trimmed at ?, and domain part is lowercase
Function FixURL($URL) {
  If ($URL.IndexOf("://") -eq -1) {$URL = "https://$URL"}
  $URL = $URL.Split("?")[0]
  If ($ForceLowercase) {
    For ($i = 1; $i -lt $URL.Length; $i++) {
      If (($URL.Substring($i,1) -eq "/") -And ($URL.Substring($i-1,1) -ne "/") -And ("$URL ".Substring($i+1,1) -ne "/")) {Break}
    }
    If ($i -eq $URL.Length) {$URL = "$URL/"}
    $L = ($URL.Substring(0,$i)).ToLower()
    $R = $URL.Substring($i)
    $URL = "$L$R"
  }
  Return $URL
}

#If $Backup flag is True, save original data to date-time-named backup file.
Function BackupPrefsFile {
  If ($Backup) {
    $DT = Get-Date -Format "yyyy-MM-dd-hhmm-ss"
    $Suffix = "-Backup-$DT"
    $Data | Out-File "$PrefsFile$Suffix" -Encoding Default
  }
}

#Remove all IE Mode pages
Function ClearEntries {
  $oData.dual_engine.user_list_data_1 | ForEach {
    $_.psobject.properties | ForEach {
      If ($_.name -ne $Null) {$oData.dual_engine.user_list_data_1.psobject.properties.remove($_.name)}
    }
  }
}

#Find and change every IE Mode page date
Function UpdateEntries {
  $oData.dual_engine.user_list_data_1 | ForEach {
    $_.psobject.properties | ForEach {
      $_.value | ForEach {
        $_.psobject.properties | ForEach {
          If ($_.name -eq 'date_added') {$_.value = $EdgeDateAdded}
        }
      }
    }
  }
}

#Remove any pages specified with the $RemovePages variable
Function RemoveEntries {
  Foreach ($Item in $aRemovePages) {
    $oData.dual_engine.user_list_data_1 | ForEach {
      $_.psobject.properties | ForEach {
        If ($_.name -eq $Item) {$oData.dual_engine.user_list_data_1.psobject.properties.remove($Item)}
      }
    }
  }
}

#Add any pages specified with the $AddPages variable
Function AddEntries {
  $ErrorActionPreference = 'SilentlyContinue'
  $oData | Add-Member -NotePropertyName 'dual_engine' -NotePropertyValue ([PSCustomObject]@{})
  $oData.dual_engine | Add-Member -NotePropertyName 'user_list_data_1' -NotePropertyValue ([PSCustomObject]@{})
    Foreach ($URL in $aAddPages) {
      $URL = FixURL($URL)
      If (-Not(BadURL($URL))) {
        $oData.dual_engine.user_list_data_1 | Add-Member -NotePropertyName $URL -NotePropertyValue ([PSCustomObject]@{})
        $oData.dual_engine.user_list_data_1.$URL | Add-Member -NotePropertyName 'date_added' -NotePropertyValue $EdgeDateAdded
        $oData.dual_engine.user_list_data_1.$URL | Add-Member -NotePropertyName 'engine' -NotePropertyValue 2
        $oData.dual_engine.user_list_data_1.$URL | Add-Member -NotePropertyName 'visits_after_expiration' -NotePropertyValue 0
      }
    }
  $ErrorActionPreference = 'Continue'
}

#Find and replace strings specified with the $FindReplace variable
Function FindReplaceAnyString {
  Foreach ($Item in $aFindReplace) {
    $Old = $Item.Split(",")[0]
    $New = $Item.Split(",")[1]
    $Script:Data = ($Data -Replace $Old,$New)
  }
}

Function EditProfile {

  #Read contents of Edge Preferences file into a variable
  $Script:Data = (Get-Content -Raw $PrefsFile).Trim()
  If ($PSVersionTable.PSVersion.Major -ge 6) {
    $oData = $Data | ConvertFrom-JSON -AsHashTable -ErrorAction SilentlyContinue
  }
  Else {
    $oData = $Data | ConvertFrom-JSON -ErrorAction SilentlyContinue
  }

  LogMsg $PrefsFile
  
  $Account = $oData.account_info.account_id

  #Exit if user is signed in
  If ($Account -ne $Null) {
    LogMsg "Edge profile sign-in detected. Profile cannot be updated."
    Return
  }

  If ($RemoveAll) {
    ClearEntries
  } 
  Else {
    RemoveEntries
    UpdateEntries
  }

  AddEntries

  #Set "Allow pages to be reloaded in Internet Explorer mode" to "Allow"
  $oData.dual_engine.consumer_mode | Add-Member -NotePropertyName 'enabled_state' -NotePropertyValue 1 -Force -ErrorAction SilentlyContinue
  
  $OriginalData = $Data.Trim()

  $Script:Data = ($oData | ConvertTo-Json -Compress -Depth 9).Trim()
  
  FindReplaceAnyString

  #Overwrite the Preferences file with the new data
  If ($Data -ne $OriginalData) {
    BackupPrefsFile
    $Data | Out-File $PrefsFile -Encoding Default
    LogMsg "Profile updated"
  }
  Else {
    LogMsg "Profile already updated"
  }
}

#Process profiles in all known Edge profile folders
Function ProcessProfiles($ProfileFolder) {
  $EdgeData = "$ProfileFolder\User Data\"
  try {
    If (Test-Path -Path $EdgeData -ErrorAction Stop) {
      Get-ChildItem $EdgeData -Directory | ForEach-Object {
        $Script:PrefsFile = "$EdgeData$_\Preferences"
        If (Test-Path -Path $PrefsFile) {EditProfile}
      }
    }
  }
  catch {
    LogMsg $EdgeData
    LogMsg "Insufficient access rights"
  }
}

Function ProcessUserProfiles($Path) {
  $Path = "$Path\AppData\Local\Microsoft"
  ProcessProfiles "$Path\Edge" #For released Edge profile
  ProcessProfiles "$Path\Edge Beta" #For Beta Edge profile
  ProcessProfiles "$Path\Edge Dev" #For Dev Edge profile
  ProcessProfiles "$Path\Edge SxS" #For Canary Edge profile
}

#Main code continues here

LogMsg "Profiles processed:"

If ($AllUsers) {
  $usersPath = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\ProfileList").ProfilesDirectory
  $folders = Get-ChildItem -Path $usersPath -Directory
  $folders | ForEach-Object {
    ProcessUserProfiles $_.FullName
  }
}
Else {
  ProcessUserProfiles $Env:UserProfile
}

ExitScript