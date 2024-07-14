'IEModeExpiryFix.vbs
'https://lesferch.github.io/IEModeExpiryFix/

'Sets the date added for all Edge IE Mode pages to any date you specify below.
'This causes the expiry dates to be the specified date plus 30 days.
'The default date added is a date in 2099, making the expiry a long way in the future.

'This script only works with completely local Edge profiles. It will not work if Edge is signed-in.

'Optionally run via CScript (i.e. CScript IEModeExpiryFix.vbs) to get console output instead of message boxes.

'Note for system administrators: If your computers are in Active Directory, please consider
'using the Enterprise Mode Site List instead of this script. See this link:
'https://docs.microsoft.com/en-us/internet-explorer/ie11-deploy-guide/what-is-enterprise-mode

'How to use:

'1. Add your IE Mode pages in Microsoft Edge
'   (or add them to the $AddPages variable below or add them via an INI file)
'2. Close Microsoft Edge (if open)
'3. Run this script

'Repeat the above steps to add more IE Mode pages.

'Variables are read from an INI file if the INI file name is provided as a parameter on the command line.
'INI file settings override settings supplied directly in the script.

Version = "1.2.0"

Silent = False 'Change to True for no prompts and no report.
AllUsers = False 'Set to True to process all user folders (typically used with SYSTEM account and $Silent = $True).
RemoveAll = False 'Set to True to remove all existing IE Mode pages.
Backup = True 'Set to False for no backup.
ForceLowercase = True 'Force domain part of URL to be lowercase.
Setlocale("en-us") 'Do NOT change unless you change the date format for "DateAdded" below.
DateAdded = "10/28/2099 10:00:00 PM" 'Specify the date here (ensure format is consistent with "Setlocale").

'To remove pages, copy, uncomment and edit the RemovePages line below. Separate each page entry with a |.
'Entries must match the existing URL EXACTLY, including case and trailing slash (if any).

'RemovePages = "http://www.example.com/|http://www.example.com/page.asp"

'To add pages, copy, uncomment and edit the AddPages line below. Separate each page entry with a |.
'Entries must end with a slash unless the URL ends with a file such as .html, .aspx, etc.
'The domain part of the entry must be all lowercase.
'Edge IE Mode will not accept URL parameters, so the script will trim URLs at the first "?" character.

'AddPages = "http://www.fiat.it/|http://www.ferrari.it/|http://www.some.com/page.asp"

'To find and replace any string in the Preferences file, copy, uncomment and edit the FindReplace line below.
'Separate find and replace strings with a comma and separate each find/replace pair with a |.
'Be very careful that you specify exact, unique, case-sensitive text! Use at your own risk!
'Be careful not to generate duplicates by making an incorrect entry the same as an existing entry!
'This feature operates on the whole file, not just the IE Mode pages section. You have been warned!
'It's safer to use RemovePages with AddPages to fix an error. Use FindReplace as a last resort!

'FindReplace = "www.oops.com,www.correct.com|www.oops.net,www.correct.net"

Const ForWriting = 2
Dim PrefsFile,MyLog,Data,OriginalData
Z = VBCRLF
ZZ = VBCRLF & VBCRLF

Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oSettings = CreateObject("Scripting.Dictionary")

'Get settings from INI file if specified on the command line.
If WScript.Arguments.Count>0 Then INIFile = WScript.Arguments.Item(0)
oWSH.CurrentDirectory = oFSO.GetParentFolderName(WScript.ScriptFullName)
If INIFile<>"" Then
  If Not oFSO.FileExists(INIFile) Then
    LogMsg "File not found: " & INIFile
    ExitScript
  End If
  INI2Dict
  If oSettings.Exists("[Options]Silent") Then Silent = (oSettings.Item("[Options]Silent")=1)
  If oSettings.Exists("[Options]AllUsers") Then AllUsers = (oSettings.Item("[Options]AllUsers")=1)
  If oSettings.Exists("[Options]RemoveAll") Then RemoveAll = (oSettings.Item("[Options]RemoveAll")=1)
  If oSettings.Exists("[Options]Backup") Then Backup = (oSettings.Item("[Options]Backup")=1)
  If oSettings.Exists("[Content]DateAdded") Then DateAdded = oSettings.Item("[Content]DateAdded")
  If oSettings.Exists("[Content]RemovePages") Then RemovePages = oSettings.Item("[Content]RemovePages")
  If oSettings.Exists("[Content]AddPages") Then AddPages = oSettings.Item("[Content]AddPages")
  If oSettings.Exists("[Content]FindReplace") Then FindReplace = oSettings.Item("[Content]FindReplace")
End If

'Convert variable lists to arrays
aRemovePages = Split(RemovePages,"|") 'Do NOT edit this!
aAddPages = Split(AddPages,"|") 'Do NOT edit this!
aFindReplace = Split(FindReplace,"|") 'Do NOT edit this!

'Convert the date
Set oDateTime = CreateObject("WbemScripting.SWbemDateTime")
Call oDateTime.SetVarDate(DateAdded,True)
EdgeDateAdded = Left(oDateTime.GetFileTime,17)

CScript = InStr(LCase(WScript.FullName),"cscript")>0

'Display list of actions to be performed
If Not Silent And Not CScript Then
  MC = "Remove ALL IE Mode pages." & ZZ
  MD = "Remove these IE Mode pages: " & Z & RemovePages & ZZ
  MA = "Add these IE Mode pages: " & Z & AddPages & ZZ
  MF = "Find and replace these strings: " & Z & FindReplace & ZZ
  MX = "Set all IE Mode pages to expire 30 days after: " & DateAdded & ZZ
  MB = "Create a Preferences backup file if any changes are made." & ZZ
  MS = "Skip any synced profiles (because Edge will reject edits)." & ZZ
  MK = "Kill any running MSEdge.exe tasks."
  If RemoveAll Then Msg = Msg & MC
  If RemovePages<>"" And Not RemoveAll Then Msg = Msg & MD
  If AddPages<>"" Then Msg = Msg & MA
  If FindReplace<>"" Then Msg = Msg & MF
  If (Not RemoveAll) Or AddPages<>"" Then Msg = Msg & MX
  If Backup Then Msg = Msg & MB
  Msg = Msg & MS & MK
  Response = MsgBox(Msg,VBOKCancel,"The following actions will be performed:")
  If Response=VBCancel Then WScript.Quit
End If

'Edge must be closed to modify the Preferences file
oWSH.Run "TaskKill /im MSEdge.exe /f",0,True

LocalAppData = oWSH.ExpandEnvironmentStrings("%LocalAppData%")

LogMsg "Profiles processed:"

If AllUsers Then
  usersPath = oWSH.RegRead("HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\ProfilesDirectory")
  usersPath = oWSH.ExpandEnvironmentStrings(usersPath)
  Set oFolder = oFSO.GetFolder(usersPath)
  For Each oFolder In oFolder.SubFolders
    Folder = oFSO.GetAbsolutePathName(oFolder)
    ProcessUserProfiles(Folder)
  Next
Else
  ProcessUserProfiles(oWSH.ExpandEnvironmentStrings("%UserProfile%"))
End If

ExitScript

'End of main code. Subs and functions below.

Sub ExitScript
  If Not Silent Then
    WScript.Echo MyLog
  End If
  WScript.Quit
End Sub

'Add a message to the MyLog variable
Sub LogMsg(Msg)
  MyLog = MyLog & Msg & ZZ
End Sub

'Check URL for too many slashes in prefix
Function BadURL(ByVal URL)
  URL = LCase(URL)
  BadURL = Instr(URL,"http:///")>0 Or Instr(URL,"https:///")>0 Or Instr(URL,"file:////")>0
End Function

'Ensure URL starts with something legit, is trimmed at ?, and domain part is lowercase
Function FixURL(byVal URL)
  If Instr(URL,"://")=0 Then URL = "https://" & URL
  URL = Split(URL,"?")(0)
  If ForceLowercase Then
    For i = 2 To Len(URL)
      If Mid(URL,i,1)="/" And Mid(URL,i-1,1)<>"/" And Mid(URL,i+1,1)<>"/" Then Exit For
    Next
    If i>Len(URL) Then URL = URL & "/"
    URL = LCase(Left(URL,i)) & Mid(URL,i+1)
  End If
  FixURL = URL
End Function

'If Backup flag is True, save original data to date-time-named backup file
Sub BackupPrefsFile
  If Backup Then
    d = Now()
    s1 = Year(d) & "-" & Right("0" & Month(d),2) & "-" & Right("0" & Day(d),2) & "-"
    s2 = Right("0" & Hour(d),2) & Right("0" & Minute(d),2) & "-" & Right("0" & Second(d),2)
    Suffix = "-Backup-" & s1 & s2
    Call oFSO.OpenTextFile(PrefsFile & Suffix,ForWriting,True).Write(OriginalData)
  End If
End Sub

'Remove all IE Mode pages
Sub ClearEntries
  FirstBlockEnd = InStr(Data,"user_list_data_1") + 18
  If FirstBlockEnd>18 Then
    SecondBlockStart = InStr(Data,"}},""edge""")
    Data = Left(Data,FirstBlockEnd) & Mid(Data,SecondBlockStart)
  End If
End Sub

'Find and change every IE Mode page date
Sub UpdateEntries
  StartPos = 1
  Do
    FoundPos = InStr(StartPos,Data,"date_added")
    If FoundPos=0 Then Exit Do
    Data = Mid(Data,1,FoundPos + 12) & EdgeDateAdded & Mid(Data,FoundPos + 30)
    StartPos = FoundPos + 1
  Loop
End Sub

'Remove any pages specified with the RemovePages variable
Sub RemoveEntries
  For i = 0 To UBound(aRemovePages)
    URL = FixURL(aRemovePages(i))
    S1 = Instr(Data,"""" & URL & """:{""date_added")
    If S1>0 Then
      E1 = InStr(S1,Data,"}")
      Data = Left(Data,S1 - 1) & Mid(Data,E1 + 1)
    End If
    Data = Replace(Data,"{,","{")
    Data = Replace(Data,",}","}")
    Data = Replace(Data,",,",",")
  Next
End Sub

'Add any pages specified with the AddPages variable
Sub AddEntries
  For i = 0 To UBound(aAddPages)
    URL = FixURL(aAddPages(i))
    If Not BadURL(URL) Then
      If Instr(Data,"user_list_data_1")=0 Then Data = Replace(Data,"},""edge"":{",",""user_list_data_1"":{}},""edge"":{")
      If Instr(Data,"""" & URL & """:{""date_added")=0 Then Data = Replace(Data,"""user_list_data_1"":{","""user_list_data_1"":{""" & URL & """:{""date_added"":""" & EdgeDateAdded & """,""engine"":2,""visits_after_expiration"":0},")
      Data = Replace(Data,",}","}")
    End If
  Next
End Sub

'Find and replace strings specified with the FindReplace variable
'This feature operates on the whole file, not just the IE Mode pages section!
Sub FindReplaceAnyString
  For i = 0 To UBound(aFindReplace)
    aFindReplacePair = Split(aFindReplace(i),",")
    Data = Replace(Data,aFindReplacePair(0),aFindReplacePair(1))
  Next
End Sub

Sub EditProfile

  'Read contents of Edge Preferences file into a variable
  Data = oFSO.OpenTextFile(PrefsFile).ReadAll

  LogMsg PrefsFile

  'Exit if user is signed in
  If InStr(Data,"""account_info"":[]")=0 And InStr(Data,"account_info")>0 Then
    LogMsg "Edge profile sign-in detected. Profile cannot be updated."
    Exit Sub
  End If

  OriginalData = Data

  If RemoveAll Then
    ClearEntries
  Else
    RemoveEntries
    UpdateEntries
  End If

  AddEntries

  FindReplaceAnyString
  
  'Set "Allow pages to be reloaded in Internet Explorer mode" to "Allow"
  Data = Replace(Data,"{""ie_user""","{""enabled_state"":1,""ie_user""")
  Data = Replace(Data,"{""enabled_state"":0,""ie_user""","{""enabled_state"":1,""ie_user""")
  Data = Replace(Data,"{""enabled_state"":2,""ie_user""","{""enabled_state"":1,""ie_user""")

  'Overwrite the Preferences file with the new data
  If Data<>OriginalData Then
    BackupPrefsFile
    Call oFSO.CreateTextFile(PrefsFile,True).Write(Data)
    LogMsg "Profile updated"
  Else
    LogMsg "Profile already updated"
  End If
End Sub

'Process profiles in all known Edge profile folders
Sub ProcessProfiles(ProfileFolder)
  EdgeData = ProfileFolder & "\User Data\"
  If oFSO.FolderExists(EdgeData) Then
    For Each oFolder In oFSO.GetFolder(EdgeData).SubFolders
      PrefsFile = oFolder.Path & "\Preferences"
      If oFSO.FileExists(PrefsFile) Then EditProfile
    Next
  End If
End Sub

'Process profiles for all Edge versions
Sub ProcessUserProfiles(Path)
  Path = Path & "\AppData\Local\Microsoft"
  ProcessProfiles(Path & "\Edge") 'For released Edge profile.
  ProcessProfiles(Path & "\Edge Beta") 'For Beta Edge profile.
  ProcessProfiles(Path & "\Edge Dev") 'For Dev Edge profile.
  ProcessProfiles(Path & "\Edge SxS") 'For Canary Edge profile.
End Sub

'Read INI file into a dictionary
Sub Ini2Dict
  If oFSO.FileExists(INIFile) Then 
    Set oFile = oFSO.OpenTextFile(INIFile)
    Do Until oFile.AtEndOfStream
      Line = Trim(oFile.ReadLine)
      If Line<>"" And Left(Line,1)<>";" Then
        If Left(Line,1)="[" Then
          Section = Line
        Else
          ArrLine = Split(Line,"=")
          If UBound(Arrline)=1 Then
            oSettings.Add Section & Trim(ArrLine(0)),Trim(ArrLine(1))
          End If
        End If
      End If
    Loop
    oFile.Close
  End If
End Sub
