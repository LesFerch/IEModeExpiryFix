'IEModeExpiryFix.vbs
'https://lesferch.github.io/IEModeExpiryFix/

'Sets the date added for all Edge IE Mode pages to any date you specify below
'This causes the expiry dates to be the specified date plus 30 days
'The default date added is a date in 2099, making the expiry a long way in the future

'This script only works with completely local Edge profiles. It will not work if Edge is signed-in.

'Optionally run via CScript (i.e. CScript IEModeExpiryFix.vbs) to get console output instead of message boxes.

'Note for system administrators: If your computers are in Active Directory, please consider
'using the Enterprise Mode Site List instead of this script. See this link:
'https://docs.microsoft.com/en-us/internet-explorer/ie11-deploy-guide/what-is-enterprise-mode

'How to use:
'1. Add your IE Mode pages in Microsoft Edge (or add them to the AddSites variable below)
'2. Close Microsoft Edge (if open)
'3. Run this script
'Repeat the above steps to add more IE Mode pages

ClearAll = False 'Set to True to clear all existing IE Mode entries
Backup = True 'Set to False for no backup
Silent = False 'Change to True for no prompts
ForceLowercase = True 'Force domain part of URL to be lowercase
Setlocale("en-us") 'Do NOT change unless you change the date format for "DateAdded" below
DateAdded = "10/28/2099 10:00:00 PM" 'Specify the date here (ensure format is consistent with "Setlocale")

'To add sites, copy, uncomment and edit the AddSites line below. Separate each page entry with a |.
'Entries must end with a slash unless the URL ends with a file such as .html, .aspx, etc.
'The domain part of the entry must be all lowercase.
'Edge IE Mode will not accept URL parameters, so the script will trim URLs at the first "?" character 
'AddSites = "http://www.fait.it/|http://www.ferari.it/"

'To find and replace any string in the Preferences file, copy, uncomment and edit the FindReplace line below.
'Be very careful that you specify exact, unique, case-sensitive text! Use at your own risk!
'Separate find and replace strings with a comma and separate each find/replace pair with a |.
'FindReplace = "www.fait.it,www.fiat.it|www.ferari.it,www.ferrari.it"

Const ForWriting = 2
Const Ansi = 0
Dim PrefsFile,MyLog,Data,OriginalData
Z = VBCRLF
ZZ = VBCRLF & VBCRLF

'Convert AddSites and FindReplace lists to arrays"
aAddSites = Split(AddSites,"|")
aFindReplace = Split(FindReplace,"|")

'Convert the date 
Set oDateTime = CreateObject("WbemScripting.SWbemDateTime")
Call oDateTime.SetVarDate(DateAdded,True)
EdgeDateAdded = Left(oDateTime.GetFileTime,17)

Set oWSH = CreateObject("WScript.Shell")
Set oFSO = CreateObject("Scripting.FileSystemObject")

CScript = InStr(LCase(WScript.FullName),"cscript")>0

If Not Silent And Not CScript Then
  MC = "Clear all existing IE Mode pages." & ZZ
  MA = "Add these URLs to IE Mode: " & Z & AddSites & ZZ
  MF = "Find and replace these strings: " & Z & FindReplace & ZZ
  MX = "Set all IE Mode pages to expire 30 days after: " & DateAdded & ZZ
  MB = "Create a Preferences backup file if any changes are made." & ZZ
  MS = "Skip any synced profiles (because Edge will reject edits)." & ZZ
  MK = "Kill any running MSEdge.exe tasks."
  If ClearAll Then Msg = Msg & MC
  If AddSites<>"" Then Msg = Msg & MA
  If FindReplace<>"" Then Msg = Msg & MF
  If (Not ClearAll) Or AddSites<>"" Then Msg = Msg & MX
  If Backup Then Msg = Msg & MB
  Msg = Msg & MS & MK
  Response = MsgBox(Msg,VBOKCancel,"The following actions will be performed:")
  If Response=VBCancel Then WScript.Quit
End If

'Edge must be closed to modify the Preferences file
oWSH.Run "TaskKill /im MSEdge.exe /f",0,True

LocalAppData = oWSH.ExpandEnvironmentStrings("%LocalAppData%")

LogMsg "Profiles processed:"

ProcessProfiles("Edge") 'For released Edge profile
ProcessProfiles("Edge Beta") 'For Beta Edge profile
ProcessProfiles("Edge Dev") 'For Dev Edge profile
ProcessProfiles("Edge SxS") 'For Canary Edge profile

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

Sub ClearEntries
  FirstBlockEnd = InStr(Data,"user_list_data_1") + 18
  SecondBlockStart = InStr(Data,"}},""edge""")
  Data = Left(Data,FirstBlockEnd) & Mid(Data,SecondBlockStart)
End Sub

'Find and change every IE Mode page entry
Sub UpdateEntries
  StartPos = 1
  Do
    FoundPos = InStr(StartPos,Data,"date_added")
    If FoundPos=0 Then Exit Do
    Data = Mid(Data,1,FoundPos + 12) & EdgeDateAdded & Mid(Data,FoundPos + 30)
    StartPos = FoundPos + 1
  Loop
End Sub

'Add any sites specified with the AddSites variable
Sub AddEntries
  For i = 0 To UBound(aAddSites)
    AddSite = FixURL(aAddSites(i))
    If Not BadURL(AddSite) Then
      AddSite = FixURL(aAddSites(i))
      If Instr(Data,"user_list_data_1")=0 Then Data = Replace(Data,"},""edge"":{",",""user_list_data_1"":{}},""edge"":{")
      If Instr(Data,AddSite)=0 Then Data = Replace(Data,"""user_list_data_1"":{","""user_list_data_1"":{""" & AddSite & """:{""date_added"":""" & EdgeDateAdded & """,""engine"":2,""visits_after_expiration"":0},")
      Data = Replace(Data,"},}},","}}},")
    End If
  Next
End Sub

'Find and replace strings specified with the FindReplace variable
Sub FindReplaceEntries
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

  If ClearAll Then Call ClearEntries Else UpdateEntries

  AddEntries

  FindReplaceEntries
  
  'Set "Allow sites to be reloaded in Internet Explorer mode" to "Allow"
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

Sub ProcessProfiles(ProfileFolder)
  EdgeData = LocalAppData & "\Microsoft\" & ProfileFolder & "\User Data\"
  If oFSO.FolderExists(EdgeData) Then
    For Each oFolder In oFSO.GetFolder(EdgeData).SubFolders
      PrefsFile = oFolder.Path & "\Preferences"
      If oFSO.FileExists(PrefsFile) Then EditProfile
    Next
  End If
End Sub

If Not Silent Then
  WScript.Echo MyLog
End If