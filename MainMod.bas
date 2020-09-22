Attribute VB_Name = "MainMod"
Public Enum ltabs
    Setup = 1
    LANConnect = 2
    Routers = 3
End Enum

Public Enum ctype
    pic = 1
    tip = 2
End Enum

Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum

Public Enum errors
    UNKNOWN_ERROR = 0
    DATABASE_NOT_FOUND = 1
    MISSING_DATA = 2
    INVALID_DATABASE = 3
    INVALID_ROUTER = 4
    INVALID_HTTP = 5
    INVALID_IRC = 6
    UNABLE_RESOLVE = 7
    APP_START = 8
    APP_QUIT = 9
    INVALID_TYPE = 10
    INVALID_NAME = 11
    MISSING_APP = 12
    INVALID_CHARACTER = 13
    HELP_FILE_MISSING = 14
End Enum

Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function gethostbyname Lib "wsock32" (ByVal hostname As String) As Long
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const REG_SZ = 1
Public Const ERROR_SUCCESS = 0&
Public Const WM_CLOSE = &H10
Public Const SYNCH_PAGE = "http://asp.crazymateo.com/synchronize.asp"
Public Const SB_TOP = 6
Public Const WM_VSCROLL& = &H115
Public Const MF_STRING = &H0&
Public Const MF_SEPARATOR = &H800&
Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111
Public Const TPM_BOTTOMALIGN = &H20&

Public Modification As String
Public TrayIcon As NOTIFYICONDATA
Public ProgramName() As String
Public ProgramhWnd() As Long
Public Proxy As String
Public playsounds As Boolean
Public Status As listview
Public Menu As Long
Public OldWindowLong As Long

Public Sub Main()
On Error GoTo err_handler
Dim temp As String
DataPath = App.path & "\Settings.mdb": DataErrors = 0: ExternalIP = "0.0.0.0"
If PathFileExists(App.path & "\Corrupt.tmp") Then Kill DataPath: Kill App.path & "\Corrupt.tmp"
If PathFileExists(DataPath) = 0 Then CreateDatabase DataPath: CreateDefaultTables DataPath
temp = ReadFromDatabase(DataPath, "Setup", "Check0", "Value", ";", "norecord", "null", True, False)
If temp = "1" Then
    If InRun("DNSUpdater") = False Then AddToRun "DNSUpdater", App.path & "\DNSUpdater.exe"
Else
    If InRun("DNSUpdater") = True Then DeleteFromRun "DNSUpdater"
End If
Set Status = MainFrm.ListView0: LoadHelp App.path & "\help.chm": LoadDefaultMenu MainFrm, MainFrm.ImageList0
Proxy = ReadFromDatabase(DataPath, "LANConnect", "Text9", "Value", ";", "norecord", "null", True, False)
If Trim(Proxy) = vbNullString Or Proxy = "norecord" Or Proxy = "null" Then Proxy = vbNullString
temp = ReadFromDatabase(DataPath, "Setup", "Check6", "Value", ";", "norecord", "null", True, False)
If temp = "1" Then playsounds = True
ReDim ProgramName(0): ReDim ProgramhWnd(0)
LoadPic DataPath, "Misc", "Default", MainFrm.Picture0, MainFrm.ImageList0, MainFrm: LoadTrayIcon MainFrm.Picture0
temp = ReadFromDatabase(DataPath, "Setup", "Check1", "Value", ";", "norecord", "null", True, False)
If temp <> "1" Then
    MainFrm.Show: ModifyMenu Menu, 0, MF_STRING, 0, "Hide"
    SetMenuItemBitmaps Menu, 0, 0, MainFrm.ImageList0.ListImages(10).Picture, MainFrm.ImageList0.ListImages(10).Picture
End If
If playsounds = True Then PlaySound DataPath, "Misc", "Start"
Exit Sub
err_handler:
    AddError APP_START
End Sub

Public Sub MainQuit()
On Error GoTo err_handler
Dim temp As String
If playsounds = True Then PlaySound DataPath, "Misc", "Quit", 2
For Each Form In Forms
    Unload Form
Next Form
MainFrm.Inet0.Cancel: MainFrm.Winsock0.Close
UnloadTrayIcon MainFrm.Picture0
ClosePrograms
Exit Sub
err_handler:
    AddError APP_QUIT
End Sub

Public Function LoadHelp(path As String)
Dim temp As String
If PathFileExists(path) = 0 Then
    For Each Control In MainFrm.Controls
        If InString(Control.name, "imagelist inet commondialog winsock timer line") = False Then
            Control.WhatsThisHelpID = 0
        End If
    Next Control
Else
    App.HelpFile = path & "::/WhatsThisHelpTopics.txt"
End If
End Function

Public Function ListAddString(list As Variant, str As String, Optional delimiter As String)
On Error GoTo err_handler
Dim temp() As String
Dim x As Integer
list.Clear
If Trim(str) = vbNullString Then Exit Function
temp = Split(str, IIf(Trim(delimiter) = vbNullString, ";", delimiter))
For x = 0 To UBound(temp)
    list.AddItem temp(x)
Next x
Exit Function
err_handler:
    AddError MISSING_DATA, "Unable To Add String To List."
End Function

Public Function ListViewAddString(listview As listview, values As String, Optional delimiter1 As String, Optional delimiter2 As String)
On Error GoTo err_handler
Dim temp1() As String
Dim temp2() As String
Dim x, y As Integer
listview.ListItems.Clear
If Trim(values) = vbNullString Then Exit Function
temp1 = Split(values, IIf(Trim(delimiter1) = vbNullString, ";", delimiter1))
For x = 0 To UBound(temp1)
    temp2 = Split(temp1(x), IIf(Trim(delimiter2) = vbNullString, ",", delimiter2))
    If listview.ColumnHeaders.Count - 1 = UBound(temp2) Then
        For y = 0 To UBound(temp2)
            If y = 0 Then listview.ListItems.Add , , temp2(y) Else: listview.ListItems.Item(x + 1).ListSubItems.Add , , temp2(y)
        Next y
    End If
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Add String To List View."
End Function

Public Function AddMessage(obj As listview, message As String, Optional error As Boolean)
On Error GoTo err_handler
Dim temp As String
Dim x As Integer
If Trim(message) = vbNullString Then AddError MISSING_DATA, "Unable To Add Message.": Exit Function
If obj.ListItems.Count = "99999" Then obj.ListItems.Clear
obj.ListItems.Add , , "[" & Time & "] " & message, , IIf(error, 8, 7)
obj.ListItems(obj.ListItems.Count).ListSubItems.Add , , AddZeros(CStr(obj.ListItems.Count), 5)
If obj.ListItems.Count > 8 Then obj.ColumnHeaders(1).Width = obj.Width - 310
obj.SortKey = 1: obj.Sorted = True
SendMessage obj.hwnd, WM_VSCROLL, SB_TOP, obj.hwnd
obj.ListItems(1).Selected = True: obj.ListItems(1).Selected = False
temp = ReadFromDatabase(DataPath, "Setup", "Check7", "Value", ";", "norecord", "null", True, True)
If temp <> "1" Then Exit Function
temp = ReadFromDatabase(DataPath, "Misc", "LogFile", "Value", ";", "norecord", "null", True, True)
temp = ApplyConstants(temp)
If temp Like "?:\*.log" Then
    x = FreeFile
    CreateDirectoryStructure temp
    Open temp For Append As x
        Print #x, "[" & Time & "] " & message
    Close x
End If
Exit Function
err_handler:
    If PathFileExists(DataPath) Then WriteToDatabase DataPath, "Setup", "Check7", "Value", "0", ";", True
    MainFrm.Check1(11).value = False
    MsgBox "Logging Has Been Disabled Because Of Errors."
End Function

Private Function AddZeros(number As String, Length As Integer) As String
AddZeros = number
Do Until Len(AddZeros) >= Length
    AddZeros = "0" & AddZeros
Loop
End Function

Public Function IsFormLoaded(name As String) As Boolean
On Error GoTo err_handler
If Trim(name) = vbNullString Then AddError MISSING_DATA, "Unable To Check If Form Loaded.": Exit Function
For Each Form In Forms
    If LCase(name) = LCase(Form.name) Then IsFormLoaded = True: Exit Function
Next Form
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Check If Form Loaded."
End Function

Public Function AddHeaders(listview As listview, str As String, Optional delimiter As String)
On Error GoTo err_handler
Dim temp1() As String
Dim temp2() As String
Dim x As Integer
If Trim(str) = vbNullString Then AddError MISSING_DATA, "Unable To Add Headers.": Exit Function
temp = Split(str, IIf(Trim(delimiter) = vbNullString, ";", delimiter))
If UBound(temp) = 1 Then
    listview.ColumnHeaders.Add , , temp(0), (listview.Width / 3) - 40
    listview.ColumnHeaders.Add , , temp(1), 2 * (listview.Width / 3) - 40
End If
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Add Headers."
End Function

Public Function FixHeaders(listview As listview, rows As Integer)
On Error GoTo err_handler
If rows = 0 Then AddError MISSING_DATA, "Unable To Fix Headers": Exit Function
listview.ColumnHeaders(1).Width = IIf(listview.ListItems.Count > rows, (listview.Width / 3) - 160, (listview.Width / 3) - 40)
listview.ColumnHeaders(2).Width = IIf(listview.ListItems.Count > rows, 2 * (listview.Width / 3) - 160, 2 * (listview.Width / 3) - 40)
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Fix Headers."
End Function

Public Function ApplyConstants(txtString As String) As String
On Error GoTo err_handler
Dim temp1() As String
Dim temp2(2) As String
Dim x As Integer
If Trim(txtString) = vbNullString Then ApplyConstants = txtString: Exit Function
temp1 = Split("%UserName%;%Password%;%RouterIP%;%RouterPassword%;%RouterPort%;%AppPath%;%Time%;%Date%;%IP%", ";")
ApplyConstants = txtString
For x = 0 To UBound(temp1)
    If InStr(LCase(ApplyConstants), LCase(temp1(x))) <> 0 Then
        Select Case LCase(temp1(x))
            Case "%username%", "%password%"
                temp2(0) = "Setup"
                If LCase(temp1(x)) = "%username%" Then temp2(1) = "Text4" Else: temp2(1) = "Text5"
            Case "%routerpassword%", "%routerip%", "%routerport%"
                temp2(0) = "LANConnect"
                If LCase(temp1(x)) = "%routerpassword%" Then temp2(1) = "Text0"
                If LCase(temp1(x)) = "%routerip%" Then temp2(1) = "Text1"
                If LCase(temp1(x)) = "%routerport%" Then temp2(1) = "Text2"
            Case "%time%"
                temp2(2) = Time: temp2(0) = vbNullString: temp2(1) = vbNullString
                If InStr(temp2(2), ":") <> 0 Then temp2(2) = Replace(temp2(2), ":", "")
            Case "%date%"
                temp2(2) = Date: temp2(0) = vbNullString: temp2(1) = vbNullString
                If InStr(temp2(2), "/") <> 0 Then temp2(2) = Replace(temp2(2), "/", "-")
            Case "%apppath%"
                temp2(2) = App.path: temp2(0) = vbNullString: temp2(1) = vbNullString
            Case "%ip%"
                temp2(2) = ExternalIP
        End Select
        If temp2(0) <> vbNullString And temp2(1) <> vbNullString Then
            temp2(2) = ReadFromDatabase(DataPath, temp2(0), temp2(1), "Value", ";", "norecord", "null", True)
            If temp2(2) = "norecord" Or temp2(2) = "null" Then temp2(2) = vbNullString
        End If
        ApplyConstants = Replace(ApplyConstants, temp1(x), LCase(temp2(2)), , , vbTextCompare)
    End If
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Apply Constants."
End Function

Private Function CreateDirectoryStructure(path As String)
On Error GoTo err_handler
Dim temp1() As String
Dim temp2 As String
Dim x As Integer
If Trim(path) = vbNullString Then AddError MISSING_DATA, "Unable To Create Directory Structure.": Exit Function
temp1 = Split(path, "\")
If PathFileExists(temp1(0)) Then
    If UBound(temp1) = 0 Then Exit Function
    If UBound(temp1) = 1 And Right(temp1(1), 4) Like ".???" Then Exit Function
    For x = 0 To UBound(temp1)
        temp2 = IIf(Trim(temp2) = vbNullString, vbNullString, temp2 & "\") & temp1(x)
        If Right(temp2, 4) Like ".???" Then Exit For
        If PathFileExists(temp2) = 0 Then MkDir temp2
    Next x
Else
    Exit Function
End If
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Create Directory Structure."
End Function

Public Function LoadTrayIcon(picturebox As picturebox)
On Error GoTo err_handler
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = picturebox.hwnd
TrayIcon.uId = 1&
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.ucallbackMessage = WM_LBUTTONDOWN
TrayIcon.hIcon = picturebox.Picture
TrayIcon.szTip = "DNSUpdater" & Chr(0)
Shell_NotifyIcon NIM_ADD, TrayIcon
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Load Tray Icon."
End Function

Public Function UnloadTrayIcon(picturebox As picturebox)
On Error GoTo err_handler
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hwnd = picturebox.hwnd
TrayIcon.uId = 1&
Shell_NotifyIcon NIM_DELETE, TrayIcon
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Unload Tray Icon."
End Function

Public Function LoadPic(path As String, table As String, name As String, picturebox As picturebox, imagelist As imagelist, frm As Form)
On Error GoTo err_handler
Dim temp As String
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Load Picture.": Exit Function
If Trim(table) = vbNullString Or Trim(name) = vbNullString Then AddError MISSING_DATA, "Unable To Load Picture.": Exit Function
temp = ReadFromDatabase(path, table, "Picture" & name, "Value", ";", "norecord", "null", True, False)
If PathFileExists(temp) And LCase(Right(temp, 4)) = ".bmp" Or PathFileExists(temp) And LCase(Right(temp, 4)) = ".ico" Then
    picturebox.Picture = LoadPicture(temp, vbLPSmall)
    If Trim(LCase(name)) = "default" Then frm.Icon = picturebox.Picture
Else
    Select Case LCase(name)
        Case "default"
            picturebox.Picture = imagelist.ListImages(1).Picture
            frm.Icon = picturebox.Picture
        Case "updating"
            picturebox.Picture = imagelist.ListImages(2).Picture
        Case "synchronizing"
            picturebox.Picture = imagelist.ListImages(3).Picture
        Case "retrying"
            picturebox.Picture = imagelist.ListImages(4).Picture
        Case "success"
            picturebox.Picture = imagelist.ListImages(5).Picture
        Case "failure"
            picturebox.Picture = imagelist.ListImages(6).Picture
        Case Else
            AddError INVALID_NAME, "Unable To Load Picture.": Exit Function
    End Select
End If
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Load Picture."
End Function

Public Function ChangeTrayIcon(ctype As ctype, tip_name As String, Optional picturebox As picturebox, Optional imagelist As imagelist, Optional frm As Form)
On Error GoTo err_handler
If Trim(tip_name) = vbNullString Then AddError MISSING_DATA, "Unable To Change Tray Icon.": Exit Function
If ctype = pic Then
    LoadPic DataPath, "Misc", tip_name, picturebox, imagelist, frm
    TrayIcon.hIcon = picturebox.Picture
ElseIf ctype = tip Then
    TrayIcon.szTip = tip_name & Chr(0)
End If
Shell_NotifyIcon NIM_MODIFY, TrayIcon
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Change Tray Icon."
End Function

Private Function SaveRegistryValue(hkey As HKeyTypes, strPath As String, strValue As String, strdata As String)
On Error GoTo err_handler
Dim keyhand As Long
Dim temp As Long
If Trim(strValue) = vbNullString Or Trim(strdata) = vbNullString Then AddError MISSING_DATA, "Unable To Save Registry Value.": Exit Function
temp = RegCreateKey(hkey, strPath, keyhand)
temp = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
temp = RegCloseKey(keyhand)
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Save Registry Value."
End Function

Private Function DeleteRegistryValue(ByVal hkey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)
On Error GoTo err_handler
Dim keyhand As Long
Dim temp As Long
If Trim(strPath) = vbNullString Or Trim(strValue) = vbNullString Then AddError MISSING_DATA, "Unable To Delete Registry Value.": Exit Function
temp = RegOpenKey(hkey, strPath, keyhand)
temp = RegDeleteValue(keyhand, strValue)
temp = RegCloseKey(keyhand)
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Delete Registry Value."
End Function

Public Function GetRegistryValue(ByVal hkey As HKeyTypes, ByVal strPath As String, strValue As String, Optional nokey As String)
On Error GoTo err_handler
Dim keyhand As Long
Dim temp As Long
Dim lValueType As Long
Dim lDataBufferSize As Long
Dim buffer As String
If Trim(strPath) = vbNullString Or Trim(strValue) = vbNullString Then AddError MISSING_DATA, "Unable To Get Registry Value.": Exit Function
temp = RegOpenKey(hkey, strPath, keyhand)
temp = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
If lValueType = REG_SZ Then
    buffer = String(lDataBufferSize, " ")
    temp = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal buffer, lDataBufferSize)
    If temp = ERROR_SUCCESS Then
        If InStr(buffer, Chr(0)) > 0 Then GetRegistryValue = Left(buffer, InStr(buffer, Chr(0)) - 1)
        Else: GetRegistryValue = buffer
    End If
End If
If Trim(GetRegistryValue) = vbNullString Then GetRegistryValue = IIf(Trim(nokey) = vbNullString, "nokey", nokey)
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get Registry Value."
End Function

Public Function InRun(ProgramName As String) As Boolean
On Error GoTo err_handler
If GetRegistryValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, "nokey") <> "nokey" Then InRun = True
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Check If In Run."
End Function

Public Function AddToRun(ProgramName As String, FileToRun As String)
On Error GoTo err_handler
SaveRegistryValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName, FileToRun
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Add To Run."
End Function

Public Function DeleteFromRun(ProgramName As String)
On Error GoTo err_handler
DeleteRegistryValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", ProgramName
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Delete From Run."
End Function

Public Function UnloadVisibleForms()
On Error GoTo err_handler
For Each Form In Forms
    If Form.Visible = True Then Unload Form
Next Form
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Unload Visible Forms."
End Function

Public Function RunSynchronize(frm As Form, path As String, animatetray As Boolean, playsounds As Boolean)
On Error GoTo err_handler
Dim temp1(2) As String
Dim temp2() As String
Dim x As Integer
If animatetray = True Then ChangeTrayIcon pic, "Synchronizing", frm.Picture0, frm.ImageList0: ChangeTrayIcon tip, "DNSUpdater: Synchronizing"
AddMessage Status, "Synchronizing"
If playsounds = True Then PlaySound DataPath, "Misc", "Synchronizing"
temp1(1) = ReadFromDatabase(path, "LANConnect", "Text9", "Value", ";", "norecord", "null", True, False)
temp1(0) = GetTableNames(path, "Services", ";", "norecord")
temp1(0) = OpenHTTP(frm, SYNCH_PAGE & "?table=services&exclude=" & temp1(0), IIf(Trim(temp1(1)) = vbNullString, vbNullString, temp1(1)))
If temp1(0) = "error" Then GoTo err_handler
temp1(2) = temp1(0): temp2 = Split(temp1(0), ";")
For x = 0 To UBound(temp2)
    If temp2(x) Like "*^*^*^*" Then WriteToDatabase path, "Services", Mid(temp2(x), 1, InStr(temp2(x), "^") - 1), "Name^Address^Fields^Keyword", temp2(x), "^", False
Next x
temp1(0) = GetTableNames(path, "Routers", ";", "norecord")
temp1(0) = OpenHTTP(frm, SYNCH_PAGE & "?table=routers&exclude=" & temp1(0), IIf(Trim(temp1(1)) = vbNullString, vbNullString, temp1(1)))
If temp1(0) = "error" Then GoTo err_handler
If Trim(temp1(2)) = vbNullString Then temp1(2) = temp1(0)
temp2 = Split(temp1(0), ";")
For x = 0 To UBound(temp2)
    If temp2(x) Like "*^*^*^*^*" Then WriteToDatabase path, "Routers", Mid(temp2(x), 1, InStr(temp2(x), "^") - 1), "Name^LogIn^LogOut^Status^Keyword", temp2(x), "^", False
Next x
For x = 2 To 5
    frm.TabStrip0.Tabs(x).Tag = False
Next x
If animatetray = True Then ChangeTrayIcon pic, "Success", frm.Picture0, frm.ImageList0: ChangeTrayIcon tip, "DNSUpdater: Success"
AddMessage Status, IIf(Trim(temp1(2)) = vbNullString, "Database Already Synchronized", "Synchronization Complete")
If playsounds = True Then PlaySound DataPath, "Misc", "Success"
Exit Function
err_handler:
    If animatetray = True Then ChangeTrayIcon pic, "Failure", frm.Picture0, frm.ImageList0: ChangeTrayIcon tip, "DNSUpdater: Failure"
    AddError UNKNOWN_ERROR, "Unable To Run Synchronize."
    If playsounds = True Then PlaySound DataPath, "Misc", "Failure"
End Function

Public Function PlaySound(path As String, table As String, name As String, Optional mode As Integer)
On Error GoTo err_handler
Dim temp As String
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Play Sound.": Exit Function
If Trim(table) = vbNullString Or Trim(name) = vbNullString Then AddError MISSING_DATA, "Unable To Play Sound.": Exit Function
temp = ReadFromDatabase(path, table, "Sound" & name, "Value", ";", "norecord", "null", True, False)
temp = ApplyConstants(temp)
If PathFileExists(temp) = 0 Or Right(temp, 4) <> ".wav" Then AddError INVALID_TYPE, "Unable To Play Sound.": Exit Function
sndPlaySound temp, IIf(mode = 0, 1, mode)
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Play Sound."
End Function

Public Function InstanceToWnd(ByVal instance As Long) As Long
On Error GoTo err_handler
Dim temp(2) As Long
Dim x As Integer
If instance <= 0 Then AddError MISSING_DATA, "Unable To Convert Instance To Wnd.": Exit Function
temp(0) = FindWindow(ByVal 0&, ByVal 0&)
Do Until x = 400
    DoEvents
    x = x + 1
Loop
Do While temp(0) <> 0
    If GetParent(temp(0)) = 0 Then
        temp(2) = GetWindowThreadProcessId(temp(0), temp(1))
        If temp(1) = instance Then InstanceToWnd = temp(0): Exit Do
    End If
    temp(0) = GetWindow(temp(0), 2)
Loop
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Convert Instance To Wnd."
End Function

Public Function AddToArray(arr As Variant, value As String)
On Error GoTo err_handler
If IsArray(arr) = False Then AddError INVALID_TYPE, "Unable To Add To Array.": Exit Function
ReDim Preserve arr(UBound(arr) + 1)
arr(UBound(arr)) = value
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Add To Array."
End Function

Public Function DeleteFromArray(arr As Variant, value As String)
On Error GoTo err_handler
Dim temp() As Variant
Dim x As Integer
If Trim(value) = vbNullString Then AddError MISSING_DATA, "Unable To Delete From Array.": Exit Function
If IsArray(arr) = False Then AddError INVALID_TYPE, "Unable To Delete From Array.": Exit Function
If InArray(arr, value) = False Then AddError MISSING_DATA, "Unable To Delet From Array.": Exit Function
ReDim temp(0)
For x = 1 To UBound(arr)
    If LCase(Trim(arr(x))) <> LCase(Trim(value)) Then
        ReDim Preserve temp(UBound(temp) + 1)
        temp(UBound(temp)) = arr(x)
    End If
Next x
ReDim arr(UBound(temp))
For x = 1 To UBound(temp)
    arr(x) = temp(x)
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Delete From Error."
End Function

Public Function InArray(arr As Variant, value As String) As Boolean
On Error GoTo err_handler
Dim x As Integer
If Trim(value) = vbNullString Then AddError MISSING_DATA, "Unable To Find In Array.": Exit Function
If IsArray(arr) = False Then Exit Function
For x = 0 To UBound(arr)
    If LCase(arr(x)) = LCase(value) Then InArray = True: Exit Function
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Find In Array.": Exit Function
End Function

Public Function RunProgram(filename As String) As Long
On Error GoTo err_handler
Dim temp As Long
If PathFileExists(filename) = 0 Then AddError MISSING_APP, "Unable To Run Program.": Exit Function
If LCase(Right(filename, 4)) <> ".exe" Then AddError INVALID_TYPE, "Unable To Run Program.": Exit Function
temp = Shell(filename)
temp = InstanceToWnd(temp)
RunProgram = temp
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Run Program."
End Function

Public Function StartPrograms(path As String, table As String)
On Error GoTo err_handler
Dim temp1() As String
Dim temp2 As String
Dim x As Integer
temp2 = ReadFromDatabase(path, "Setup", "Check2", "Value", ";", "norecord", "null", True, False)
If Trim(table) = vbNullString Then AddError MISSING_DATA, "Unable To Start Programs.": Exit Function
If temp2 <> "1" Then Exit Function
RefreshProgramArray
temp1 = Split(GetTableNames(path, table, ";", "norecord"), ";")
If temp1(0) = "norecord" Then Exit Function
For x = 0 To UBound(temp1)
    temp2 = ReadFromDatabase(path, table, temp1(x), "Value", ";", "norecord", "null", False, False)
    If PathFileExists(temp2) And LCase(Right(temp2, 4)) = ".exe" And InArray(ProgramName, temp1(x)) = False Then
        temp2 = RunProgram(temp2)
        AddToArray ProgramName, temp1(x): AddToArray ProgramhWnd, temp2
    End If
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Start Programs."
End Function

Public Function ClosePrograms(Optional name As String)
On Error GoTo err_handler
Dim temp
Dim tempName() As String
Dim temphWnd() As Long
Dim x As Integer
temp = ReadFromDatabase(DataPath, "Setup", "Check3", "Value", ";", "norecord", "null", True, False)
If temp <> "1" Then Exit Function
ReDim tempName(0): ReDim temphWnd(0)
RefreshProgramArray
For x = 1 To UBound(ProgramName)
    If Trim(name) <> vbNullString Then
        SendMessage ProgramhWnd(x), WM_CLOSE, 0, 0
        AddToArray tempName, ProgramName(x)
        AddToArray temphWnd, CStr(ProgramhWnd(x))
        Exit For
    Else
        SendMessage ProgramhWnd(x), WM_CLOSE, 0, 0
    End If
Next x
For x = 1 To UBound(tempName)
    DeleteFromArray ProgramName, tempName(x)
    DeleteFromArray ProgramhWnd, CStr(temphWnd(x))
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Close Program" & IIf(Trim(name) = vbNullString, "s", vbNullString) & "."
End Function

Public Function RefreshProgramArray()
On Error GoTo err_handler
Dim tempName() As String
Dim temphWnd() As Long
Dim x As Integer
ReDim tempName(0)
ReDim temphWnd(0)
For x = 1 To UBound(ProgramhWnd)
    If IsWindow(ProgramhWnd(x)) = 0 Then
        AddToArray tempName, ProgramName(x)
        AddToArray temphWnd, CStr(ProgramhWnd(x))
    End If
Next x
For x = 1 To UBound(tempName)
    DeleteFromArray ProgramName, tempName(x)
    DeleteFromArray ProgramhWnd, CStr(temphWnd(x))
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Refresh Program Array."
End Function

Public Function AddError(message As errors, Optional txt As String)
On Error Resume Next
Dim temp As String
Select Case message
    Case 0
        temp = "Unknown Error Occurred"
    Case 1
        temp = "Database Not Found"
    Case 2
        temp = "Missing Recquired Data"
    Case 3
        temp = "Database Is Invalid"
    Case 4
        temp = "Invalid Router"
    Case 5
        temp = "Invalid HTTP Page"
    Case 6
        temp = "Invalid IRC Server"
    Case 7
        temp = "Unable To Resolve Host"
    Case 8
        temp = "Problems Occurred Starting"
    Case 9
        temp = "Problems Occurred Quiting"
    Case 10
        temp = "Invalid Type Selected"
    Case 11
        temp = "Invalid Name"
    Case 12
        temp = "Appication Does Not Exist"
    Case 13
        temp = "Invalid Character Detected"
    Case 14
        temp = "Help File Not Found"
    Case Else
        temp = "Unknown Error Occurred"
End Select
If Trim(txt) <> vbNullString Then temp = temp & ". " & txt
AddMessage Status, temp, True
End Function

Public Function InString(str As String, characters As String, Optional delimiter As String) As Boolean
Dim temp() As String
Dim x As Integer
If Trim(str) = vbNullString Or Trim(characters) = vbNullString Then Exit Function
temp = Split(characters, IIf(Trim(delimiter) = vbNullString, " ", delimiter))
For x = 0 To UBound(temp)
    If InStr(LCase(str), LCase(temp(x))) <> 0 Then InString = True: Exit For
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Tell If In String.": Exit Function
End Function

Public Function OpenHelp(frm As Form, path As String)
On Error GoTo err_handler
If PathFileExists(path) = 0 Then AddError HELP_FILE_MISSING, "Unable To Open Help.": Exit Function
HTMLHelp frm.hwnd, path, &H1, 0
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Open Help."
End Function

Public Function BuildMenu(items As String, Optional delimiter As String, Optional imagelist As imagelist, Optional pics As String) As Long
On Error GoTo err_handler
Dim temp1() As String
Dim temp2() As String
Dim x As Integer
If Trim(items) = vbNullString Then AddError MISSING_DATA, "Unable To Build Menu.": Exit Function
BuildMenu = CreatePopupMenu
temp1 = Split(items, IIf(Trim(delimiter) = vbNullString, ";", delimiter))
If Trim(pics) <> vbNullString Then temp2 = Split(pics, IIf(Trim(delimiter) = vbNullString, ";", delimiter)) Else: ReDim temp2(0)
For x = 0 To UBound(temp1)
    If Trim(temp1(x)) = "-" Then
        InsertMenu BuildMenu, x, MF_SEPARATOR Or MF_BYPOSITION Or MF_POPUP, x, temp1(x)
    Else
        InsertMenu BuildMenu, x, MF_STRING Or MF_BYPOSITION Or MF_POPUP, x, temp1(x)
        If x <= UBound(temp2) Then If temp2(x) >= 1 And temp2(x) <= imagelist.ListImages.Count Then SetMenuItemBitmaps BuildMenu, x, 0, imagelist.ListImages(CInt(temp2(x))).Picture, imagelist.ListImages(CInt(temp2(x))).Picture
    End If
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Build Menu."
End Function

Public Function LoadDefaultMenu(frm As Form, imagelist As imagelist)
On Error GoTo err_handler
Menu = BuildMenu("Show;-;Synchronize;Update;-;Help;Quit", ";", imagelist, "9;;11;12;;13;14")
OldWindowLong = GetWindowLong(frm.hwnd, GWL_WNDPROC)
SetWindowLong frm.hwnd, GWL_WNDPROC, AddressOf MenuCommand
Exit Function
err_handler:
    If OldWindowLong <> 0 Then SetWindowLong frm.hwnd, GWL_WNDPROC, OldWindowLong
    AddError UNKNOWN_ERROR, "Unable To Load Default Menu."
End Function

Public Function MenuCommand(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo err_handler
Select Case wMsg
    Case WM_CLOSE
        SetWindowLong hwnd, GWL_WNDPROC, OldWindowLong
    Case WM_COMMAND
        Select Case CInt(wParam)
            Case 0: MainFrm.Command0_Click 3
            Case 2: MainFrm.Command0_Click 1
            Case 3: MainFrm.Command0_Click 0
            Case 5: MainFrm.Command0_Click 2
            Case 6: MainFrm.Command0_Click 4
        End Select
End Select
MenuCommand = CallWindowProc(OldWindowLong, hwnd, wMsg, wParam, lParam)
Exit Function
err_handler:
    MenuCommand = CallWindowProc(OldWindowLong, hwnd, wMsg, wParam, lParam)
    SetWindowLong hwnd, GWL_WNDPROC, OldWindowLong
End Function
