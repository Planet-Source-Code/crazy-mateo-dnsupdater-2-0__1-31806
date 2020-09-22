Attribute VB_Name = "UpdateMod"
Public Declare Function RasEnumConnectionsA& Lib "RasApi32.DLL" (lprasconn As Any, lpcb&, lpcConnections&)
Public Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Public Declare Function WSACleanup Lib "wsock32" () As Long
Public Declare Function WSAStartup Lib "wsock32" (ByVal VersionReq As Long, WSADataReturn As WSAData) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type WSAData
    wversion As Integer
    wHighVersion As Integer
    szDescription(256) As Byte
    szSystemStatus(128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Public Type RASCON
    dwSize As Long
    hRasConn As Long
    szEntryName(256) As Byte
    szDeviceType(16) As Byte
    szDeviceName(128) As Byte
End Type

Public Type IPInfo
    address As Long
    Index As Long
    subnetmask As Long
    broadcastaddress As Long
    assemblysize  As Long
    unused1 As Integer
    unused2 As Integer
End Type

Public Type IPArray
    nEntries As Long
    IPInfo(10) As IPInfo
End Type

Public ExternalIP As String

Public Function RASIP(frm As Form) As String
On Error GoTo err_handler
Dim RAS As RASCON
Dim temp1 As Long
Dim temp2 As String
RAS.dwSize = 412
If RasEnumConnectionsA(RAS, RAS.dwSize, temp1) = 0 Then
    temp2 = ReadFromDatabase(DataPath, "Setup", "Combo1", "Value", ";", "norecord", "null", True, False)
    RASIP = IIf(temp1 = 0, "0.0.0.0", IIf(ValidIP(GetFullIP(temp2)), GetFullIP(temp2), IIf(ValidIP(ExternalIP) = False, "1.1.1.1", ExternalIP)))
Else
    RASIP = "0.0.0.0"
End If
clean:
    If Err.number <> 0 And ValidIP(RASIP) = False And RASIP <> "0.0.0.0" Then RASIP = ExternalIP
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get RASIP."
    Resume clean
End Function

Public Function GetFullIP(partip As String) As String
On Error GoTo err_handler
Dim temp1 As String
Dim temp2() As String
Dim x As Integer
If Trim(partip) = vbNullString Then AddError MISSING_DATA, "Unable To Get Full IP.": Exit Function
temp1 = GetLocalIPs
If Trim(temp1) = vbNullString Then GetFullIP = "0.0.0.0": Exit Function
temp2 = Split(temp1, ";")
For x = 0 To UBound(temp2)
    If temp2(x) Like partip Then GetFullIP = temp2(x): Exit For
Next x
clean:
    If Not Trim(GetFullIP) Like "*.*.*.*" Then GetFullIP = "0.0.0.0"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get Full IP."
    Resume clean
End Function

Public Function ValidIP(ip As String) As Boolean
On Error GoTo err_handler
Dim temp() As String
Dim x As Integer
If Not ip Like "*.*.*.*" Then Exit Function
temp = Split(ip, ".")
If UBound(temp) = 3 Then
    For x = 0 To UBound(temp)
        If Trim(temp(x)) = vbNullString Or Not IsNumeric(temp(x)) Then Exit Function
        Select Case x
            Case 0
                If temp(x) > 223 Or temp(x) <= 0 Then Exit Function
            Case 1 To 3
                If temp(x) > 255 Or temp(x) < 0 Then Exit Function
        End Select
    Next x
End If
ValidIP = True
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Verify IP."
End Function

Public Function GetLocalIPs(Optional delimiter As String) As String
On Error GoTo err_handler
Dim temp As Long
Dim buffer() As Byte
Dim IPArray As IPArray
Dim x As Integer
GetIpAddrTable ByVal 0&, temp, True
If temp > 0 Then
    ReDim buffer(temp - 1)
    GetIpAddrTable buffer(0), temp, False
    CopyMemory IPArray.nEntries, buffer(0), 4
    For x = 0 To IPArray.nEntries - 1
        CopyMemory IPArray.IPInfo(x), buffer(4 + (x * Len(IPArray.IPInfo(0)))), Len(IPArray.IPInfo(x))
        GetLocalIPs = IIf(Trim(GetLocalIPs) = vbNullString, vbNullString, GetLocalIPs & IIf(Trim(delimiter) = vbNullString, ";", delimiter)) & ConvertLongAddressToString(IPArray.IPInfo(x).address)
    Next x
End If
clean:
    If Not Trim(GetLocalIPs) Like "*.*.*.*" Then GetLocalIPs = "0.0.0.0"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get Local IPs."
    Resume clean
End Function

Public Function ConvertLongAddressToString(address As Long) As String
On Error GoTo err_handler
Dim buffer(3) As Byte
Dim x As Long
CopyMemory buffer(0), address, 4
For x = 0 To 3
    ConvertLongAddressToString = IIf(Trim(ConvertLongAddressToString) = vbNullString, vbNullString, ConvertLongAddressToString & ".") & CStr(buffer(x))
Next x
clean:
    If Not Trim(ConvertLongAddressToString) Like "*.*.*.*" Then ConvertLongAddressToString = "0.0.0.0"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Convert Long Address To String."
    Resume clean
End Function

Public Function RouterIP(frm As Form) As String
On Error GoTo err_handler
Dim temp1 As String
Dim temp2() As String
temp1 = ReadFromDatabase(DataPath, "LANConnect", "Combo", "Value", ";", "norecord", "null", True, False)
If IsName(temp1, Router) = False Then AddError INVALID_ROUTER, "Unable To Get Router IP.": RouterIP = ExternalIP: GoTo clean
temp2 = Split(ReadFromDatabase(DataPath, "Routers", temp1, "Name;LogIn;LogOut;Status;Keyword", ";", "norecord", "null", False, False), ";")
If Not temp2(3) Like "*://*" Then AddError INVALID_HTTP, "Unable To Get Router IP.": RouterIP = ExternalIP: GoTo clean
If temp2(1) Like "*://*" Then
    temp1 = OpenHTTP(frm, ApplyConstants(temp2(1)))
    If InStr(LCase(temp1), "incorrect") <> 0 Or temp1 = "error" Then RouterIP = ExternalIP: GoTo clean
End If
temp1 = OpenHTTP(frm, ApplyConstants(temp2(3)))
If temp1 = "error" Then RouterIP = ExternalIP: GoTo clean
RouterIP = FindIPInString(temp1, Trim(temp2(4)))
If temp2(2) Like "*://*" Then OpenHTTP frm, ApplyConstants(temp2(2))
clean:
    If RouterIP <> ExternalIP And ValidIP(RouterIP) = False And Err.number = 0 Then RouterIP = "0.0.0.0"
    Exit Function
err_handler:
    AddError MISSING_DATA, "Unable To Get Router IP."
    Resume clean
End Function

Public Function FindIPInString(str As String, Optional keyword As String) As String
On Error GoTo err_handler
Dim temp() As String
Dim x, y As Integer
If Trim(str) = vbNullString Then GoTo clean
If Trim(keyword) <> vbNullString Then x = InStr(str, keyword)
If x = 0 Then x = 1
Do Until x > Len(str)
    Do Until Mid(str, x, 15) Like "*.*.*.*" And IsNumeric(Mid(str, x, 1)) Or x > Len(str)
        x = x + 1
    Loop
    If x > Len(str) Then GoTo clean
    temp = Split(Mid(str, x, 15), "."): y = 1
    Do Until Not IsNumeric(Mid(temp(3), y, 1)) Or y > Len(temp(3))
        y = y + 1
    Loop
    temp(3) = Mid(temp(3), 1, y - 1)
    For y = 0 To 3
        FindIPInString = IIf(Trim(FindIPInString) = vbNullString, vbNullString, FindIPInString & ".") & temp(y)
        If Not IsNumeric(temp(y)) Then Exit For
    Next y
    If ValidIP(FindIPInString) Or FindIPInString = "0.0.0.0" Then Exit Do Else FindIPInString = vbNullString
    x = x + 1
Loop
clean:
    If Not Trim(FindIPInString) Like "*.*.*.*" Then FindIPInString = "0.0.0.0"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Find IP."
    Resume clean
End Function

Public Function OpenHTTP(frm As Form, url As String, Optional Proxy As String) As String
On Error GoTo err_handler
If frm.Inet0.StillExecuting = True Then GoTo clean
If Not url Like "*://*" Then AddError INVALID_HTTP, "Unable To Open HTTP.": GoTo clean
If Trim(Proxy) <> vbNullString Then frm.Inet0.Proxy = Proxy
frm.Inet0.Tag = vbNullString
frm.Inet0.RequestTimeout = 30
frm.Inet0.Execute url, "GET", , "User-Agent: DNSUpdater"
Do Until frm.Inet0.StillExecuting = False
    DoEvents
Loop
OpenHTTP = IIf(Trim(frm.Inet0.Tag) = vbNullString, vbNullString, frm.Inet0.Tag)
clean:
    frm.Inet0.Tag = vbNullString
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Open HTTP."
    Resume clean
End Function

Public Function HTTPIP(frm As Form) As String
On Error GoTo err_handler
Dim temp(2) As String
temp(0) = ReadFromDatabase(DataPath, "LANConnect", "Text3", "Value", ";", "norecord", "null", True, False)
If Not temp(0) Like "*://*" Then AddError INVALID_HTTP, "Unable To Get HTTP IP.": HTTPIP = ExternalIP: GoTo clean
temp(1) = ReadFromDatabase(DataPath, "LANConnect", "Text4", "Value", ";", "norecord", "null", True, False)
temp(2) = OpenHTTP(frm, ApplyConstants(temp(0)), Proxy)
If temp(2) = "error" Then HTTPIP = "0.0.0.0": GoTo clean
temp(2) = FindIPInString(temp(2), temp(1)): HTTPIP = temp(2)
clean:
    If HTTPIP <> ExternalIP And ValidIP(HTTPIP) = False And Err.number = 0 Then HTTPIP = "0.0.0.0"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get HTTPIP.": HTTPIP = ExternalIP
    Resume clean
End Function

Public Function IRCIP(frm As Form) As String
On Error GoTo err_handler
Dim temp(2) As String
Dim x As Long
If frm.Winsock0.State <> 0 Then IRCIP = ExternalIP: GoTo clean
frm.Winsock0.Tag = vbNullString
temp(0) = ReadFromDatabase(DataPath, "LANConnect", "Text5", "Value", ";", "norecord", "null", True, False)
If Not temp(0) Like "*.*" Or temp(0) = "norecord" Or temp(0) = "null" Then AddError INVALID_IRC, "Unable To Get IRC IP.": IRCIP = ExternalIP: GoTo clean
temp(1) = ReadFromDatabase(DataPath, "LANConnect", "Text6", "Value", ";", "norecord", "null", True, False)
If temp(1) = "norecord" Or temp(1) = "null" Or Not IsNumeric(temp(1)) Then temp(1) = "6667"
If IsNumeric(temp(1)) Then If temp(1) <= 0 Then temp(1) = "6667"
frm.Winsock0.Connect temp(0), temp(1)
Do Until Trim(frm.Winsock0.Tag) <> vbNullString Or x = 500000
    DoEvents
    x = x + 1
Loop
temp(2) = frm.Winsock0.Tag: frm.Winsock0.Tag = vbNullString
If x = 500000 Or temp(2) Like "error?" Or Trim(temp(2)) = vbNullString Then
    If Trim(temp(2)) = vbNullString Then IRCIP = ExternalIP Else IRCIP = "0.0.0.0"
    GoTo clean
Else
    IRCIP = HostToIP(temp(2)): IRCIP = IIf(ValidIP(IRCIP), IRCIP, ExternalIP)
End If
clean:
    If IRCIP <> ExternalIP And ValidIP(IRCIP) = False And Err.number = 0 Then IRCIP = "0.0.0.0"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get IRC IP."
    Resume clean
End Function

Public Function HostToIP(hostname As String) As String
On Error GoTo err_handler
Dim WSAData As WSAData
Dim host As HOSTENT
Dim temp1(1) As Long
Dim temp2() As Byte
Dim x As Integer
If Trim(hostname) = vbNullString Then AddError MISSING_DATA, "Unable To Get Host IP.": GoTo clean
If WSAStartup(257, WSAData) Then GoTo err_handler
temp1(0) = gethostbyname(hostname)
If temp1(0) = 0 Then AddError UNABLE_RESOLVE: GoTo clean
RtlMoveMemory host, temp1(0), LenB(host)
RtlMoveMemory temp1(1), host.hAddrList, 4
ReDim temp2(host.hLength)
RtlMoveMemory temp2(1), temp1(1), host.hLength
For x = 1 To host.hLength
    HostToIP = IIf(Trim(HostToIP) = vbNullString, vbNullString, HostToIP & ".") & temp2(x)
Next x
clean:
    WSACleanup
    If Not HostToIP Like "*.*.*.*" Then HostToIP = "0.0.0.0"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get Host IP."
    Resume clean
End Function

Public Function NewIP(frm As Form) As Boolean
On Error GoTo err_handler
Dim temp1 As String
Dim temp2 As String
temp1 = ReadFromDatabase(DataPath, "Setup", "Check5", "Value", ";", "norecord", "null", True, False)
If temp1 = "1" Then
    temp1 = ReadFromDatabase(DataPath, "LANConnect", "Option", "Value", ";", "norecord", "null", True, False)
    Select Case temp1
        Case 1: temp2 = IRCIP(frm)
        Case 2: temp2 = HTTPIP(frm)
        Case Else: temp2 = RouterIP(frm)
    End Select
Else
    temp2 = RASIP(frm)
End If
clean:
    If Err.number = 0 And ValidIP(temp2) And Trim(temp2) <> ExternalIP Or temp2 = "0.0.0.0" And Trim(temp2) <> ExternalIP Then NewIP = True: ExternalIP = temp2
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Determine If New IP."
    Resume clean
End Function

Public Function RunUpdate(frm As Form, animatetray As Boolean, Optional retrynumber As Integer, Optional playsounds As Boolean) As Boolean
On Error GoTo err_handler
Dim temp1 As String
Dim temp2() As String
Dim x As Integer
temp1 = ReadFromDatabase(DataPath, "Setup", "Combo0", "Value", ";", "norecord", "null", True, False)
If IsName(temp1, service) Then
    temp2 = Split(ReadFromDatabase(DataPath, "Services", temp1, "Name;Address;Fields;Keyword", ";", "norecord", "null", False, False), ";")
    If temp2(0) = "norecord" Then AddError MISSING_DATA, "Unable To Run Update.": Exit Function
    If Not temp2(1) Like "*://*" Then AddError INVALID_HTTP, "Unable To Run Update.": Exit Function
End If
temp1 = temp2(1) & IIf(Trim(temp2(2)) <> vbNullString And temp2(2) <> "null", "?" & LCase(temp2(2)), vbNullString)
temp1 = ApplyConstants(temp1)
If animatetray = True Then ChangeTrayIcon pic, IIf(retrynumber = 0, "Updating", "Retrying"), frm.Picture0, frm.ImageList0: ChangeTrayIcon tip, "DNSUpdater: " & IIf(retrynumber = 0, "Updating", "Retrying")
AddMessage Status, IIf(retrynumber = 0, "Updating", "Retrying")
If playsounds = True Then PlaySound DataPath, "Misc", "Updating"
temp1 = OpenHTTP(frm, temp1, Proxy)
If temp1 = "error" Then GoTo clean
temp2 = Split(temp2(3), ",")
For x = 0 To UBound(temp2)
    If InStr(LCase(temp1), LCase(temp2(x))) <> 0 Then
        If animatetray = True Then ChangeTrayIcon pic, "Success", frm.Picture0, frm.ImageList0: ChangeTrayIcon tip, "DNSUpdater: Success"
        AddMessage Status, IIf(retrynumber = 0, "Update", "Retry " & retrynumber) & " Succeeded": RunUpdate = True
        If playsounds = True Then PlaySound DataPath, "Misc", "Success"
        temp1 = ReadFromDatabase(DataPath, "Setup", "Check8", "Value", ";", "norecord", "null", True, False)
        If temp1 = "1" Then Unload MainFrm
        Exit Function
    End If
Next x
clean:
    If animatetray = True Then ChangeTrayIcon pic, "Failure", frm.Picture0, frm.ImageList0: ChangeTrayIcon tip, "DNSUpdater: Failure"
    If Err.number = 0 Then AddMessage Status, IIf(retrynumber = 0, "Update", "Retry " & retrynumber) & " Failed"
    If playsounds = True Then PlaySound DataPath, "Misc", "Failure"
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Run Update."
    Resume clean
End Function
