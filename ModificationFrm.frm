VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ModificationFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modification"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComDlg.CommonDialog CommonDialog00 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame0 
      Height          =   1455
      Index           =   3
      Left            =   120
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command01 
         Caption         =   "Cancel"
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   35
         Top             =   1080
         WhatsThisHelpID =   21013
         Width           =   735
      End
      Begin VB.CommandButton Command00 
         Caption         =   "Save"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         WhatsThisHelpID =   21003
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   33
         Top             =   600
         WhatsThisHelpID =   2331
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   32
         Top             =   240
         WhatsThisHelpID =   2330
         Width           =   2775
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   3360
         Picture         =   "ModificationFrm.frx":0000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   240
      End
      Begin VB.Line Line3 
         X1              =   3600
         X2              =   120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Target:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   600
         WhatsThisHelpID =   2331
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   240
         WhatsThisHelpID =   2330
         Width           =   735
      End
   End
   Begin VB.Frame Frame0 
      Height          =   2895
      Index           =   2
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   2880
         TabIndex        =   39
         Top             =   240
         WhatsThisHelpID =   2321
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   29
         Top             =   1080
         WhatsThisHelpID =   2122
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   28
         Top             =   1080
         WhatsThisHelpID =   2121
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         WhatsThisHelpID =   2120
         Width           =   615
      End
      Begin VB.CommandButton Command01 
         Caption         =   "Cancel"
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   24
         Top             =   2520
         WhatsThisHelpID =   21012
         Width           =   735
      End
      Begin VB.CommandButton Command00 
         Caption         =   "Save"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         WhatsThisHelpID =   21002
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   22
         Text            =   "http://"
         Top             =   600
         WhatsThisHelpID =   2322
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   21
         Top             =   240
         WhatsThisHelpID =   2320
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   855
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         WhatsThisHelpID =   2020
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   600
         WhatsThisHelpID =   2322
         Width           =   735
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   3600
         X2              =   120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   3600
         X2              =   120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword:"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         WhatsThisHelpID =   2321
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         WhatsThisHelpID =   2320
         Width           =   735
      End
   End
   Begin VB.Frame Frame0 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   240
         WhatsThisHelpID =   2310
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   16
         Top             =   600
         WhatsThisHelpID =   2311
         Width           =   2775
      End
      Begin VB.CommandButton Command00 
         Caption         =   "Save"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         WhatsThisHelpID =   21001
         Width           =   735
      End
      Begin VB.CommandButton Command01 
         Caption         =   "Cancel"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   1080
         WhatsThisHelpID =   21011
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   240
         WhatsThisHelpID =   2310
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   600
         WhatsThisHelpID =   2311
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   3600
         X2              =   120
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.Frame Frame0 
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command01 
         Caption         =   "Cancel"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   12
         Top             =   2400
         WhatsThisHelpID =   21010
         Width           =   735
      End
      Begin VB.CommandButton Command00 
         Caption         =   "Save"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         WhatsThisHelpID =   21000
         Width           =   735
      End
      Begin VB.TextBox Text0 
         Height          =   285
         Index           =   4
         Left            =   840
         TabIndex        =   10
         Top             =   1920
         WhatsThisHelpID =   2304
         Width           =   2775
      End
      Begin VB.TextBox Text0 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   8
         Text            =   "http://"
         Top             =   1080
         WhatsThisHelpID =   2302
         Width           =   2775
      End
      Begin VB.TextBox Text0 
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   7
         Text            =   "http://"
         Top             =   1560
         WhatsThisHelpID =   2303
         Width           =   2775
      End
      Begin VB.TextBox Text0 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Text            =   "http://"
         Top             =   720
         WhatsThisHelpID =   2301
         Width           =   2775
      End
      Begin VB.TextBox Text0 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   3
         Top             =   240
         WhatsThisHelpID =   2300
         Width           =   2775
      End
      Begin VB.Line Line0 
         Index           =   2
         X1              =   3600
         X2              =   120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line0 
         Index           =   1
         X1              =   3600
         X2              =   120
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line0 
         Index           =   0
         X1              =   3600
         X2              =   120
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         WhatsThisHelpID =   2304
         Width           =   735
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Log Out:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         WhatsThisHelpID =   2302
         Width           =   735
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         WhatsThisHelpID =   2303
         Width           =   735
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Log In:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         WhatsThisHelpID =   2301
         Width           =   735
      End
      Begin VB.Label Label0 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         WhatsThisHelpID =   2300
         Width           =   735
      End
   End
End
Attribute VB_Name = "ModificationFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frm As Form
Dim temp() As String
Dim tempfields As String

Private Sub Command00_Click(Index As Integer)
On Error GoTo err_handler
Dim fields As String
Dim X As Integer
Select Case CInt(temp(3))
    Case 0 To 1
        If Trim(Text0(0)) = vbNullString Then MsgBox "You Must Enter A Name.": Exit Sub
        If Trim(Text0(3)) = vbNullString Then MsgBox "You Must Enter A Status Page.": Exit Sub
        For X = 0 To 4
            If InString(Text0(X), "^ ;") Then MsgBox "Invalid Character Found.": Exit Sub
        Next X
        If Trim(Text0(1)) <> vbNullString And Not Text0(1) Like "*://*" Then MsgBox "You Must Enter A Valid Log In Page.": Exit Sub
        If Trim(Text0(2)) <> vbNullString And Not Text0(2) Like "*://*" Then MsgBox "You Must Enter A Valid Log Out Page.": Exit Sub
        If Not Text0(3) Like "*://*" Then MsgBox "You Must Enter A Valid Status Page.": Exit Sub
        If temp(3) = 0 And IsName(Text0(0), Router) Then MsgBox "Router Name Already In Use.": Exit Sub
        If temp(3) = 1 And LCase(Text0(0)) <> LCase(frm(temp(1))) And IsName(Text0(0), Router) Then MsgBox "Router Name Already In Use.": Exit Sub
        WriteToDatabase DataPath, temp(2), IIf(temp(3) = 1, frm(temp(1)), Text0(0)), "Name;LogIn;LogOut;Status;Keyword", Text0(0) & ";" & Text0(1) & ";" & Text0(2) & ";" & Text0(3) & ";" & Text0(4), ";"
        If temp(3) = 0 Then frm(temp(1)).AddItem Text0(0): frm(temp(1)).Selected(frm(temp(1)).ListCount - 1) = True
        If temp(3) = 1 Then
            If LCase(Text0(0)) <> LCase(frm(temp(1))) Then frm.RouterNameEdit frm(temp(1)), Text0(0)
            frm(temp(1)).list(frm(temp(1)).ListIndex) = Text0(0)
        End If
        frm.LoadRouterOptions: frm.LoadRouter
    Case 2
        If InString(Text1(1), "^ ;") Then MsgBox "Invalid Character Found.": Exit Sub
        If Text1(0) = "Name" And Trim(Text1(1)) = vbNullString Or Text1(0) = "Status" And Trim(Text1(1)) = vbNullString Then MsgBox "You Must Enter A Value.": Exit Sub
        If Text1(0) = "Name" And LCase(frm(temp(5))) <> LCase(Text1(1)) And IsName(Text1(1), Router) Then MsgBox "Router Name Already In Use.": Exit Sub
        If Text1(0) = "Status" And Not Text1(1) Like "*://*" Then MsgBox "You Must Enter A Valid Status Page.": Exit Sub
        If Text1(0) = "Log In" And Trim(Text1(1)) <> vbNullString And Not Text1(1) Like "*://*" Then MsgBox "You Must Enter A Valid Log In Page.": Exit Sub
        If Text1(0) = "Log Out" And Trim(Text1(1)) <> vbNullString And Not Text1(1) Like "*://*" Then MsgBox "You Must Enter A Valid Log Out Page.": Exit Sub
        WriteToDatabase DataPath, temp(2), frm(temp(5)), Replace(Text1(0), " ", ""), Text1(1), ";"
        If LCase(Text1(0)) = "name" Then
            If LCase(frm(temp(5))) <> LCase(Text1(1)) Then frm.RouterNameEdit frm(temp(5)), Text1(1)
            frm(temp(5)).list(frm(temp(5)).ListIndex) = Text1(1)
        End If
        frm.LoadRouterOptions: frm.LoadRouter
    Case 3 To 4
        If Trim(Text2(0)) = vbNullString Then MsgBox "You Must Enter A Name.": Exit Sub
        If Trim(Text2(2)) = vbNullString Then MsgBox "You Must Enter A Address.": Exit Sub
        For X = 0 To 2
            If InString(Text2(X), ";") Then MsgBox "Invalid Character Found.": Exit Sub
        Next X
        If Not Text2(2) Like "*://*" Then MsgBox "You Must Enter A Valid Address.": Exit Sub
        If temp(3) = 3 And IsName(Text2(0), service) Then MsgBox "Service Name Already In Use.": Exit Sub
        If temp(3) = 4 And LCase(Text2(0)) <> LCase(frm(temp(1))) And IsName(Text2(0), service) Then MsgBox "Service Already In Use.": Exit Sub
        WriteToDatabase DataPath, temp(2), IIf(temp(3) = 4, frm(temp(1)), Text2(0)), "Name;Address;Fields;Keyword", Text2(0) & ";" & Text2(2) & ";" & tempfields & ";" & Text2(1), ";"
        If temp(3) = 3 Then frm(temp(1)).AddItem Text2(0): frm(temp(1)).Selected(frm(temp(1)).ListCount - 1) = True
        If temp(3) = 4 Then
            If frm(temp(1)) <> Text2(0) Then frm.ServiceNameEdit frm(temp(1)), Text2(0)
            frm(temp(1)).list(frm(temp(1)).ListIndex) = Text2(0)
        End If
        frm.LoadServiceOptions: frm.LoadService
    Case 5 To 6
        If Trim(Text1(0)) = vbNullString Then MsgBox "You Must Enter A Valid Name.": Exit Sub
        If Trim(Text1(1)) = vbNullString Then MsgBox "You Must Enter A Valid Value.": Exit Sub
        For X = 0 To 1
            If InString(Text1(X), "; & =") Then MsgBox "Invalid Character Found.": Exit Sub
        Next X
        If temp(3) = 5 And IsName(frm(temp(5)), field, Text1(0)) Then MsgBox "Field Already In Use.": Exit Sub
        If temp(3) = 6 Then If LCase(Text1(0)) <> LCase(frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index)) And IsName(frm(temp(5)), field, Text1(0)) Then MsgBox "Field Already In Use.": Exit Sub
        fields = ReadFromDatabase(DataPath, temp(2), frm(temp(5)), "Fields", ";", "norecord", "null", False)
        If fields = "norecord" Then GoTo err_handler
        If fields = "null" Then fields = vbNullString
        If temp(3) = 5 Then fields = IIf(Trim(fields) = vbNullString, vbNullString, fields & "&") & Text1(0) & "=" & Text1(1)
        If temp(3) = 6 Then fields = Replace(fields, frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index) & "=" & frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index).SubItems(1), Text1(0) & "=" & Text1(1))
        WriteToDatabase DataPath, temp(2), frm(temp(5)), "Fields", fields, ";"
        frm.LoadServiceOptions: frm.LoadService
    Case 7 To 8
        If Trim(Text3(0)) = vbNullString Then MsgBox "You Must Enter A Valid Name.": Exit Sub
        If Trim(Text3(1)) = vbNullString Or PathFileExists(ApplyConstants(Text3(1))) = 0 Or LCase(Right(Text3(1), 4)) <> ".exe" Then MsgBox "You Must Enter A Valid Value.": Exit Sub
        For X = 0 To 1
            If InString(Text3(X), ";") Then MsgBox "Invalid Character Found.": Exit Sub
        Next X
        If temp(3) = 7 And IsName(Text3(0), program) Then MsgBox "Program Name Already In Use.": Exit Sub
        If temp(3) = 8 Then If LCase(Text3(0)) <> LCase(frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index)) And IsName(Text3(0), program) Then MsgBox "Program Name Already In Use.": Exit Sub
        If temp(3) = 7 Then WriteToDatabase DataPath, temp(2), "Program" & Text3(0), "Name;Value", "Program" & Text3(0) & ";" & Text3(1), ";"
        If temp(3) = 8 Then WriteToDatabase DataPath, temp(2), "Program" & frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index), "Name;Value", "Program" & Text3(0) & ";" & Text3(1), ";"
        frm.LoadPrograms
    Case 9 To 10
        If Trim(Text1(0)) = vbNullString Then MsgBox "You Must Enter A Valid Field.": Exit Sub
        If Trim(Text1(1)) = vbNullString Then MsgBox "You Must Enter A Valid Value.": Exit Sub
        For X = 0 To 1
            If InString(Text1(X), "; & =") Then MsgBox "Invalid Character Found.": Exit Sub
        Next X
        If temp(3) = 9 Then
            If frm.IsTempField(Text1(0)) Then MsgBox "Field Already In Use.": Exit Sub
            frm.AddField frm(temp(1)), Text1(0), Text1(1)
        Else
            If LCase(Text1(0)) <> LCase(temp(2)) And frm.IsTempField(Text1(0)) Then MsgBox "Field Already In Use.": Exit Sub
            frm.EditField frm(temp(1)), frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index), frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index).ListSubItems(1), Text1(0), Text1(1)
        End If
End Select
clean:
    Unload Me
    Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Make Modifications."
    Resume clean
End Sub

Private Sub Command01_Click(Index As Integer)
On Error Resume Next: Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)
Dim frm As New ModificationFrm
Dim replacestr As String
If ListView2.ListItems.Count = 0 And Index <> 0 Then MsgBox "You Must Select A Field.": Exit Sub
If Index = 1 Then
    If ListView2.SelectedItem.Index <> 1 Then replacestr = "&"
    replacestr = replacestr & ListView2.ListItems(ListView2.SelectedItem.Index) & "=" & ListView2.ListItems(ListView2.SelectedItem.Index).ListSubItems(1)
    If ListView2.SelectedItem.Index = 1 And ListView2.ListItems.Count <> 1 Then replacestr = replacestr & "&"
    tempfields = Replace(tempfields, replacestr, "")
    ListView2.ListItems.Remove ListView2.SelectedItem.Index
Else
    If Index = 0 Then Modification = hwnd & ";ListView2;;9;Add Field"
    If Index = 2 Then Modification = hwnd & ";ListView2;" & ListView2.ListItems(ListView2.SelectedItem.Index) & ";10;Edit Field"
    frm.Show vbModal
End If
FixHeaders ListView2, 2
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Modify Field."
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err_handler
Dim values() As String
Dim tempfield() As String
Dim frame, fheight As Integer
Dim X As Integer
If Trim(Modification) = vbNullString Then GoTo err_handler
temp = Split(Modification, ";"): Modification = vbNullString: tempfields = vbNullString
For Each Form In Forms
    If Form.name = temp(0) Or Form.hwnd = temp(0) Then Set frm = Form
Next Form
If UBound(temp) <> 4 And UBound(temp) <> 5 Or Not IsNumeric(temp(3)) Then GoTo err_handler
Caption = temp(4)
Select Case CInt(temp(3))
    Case 0 To 1
        frame = 0: fheight = 3310
        If temp(3) = 1 Then values = Split(ReadFromDatabase(DataPath, temp(2), frm(temp(1)), "Name;LogIn;LogOut;Status;Keyword", ";", "norecord", "null", False), ";") Else: ReDim values(0)
    Case 2
        frame = 1: fheight = 1995
        Label1(0).Caption = "Property": Label1(1).Caption = "Value": Text1(0).Enabled = False
        values = Split(ReadFromDatabase(DataPath, temp(2), frm(temp(5)), Replace(frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index), " ", ""), ";", "norecord", "null", False), ";")
    Case 3 To 4
        frame = 2: fheight = 3495
        If temp(3) = 4 Then values = Split(ReadFromDatabase(DataPath, temp(2), frm(temp(1)), "Name;Address;Fields;Keyword", ";", "norecord", "null", False), ";") Else: ReDim values(0)
    Case 5 To 6
        frame = 1: fheight = 1995
        Label1(0).Caption = "Field": Label1(1).Caption = "Value"
        If temp(3) = 6 Then values = Split(ReadFromDatabase(DataPath, temp(2), frm(temp(5)), "Fields", ";", "norecord", "null", False), "&") Else: ReDim values(0)
    Case 7 To 8
        frame = 3: fheight = 1995
        If temp(3) = 8 Then values = Split(ReadFromDatabase(DataPath, temp(2), "Program" & frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index), "Name;Value", ";", "norecord", "null", False), ";") Else: ReDim values(0)
    Case 9 To 10
        frame = 1: fheight = 1995
        Label1(0).Caption = "Field": Label1(1).Caption = "Value": ReDim values(0)
    Case Else
        frame = 1: fheight = 1995
        ReDim values(0)
End Select
Frame0(frame).Top = 0: Frame0(frame).Visible = True: Height = fheight
AddHeaders ListView2, "Name;Value", ";"
If values(0) = "norecord" Then GoTo err_handler
If temp(3) = 1 Then
    If UBound(values) <> 4 Then GoTo err_handler
    For X = 0 To UBound(values)
        If X >= 1 And X <= 3 Then
            Text0(X) = IIf(Not values(X) Like "*://*", "http://", values(X))
        Else
            Text0(X) = IIf(values(X) = "null" Or Trim(values(X)) = vbNullString, vbNullString, values(X))
        End If
    Next X
ElseIf temp(3) = 2 Then
    If UBound(values) <> 0 Then GoTo err_handler
    Text1(0) = frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index)
    Select Case LCase(Replace(frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index), " ", ""))
        Case "logout", "login", "status"
            Text1(1) = IIf(Not values(0) Like "*://*", "http://", values(0))
        Case Else
            Text1(1) = IIf(values(0) = "null" Or Trim(values(0)) = vbNullString, vbNullString, values(0))
    End Select
ElseIf temp(3) = 4 Then
    If UBound(values) <> 3 Then GoTo err_handler
    Text2(0) = IIf(values(0) = "null", vbNullString, values(0)): Text2(1) = IIf(values(3) = "null", vbNullString, values(3))
    Text2(2) = IIf(Not values(1) Like "*://*", "http://", values(1))
    ListViewAddString ListView2, values(2), "&", "=": tempfields = values(2)
    FixHeaders ListView2, 2
ElseIf temp(3) = 6 Then
    For X = 0 To UBound(values)
        tempfield = Split(values(X), "=")
        If LCase(tempfield(0)) = LCase(frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index)) Then Text1(0) = tempfield(0): Text1(1) = tempfield(1): Exit For
    Next X
ElseIf temp(3) = 8 Then
    If UBound(values) <> 1 Then GoTo err_handler
    For X = 0 To UBound(values)
        If X = 0 Then
            Text3(X) = Mid(values(X), 8, Len(values(X)) - 7)
        ElseIf X = 1 Then
            Text3(X) = IIf(PathFileExists(values(X)) = 0, vbNullString, values(X))
        End If
    Next X
ElseIf temp(3) = 10 Then
    Text1(0) = frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index)
    Text1(1) = frm(temp(1)).ListItems(frm(temp(1)).SelectedItem.Index).ListSubItems(1)
End If
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Modify Property."
    Unload Me
End Sub

Public Sub AddField(listview As listview, field As String, value As String)
On Error GoTo err_handler
If Trim(field) = vbNullString Or Trim(value) = vbNullString Then GoTo err_handler
tempfields = IIf(Trim(tempfields) = vbNullString, vbNullString, tempfields & "&") & field & "=" & value
ListViewAddString listview, tempfields, "&", "="
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Add Field."
    Unload Me
End Sub

Public Sub EditField(listview As listview, ofield As String, ovalue As String, nfield As String, nvalue As String)
On Error GoTo err_handler
If Trim(ofield) = vbNullString Or Trim(ovalue) = vbNullString Or Trim(nfield) = vbNullString Or Trim(nvalue) = vbNullString Then GoTo err_handler
tempfields = Replace(tempfields, ofield & "=" & ovalue, nfield & "=" & nvalue)
ListViewAddString listview, tempfields, "&", "="
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Edit Field."
    Unload Me
End Sub

Public Function IsTempField(name As String) As Boolean
On Error GoTo err_handler
Dim temp() As String
Dim X As Integer
If Trim(tempfields) = vbNullString Then Exit Function
temp = Split(tempfields, "&")
For X = 0 To UBound(temp)
    If LCase(Left(temp(X), Len(name) + 1)) = LCase(name) & "=" Then IsTempField = True: Exit For
Next X
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Failed Checking If In Temp Field."
    Unload Me
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next: tempfields = vbNullString
End Sub

Private Sub Image3_Click()
On Error GoTo err_handler
CommonDialog00.DialogTitle = "Browse Program": CommonDialog00.Filter = "Executable (*.exe)|*.exe"
CommonDialog00.InitDir = "C:\": CommonDialog00.ShowOpen
If Trim(CommonDialog00.filename) = vbNullString Then Exit Sub
Text3(1) = CommonDialog00.filename: CommonDialog00.filename = vbNullString
Exit Sub
err_handler:
    AddError UNKNOWN_ERROR, "Failed To Browse."
    Unload Me
End Sub

Private Sub ListView2_DblClick()
On Error Resume Next: Command2_Click 2
End Sub
