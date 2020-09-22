Attribute VB_Name = "DataMod"
Public Enum ntype
    service = 1
    Router = 2
    program = 3
    field = 4
End Enum

Public Const MaxDataErrors = 50

Public DataErrors As Integer
Public DataPath As String

Public Function LoadSettings(frm As Form, path As String, ltabs As ltabs, table As String)
On Error GoTo err_handler
Dim db As Database
Dim rs As Recordset
Dim Index As Integer
Dim value As String
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Load Settings.": GoTo clean
If Trim(table) = vbNullString Then AddError MISSING_DATA, "Unable To Load Settings.": GoTo clean
Set db = DBEngine.OpenDatabase(path)
Set rs = db.OpenRecordset(table)
Do Until rs.EOF
    Index = IIf(IsNumeric(Right(rs!name, 2)), Right(rs!name, 2), IIf(IsNumeric(Right(rs!name, 1)), Right(rs!name, 1), -1))
    If LCase(Left(rs!name, 5)) = "check" Then
        value = IIf(rs!value = 1, 1, 0)
        If Index = -1 Then frm("Check" & ltabs).value = value Else: If Index <= frm("Check" & ltabs).Count - 1 Then frm("Check" & ltabs)(Index).value = value
    ElseIf LCase(Left(rs!name, 4)) = "text" Then
        If ltabs = Setup Then
            Select Case Index
                Case 0
                    value = IIf(Not IsNumeric(rs!value) Or rs!value < 1 Or rs!value > 60, 1, rs!value)
                Case 1
                    value = IIf(Not IsNumeric(rs!value) Or rs!value < 30 Or rs!value > 120, 30, rs!value)
                Case 2 To 3
                    value = IIf(Not IsNumeric(rs!value) Or rs!value < 1 Or rs!value > 120, 1, rs!value)
                Case Else
                    value = rs!value
            End Select
        ElseIf ltabs = LANConnect Then
            Select Case Index
                Case 1
                    value = IIf(ValidIP(rs!value) = False, "0.0.0.0", rs!value)
                Case 2
                    value = IIf(Not IsNumeric(rs!value) Or rs!value <= 0, 80, rs!value)
                Case 3
                    value = IIf(Not rs!value Like "*://*", "http://", rs!value)
                Case 5, 9
                    value = IIf(ValidIP(rs!value) = False And Not rs!value Like "*.*", vbNullString, rs!value)
                Case 6
                    value = IIf(Not IsNumeric(rs!value) Or rs!value < 1, 6667, rs!value)
                Case Else
                    value = rs!value
            End Select
        Else
            value = rs!value
        End If
        If Index = -1 Then frm("Text" & ltabs) = value Else: If Index <= frm("Text" & ltabs).Count - 1 And Index >= 0 Then frm("Text" & ltabs)(Index) = value
    ElseIf LCase(Left(rs!name, 6)) = "option" Then
        If Not IsNumeric(rs!value) Then value = 0 Else: value = IIf(frm("Option" & ltabs).Count - 1 < CInt(rs!value) Or CInt(rs!value) < 0, 0, rs!value)
        frm("Option" & ltabs)(value).value = 1
    ElseIf LCase(Left(rs!name, 5)) = "combo" Then
        If ltabs = Setup Then
            If Index = 0 Then
                value = IIf(IsName(rs!value, service), rs!value, vbNullString)
            ElseIf Index = 1 Then
                value = IIf(rs!value Like "*.*.*.*", rs!value, vbNullString)
            End If
        ElseIf ltabs = LANConnect Then
            value = IIf(IsName(rs!value, Router), rs!value, vbNullString)
        Else
            value = rs!value
        End If
        If Index = -1 Then frm("Combo" & ltabs) = value Else: If Index <= frm("Combo" & ltabs).Count - 1 Then frm("Combo" & ltabs)(Index) = value
    End If
    rs.MoveNext
Loop
clean:
    Set rs = Nothing
    Set db = Nothing
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Load Settings."
    Resume clean
End Function

Public Function IsName(name As String, ntype As ntype, Optional field As String) As Boolean
On Error GoTo err_handler
Dim temp() As String
Dim x As Integer
If Trim(name) = vbNullString Then AddError MISSING_DATA, "Unable To Determine If Is Name.": Exit Function
Select Case ntype
    Case Router
        temp = Split(GetTableNames(DataPath, "Routers", ";", "norecord"), ";")
    Case service
        temp = Split(GetTableNames(DataPath, "Services", ";", "norecord"), ";")
    Case program
        temp = Split(GetTableNames(DataPath, "Misc", ";", "norecord"), ";")
    Case field
        temp = Split(ReadFromDatabase(DataPath, "Services", name, "Fields", ";", "norecord", "null", False), "&")
    Case Else
        AddError INVALID_TYPE, "Unable To Determine If Is Name.": Exit Function
End Select
If temp(0) = "norecords" Or temp(0) = "null" Or temp(0) = "norecord" Then Exit Function
For x = 0 To UBound(temp)
    Select Case ntype
        Case 1 To 2
            If LCase(temp(x)) = LCase(name) Then IsName = True: Exit For
        Case 3
            If LCase(temp(x)) = "program" & LCase(name) Then IsName = True: Exit For
        Case 4
            If Left(LCase(temp(x)), Len(field) + 1) = LCase(field) & "=" Then IsName = True: Exit For
    End Select
Next x
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Determine If Is Name."
End Function

Public Function WriteToDatabase(path As String, table As String, name As String, fields As String, values As String, Optional delimiter As String, Optional exitonerror As Boolean)
On Error GoTo err_handler
Dim db As Database
Dim rs As Recordset
Dim temp1() As String
Dim temp2() As String
Dim x As Integer
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Write Record.": GoTo clean
If Trim(table) = vbNullString Or Trim(name) = vbNullString Or Trim(fields) = vbNullString Then AddError MISSING_DATA, "Unable To Write Record.": GoTo clean
Set db = DBEngine.OpenDatabase(path)
Set rs = db.OpenRecordset("SELECT * FROM " & table & " WHERE Name = '" & name & "'")
temp1 = Split(fields, IIf(Trim(delimiter) = vbNullString, ";", delimiter))
temp2 = Split(values, IIf(Trim(delimiter) = vbNullString, ";", delimiter))
If Trim(values) = vbNullString And UBound(temp1) = 0 Then ReDim temp2(0)
If UBound(temp1) <> UBound(temp2) Then AddError MISSING_DATA, "Fields And Values Don't Match.": GoTo clean
If rs.RecordCount = 0 Then
    rs.AddNew
    rs!name = name
Else
    If rs.RecordCount > 1 Then
        Do Until rs.RecordCount = 1
            rs.MoveLast
            rs.Delete
        Loop
    End If
    rs.Edit
End If
For x = 0 To UBound(temp1)
    rs(temp1(x)) = IIf(Trim(temp2(x)) = vbNullString, Null, temp2(x))
Next x
rs.Update
rs.Close
clean:
    Set rs = Nothing
    Set db = Nothing
    Exit Function
err_handler:
    If exitonerror = False Then AddError UNKNOWN_ERROR, "Unable To Write Record."
    DataError
    Resume clean
End Function

Public Function ReadFromDatabase(path As String, table As String, name As String, fields As String, Optional delimiter As String, Optional norecord As String, Optional nullvalue As String, Optional disablenorecord As Boolean, Optional exitonerror As Boolean) As String
On Error GoTo err_handler
Dim db As Database
Dim rs As Recordset
Dim temp() As String
Dim x As Integer
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Read Record.": GoTo clean
If Trim(table) = vbNullString Or Trim(name) = vbNullString Or Trim(fields) = vbNullString Then AddError MISSING_DATA, "Unable To Read Record.": GoTo clean
Set db = DBEngine.OpenDatabase(path)
Set rs = db.OpenRecordset("SELECT * FROM " & table & " WHERE Name = '" & name & "'")
temp = Split(fields, IIf(Trim(delimiter) = vbNullString, ";", delimiter))
If rs.RecordCount <= 0 Then
    If disablenorecord = False Then AddError MISSING_DATA, "Record Does Not Exist."
    ReadFromDatabase = IIf(Trim(norecord) = vbNullString, "norecord", norecord): GoTo clean
Else
    Do Until rs.RecordCount = 1
        rs.MoveLast
        rs.Delete
    Loop
    For x = 0 To UBound(temp)
        If Trim(ReadFromDatabase) = vbNullString Then
            ReadFromDatabase = IIf(IsNull(rs(temp(x))), IIf(Trim(nullvalue) = vbNullString, "null", nullvalue), rs(temp(x)))
        Else
            ReadFromDatabase = ReadFromDatabase & IIf(Trim(delimiter) = vbNullString, ";", delimiter) & IIf(IsNull(rs(temp(x))), IIf(Trim(nullvalue) = vbNullString, "null", nullvalue), rs(temp(x)))
        End If
    Next x
End If
clean:
    Set rs = Nothing
    Set db = Nothing
    Exit Function
err_handler:
    If exitonerror = False Then AddError UNKNOWN_ERROR, "Unable To Read Record."
    DataError
    Resume clean
End Function

Public Function DeleteFromDatabase(path As String, table As String, name As String)
On Error GoTo err_handler
Dim db As Database
Dim rs As Recordset
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Delete Record.": GoTo clean
If Trim(table) = vbNullString Or Trim(name) = vbNullString Then AddError MISSING_DATA, "Unable To Delete Record.": GoTo clean
Set db = DBEngine.OpenDatabase(path)
Set rs = db.OpenRecordset("SELECT * FROM " & table & " WHERE Name = '" & name & "'")
If rs.RecordCount = 0 Then AddError MISSING_DATA, "Record Does Not Exist.": GoTo clean
Do Until rs.EOF
    rs.Delete
    rs.MoveNext
Loop
clean:
    Set rs = Nothing
    Set db = Nothing
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Delete Record."
    Resume clean
End Function

Public Function GetTableNames(path As String, table As String, Optional delimiter As String, Optional norecords As String) As String
On Error GoTo err_handler
Dim db As Database
Dim rs As Recordset
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Get Table Names.": GoTo clean
If Trim(table) = vbNullString Then AddError MISSING_DATA, "Unable To Get Table Names.": GoTo clean
Set db = DBEngine.OpenDatabase(path)
Set rs = db.OpenRecordset(table)
Do Until rs.EOF
    GetTableNames = IIf(Trim(GetTableNames) = vbNullString, rs!name, GetTableNames & IIf(Trim(delimiter) = vbNullString, ";", delimiter) & rs!name)
    rs.MoveNext
Loop
If Trim(GetTableNames) = vbNullString Then GetTableNames = IIf(Trim(norecords) = vbNullString, "norecords", norecords)
clean:
    Set rs = Nothing
    Set db = Nothing
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Get Table Names."
    Resume clean
End Function

Public Function CreateDatabase(path As String)
On Error GoTo err_handler
Dim db As Database
Dim temp As String
If PathFileExists(path) Then
    AddError INVALID_DATABASE, "Database Already Exists."
    If MsgBox("Specified Database Already Exists. Would You Like To Delete It?", vbYesNo, "Delete") = vbYes Then
        Kill path
    Else
        GoTo clean
    End If
End If
If Not path Like "*.mdb" Then AddError INVALID_DATABASE, "Unable To Create Database.": GoTo clean
Set db = DBEngine.CreateDatabase(path, dbLangGeneral, dbVersion30)
clean:
    Set db = Nothing
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Create Database."
End Function

Private Function CreateTable(path As String, name As String, fields As String, Optional delimiter1 As String, Optional delimiter2 As String)
On Error GoTo err_handler
Dim db As Database
Dim tdf As TableDef
Dim temp1() As String
Dim temp2() As String
Dim x As Integer
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Create Table.": GoTo clean
If Trim(name) = vbNullString Or Trim(fields) = vbNullString Then AddError MISSING_DATA, "Unable To Create Table.": GoTo clean
Set db = DBEngine.OpenDatabase(path)
temp1 = Split(fields, IIf(dilimiter1 = vbNullString, ";", dilimiter1))
Set tdf = db.CreateTableDef(name)
For x = 0 To UBound(temp1)
    If Trim(temp1(x)) <> vbNullString Then
        temp2 = Split(temp1(x), IIf(delimiter2 = vbNullString, ",", delimiter2))
        If UBound(temp2) = 1 And Trim(temp2(1)) <> vbNullString Then tdf.fields.Append tdf.CreateField(temp2(0), temp2(1))
    End If
Next x
db.TableDefs.Append tdf
clean:
    Set db = Nothing
    Set tdf = Nothing
    Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Create Table."
End Function

Public Function CreateDefaultTables(path As String)
On Error GoTo err_handler
If PathFileExists(path) = 0 Then AddError DATABASE_NOT_FOUND, "Unable To Create Default Tables.": Exit Function
CreateTable path, "Setup", "Name," & dbText & ";Value," & dbText
CreateTable path, "LANConnect", "Name," & dbText & ";Value," & dbText
CreateTable path, "Routers", "Name," & dbText & ";LogIn," & dbMemo & ";LogOut," & dbMemo & ";Status," & dbMemo & ";Keyword," & dbText
CreateTable path, "Services", "Name," & dbText & ";Address," & dbMemo & ";Fields," & dbMemo & ";Keyword," & dbText
CreateTable path, "Misc", "Name," & dbText & ";Value," & dbMemo
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Unable To Create Default Tables."
End Function

Public Function DataError()
On Error GoTo err_handler
Dim x As Integer
If DataErrors = -1 Then Exit Function
DataErrors = DataErrors + 1
If DataErrors >= MaxDataErrors Then
    If MsgBox("It Appears That Your Database Has Become Corrupted And Unusable. Would You Like To Delete The Database And Start Over?", vbYesNo, "Delete?") = vbYes Then
        If PathFileExists(App.path & "\Corrupt.tmp") = 0 Then
            x = FreeFile
            Open App.path & "\Corrupt.tmp" For Append As x
                Print #x, "Do Not Delete Unless You Want To Cancel The Deletion Of Your Database"
            Close #1
        End If
        MsgBox "Please Exit Out Now To Delete Your Database."
        DataErrors = -1
    Else
        If PathFileExists(App.path & "\Corrupt.tmp") Then Kill App.path & "\Corrupt.tmp"
    End If
    DataErrors = 0
Else
    DataErrors = 0
End If
Exit Function
err_handler:
    AddError UNKNOWN_ERROR, "Writing Corrupt.tmp."
End Function
