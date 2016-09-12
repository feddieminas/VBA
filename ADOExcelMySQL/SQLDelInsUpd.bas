Attribute VB_Name = "SQLDelInsUpd"
Option Explicit

''' are the Comments
'   are the Comments to uncomment depending on your selection

Sub MySQLDelQueryCOM() '''Delete a Record on the Database filtering a Date and an Hour

Dim objCommand As ADODB.Command
Dim sConnect As String
Dim DelSQL As String
Dim start As Double, finish As Double, totaltime As Double

start = Timer

MyDatabaseLogin

On Error GoTo myend

'''ConnectionString to MySQL Database
sConnect = "Driver={" & SQLODBCDriver & "};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"
'sConnect = "Driver={MySQL ODBC 5.3 ANSI Driver};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"

Set objCommand = New ADODB.Command
objCommand.ActiveConnection = sConnect

Dim idate As String
idate = Format(Now, "YYYYMMDD")

'''1)
'''DDate declared as Number to MySQL Table

DelSQL = "DELETE " & _
"FROM `test`.`CrossBorder` WHERE DDate=" & CLng(idate) & " AND Hour=" & CInt(25)

'''MsgBox DelSQL

'''OR

'''2)
'''DDate declared as Date to MySQL Table

'DelSQL = "DELETE " & _
"FROM `test`.`CrossBorder` WHERE DDate=DATE_FORMAT('" & (Left(idate, 4) & "/" & Mid(idate, 5, 2) & "/" & Right(idate, 2)) & "', '%Y-%m-%d')" & _
" AND Hour=" & CInt(25)

'''MsgBox DelSQL

objCommand.CommandText = DelSQL
'On Error Resume Next
objCommand.Execute
'On Error GoTo 0

myend:

'''Kill Command Object
If Not objCommand.ActiveConnection Is Nothing Then
If CBool(objCommand.ActiveConnection.State And adStateOpen) Then objCommand.ActiveConnection.Close
End If

Set objCommand = Nothing

finish = Timer
totaltime = Format(finish - start, "0.00")

'''Insert the number of seconds
If Err.Number = 0 Then
MsgBox "MySQL Records Deleted"
ThisWorkbook.Worksheets("Dashboard").Range("E15").Value = totaltime
Else
MsgBox "MySQL Records Not Deleted", vbCritical
ThisWorkbook.Worksheets("Dashboard").Range("E15").Value = ""
End If
End Sub


Sub MySQLInsQueryCOM() '''Insert a Record on the Database

Dim objCommand As ADODB.Command
Dim sConnect As String
Dim InsSQL As String
Dim start As Double, finish As Double, totaltime As Double

start = Timer

MyDatabaseLogin

On Error GoTo myend

'''ConnectionString to MySQL Database
sConnect = "Driver={" & SQLODBCDriver & "};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"
'sConnect = "Driver={MySQL ODBC 5.3 ANSI Driver};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"

Set objCommand = New ADODB.Command
objCommand.ActiveConnection = sConnect

Dim arrInsVal(0, 6) As Variant
arrInsVal(0, 0) = CInt(1)

'''1)
'''DDate declared as Number to MySQL Table

arrInsVal(0, 1) = CLng(Format(Now, "YYYYMMDD"))

'''2)
'''DDate declared as Date to MySQL Table

'arrInsVal(0, 1) = CLng(Format(Now, "YYYYMMDD"))         '''Date
'arrInsVal(0, 1) = CVar(Left(arrInsVal(0, 1), 4) & "-" & Mid(arrInsVal(0, 1), 5, 2) & "-" & Right(arrInsVal(0, 1), 2))


arrInsVal(0, 2) = CInt(25)
arrInsVal(0, 3) = "DECH"
arrInsVal(0, 4) = "SELL"

Dim arrQtyPri() As Variant
arrQtyPri = Array(10, 30.45)

If Application.DecimalSeparator = "," And Int(arrQtyPri(0)) <> arrQtyPri(0) Then
arrInsVal(0, 5) = Replace(Format(arrQtyPri(0), "##0.00"), ",", ".")
Else
arrInsVal(0, 5) = arrQtyPri(0)
End If

If Application.DecimalSeparator = "," And Int(arrQtyPri(1)) <> arrQtyPri(1) Then
arrInsVal(0, 6) = Replace(Format(arrQtyPri(1), "##0.00"), ",", ".")
Else
arrInsVal(0, 6) = arrQtyPri(1)
End If


'''***********************************************************************
'''Check whether a Record Exists on the Database to avoid duplicates

'''1)
'''DDate declared as Number to MySQL Table

objCommand.CommandText = "SELECT * " & _
"FROM `test`.`CrossBorder` " & _
"WHERE DDate=" & arrInsVal(0, 1) & " AND Hour = " & arrInsVal(0, 2)

'''2)
'''DDate declared as Date to MySQL Table

'objCommand.CommandText = "SELECT * " & _
"FROM `test`.`CrossBorder` " & _
"WHERE DDate='" & arrInsVal(0, 1) & "' AND Hour = " & arrInsVal(0, 2)

Dim Rs1 As ADODB.Recordset
Set Rs1 = objCommand.Execute()
Dim IDRet As Boolean
IDRet = Rs1.EOF  '''End of File EOF is False if Record exists, True if Record not exists
Set Rs1 = Nothing
'***********************************************************************

If IDRet = True Then  '''If Not Exists then insert on the database

'''1)
'''DDate declared as Number to MySQL Table

InsSQL = "INSERT " & _
"INTO `test`.`CrossBorder` " & "(IDIndex,DDate,Hour,Border,Purpose,Qty,Price) " & _
"VALUES (" & arrInsVal(0, 0) & "," & arrInsVal(0, 1) & "," & arrInsVal(0, 2) & ",'" & arrInsVal(0, 3) & "','" _
& arrInsVal(0, 4) & "'," & arrInsVal(0, 5) & "," & arrInsVal(0, 6) & ")"

'''2)
'''DDate declared as Date to MySQL Table

'InsSQL = "INSERT " & _
"INTO `test`.`CrossBorder` " & "(IDIndex,DDate,Hour,Border,Purpose,Qty,Price) " & _
"VALUES (" & arrInsVal(0, 0) & ",'" & arrInsVal(0, 1) & "'," & arrInsVal(0, 2) & ",'" & arrInsVal(0, 3) & "','" _
& arrInsVal(0, 4) & "'," & arrInsVal(0, 5) & "," & arrInsVal(0, 6) & ")"

'MsgBox InsSQL

objCommand.CommandText = InsSQL
'On Error Resume Next
objCommand.Execute
'On Error GoTo 0

End If

myend:
Erase arrQtyPri(): Erase arrInsVal()

'''Kill Command Object
If Not objCommand.ActiveConnection Is Nothing Then
If CBool(objCommand.ActiveConnection.State And adStateOpen) Then objCommand.ActiveConnection.Close
End If

Set objCommand = Nothing

finish = Timer
totaltime = Format(finish - start, "0.00")

'''Insert the number of seconds
If Err.Number = 0 And IDRet = False Then
MsgBox "MySQL Records Existed", vbInformation
ThisWorkbook.Worksheets("Dashboard").Range("E18").Value = ""
ElseIf Err.Number = 0 And IDRet = True Then
MsgBox "MySQL Records Inserted"
ThisWorkbook.Worksheets("Dashboard").Range("E18").Value = totaltime
Else
MsgBox "MySQL Records Not Inserted", vbCritical
ThisWorkbook.Worksheets("Dashboard").Range("E18").Value = ""
End If
End Sub


Sub MySQLUpdQueryCOM() '''Update a Record on the Database

Dim objCommand As ADODB.Command
Dim sConnect As String
Dim UpdSQL As String
Dim start As Double, finish As Double, totaltime As Double

start = Timer

MyDatabaseLogin

On Error GoTo myend

'''ConnectionString to MySQL Database
sConnect = "Driver={" & SQLODBCDriver & "};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"
'sConnect = "Driver={MySQL ODBC 5.3 ANSI Driver};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"

Set objCommand = New ADODB.Command
objCommand.ActiveConnection = sConnect

Dim arrUpdVal(0, 6) As Variant
arrUpdVal(0, 0) = 1

'''1)
'''DDate declared as Number to MySQL Table

arrUpdVal(0, 1) = CLng(Format(Now, "YYYYMMDD"))

'''2)
'''DDate declared as Date to MySQL Table

'arrUpdVal(0, 1) = CLng(Format(Now, "YYYYMMDD"))         '''Date
'arrUpdVal(0, 1) = CVar(Left(arrUpdVal(0, 1), 4) & "-" & Mid(arrUpdVal(0, 1), 5, 2) & "-" & Right(arrUpdVal(0, 1), 2))

arrUpdVal(0, 2) = CInt(25)
arrUpdVal(0, 3) = "DECH"
arrUpdVal(0, 4) = "SELL"

Dim arrQtyPri() As Variant
arrQtyPri = Array(0, 0)

If Application.DecimalSeparator = "," And Int(arrQtyPri(0)) <> arrQtyPri(0) Then
arrUpdVal(0, 5) = Replace(Format(arrQtyPri(0), "##0.00"), ",", ".")
Else
arrUpdVal(0, 5) = arrQtyPri(0)
End If

If Application.DecimalSeparator = "," And Int(arrQtyPri(1)) <> arrQtyPri(1) Then
arrUpdVal(0, 6) = Replace(Format(arrQtyPri(1), "##0.00"), ",", ".")
Else
arrUpdVal(0, 6) = arrQtyPri(1)
End If


'''***********************************************************************
'''Check whether a Record Exists on the Database so you can update it

'''1)
'''DDate declared as Number to MySQL Table

objCommand.CommandText = "SELECT * " & _
"FROM `test`.`CrossBorder` " & _
"WHERE DDate=" & arrUpdVal(0, 1) & " AND Hour = " & arrUpdVal(0, 2)

'''2)
'''DDate declared as Date to MySQL Table

'objCommand.CommandText = "SELECT * " & _
"FROM `test`.`CrossBorder` " & _
"WHERE DDate='" & arrUpdVal(0, 1) & "' AND Hour = " & arrUpdVal(0, 2)

Dim Rs1 As ADODB.Recordset
Set Rs1 = objCommand.Execute()
Dim IDRet As Boolean
IDRet = Rs1.EOF   '''End of File EOF is False if Record exists, True if Record not exists
Set Rs1 = Nothing
'''***********************************************************************

If IDRet = False Then  '''If Exists then update on the database

'''1)
'''DDate declared as Number to MySQL Table

UpdSQL = "UPDATE " & _
"`test`.`CrossBorder` SET " & "IDIndex=" & arrUpdVal(0, 0) & ",DDate=" & arrUpdVal(0, 1) & _
",Hour=" & arrUpdVal(0, 2) & ",Border='" & arrUpdVal(0, 3) & _
"',Purpose='" & arrUpdVal(0, 4) & "',Qty=" & arrUpdVal(0, 5) & ",Price=" & arrUpdVal(0, 6) & _
" WHERE DDate=" & arrUpdVal(0, 1) & " AND Hour = " & arrUpdVal(0, 2)

'''2)
'''DDate declared as Date to MySQL Table

'UpdSQL = "UPDATE " & _
"`test`.`CrossBorder` SET " & "IDIndex=" & arrUpdVal(0, 0) & ",DDate='" & arrUpdVal(0, 1) & _
"',Hour=" & arrUpdVal(0, 2) & ",Border='" & arrUpdVal(0, 3) & _
"',Purpose='" & arrUpdVal(0, 4) & "',Qty=" & arrUpdVal(0, 5) & ",Price=" & arrUpdVal(0, 6) & _
" WHERE DDate='" & arrUpdVal(0, 1) & "' AND Hour = " & arrUpdVal(0, 2)

'''MsgBox UpdSQL

objCommand.CommandText = UpdSQL
'On Error Resume Next
objCommand.Execute
'On Error GoTo 0

End If

myend:
Erase arrQtyPri(): Erase arrUpdVal()

'''Kill Command Object
If Not objCommand.ActiveConnection Is Nothing Then
If CBool(objCommand.ActiveConnection.State And adStateOpen) Then objCommand.ActiveConnection.Close
End If

Set objCommand = Nothing

finish = Timer
totaltime = Format(finish - start, "0.00")

'''Insert the number of seconds
If Err.Number = 0 And IDRet = True Then
MsgBox "MySQL Records Not Existed", vbInformation
ThisWorkbook.Worksheets("Dashboard").Range("E21").Value = ""
ElseIf Err.Number = 0 And IDRet = False Then
MsgBox "MySQL Records Updated"
ThisWorkbook.Worksheets("Dashboard").Range("E21").Value = totaltime
Else
MsgBox "MySQL Records Not Updated", vbCritical
ThisWorkbook.Worksheets("Dashboard").Range("E21").Value = ""
End If
End Sub











































