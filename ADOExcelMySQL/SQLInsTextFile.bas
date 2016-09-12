Attribute VB_Name = "SQLInsTextFile"
Option Explicit

''' are the Comments
'   are the Comments to uncomment depending on your selection

Sub MySQLInsTxtFileCOM() '''Insert Records to the Database via Text File

Dim objCommand As ADODB.Command
Dim sConnect As String
Dim InsSQL As String
Dim MyTxtFile As String
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

MyTxtFile = ThisWorkbook.Path & "\" & "MyTableData" & ".txt"
MyTxtFile = Replace(MyTxtFile, "\", "/")
InsSQL = "LOAD DATA LOCAL INFILE " & Chr(34) & MyTxtFile & Chr(34) & "INTO TABLE `test`.`CrossBorder`"

'''MsgBox InsSQL

objCommand.CommandText = InsSQL
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
MsgBox "MySQL Text File Records Inserted"
ThisWorkbook.Worksheets("Dashboard").Range("E10").Value = totaltime
Else
MsgBox "MySQL Text File Records Not Inserted", vbCritical
ThisWorkbook.Worksheets("Dashboard").Range("E10").Value = ""
End If
End Sub

