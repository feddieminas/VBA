Attribute VB_Name = "SQLTable"
Option Explicit

''' are the Comments
'   are the Comments to uncomment depending on your selection

Sub MySQLTableQueryCOM() '''Create a MySQL Table and Insert it on the Database

Dim objCommand As ADODB.Command
Dim sConnect As String
Dim TableSQL As String
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

'On Error Resume Next
objCommand.CommandText = "DROP TABLE IF EXISTS `crossborder`"
'On Error GoTo 0

'On Error Resume Next
objCommand.Execute
'On Error GoTo 0

'''Below on IDIndex and Hour column, one can also declare them as int(11)
'''`DDate` can be set as number or as date. If number you have faster queries

'''1)DDate declared as Number to MySQL Table

TableSQL = _
        "CREATE TABLE IF NOT EXISTS `CrossBorder` (" & _
            "`IDIndex` tinyint unsigned DEFAULT NULL," & _
            "`DDate` int DEFAULT NULL," & _
            "`Hour` smallint unsigned DEFAULT NULL," & _
            "`Border` char(4) DEFAULT NULL," & _
            "`Purpose` varchar(4) DEFAULT NULL," & _
            "`Qty` decimal(5,2) DEFAULT NULL," & _
            "`Price` float(7,2) DEFAULT NULL)" & _
            " ENGINE=InnoDB DEFAULT CHARSET=utf8;"

'''MsgBox TableSQL

'''OR

'''2)DDate declared as Date to MySQL Table

'TableSQL = _
        "CREATE TABLE IF NOT EXISTS `CrossBorder` (" & _
            "`IDIndex` tinyint unsigned DEFAULT NULL," & _
            "`DDate` date DEFAULT NULL," & _
            "`Hour` smallint unsigned DEFAULT NULL," & _
            "`Border` char(4) DEFAULT NULL," & _
            "`Purpose` varchar(4) DEFAULT NULL," & _
            "`Qty` decimal(5,2) DEFAULT NULL," & _
            "`Price` float(7,2) DEFAULT NULL)" & _
            " ENGINE=InnoDB DEFAULT CHARSET=utf8;"

'''MsgBox TableSQL

objCommand.CommandText = TableSQL
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
MsgBox "MySQL Table Created"
ThisWorkbook.Worksheets("Dashboard").Range("E4").Value = totaltime
Else
MsgBox "MySQL Table Not Created", vbCritical
ThisWorkbook.Worksheets("Dashboard").Range("E4").Value = ""
End If
End Sub
