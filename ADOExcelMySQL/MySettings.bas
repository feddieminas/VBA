Attribute VB_Name = "MySettings"
Option Explicit

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Worked the project in Windows 7 and Excel 2007, 2010
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'''INSTRUCTIONS

'''1) Add VBA Library reference :
'''Need to add the MDAC Library named Microsoft ActiveX Data Objects 2.8 Library
'''VBA Editor --> Tools --> References --> Microsoft ActiveX Data Objects 2.8 Library
'''Alternatively you can find it VBA Editor --> Tools --> References --> Browse
'''C:\Program Files (x86)\Common Files\System\ado\msado28.tlb

'''2)'Install MySQL ODBC connector client

'''ODBC Data Source Administrator
'''Need to open and check that on tab User DSN tab is registered your Database (for ex. Test) and the name
'''of the SQL Driver.
'''The MySQL ODBC 5.2 ANSI Driver at the ODBC Data Source Administrator is the MySQL ODBC 5.3 ANSI Driver on ADO
'''connection Strings on your vba macros. On this workbook it's on this ways
'''The MySQL ODBC 5.1 ANSI Driver at the ODBC Data Source Administrator is the MySQL ODBC 5.2 ANSI Driver on ADO
'''connection Strings on your vba macros

'''This is an excellent article that explains the interface between MySQL and Excel and how to
'''install it
'''http://www.heritage-tech.net/908/inserting-data-into-mysql-from-excel-using-vba/

'''3)Insert below your global settings

''' are the Comments
'   are the Comments to uncomment depending on your selection

Global ServerName As String
Global DatabaseName As String
Global UserName As String
Global Password As String
Global SQLODBCDriver As String

Sub MyDatabaseLogin() '''Store once Database ServerName, Name, UserName and Password

ServerName = ""
DatabaseName = ""
UserName = ""
Password = ""
SQLODBCDriver = "MySQL ODBC 5.3 ANSI Driver" '''The MySQL ODBC 5.2 Driver at the ODBC Data Source Administrator

End Sub

Sub Testconnection() '''Test your connection String that works. We chose MySQL ODBC 5.2 ANSI Driver
                     '''thus on the connection strings mentioned as per ex. below u see set MySQL ODBC 5.3 ANSI Driver

MyDatabaseLogin

Dim sConnect As String

Dim start As Double, finish As Double, totaltime As Double

start = Timer

On Error GoTo myend

'''ConnectionString to MySQL Database
sConnect = "Driver={" & SQLODBCDriver & "};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"
'sConnect = "Driver={MySQL ODBC 5.3 ANSI Driver};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"

Dim oConn As ADODB.Connection
Set oConn = New ADODB.Connection

'''Connect with the Database
oConn.Open sConnect
'''Close the Connection with the Database
oConn.Close

myend:
'''Destroy the connection with the Database
Set oConn = Nothing

finish = Timer
totaltime = Format(finish - start, "0.00")

'''Insert the number of seconds
If Err.Number <> 0 Then
MsgBox "MySQL Database Connection Test Not Occured", vbCritical
ThisWorkbook.Worksheets("Dashboard").Range("E1").Value = ""
End If

ThisWorkbook.Worksheets("Dashboard").Range("E1").Value = totaltime
MsgBox "MySQL Database Connection Test Occured"
End Sub

