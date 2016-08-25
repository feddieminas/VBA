Attribute VB_Name = "GH_ADOExcel_NewWBWS"
Option Explicit
Option Base 1

'Tried in Windows 7 and Excel 2007, 2010

'******************************************************************
'Example of an Excel Database Creation with Closed Workbooks
'******************************************************************

'1)Need to check in your ODBC Data Source Administrator
'on User DSN tab that you have :
'User Data Sources Name = Excel Files
'User Data Sources Driver = Microsoft Excel Driver(*xls,*xlsx,*xlsm,*xlsb)
'2)Add VBA Library reference :
'Need to add the MDAC Library named Microsoft ActiveX Data Objects 2.8 Library
'VBA Editor --> Tools --> References --> Microsoft ActiveX Data Objects 2.8 Library
'Alternatively you can find it VBA Editor --> Tools --> References --> Browse
'C:\Program Files (x86)\Common Files\System\ado\msado28.tlb

'Instructions :
'Below there is an example i have created. To get an idea and see that it works for you, you can adjust the macro Settings
'in filepaths, folder and file template name of your preference. If you would not insert them, by default it will create
'the variables with its names of my imagination
'The example further for creating excel files with numerical random values, I demonstrate an example with regards to the sector i work, the energy sector
'Run CreateFileTemplate and ADOProc macro in order

Dim MyWorkingFilePath As String
Dim DestFolder As String
Dim DestFolderFileTemplate As String

Sub MySettings()
'Example :
'MyWorkingFilePath ="C:\Users\faidon.dermesonoglou\Desktop"
'DestFolder = "CrossBorder"
'DestFolderFileTemplate = "DECHTemplate.xls"
'It will create a folder inside your WorkingFilePath to be like "C:\Users\faidon.dermesonoglou\Desktop\CrossBorder"
'Inside this new path it will create your template file needed as an object to create sub files for your Excel Database

MyWorkingFilePath = ""                                                          'INSERT or keep it your Filepath you want to create your Folder for your Database
If MyWorkingFilePath = "" Then MyWorkingFilePath = ThisWorkbook.Path

DestFolder = ""                                                                 'INSERT or keep it the Name of the Folder to Create your Database
If DestFolder = "" Then DestFolder = "CrossBorder"

DestFolderFileTemplate = ""                                                     'INSERT or keep it Recommendeded to be an xls file type
If DestFolderFileTemplate = "" Then DestFolderFileTemplate = "DECHTemplate.xls"

End Sub

Sub CreateFileTemplate() 'Need to Run only once to create your Template file

MySettings

If FolderFileExists(MyWorkingFilePath & SChar(MyWorkingFilePath) & DestFolder _
& SChar(DestFolder) & DestFolderFileTemplate) = True Then Exit Sub

Application.ScreenUpdating = False
Application.DisplayAlerts = False

On Error Resume Next
MkDir MyWorkingFilePath & SChar(MyWorkingFilePath) & "CrossBorder"
On Error GoTo 0

Dim WB As Workbook
Set WB = Workbooks.Add

With WB

With .Worksheets(1)  'DECH
.Name = "DECH"
.Range("A1") = "Hour"
.Range("B1:F1").Value = Array("RESDECHY", "RESDECHM", "NOMDECHY", "NOMDECHM", "NOMDECHD")
.Range("A2:A26").Value = _
Application.Transpose(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25))
'.Range("B2:F26").NumberFormat = "0"  'Specify Number Format to zero decimal places
.Range("B2:F26").NumberFormat = "0.0"  'Specify Number Format to one decimal place
'.Range("B2:F26").NumberFormat = "0.00"  'Specify Number Format to two decimal places
.Range("A:F").EntireColumn.AutoFit
End With

Dim WS As Worksheet  'Count Wsts if needed a second one
If .Worksheets.Count = 1 Then
.Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = "Sheet2"
End If

With .Worksheets(2)  'CHDE
.Name = "CHDE"
.Range("A1") = "Hour"
.Range("B1:F1").Value = Array("RESCHDEY", "RESCHDEM", "NOMCHDEY", "NOMCHDEM", "NOMCHDED")
.Range("A2:A26").Value = _
Application.Transpose(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25))
'.Range("B2:F26").NumberFormat = "0"  'Specify Number Format to zero decimal places
.Range("B2:F26").NumberFormat = "0.0"  'Specify Number Format to one decimal place
'.Range("B2:F26").NumberFormat = "0.00"  'Specify Number Format to two decimal places
.Range("A:F").EntireColumn.AutoFit
End With

.SaveAs MyWorkingFilePath & SChar(MyWorkingFilePath) & DestFolder _
& SChar(DestFolder) & DestFolderFileTemplate, FileFormat:=56
.Close savechanges:=False

End With

Set WB = Nothing

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub ADOProc()

MySettings

Application.StatusBar = False
Application.StatusBar = "File Archive Starts"  'Status for the archive

Dim DDate As String
DDate = Format(Now, "YYYYMMDD")

Dim MyNewFile As String, MyFileTemplate As String
MyFileTemplate = MyWorkingFilePath & SChar(MyWorkingFilePath) & DestFolder _
& SChar(DestFolder) & DestFolderFileTemplate   'My Template file and its Filepath


MyNewFile = "CBDECH_" & DDate & ".xls"         'My new File I want to create basis the Template
MyNewFile = MyWorkingFilePath & SChar(MyWorkingFilePath) & DestFolder _
& SChar(DestFolder) & MyNewFile                'My New File and its Filepath


ADOCreate MyNewFile:=MyNewFile, MyFileTemplate:=MyFileTemplate 'Create MyNewFile with closed workbooks
Application.StatusBar = "File Archive Created"

    'Below in ''' comments is a function called WaitForFileClose written by CPearson
    'Below is the link with its code. You can copy it on a separate module
    'http://www.cpearson.com/excel/WaitForFileClose.htm
    
    'WaitForFileClose is an excellent function to specify whether an Excel file
    'has been closed or currently used by another person
    
    '''Dim IsClosed As Boolean
    '''IsClosed = WaitForFileClose(FileName:=MyNewFile, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
    '''If IsClosed = True Then
    
        '''''''''''''''''''''''''''''''
        ''''' The file was closed before
        ''''' the time out expired.
        '''''''''''''''''''''''''''''''

ADOUpdate MyNewFile:=MyNewFile, MyFileTemplate:=MyFileTemplate

Application.StatusBar = "File Archive Created and Updated"
MsgBox "File Archive Created and Updated"
Application.StatusBar = False
                          
    '''Else
        '''''''''''''''''''''''''''''''
        ''''' The procedure timed out.
        '''''''''''''''''''''''''''''''
        '''MsgBox "TimeOut. File CBDECH_" & DeliveryDate & " is still open."
    '''End If

End Sub

Sub ADOCreate(MyNewFile As String, MyFileTemplate As String)

If Not FolderFileExists(MyNewFile) = True Then

Dim oConn As New ADODB.Connection
Set oConn = New ADODB.Connection
oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
           "Data Source=" & MyFileTemplate & ";" & _
           "Extended Properties=""Excel 12.0 Macro;HDR=YES;"""
           
'The Right hand side is the range from your template file which copies its objects to create the
'new file "CBDECH_" & DDate & ".xls" on the filepath designated
oConn.Execute "SELECT * INTO [Excel 8.0;DATABASE=" & MyNewFile & "].[DECH] FROM [DECH$A1:F26]"
oConn.Execute "SELECT * INTO [Excel 8.0;DATABASE=" & MyNewFile & "].[CHDE] FROM [CHDE$A1:F26]"
If CBool(oConn.State And adStateOpen) Then oConn.Close
Set oConn = Nothing

End If

End Sub

Sub ADOUpdate(MyNewFile As String, MyFileTemplate As String)

Dim sConnect As String
Dim objCommand As ADODB.Command
Dim lRecordsAffected As Long

'Your Destination Worksheets
Dim MyDestWst As String

'Create arrMyBorders
Dim b As Integer
Dim arrMyBorders() As Variant

'Need to insert your Number of Column Names
ReDim arrMyBorders(1 To 10)

'Need to insert your desired Column Names
arrMyBorders = Array("RESDECHY", "RESDECHM", "NOMDECHY", "NOMDECHM", "NOMDECHD", "RESCHDEY", "RESCHDEM", "NOMCHDEY", "NOMCHDEM", "NOMCHDED")

'Need to insert your Number of Column Names
For b = 1 To 10
arrMyBorders(b) = CStr(arrMyBorders(b))
Next b

'Create arrMyHours
Dim arrMyHours() As Variant
Dim h As Integer
ReDim arrMyHours(1 To 24)
arrMyHours = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)

'Create arrMyTradeData_data
Dim arrMyTradeData(1 To 24, 1 To 10) As Variant
Dim ValuesInc As Double

'*******       'Randomly create an array of Trade Data with Numerical Values
ValuesInc = 0
For b = 1 To 10
For h = 1 To 24
If b > 5 Then
arrMyTradeData(h, b) = CLng(ValuesInc + 1)   'Create CHDE worksheet with no decimal values
Else
arrMyTradeData(h, b) = CDbl(ValuesInc + 1.1) 'Create DECH worksheet with decimal values
End If
ValuesInc = ValuesInc + 1
Next h
Next b
'*******

Dim arrMyTradeExpData(1 To 24, 1 To 5) As Double, arrMyTradeImpData(1 To 24, 1 To 5) As Double
Dim MyTradeCols As Integer, MyTradeRows As Integer

MyTradeRows = UBound(arrMyTradeData, 1) 'h
MyTradeCols = UBound(arrMyTradeData, 2) 'b

For h = 1 To MyTradeRows

For b = 1 To MyTradeCols               'If there is an empty Value then modify it to Zero
If IsEmpty(arrMyTradeData(h, b)) Then
arrMyTradeData(h, b) = 0
Else

arrMyTradeData(h, b) = Format(arrMyTradeData(h, b), "##0.0")  'Insert your desired decimal places
'OR
'If you experience issues with local settings having comma as decimal separatos
'arrMyTradeData(h, b) = Replace(Format(arrMyTradeData(h, b), "##0.0"), ",", ".")
'or dot decimal separators
'arrMyTradeData(h, b) = Replace(Format(arrMyTradeData(h, b), "##0.0"), ".", ",")
End If
Next b

'Separate DECH from your Array Trade Data
arrMyTradeExpData(h, 1) = arrMyTradeData(h, 1)
arrMyTradeExpData(h, 2) = arrMyTradeData(h, 2)
arrMyTradeExpData(h, 3) = arrMyTradeData(h, 3)
arrMyTradeExpData(h, 4) = arrMyTradeData(h, 4)
arrMyTradeExpData(h, 5) = arrMyTradeData(h, 5)

'Separate CHDE from your Array Trade Data
arrMyTradeImpData(h, 1) = arrMyTradeData(h, 6)
arrMyTradeImpData(h, 2) = arrMyTradeData(h, 7)
arrMyTradeImpData(h, 3) = arrMyTradeData(h, 8)
arrMyTradeImpData(h, 4) = arrMyTradeData(h, 9)
arrMyTradeImpData(h, 5) = arrMyTradeData(h, 10)

Next h
Erase arrMyTradeData()

'************************

sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & MyNewFile & ";" & _
           "Extended Properties=""Excel 8.0;HDR=Yes;"";"
           
Set objCommand = New ADODB.Command
objCommand.ActiveConnection = sConnect

'ArrayLoop
MyTradeRows = UBound(arrMyHours) 'h

For b = 1 To 5
For h = 1 To MyTradeRows

'Notes :
'Note that on String type values you need to open and close your setting variable value with ' ex. arrMyBorders(b)
'after WHERE you can also use a second filtering with the and function. Make sure you have spaces when you concatenate

'My Export Array Insert Data by means of SQL Update
MyDestWst = "DECH"
objCommand.CommandText = "UPDATE [" & MyDestWst & "$A1:F26] " & _
"SET [" & arrMyBorders(b) & "]='" & arrMyTradeExpData(h, b) & "' WHERE Hour=" & arrMyHours(h)
On Error Resume Next
objCommand.Execute RecordsAffected:=lRecordsAffected, Options:=adCmdText Or adExecuteNoRecords
On Error GoTo 0

'My Import Array Insert Data by means of SQL Update
MyDestWst = "CHDE"
objCommand.CommandText = "UPDATE [" & MyDestWst & "$A1:F26] " & _
"SET [" & arrMyBorders(b + 5) & "]='" & arrMyTradeImpData(h, b) & "' WHERE Hour=" & arrMyHours(h)
On Error Resume Next
objCommand.Execute RecordsAffected:=lRecordsAffected, Options:=adCmdText Or adExecuteNoRecords
On Error GoTo 0

Next h
Next b

'Kill Command Object
If CBool(objCommand.ActiveConnection.State And adStateOpen) Then objCommand.ActiveConnection.Close
Set objCommand = Nothing

Erase arrMyTradeImpData()
Erase arrMyTradeData()
Erase arrMyHours()
Erase arrMyBorders()
End Sub

Private Function FolderFileExists(FFName As String) As Boolean
Dim FFNameThere As String
FFNameThere = Dir(FFName, vbDirectory)
If FFNameThere = "" Then FFNameThere = Dir(FFName)
If FFNameThere = "" Then
FolderFileExists = False
Else
FolderFileExists = True
End If
End Function

Private Function SChar(FPath As String) As String
SChar = IIf(Right(FPath, 1) = "\", "", "\")
End Function
