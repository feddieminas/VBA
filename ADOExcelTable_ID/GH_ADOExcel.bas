Attribute VB_Name = "GH_ADOExcel"
Option Explicit
Option Base 1

'Tried in Windows 7 and Excel 2007, 2010

'****************************************
'ADO ID RETRIEVE, UPDATE, INSERT, DELETE
'****************************************

'1)Need to check in your ODBC Data Source Administrator
'on User DSN tab that you have :
'User Data Sources Name = Excel Files
'User Data Sources Driver = Microsoft Excel Driver(*xls,*xlsx,*xlsm,*xlsb)
'2)Add VBA Library reference :
'Need to add the MDAC Library named Microsoft ActiveX Data Objects 2.8 Library
'VBA Editor --> Tools --> References --> Microsoft ActiveX Data Objects 2.8 Library
'Alternatively you can find it VBA Editor --> Tools --> References --> Browse
'C:\Program Files (x86)\Common Files\System\ado\msado28.tlb

'Instructions
'Below there is an example i have created. To get an idea and see that it works for you, you can adjust the macro Settings
'in filepaths, folder and file template name of your preference. If you would not insert them, by default it will create
'the variables with its names of my imagination
'Run CreateFileTemplate, RunProcIDUpd, RunProcIDRet, RunProcIDIns and RunProcIDDel macro in order

Dim MyWorkingFilePath As String
Dim FileTemplate As Variant
Dim myYear As Long

Sub MySettings()
'Example :
'MyWorkingFilePath ="C:\Users\faidon.dermesonoglou\Desktop"
'FileTemplate = "ID.xls"
'It will create a template file inside your WorkingFilePath to be like "C:\Users\faidon.dermesonoglou\Desktop\ID.xls"

MyWorkingFilePath = ""                                                          'INSERT or keep it your Filepath you want to create your Folder for your Database
If MyWorkingFilePath = "" Then MyWorkingFilePath = ThisWorkbook.Path

myYear = 0                                                                      'INSERT or keep it your Year
If myYear = 0 Then myYear = 2016

FileTemplate = ""                                                               'INSERT or keep it Recommended to be an xls file type
If FileTemplate = "" Then FileTemplate = "ID.xls"

End Sub

Sub CreateFileTemplate() 'Need to Run only once to create your ID Template file

MySettings

If FolderFileExists(MyWorkingFilePath & SChar(MyWorkingFilePath) & FileTemplate) = True Then Exit Sub

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim WB As Workbook
Set WB = Workbooks.Add

With WB  'ID Workbook

With .Worksheets(1)
.Name = "Sheet1"
.Range("A1").Value = "DDATE"
.Range("B1").Value = "ID"

Dim arrDates() As Variant
Dim myDays As Long, i As Integer

myDays = 365
If myYear Mod 4 = 0 Then myDays = 366 'If Leap Year
ReDim arrDates(myDays)

For i = 1 To myDays
arrDates(i) = Format(CDate("01/01/" & myYear) + i - 1, "YYYYMMDD")
Next i

.Range("A2").Resize(myDays, 1).Value = Application.Transpose(arrDates)
Erase arrDates()

.Range("B2:B" & myDays + 1).Value = 0
.Range("B2:B" & myDays + 1).NumberFormat = "0"  'Specify Number Format to zero decimal places
'.Range("B2:B" & myDays + 1).NumberFormat = "0.0"  'Specify Number Format to one decimal place
'.Range("B2:B" & myDays + 1).NumberFormat = "0.00"  'Specify Number Format to two decimal places

.Range("B2:B" & myDays + 1).NumberFormat = "0_ ;[Red]-0 "      'zero decimal with negative numbers as red
'.Range("B2:B" & myDays + 1).NumberFormat = "0.0_ ;[Red]-0.0 "
'.Range("B2:B" & myDays + 1).NumberFormat = "0.00_ ;[Red]-0.00 "

.Range("A:B").EntireColumn.AutoFit
.Columns("B:B").ColumnWidth = 5.43

End With

.SaveAs MyWorkingFilePath & SChar(MyWorkingFilePath) & FileTemplate, FileFormat:=56
.Close savechanges:=False

End With

Set WB = Nothing

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

'********************************
'ADO ID UPDATE
'********************************

Sub RunProcIDUpd() 'Update 20160101 with a value of 1908
MySettings

Dim IDRetrieval As Long, MyIDNum As Long
Dim myMonth As String, myDay As String
myMonth = Format(CDate("01/01/" & myYear), "MM")
myDay = Format(CDate("01/01/" & myYear), "DD")

MyIDNum = 1908

If IDUpd(myYear & myMonth & myDay, CStr(MyIDNum)) Then
MsgBox "ID Updated = " & MyIDNum
End If

End Sub

Function IDUpd(idate As String, IDNumber As String) As Boolean
IDUpd = False
On Error GoTo ErrHandler:

Dim IDData As cID
Set IDData = New cID

With IDData
.sDestWBName = MyWorkingFilePath & SChar(MyWorkingFilePath) & FileTemplate 'ID.xls file
.sDestWSName = "Sheet1"
.sDestWSRange = "$" & "A1:B380" 'Max Days of a Year 366 plus 14. Trials for 12 new insert rows
.sDDate = idate
.sID = IDNumber

IDUpd = .fUpdWS
End With

ErrHandler:
Set IDData = Nothing
End Function


'********************************
'ADO ID RETRIEVE
'********************************

Sub RunProcIDRet() 'Retrieve ID value of 20160101
MySettings

Dim IDRetrieval As Long
Dim myMonth As String, myDay As String
myMonth = Format(CDate("01/01/" & myYear), "MM")
myDay = Format(CDate("01/01/" & myYear), "DD")

IDRetrieval = CLng(IDRet(myYear & myMonth & myDay))

MsgBox "ID Retrieved = " & IIf(IDRetrieval = -1, "NULL", IDRetrieval)
End Sub

Function IDRet(idate As String) As String
Dim IDData As cID
Set IDData = New cID

With IDData
.sSourceWBName = MyWorkingFilePath & SChar(MyWorkingFilePath) & FileTemplate 'ID.xls file
.sSourceWSName = "Sheet1"
.sSourceWSRange = "A1:B380" 'Max Days of a Year 366 plus 14. Trials for 12 new insert rows
.sDDate = idate

IDRet = .fRetWS
End With

If IDRet = "NULL" Then IDRet = "-1" 'Set first negative number if no IDRet is Retrieved

Set IDData = Nothing
End Function


'********************************
'ADO ID INSERT
'********************************

Sub RunProcIDIns() 'Insert 20170101 record with an ID value of 2017
                   'Assume you want unique rather than duplicate values
MySettings

Dim MySourceWB As String, MySourceWS As String, MySourceWRng As String
MySourceWB = ThisWorkbook.FullName
MySourceWS = ThisWorkbook.Worksheets(1).Name

Dim myNewYear As Long
myNewYear = myYear + 1

With ThisWorkbook.Worksheets(MySourceWS)
.Range("A1").Value = "Header"
.Range("B1").Value = "DDATE"            'Column Header 1
.Range("B2").Value = myNewYear & "0101" 'Column Value 1 ex. if myYear=2016 then .Range("A2").Value = 20170101

.Range("A2").Value = "IDInsert"
.Range("C1").Value = "ID"               'Column Header 2
.Range("C2").Value = myNewYear          'Column Value 2
End With

MySourceWRng = "B1:C2"

If IDIns(MySourceWB, MySourceWS, MySourceWRng, ThisWorkbook.Worksheets(MySourceWS).Range("B2").Value) Then
MsgBox "ID Inserted=" & ThisWorkbook.Worksheets(MySourceWS).Range("C2").Value
End If

'ThisWorkbook.Worksheets(MySourceWS).Range("A1").Resize(2, 3).Cells.Clear

End Sub

Function IDIns(MySourceWB As String, MySourceWS As String, MySourceWRng As String, idate As String) As Boolean
IDIns = False
On Error GoTo ErrHandler:

Dim MyInsID As Boolean
Dim IDRetrieval As String

Dim IDData As cID
Set IDData = New cID

With IDData
'Check if a Row already exists
.sSourceWBName = MyWorkingFilePath & SChar(MyWorkingFilePath) & FileTemplate 'ID.xls file
.sSourceWSName = "Sheet1"
.sSourceWSRange = "A1:B380" 'Max Days of a Year 366 plus 14. Trials for 12 new insert rows
.sDDate = idate

IDRetrieval = .fRetWS

'Inserting a Row if it does not exist
If IDRetrieval = "NULL" Then
.sSourceWBName = MySourceWB
.sSourceWSName = MySourceWS
.sSourceWSRange = MySourceWRng
.sDestWBName = MyWorkingFilePath & SChar(MyWorkingFilePath) & FileTemplate 'ID.xls file
.sDestWSName = "Sheet1"

MyInsID = .fInsWS
If MyInsID = True Then IDIns = True 'or IDIns = MyInsID
End If

ErrHandler:
End With

Set IDData = Nothing
End Function


'********************************
'ADO ID DELETE
'********************************

'Notes
'Process of ADO process deletion on an Excel Database can occur only through process of
'updating the current record to a value that is considered a deleted record
'As per default, a negative ID value in this example is considered a deleted record

Sub RunProcIDDel() 'Delete 20170101 record with an ID value of -1925 (number of personal choice)
MySettings

Dim IDRetrieval As Long, myNewYear As Long
Dim myMonth As String, myDay As String

myNewYear = myYear + 1

myMonth = Format(CDate("01/01/" & myNewYear), "MM") 'Ex. 20170101
myDay = Format(CDate("01/01/" & myNewYear), "DD")

If IDDel(myNewYear & myMonth & myDay) Then
MsgBox "ID Deleted = " & myNewYear & myMonth & myDay
End If

End Sub

Function IDDel(idate As String) As Boolean
IDDel = False
On Error GoTo ErrHandler:

Dim IDData As cID
Set IDData = New cID

With IDData
.sDestWBName = MyWorkingFilePath & SChar(MyWorkingFilePath) & FileTemplate 'ID.xls file
.sDestWSName = "Sheet1"
.sDestWSRange = "$" & "A1:B380" 'Max Days of a Year 366 plus 14. Trials for 12 new insert rows
.sDDate = idate

IDDel = .fDelWS
End With

ErrHandler:
Set IDData = Nothing
End Function


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
