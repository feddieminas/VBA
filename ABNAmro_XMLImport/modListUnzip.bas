Attribute VB_Name = "modListUnzip"
Option Explicit
Option Base 1

Dim SampleFile As Boolean

Public Function FileFolderExists(strFullPath As String) As Boolean
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
EarlyExit:
    On Error GoTo 0
End Function

Sub MyClearDash()

ThisWorkbook.Worksheets("Dashboard").Range("A3:C100").Cells.Clear
ThisWorkbook.Worksheets("Dashboard").Range("A1").Value = "Zips"
ThisWorkbook.Worksheets("Dashboard").Range("A2").ClearContents
ThisWorkbook.Worksheets("Dashboard").Range("C1").Value = "Files"
ThisWorkbook.Worksheets("Dashboard").Range("C2").ClearContents

End Sub

Sub MyClearWS()
Dim WS As Worksheet
    
  Application.DisplayAlerts = False
  For Each WS In ThisWorkbook.Worksheets
    If Not WS.Name = "Dashboard" Then WS.Delete
  Next WS
  Application.DisplayAlerts = True

End Sub


Sub ListFiles(targetFolder As String)
Dim objFSO As Object
 Dim objFolder As Object
 Dim objFile As Object
 Dim i As Integer

ThisWorkbook.Worksheets("Dashboard").Range("A3:A100").Cells.Clear
ThisWorkbook.Worksheets("Dashboard").Range("A2").ClearContents
ThisWorkbook.Worksheets("Dashboard").Range("A1").Value = "Zips"

Set objFSO = CreateObject("Scripting.FileSystemObject")
 'Get the folder object
Set objFolder = objFSO.GetFolder(targetFolder)
 i = 1

SampleFile = False

'loops through each file in the directory and prints their names and path
For Each objFile In objFolder.Files
     'print file name
    ThisWorkbook.Worksheets("Dashboard").Range("A" & 2 + i).Value = objFile.Name
    If InStr(1, objFile.Name, "Sample", vbTextCompare) > 0 Then SampleFile = True
    ThisWorkbook.Worksheets("Dashboard").Range("A" & 2 + i).Interior.ColorIndex = 6
    'Kill targetFolder & ThisWorkbook.Worksheets("Dashboard").Range("A" & 2 + i).Value
    i = i + 1
Next objFile

ThisWorkbook.Worksheets("Dashboard").Range("C2").Value = i - 1
If SampleFile = True Then ThisWorkbook.Worksheets("Dashboard").Range("C2").Value = 7

Set objFolder = Nothing
Set objFSO = Nothing
 End Sub

Sub MyUnzip()
Dim targetFolder As String, targetFileZip As String
Dim Fname As Variant
Dim i As Integer

'Insert target folder at Sheet Dashboard on Range E15
targetFolder = ThisWorkbook.Worksheets("Dashboard").Range("E15").Value & IIf(Right(ThisWorkbook.Worksheets("Dashboard").Range("E15").Value, 1) = "\", "", "\")

ListFiles targetFolder

ThisWorkbook.Worksheets("Dashboard").Range("C3:C100").Cells.Clear
ThisWorkbook.Worksheets("Dashboard").Range("C1").Value = "Files"

If FileFolderExists(targetFolder & "Unzipped\") Then
Else
On Error Resume Next
MkDir targetFolder & "Unzipped\"
On Error GoTo 0
End If

Dim LoopTimes As Integer, SampleLoopTimes As Integer
If SampleFile = True Then
LoopTimes = 1
Else
LoopTimes = ThisWorkbook.Worksheets("Dashboard").Range("C2").Value
End If

SampleLoopTimes = -1

For i = 1 To LoopTimes

targetFileZip = targetFolder & ThisWorkbook.Worksheets("Dashboard").Range("A" & 2 + i).Value

'Retrieve Name of the Files
Dim o As Object, ofile As Variant, ZipFiles
Set o = CreateObject("Shell.Application")

ZipFiles = targetFileZip
For Each ofile In o.Namespace(ZipFiles).Items
If SampleFile = True Then SampleLoopTimes = SampleLoopTimes + 1
ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i + SampleLoopTimes).Value = ofile.Name & ".xml"
Next ofile
Set o = Nothing

'MKDir
targetFolder = targetFolder & "Unzipped\" & Left(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i).Value, 8)
If Not FileFolderExists(targetFolder) Then
MkDir (targetFolder)
End If

Fname = targetFileZip
    
MyUnzip7 FileNameZip:=Fname, NameUnzipFolder:=targetFolder

'Kill Fname

Next i

'ThisWorkbook.Worksheets("Dashboard").Range("C2").Value = i - 1

End Sub



'************************************* BACKUP XMLMAP ******************************************************
Sub LoopMyFiles()
End Sub

Private Sub Delete_XMLMaps()
'If any of the import are considered as XML Maps
Dim XMLMap
For Each XMLMap In ThisWorkbook.XmlMaps
XMLMap.Delete
Next
End Sub

Public Sub XMLIMport()
Delete_XMLMaps

'Dim MyExtrZipFiles(7) As String

Dim rng As Range
Set rng = ThisWorkbook.Worksheets("CA").Range("A2")

Dim MyCellInTable As Boolean
On Error Resume Next
MyCellInTable = (ThisWorkbook.Worksheets("CA").Range("A2").ListObject.Name <> "")
On Error GoTo 0

Dim My_Table As ListObject
If MyCellInTable = True Then
Set My_Table = rng.ListObject
My_Table.Unlist
Set My_Table = Nothing
End If

ThisWorkbook.Worksheets("CA").Cells.Clear

Dim strXML As String
strXML = "C:\ABNAmbroXML\Unzipped\20150525\20150525-2141-C2141-BM (L)-16628583.xml"

ThisWorkbook.XMLIMport url:=strXML, ImportMap:=Nothing, Overwrite:=True, Destination:=rng
Set rng = Nothing

Delete_Connection
End Sub

Private Sub Delete_Connection()
Dim Conn As Object        '* connection collection
For Each Conn In ThisWorkbook.Connections
Conn.Delete
Next Conn
End Sub
'************************************* BACKUP XMLMAP ******************************************************


