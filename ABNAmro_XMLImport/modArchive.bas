Attribute VB_Name = "modArchive"
Option Explicit
Option Base 1

Sub MyArchive()
Dim DestWB() As Workbook, WBC As Integer
Dim i As Integer
Dim MySelectedDate As String
Application.ScreenUpdating = False
    
'Create the single Workbook
WBC = 1        'Assume you Create at least 1
Workbooks.Add
ReDim Preserve DestWB(1 To WBC)
Set DestWB(WBC) = ActiveWorkbook

'Create any extra Workbook
For i = 1 To ThisWorkbook.Worksheets("Dashboard").Range("C2").Value
If Left(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i).Value, 8) <> _
Left(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i + 1).Value, 8) And _
ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i + 1).Value <> "" Then
WBC = WBC + 1
Workbooks.Add
ReDim Preserve DestWB(1 To WBC)
Set DestWB(WBC) = ActiveWorkbook
End If
Next i

Application.DisplayAlerts = False
Dim WS As Worksheet, w As Integer, WSC As Integer, WSComp As Long
For w = 1 To WBC
WSC = 1
  For Each WS In ThisWorkbook.Worksheets
    If Not WS.Name = "Dashboard" Then
    For i = 1 To ThisWorkbook.Worksheets("Dashboard").Range("C2").Value
    WSComp = InStr(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i).Value, Left(WS.Name, 2))
    If WSComp > 0 Then
    If Left(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i).Value, 8) = _
    Left(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i + 1).Value, 8) And _
    ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i + 1).Value <> "" Then
    
    With DestWB(w)
    If WSC > .Sheets.Count Then .Sheets.Add After:=.Sheets(.Sheets.Count)
    .Worksheets(WSC).Name = WS.Name
    .Worksheets(WSC).Range("A1:AD6000").Value = WS.Range("A1:AD6000").Value
    .Worksheets(WSC).Columns("A:AD").AutoFit
    WSC = WSC + 1
    End With
    
    End If
    End If
    Next i
    End If
  Next WS
  
MySelectedDate = Left(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i - 1).Value, 8)
With DestWB(w)
    .SaveAs ThisWorkbook.Worksheets("Dashboard").Range("E15").Value & _
    IIf(Right(ThisWorkbook.Worksheets("Dashboard").Range("E15").Value, 1) = "\", "", "\") & _
    "Unzipped\" & MySelectedDate & "\" & "ABNAmbro_" & MySelectedDate & ".xls", _
    FileFormat:=-4143
    .Close savechanges:=False
End With
Next w
    
fine:
Erase DestWB
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
