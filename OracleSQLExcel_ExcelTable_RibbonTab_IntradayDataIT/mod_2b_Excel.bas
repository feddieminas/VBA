Attribute VB_Name = "mod_2b_Excel"
Option Explicit

'Worksheet : Sheet (Sheet 6)
'Creates a copy of this sheet and saves the file on the specified folder indicated on sheet Settings

'MI is the Intraday Market Exchange of the Power Market in Italy

Sub crtMISheet()

Dim Destwb As Workbook
Dim anno As Integer, y As Integer, m As Integer, DatesC As Integer, MIsC As Integer, DestwbStartrow As Integer, RowsC As Integer
Dim Filename As String, Filepath As String, isitthere As String, MImarket As String, Sourcename As String
Dim ChooseData As Date
Dim WbkExists As Boolean

intro 'procedure is on sheet mod_2_MIDatabaseRun

Application.DisplayAlerts = False

'Count the number of Delivery Dates
DatesC = ThisWorkbook.Worksheets("MIQty").Range("AK5").Value

'Count MI Markets
MIsC = ThisWorkbook.Worksheets("MIQty").Range("AK6").Value

For y = 1 To DatesC

For m = 1 To MIsC

'ClearContents of existing Data
ThisWorkbook.Worksheets("Sheet").Range("A:F").ClearContents

'Format Date for Filtering
'ChooseData = Format(ThisWorkbook.Worksheets("MIQty").Range("Dates").Cells(y, 1).Value, "mm/dd/yyyy")
With ThisWorkbook.Worksheets("MIQty").Range("Dates")
ChooseData = DateSerial(Year(.Cells(y, 1)), Month(.Cells(y, 1)), Day(.Cells(y, 1)))
End With

'Determine your market
MImarket = ThisWorkbook.Worksheets("MIQty").Range("MIs").Cells(m, 1).Value

'Determine your Area
With ThisWorkbook.Sheets("Database")
If .AutoFilterMode = False Then .Cells(1, 1).AutoFilter
.Range("A1").AutoFilter Field:=1, Criteria1:=ChooseData
.Range("A1").AutoFilter Field:=3, Criteria1:=MImarket
.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Rows.Copy
With ThisWorkbook.Worksheets("Sheet").Range("A1")
.PasteSpecial xlPasteValues
Application.CutCopyMode = False
End With
.AutoFilterMode = False
End With

Dim MyYear As Long
MyYear = Year(ThisWorkbook.Worksheets("MIQty").Range("Dates").Cells(y, 1).Value)

'Filepath and Filename
Filepath = ThisWorkbook.Worksheets("Settings").Range("G6").Value
Filename = "MI_" & Format(ThisWorkbook.Worksheets("MIQty").Range("Dates").Cells(y, 1).Value, "YYYYMMDD")
Sourcename = Filepath & Filename

'Check if there exist an MI worksheet already there
isitthere = Dir(Filepath & Filename & ".xls")  'one can also choose to save a file to an xlsx. If so u then need to change the FileFormat num below

If isitthere <> "" Then

ChDir Filepath

tryagain:

On Error Resume Next
Set Destwb = Workbooks.Open(Sourcename & ".xls", True, False) 'one can also choose to save a file to an xlsx. If so u then need to change the FileFormat num below
On Error GoTo 0

WbkExists = isworkbookopen(Filename & ".xls") 'one can also choose to save a file to an xlsx. If so u then need to change the FileFormat num below
If WbkExists = False Then GoTo tryagain:

'Count Rows of Sourcewb minus 1 which is the header row
With ThisWorkbook.Worksheets("Sheet")
RowsC = .Cells(.Rows.count, "A").End(xlUp).Row - 1
End With

With Destwb

If RowsC = 0 Then GoTo mynext

'delete any existing Values for this Market
If .Worksheets(1).AutoFilterMode = False Then .Worksheets(1).Cells(1, 1).AutoFilter
.Worksheets(1).Range("A1").AutoFilter Field:=3, Criteria1:=MImarket
.Worksheets(1).Range("A1").CurrentRegion.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
.Worksheets(1).AutoFilterMode = False

'Then reinsert values
DestwbStartrow = .Worksheets(1).Cells(.Worksheets(1).Rows.count, "A").End(xlUp).Row + 1
'Copy paste from SourceWb Cells apart from Header Row to DestWb StartRow (last non filled value
'Plus rows count of SourceWb minus 1 which is the header row of DestWb
.Worksheets(1).Range("A" & DestwbStartrow & ":F" & DestwbStartrow + RowsC - 1).Value = _
ThisWorkbook.Worksheets("Sheet").Range("A2:F" & RowsC + 1).Value

'Save the File
.Save

mynext:

'Close File
.Close

End With

Set Destwb = Nothing

Else

'Create a Copy of your Sheet
ThisWorkbook.Worksheets("Sheet").Copy

'Save it to a directory
ActiveWorkbook.SaveAs Filepath & Filename, FileFormat:=56 'file xlsx FileFormat:=51, file xls FileFormat:=56 - xlsx file is lighter than xls

'Close Workbook
ActiveWorkbook.Close

End If

Next m 'Market

Next y 'Dates

conclusion 'procedure is on sheet mod_2_MIDatabaseRun

Application.DisplayAlerts = True

End Sub







