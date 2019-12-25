Attribute VB_Name = "mod_2_MIDatabaseRun"
Option Explicit

'MGP is the Day-Ahead Market Exchange of the Power Market in Italy
'MI is the Intraday Market Exchange of the Power Market in Italy

Sub intro()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

End Sub

Sub conclusion()

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub MIunits()

Dim ValidaSt As Integer, MISt As Integer, x As Integer, y As Integer, z As Integer, f As Integer, h As Integer, m As Integer
Dim arethererows As Integer, UnitsC As Integer, DatesC As Integer, oAreasC As Integer, oRowsC As Integer, MIsC As Integer
Dim SourceStartRow As Long, SourceLastRow As Long, DestStartRow As Long, DestLastRow As Long, lct As Long
Dim MyFinalQty As Double, MyFinalPrice As Double
Dim ClearDirtyStuff As Variant
Dim SourceRng, DestRng, oArea As Range
Dim MySourceTable, MyDestTable As ListObject
Dim DeliveryDate As Date
Dim OracleProgram As Boolean

intro

OracleProgram = False

With ThisWorkbook.Worksheets("MIQty")

'Check There is no State called Valida
.Range("AK1").Calculate
ValidaSt = .Range("AK1").Value
If ValidaSt > 0 Then
MsgBox "There are inside Stato Valida Values.Try Again"
GoTo fine
End If
SourceStartRow = 2
SourceLastRow = .Cells(.Rows.count, "A").End(xlUp).Row

'Check There are MI Values
.Range("AK2").Calculate
MISt = .Range("AK2").Value
If SourceLastRow <> MISt Then
MsgBox "There are inside MGP values you are trying to upload" 'We try to upload Intraday MI Market values, thus any Day-Ahead MGP values are not accepted.
GoTo fine
End If

'Trim all Units
.Range("AH2:AH1000").ClearContents
.Range("AH" & SourceStartRow & ":AH" & SourceLastRow).FormulaR1C1 = _
"=TRIM(CLEAN(SUBSTITUTE(LEFT(TRIM(RC[-33]),LEN(TRIM(RC[-33]))-OR(RIGHT(TRIM(RC[-33]))={""?"";""!"";"".""})),CHAR(160),"" "")))"
.Range("A" & SourceStartRow & ":A" & SourceLastRow).Value = _
.Range("AH" & SourceStartRow & ":AH" & SourceLastRow).Value
.Range("AH2:AH1000").ClearContents

'Source Table List
'******************************************************************************

.ListObjects.Add(xlSrcRange, .Range("$A$1:$AB$" & SourceLastRow), , xlYes).Name = "MIList"
Set SourceRng = .Range("A1")
Set MySourceTable = SourceRng.ListObject

'******************************************************************************

'Count the number of Delivery Days to insert, the Units and MI markets you would like to insert
'Units
.Range("AM1:AM51").ClearContents
MySourceTable.ListColumns(1).Range.AdvancedFilter Action:=xlFilterCopy, _
CopyToRange:=.Range("AM1"), Unique:=True
'UnitsCount Calculate
.Range("AK4").Calculate
UnitsC = .Range("AK4").Value

'Date
.Range("AN1:AN51").ClearContents
MySourceTable.ListColumns(3).Range.AdvancedFilter Action:=xlFilterCopy, _
CopyToRange:=.Range("AN1"), Unique:=True
'DeliveryDaysCount Calculate
.Range("AK5").Calculate
DatesC = .Range("AK5").Value

'MIMarkets
.Range("AO1:AO51").ClearContents
MySourceTable.ListColumns(4).Range.AdvancedFilter Action:=xlFilterCopy, _
CopyToRange:=.Range("AO1"), Unique:=True
'MImarkets Calculate
.Range("AK6").Calculate
MIsC = .Range("AK6").Value

OracleProgram = True

'Destination Table List
''''******************************************************************************''''

With ThisWorkbook.Worksheets("Database")
'Clear Previous Contents of Database Quantity
.Range("A2:F1000").ClearContents
'Find Last row
DestLastRow = .Cells(.Rows.count, "A").End(xlUp).Row + 1
.ListObjects.Add(xlSrcRange, .Range("$A$1:$F$" & DestLastRow), , xlYes).Name = "MIListDtbs"
.ListObjects("MIListDtbs").TableStyle = "TableStyleMedium11"
Set DestRng = .Range("A1")
Set MyDestTable = DestRng.ListObject
End With

''''******************************************************************************''''

'Check if there are already existing Values on Destination Sheet
'If there are values for this Unit, for this Date, for this Market on database
'Lets Delete them
For y = 1 To DatesC

For m = 1 To MIsC

For x = 1 To UnitsC

'Delivery Date
DeliveryDate = .Range("Dates").Cells(y, 1).Value
Debug.Print DeliveryDate
MyDestTable.Range.AutoFilter Field:=1, Criteria1:=DeliveryDate
'Market
Debug.Print .Range("MIs").Cells(m, 1).Value
MyDestTable.Range.AutoFilter Field:=3, Criteria1:="=" & .Range("MIs").Cells(m, 1).Value
'Unit
Debug.Print .Range("Units").Cells(x, 1).Value
MyDestTable.Range.AutoFilter Field:=4, Criteria1:="=" & .Range("Units").Cells(x, 1).Value

arethererows = MyDestTable.Range.SpecialCells(xlCellTypeVisible).Offset(1, 0).Cells.count

'Deletion Process
If arethererows > 6 And DestLastRow > 1 Then 'the const number 6 are the number of columns on sheet Database. If there are values more than the headers then proceed
Debug.Print MyDestTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas.count
For Each oArea In MyDestTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas
For lct = oArea.Rows.count To 1 Step -1
oArea.EntireRow.Rows(lct).Delete
Next
Next
End If

Next x 'Units

Next m 'Market

Next y 'Dates

'Reset to Default Fields
MyDestTable.Range.AutoFilter Field:=1
MyDestTable.Range.AutoFilter Field:=3
MyDestTable.Range.AutoFilter Field:=4

'******************************************************************************

'DetermineLastRow in order to Find the StartRow to input values
With ThisWorkbook.Worksheets("Database")
DestLastRow = .Cells(.Rows.count, "A").End(xlUp).Row
End With

'Insert New Values on the Table Database
For y = 1 To DatesC

For m = 1 To MIsC

For x = 1 To UnitsC

'Unit
Debug.Print .Range("Units").Cells(x, 1).Value
MySourceTable.Range.AutoFilter Field:=1, Criteria1:="=" & .Range("Units").Cells(x, 1).Value
'Delivery Date
DeliveryDate = .Range("Dates").Cells(y, 1).Value
Debug.Print DeliveryDate
MySourceTable.Range.AutoFilter Field:=3, Criteria1:=DeliveryDate
'Market
Debug.Print .Range("MIs").Cells(m, 1).Value
MySourceTable.Range.AutoFilter Field:=4, Criteria1:="=" & .Range("MIs").Cells(m, 1).Value
'State Accepted
''MySourceTable.Range.AutoFilter Field:=5, Criteria1:="Accettata"
If ThisWorkbook.Worksheets("ExchRes").Range("A1") = "Unit" Then
MySourceTable.Range.AutoFilter Field:=5, Criteria1:="Accepted"
ElseIf ThisWorkbook.Worksheets("ExchRes").Range("A1") = "Unità" Then
MySourceTable.Range.AutoFilter Field:=5, Criteria1:="Accettato"
End If

For h = 1 To 25 'Hours

MySourceTable.Range.AutoFilter Field:=9, Criteria1:=h 'Hour

arethererows = MySourceTable.Range.SpecialCells(xlCellTypeVisible).Offset(1, 0).Cells.count

If arethererows > 28 Then 'the const number 28 are the number of columns on sheet MIQty. If there are values more than the headers then proceed

'count Areas
Debug.Print MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas.count
oAreasC = MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas.count

MyFinalQty = 0

'copy and paste the values on the destination table

'Loop for Areas
For f = 1 To oAreasC

'Data
Debug.Print MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Row

ThisWorkbook.Worksheets("Database").Range("A" & DestLastRow).Value = _
.Range("C" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Row).Value

'Ora

ThisWorkbook.Worksheets("Database").Range("B" & DestLastRow).Value = _
.Range("I" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Row).Value

'Mercato

ThisWorkbook.Worksheets("Database").Range("C" & DestLastRow).Value = _
.Range("D" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Row).Value

'Unita

ThisWorkbook.Worksheets("Database").Range("D" & DestLastRow).Value = _
.Range("A" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Row).Value

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MIaccettatto.. We have to loop the rows because we might have more offers accepted for an hour

'the listrows count of your source table
Debug.Print MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Rows.count
oRowsC = MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Rows.count

Debug.Print MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Cells.count

'MyFinalQty = 0

'Loop for ListRows
For z = 1 To oRowsC

'If .Range("G" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Offset(z - 1, 0).Row).Value = "Acq (Buy)" Then
If (.Range("G" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Offset(z - 1, 0).Row).Value = "Acquisto") Or _
(.Range("G" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Offset(z - 1, 0).Row).Value = "Buy") Then
MyFinalQty = MyFinalQty + .Range("O" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Offset(z - 1, 0).Row).Value * -1
Else
MyFinalQty = MyFinalQty + .Range("O" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Offset(z - 1, 0).Row).Value
End If

MyFinalPrice = Format(.Range("R" & MySourceTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Areas(f).Offset(z - 1, 0).Row).Value, "##0.00")

'VisibleRows
Next z

ThisWorkbook.Worksheets("Database").Range("E" & DestLastRow).Value = MyFinalQty
ThisWorkbook.Worksheets("Database").Range("F" & DestLastRow).Value = MyFinalPrice

'DestLastRow = DestLastRow + 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Areas
Next f

DestLastRow = DestLastRow + 1

End If

Next h 'Hours

Next x 'Units

Next m 'MIMarkets

Next y 'Days

'Resize your Table with the Last Database Start Row
'MyDestTable.Resize Range("$A$1:$E$" & DestLastRow)

'Reset to Default Fields
MySourceTable.Range.AutoFilter Field:=1
MySourceTable.Range.AutoFilter Field:=3
MySourceTable.Range.AutoFilter Field:=4
MySourceTable.Range.AutoFilter Field:=5

'***********************************************************************************************************

'Convert To Range Your Table
MySourceTable.Unlist
MyDestTable.Unlist

End With
       
'***********************************************************************************************************

fine:
     
'Clear stuff that will lighter the file
With ThisWorkbook.Worksheets("MIQty").Range("A2:AB1000")
.ClearContents
.Interior.ColorIndex = 2
ClearDirtyStuff = Cl_G_H("MIQty", 0)
.Borders(xlDiagonalDown).LineStyle = xlNone
.Borders(xlDiagonalUp).LineStyle = xlNone
.Borders(xlEdgeLeft).LineStyle = xlNone
.Borders(xlEdgeTop).LineStyle = xlNone
.Borders(xlEdgeBottom).LineStyle = xlNone
.Borders(xlEdgeRight).LineStyle = xlNone
.Borders(xlInsideVertical).LineStyle = xlNone
.Borders(xlInsideHorizontal).LineStyle = xlNone
End With

'Autofit Column of Sheet Database
ThisWorkbook.Worksheets("Database").Columns("A:F").AutoFit

'***********************************************************************************************************
 
  If OracleProgram = True Then
  MIBordersTable 'Prepare Oracle Spreadsheet
  OracleUpload   'Oracle Database Values created
  crtMISheet     'Excel Database Sheet created
  End If

'***********************************************************************************************************

 'ThisWorkbook.Save
  
Application.Goto Reference:=Worksheets("MIQty").Range("A2")

Set SourceRng = Nothing
Set MySourceTable = Nothing
Set DestRng = Nothing
Set MyDestTable = Nothing

conclusion

Application.Wait (Now + #12:00:04 AM#) 'wait 4 secs

MsgBox ("Operazione Conclusa")

End Sub


'******************************
'Excel Table Codes Archive
'rownum = My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Row
'listr = My_Table.ListRows.Count + 2
'Rowc = My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count - 1
'MyDestTable.ListRows(1).DataBodyRange.EntireRow.Delete
'Debug.Print MyDestTable.ListRows.Count
'Debug.Print MyDestTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
'DestLastRow = MyDestTable.DataBodyRange.Rows(MyDestTable.DataBodyRange.Rows.Count).Count + 1

''''''''''''.ListObjects("List1").Resize Range("$A$1:$E$" & DatabaseLastRow)
'******************************





