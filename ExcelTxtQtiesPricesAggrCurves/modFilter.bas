Attribute VB_Name = "modFilter"
Option Base 1
Option Explicit

Sub MyAFFunction(MyHour As Long, MyZone As Long, MyType As String)
With Worksheets("ImportedData")
    .AutoFilterMode = False
        With .Range("A3:H3")
        .AutoFilter
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Hour. We indicate Hour 0 as No need to Filter the Criteria
         If MyHour = 0 Then
        .AutoFilter Field:=1, VisibleDropDown:=False
         Else
        .AutoFilter Field:=1, Criteria1:=MyHour, VisibleDropDown:=False
         End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        .AutoFilter Field:=2, VisibleDropDown:=False
        .AutoFilter Field:=3, Criteria1:=MyZone, VisibleDropDown:=False
        .AutoFilter Field:=4, VisibleDropDown:=False
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Type. We indicate Type 0 as No need to Filter the Criteria
         If MyType = "0" Then
        .AutoFilter Field:=5, VisibleDropDown:=False
         Else
        .AutoFilter Field:=5, Criteria1:=MyType, VisibleDropDown:=False
         End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        .AutoFilter Field:=6, VisibleDropDown:=False
        .AutoFilter Field:=7, VisibleDropDown:=False
        .AutoFilter Field:=8, Criteria1:="O", VisibleDropDown:=False
        End With
End With
End Sub

Sub CreateACResults(MyHour As Long, MyZone As Long, MyType As String)

Dim SourceRange As Range
Dim TestRange As Range
Dim FilteredRecords As Long

With ThisWorkbook.Worksheets("ImportedData").AutoFilter.Range
 On Error Resume Next
 Set TestRange = .Offset(1, 0).Resize(.Rows.Count - 1, 1) _
       .SpecialCells(xlCellTypeVisible)
 On Error GoTo 0
End With
If TestRange Is Nothing Then
'MsgBox "TARA: Error has occured in the filtering process."
Else
   Worksheets("NotePad").Cells.Clear
   Set SourceRange = ThisWorkbook.Worksheets("ImportedData").AutoFilter.Range
   FilteredRecords = SourceRange.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1 'Count Visible rows
   
   SourceRange.Offset(1, 0).Resize(SourceRange.Rows.Count - 1).Copy _
     Destination:=Worksheets("NotePad").Range("A1")
End If

  ThisWorkbook.Worksheets("ImportedData").ShowAllData

  'No ACCurve per Hour and per Zone
  If FilteredRecords = 0 Then GoTo nextarrayfilter
  
  Dim arrTemp() As Variant
  arrTemp() = ThisWorkbook.Worksheets("NotePad").Range("A1:G" & FilteredRecords).Value
  ThisWorkbook.Worksheets("NotePad").Range("P1").Resize(UBound(arrTemp(), 1), UBound(arrTemp(), 2)) = arrTemp
    If MyType = "C" Then
    MyCreateBidAC MyHourA:=MyHour, MyZoneA:=MyZone, MyDataA:=arrTemp
    Else
    MyCreateOfferAC MyHourA:=MyHour, MyZoneA:=MyZone, MyDataA:=arrTemp
    End If

nextarrayfilter:
Erase arrTemp
Set SourceRange = Nothing
Set TestRange = Nothing
End Sub


Sub MyCreateOfferAC(MyHourA As Long, MyZoneA As Long, MyDataA() As Variant)
Dim BidsNo As Long, NewSize As Long, i As Long
Dim arrAC() As Variant
Dim DestWS As Worksheet, DestWSName As String

BidsNo = UBound(MyDataA)
NewSize = (2 * BidsNo) + 2

ReDim arrAC(1 To NewSize, 1 To 3)
arrAC(1, 1) = 0
arrAC(2, 1) = 0

'Quantity
For i = 2 To BidsNo + 1
arrAC((i * 2) - 1, 1) = CDbl(MyDataA(BidsNo + 2 - i, 6))
arrAC((i * 2), 1) = CDbl(MyDataA(BidsNo + 2 - i, 6))
Next i


For i = 4 To NewSize
arrAC(i, 1) = arrAC(i, 1) + arrAC(i - 2, 1)
Next i


'Prices
For i = 1 To BidsNo
arrAC((i * 2), 3) = CDbl(MyDataA(BidsNo + 1 - i, 7))
arrAC((i * 2) + 1, 3) = CDbl(MyDataA(BidsNo + 1 - i, 7))
arrAC((i * 2), 2) = "Sell"
arrAC((i * 2) + 1, 2) = "Sell"
Next i
arrAC(1, 3) = -500
arrAC(NewSize, 3) = 3000
arrAC(1, 2) = "Sell"
arrAC(NewSize, 2) = "Sell"


Dim arrMyZones As Variant
ReDim arrMyZones(1 To 2)
arrMyZones = Array("ES", "PT")
DestWSName = arrMyZones(MyZoneA)
Set DestWS = ThisWorkbook.Worksheets(DestWSName)
DestWS.Cells(2, 1).Offset(0, (MyHourA - 1) * 6).Resize(NewSize, 3) = arrAC
'***********************************************************************************

With DestWS
'.Range(.Cells(2, 3 + ((MyHourA - 1) * 6)), .Cells(MPAdd + 2 + .Cells(3, 6 + (6 * (MyHourA - 1))).Value - 1, 3 + ((MyHourA - 1) * 6))).Select
.Cells(5, 6 + (6 * (MyHourA - 1))).Value = Application.WorksheetFunction.Match(.Cells(7, 6 + ((MyHourA - 1) * 6)).Value, .Range(.Cells(2, 3 + ((MyHourA - 1) * 6)), .Cells(0 + 2 + .Cells(3, 6 + (6 * (MyHourA - 1))).Value - 1, 3 + ((MyHourA - 1) * 6))), 1)
End With
'***********************************************************************************

Set DestWS = Nothing
Erase arrAC
Erase arrMyZones
End Sub

Sub MyCreateBidAC(MyHourA As Long, MyZoneA As Long, MyDataA() As Variant)
Dim BidsNo As Long, NewSize As Long, i As Long
Dim arrAC() As Variant
Dim DestWSName As String
Dim DestWS As Worksheet

BidsNo = UBound(MyDataA)
NewSize = (2 * BidsNo) + 2

ReDim arrAC(1 To NewSize, 1 To 3)

arrAC(1, 1) = CDbl(MyDataA(BidsNo, 6))
arrAC(2, 1) = CDbl(MyDataA(BidsNo, 6))

For i = 2 To BidsNo
arrAC((i * 2) - 1, 1) = CDbl(MyDataA(BidsNo - i + 1, 6))
arrAC((i * 2), 1) = CDbl(MyDataA(BidsNo - i + 1, 6))
Next i
arrAC(NewSize - 1, 1) = 0
arrAC(NewSize, 1) = 0

For i = 2 To NewSize - 1
arrAC(NewSize - i, 1) = arrAC(NewSize - i + 2, 1) + arrAC(NewSize - i, 1)
Next i

For i = 1 To BidsNo
arrAC((i * 2), 3) = CDbl(MyDataA(BidsNo + 1 - i, 7))
arrAC((i * 2) + 1, 3) = CDbl(MyDataA(BidsNo + 1 - i, 7))
arrAC((i * 2), 2) = "Purchase"
arrAC((i * 2) + 1, 2) = "Purchase"
Next i
arrAC(1, 3) = -500
arrAC(NewSize, 3) = 3000
arrAC(1, 2) = "Purchase"
arrAC(NewSize, 2) = "Purchase"

Dim arrMyZones As Variant
ReDim arrMyZones(1 To 2)
arrMyZones = Array("ES", "PT")
DestWSName = arrMyZones(MyZoneA)
Set DestWS = ThisWorkbook.Worksheets(DestWSName)
Dim NoofSales As Long
NoofSales = (DestWS.Cells(3, (MyHourA * 6)).Value) + 2
DestWS.Cells(2, 1).Offset(NoofSales, (MyHourA - 1) * 6).Resize(NewSize, 3) = arrAC
'***********************************************************************************

With DestWS
'.Range(.Cells(2 + 2 + .Cells(3, 6 + (6 * (MyHourA - 1))).Value, 3 + ((MyHourA - 1) * 6)), .Cells(3 + 1 + .Cells(4, 6 + (6 * (MyHourA - 1))).Value, 3 + ((MyHourA - 1) * 6))).Select
.Cells(11, 6 + (6 * (MyHourA - 1))).Value = Application.WorksheetFunction.Match(.Cells(7, 6 + ((MyHourA - 1) * 6)).Value, .Range(.Cells(2 + 2 + .Cells(3, 6 + (6 * (MyHourA - 1))).Value, 3 + ((MyHourA - 1) * 6)), .Cells(3 + 1 + .Cells(4, 6 + (6 * (MyHourA - 1))).Value, 3 + ((MyHourA - 1) * 6))), 1)
End With
'***********************************************************************************

Erase arrAC
Erase arrMyZones
Set DestWS = Nothing
End Sub

Sub MyAFClose()
With Worksheets("ImportedData")
    .AutoFilterMode = False
End With
End Sub
