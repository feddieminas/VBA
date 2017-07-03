Attribute VB_Name = "MainPBO"
Option Explicit
Option Base 1

Sub TableDescOrder(FilteredRecords As Long)
ThisWorkbook.Worksheets("Notepad").Select
ThisWorkbook.Worksheets("Notepad").Range("A4").Select

With ThisWorkbook.Worksheets("Notepad")
    .ListObjects.Add(xlSrcRange, Range("$A$4:$H$" & FilteredRecords), , xlNo).Name = _
        "Table1"
End With

Dim objListObj As ListObject
Set objListObj = ThisWorkbook.Worksheets("Notepad").ListObjects("Table1")

With objListObj
'Price Sort
    .Sort.SortFields.Clear
    .Sort.SortFields.Add _
        Key:=Range("Table1[[#All],[Column7]]"), SortOn:=xlSortOnValues, _
        Order:=xlDescending, DataOption:=xlSortNormal
        .Sort.Header = xlYes
        .Sort.MatchCase = False
        .Sort.Orientation = xlTopToBottom
        .Sort.SortMethod = xlPinYin
        .Sort.Apply
        
        .Unlist
End With
        
Set objListObj = Nothing
End Sub

Sub MySourceDataResort()
Dim SourceRange As Range
Dim TestRange As Range
Dim FilteredRecords As Long

'MI=1 ES=1 PT=2
Dim arrHubs() As Variant
ReDim arrHubs(1 To 2)
arrHubs = Array("ES", "PT")

Dim i As Long
For i = 1 To 2
ThisWorkbook.Worksheets(arrHubs(i)).Cells.Clear

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'A)
MyAFFunction 0, i, "0"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'B)
With ThisWorkbook.Worksheets("ImportedData").AutoFilter.Range
On Error Resume Next
Set TestRange = .Offset(1, 0).Resize(.Rows.Count - 1, 1).SpecialCells(xlCellTypeVisible)
On Error GoTo 0
End With
If TestRange Is Nothing Then
GoTo MyNextHub
Else
ThisWorkbook.Worksheets("NotePad").Cells.Clear
Set SourceRange = ThisWorkbook.Worksheets("ImportedData").AutoFilter.Range
FilteredRecords = SourceRange.Columns(1).SpecialCells(xlCellTypeVisible).Count - 1 'Count Visible rows
If FilteredRecords = 0 Then GoTo MyNextHub

SourceRange.Offset(1, 0).Resize(SourceRange.Rows.Count - 1).Copy _
Destination:=Worksheets("NotePad").Range("A4")
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Table and Descending Order to Count all distinct Bids
'C)
TableDescOrder FilteredRecords + 3
'Since there are no Headers we Delete the Row 4
ThisWorkbook.Worksheets("NotePad").Rows(4).Delete

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'D)
MyFilterBO arrHubs(i)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MyNextHub:
ThisWorkbook.Worksheets("ImportedData").ShowAllData
Next i
Erase arrHubs()

'E)
'Bring Back the Accumulated Results to the ImportedData Sheet
Dim ESRows As Long, PTRows As Long
With ThisWorkbook.Worksheets("ImportedData")
.Range("A4:H65000").Cells.Clear
ESRows = ThisWorkbook.Worksheets("Dashboard").Range("FD15").Value
PTRows = ThisWorkbook.Worksheets("Dashboard").Range("FD16").Value
.Range("A4:H" & ESRows + 3).Value = ThisWorkbook.Worksheets("ES").Range("A4:H" & ESRows + 3).Value
.Range("A4").Offset(ESRows).Resize(PTRows, 8).Value = ThisWorkbook.Worksheets("PT").Range("A4:H" & PTRows + 3).Value

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'F)
'Delete rows with zero Values as Energía Compra/Venta
.Range("A3:H3").AutoFilter
.Range("A3:H3").AutoFilter Field:=6, Criteria1:="0", VisibleDropDown:=False
With .AutoFilter.Range
On Error Resume Next
Application.DisplayAlerts = False
.Offset(1, 0).Resize(.Rows.Count - 1, 1).SpecialCells(xlCellTypeVisible).Delete
On Error GoTo 0
End With

.AutoFilterMode = False
End With
End Sub

Sub MyFilterBO(MyHub As Variant)
Dim WB As Workbook
Dim WS As Worksheet
Set WB = ThisWorkbook
Set WS = WB.Worksheets("Notepad")

'Step 1
Dim Lastrow As Long
Dim arrBO() As Variant
Lastrow = WS.Cells(WS.Rows.Count, "A").End(xlUp).Row
arrBO = WS.Range("A4:H" & Lastrow)


'Step 2
'****Count your Total Bids and Offers / Count that the Bid and Offer Prices are set to correct decimals****'
 Dim counterbid1 As Long, counterbid2 As Long, counterbid3 As Long, counterbid4 As Long, counterbid5 As Long, counterbid6 As Long, _
    counterbid7 As Long, counterbid8 As Long, counterbid9 As Long, counterbid10 As Long, counterbid11 As Long, counterbid12 As Long, _
    counterbid13 As Long, counterbid14 As Long, counterbid15 As Long, counterbid16 As Long, counterbid17 As Long _
    , counterbid18 As Long, counterbid19 As Long, counterbid20 As Long, counterbid21 As Long, counterbid22 As Long, counterbid23 As Long, counterbid24 As Long
    Dim counterOff1 As Long, counterOff2 As Long, counterOff3 As Long, counterOff4 As Long, counterOff5 As Long, counterOff6 As Long, _
    counterOff7 As Long, counterOff8 As Long, counterOff9 As Long, counterOff10 As Long, counterOff11 As Long, counterOff12 As Long, _
    counterOff13 As Long, counterOff14 As Long, counterOff15 As Long, counterOff16 As Long, counterOff17 As Long _
    , counterOff18 As Long, counterOff19 As Long, counterOff20 As Long, counterOff21 As Long, counterOff22 As Long, counterOff23 As Long, counterOff24 As Long

'We insert default value as 1 instead of zero because it will create issues with the ReDim statement after
counterbid1 = 1: counterbid2 = 1: counterbid3 = 1: counterbid4 = 1: counterbid5 = 1: counterbid6 = 1
counterbid7 = 1: counterbid8 = 1: counterbid9 = 1: counterbid10 = 1: counterbid11 = 1: counterbid12 = 1
counterbid13 = 1: counterbid14 = 1: counterbid15 = 1: counterbid16 = 1: counterbid17 = 1: counterbid18 = 1
counterbid19 = 1: counterbid20 = 1: counterbid21 = 1: counterbid22 = 1: counterbid23 = 1: counterbid24 = 1

counterOff1 = 1: counterOff2 = 1: counterOff3 = 1: counterOff4 = 1: counterOff5 = 1: counterOff6 = 1
counterOff7 = 1: counterOff8 = 1: counterOff9 = 1: counterOff10 = 1: counterOff11 = 1: counterOff12 = 1
counterOff13 = 1: counterOff14 = 1: counterOff15 = 1: counterOff16 = 1: counterOff17 = 1: counterOff18 = 1
counterOff19 = 1: counterOff20 = 1: counterOff21 = 1: counterOff22 = 1: counterOff23 = 1: counterOff24 = 1
    
Dim introws As Long
introws = UBound(arrBO, 1)

Dim i As Long
'C=Buy V=Sell
'introws-1 since you start from i=0 instead of i=1
For i = 0 To introws - 1
Select Case arrBO(i + 1, 1)

Case 1
If arrBO(i + 1, 5) = "C" Then
counterbid1 = counterbid1 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff1 = counterOff1 + 1
End If

Case 2
If arrBO(i + 1, 5) = "C" Then
counterbid2 = counterbid2 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff2 = counterOff2 + 1
End If

Case 3
If arrBO(i + 1, 5) = "C" Then
counterbid3 = counterbid3 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff3 = counterOff3 + 1
End If

Case 4
If arrBO(i + 1, 5) = "C" Then
counterbid4 = counterbid4 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff4 = counterOff4 + 1
End If

Case 5
If arrBO(i + 1, 5) = "C" Then
counterbid5 = counterbid5 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff5 = counterOff5 + 1
End If

Case 6
If arrBO(i + 1, 5) = "C" Then
counterbid6 = counterbid6 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff6 = counterOff6 + 1
End If

Case 7
If arrBO(i + 1, 5) = "C" Then
counterbid7 = counterbid7 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff7 = counterOff7 + 1
End If

Case 8
If arrBO(i + 1, 5) = "C" Then
counterbid8 = counterbid8 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff8 = counterOff8 + 1
End If

Case 9
If arrBO(i + 1, 5) = "C" Then
counterbid9 = counterbid9 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff9 = counterOff9 + 1
End If

Case 10
If arrBO(i + 1, 5) = "C" Then
counterbid10 = counterbid10 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff10 = counterOff10 + 1
End If

Case 11
If arrBO(i + 1, 5) = "C" Then
counterbid11 = counterbid11 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff11 = counterOff11 + 1
End If

Case 12
If arrBO(i + 1, 5) = "C" Then
counterbid12 = counterbid12 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff12 = counterOff12 + 1
End If

Case 13
If arrBO(i + 1, 5) = "C" Then
counterbid13 = counterbid13 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff13 = counterOff13 + 1
End If

Case 14
If arrBO(i + 1, 5) = "C" Then
counterbid14 = counterbid14 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff14 = counterOff14 + 1
End If

Case 15
If arrBO(i + 1, 5) = "C" Then
counterbid15 = counterbid15 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff15 = counterOff15 + 1
End If

Case 16
If arrBO(i + 1, 5) = "C" Then
counterbid16 = counterbid16 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff16 = counterOff16 + 1
End If

Case 17
If arrBO(i + 1, 5) = "C" Then
counterbid17 = counterbid17 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff17 = counterOff17 + 1
End If

Case 18
If arrBO(i + 1, 5) = "C" Then
counterbid18 = counterbid18 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff18 = counterOff18 + 1
End If

Case 19
If arrBO(i + 1, 5) = "C" Then
counterbid19 = counterbid19 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff19 = counterOff19 + 1
End If

Case 20
If arrBO(i + 1, 5) = "C" Then
counterbid20 = counterbid20 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff20 = counterOff20 + 1
End If

Case 21
If arrBO(i + 1, 5) = "C" Then
counterbid21 = counterbid21 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff21 = counterOff21 + 1
End If

Case 22
If arrBO(i + 1, 5) = "C" Then
counterbid22 = counterbid22 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff22 = counterOff22 + 1
End If

Case 23
If arrBO(i + 1, 5) = "C" Then
counterbid23 = counterbid23 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff23 = counterOff23 + 1
End If

Case 24
If arrBO(i + 1, 5) = "C" Then
counterbid24 = counterbid24 + 1
ElseIf arrBO(i + 1, 5) = "V" Then
counterOff24 = counterOff24 + 1
End If

End Select

'****Verify That Price Bids and Offer are set to correct decimals****'
arrBO(i + 1, 7) = CDbl(arrBO(i + 1, 7) * 10000)
arrBO(i + 1, 7) = Int(arrBO(i + 1, 7))
arrBO(i + 1, 7) = CDbl(arrBO(i + 1, 7) / 10000)

Next i 'Loop for rows

'Step 3
'Create array for Public Bids and offers
Dim arrBid1DataTmp() As Double, arrOff1DataTmp() As Double, arrBid2DataTmp() As Double, arrOff2DataTmp() As Double
Dim arrBid3DataTmp() As Double, arrOff3DataTmp() As Double, arrBid4DataTmp() As Double, arrOff4DataTmp() As Double
Dim arrBid5DataTmp() As Double, arrOff5DataTmp() As Double, arrBid6DataTmp() As Double, arrOff6DataTmp() As Double
Dim arrBid7DataTmp() As Double, arrOff7DataTmp() As Double, arrBid8DataTmp() As Double, arrOff8DataTmp() As Double
Dim arrBid9DataTmp() As Double, arrOff9DataTmp() As Double, arrBid10DataTmp() As Double, arrOff10DataTmp() As Double
Dim arrBid11DataTmp() As Double, arrOff11DataTmp() As Double, arrBid12DataTmp() As Double, arrOff12DataTmp() As Double
Dim arrBid13DataTmp() As Double, arrOff13DataTmp() As Double, arrBid14DataTmp() As Double, arrOff14DataTmp() As Double
Dim arrBid15DataTmp() As Double, arrOff15DataTmp() As Double, arrBid16DataTmp() As Double, arrOff16DataTmp() As Double
Dim arrBid17DataTmp() As Double, arrOff17DataTmp() As Double, arrBid18DataTmp() As Double, arrOff18DataTmp() As Double
Dim arrBid19DataTmp() As Double, arrOff19DataTmp() As Double, arrBid20DataTmp() As Double, arrOff20DataTmp() As Double
Dim arrBid21DataTmp() As Double, arrOff21DataTmp() As Double, arrBid22DataTmp() As Double, arrOff22DataTmp() As Double
Dim arrBid23DataTmp() As Double, arrOff23DataTmp() As Double, arrBid24DataTmp() As Double, arrOff24DataTmp() As Double
'Redefine your Bid and Off arrays
    ReDim arrBid1DataTmp(1 To counterbid1, 1 To 2)
    ReDim arrOff1DataTmp(1 To counterOff1, 1 To 2)
    ReDim arrBid2DataTmp(1 To counterbid2, 1 To 2)
    ReDim arrOff2DataTmp(1 To counterOff2, 1 To 2)
    ReDim arrBid3DataTmp(1 To counterbid3, 1 To 2)
    ReDim arrOff3DataTmp(1 To counterOff3, 1 To 2)
    ReDim arrBid4DataTmp(1 To counterbid4, 1 To 2)
    ReDim arrOff4DataTmp(1 To counterOff4, 1 To 2)
    ReDim arrBid5DataTmp(1 To counterbid5, 1 To 2)
    ReDim arrOff5DataTmp(1 To counterOff5, 1 To 2)
    ReDim arrBid6DataTmp(1 To counterbid6, 1 To 2)
    ReDim arrOff6DataTmp(1 To counterOff6, 1 To 2)
    ReDim arrBid7DataTmp(1 To counterbid7, 1 To 2)
    ReDim arrOff7DataTmp(1 To counterOff7, 1 To 2)
    ReDim arrBid8DataTmp(1 To counterbid8, 1 To 2)
    ReDim arrOff8DataTmp(1 To counterOff8, 1 To 2)
    ReDim arrBid9DataTmp(1 To counterbid9, 1 To 2)
    ReDim arrOff9DataTmp(1 To counterOff9, 1 To 2)
    ReDim arrBid10DataTmp(1 To counterbid10, 1 To 2)
    ReDim arrOff10DataTmp(1 To counterOff10, 1 To 2)
    ReDim arrBid11DataTmp(1 To counterbid11, 1 To 2)
    ReDim arrOff11DataTmp(1 To counterOff11, 1 To 2)
    ReDim arrBid12DataTmp(1 To counterbid12, 1 To 2)
    ReDim arrOff12DataTmp(1 To counterOff12, 1 To 2)
    ReDim arrBid13DataTmp(1 To counterbid13, 1 To 2)
    ReDim arrOff13DataTmp(1 To counterOff13, 1 To 2)
    ReDim arrBid14DataTmp(1 To counterbid14, 1 To 2)
    ReDim arrOff14DataTmp(1 To counterOff14, 1 To 2)
    ReDim arrBid15DataTmp(1 To counterbid15, 1 To 2)
    ReDim arrOff15DataTmp(1 To counterOff15, 1 To 2)
    ReDim arrBid16DataTmp(1 To counterbid16, 1 To 2)
    ReDim arrOff16DataTmp(1 To counterOff16, 1 To 2)
    ReDim arrBid17DataTmp(1 To counterbid17, 1 To 2)
    ReDim arrOff17DataTmp(1 To counterOff17, 1 To 2)
    ReDim arrBid18DataTmp(1 To counterbid18, 1 To 2)
    ReDim arrOff18DataTmp(1 To counterOff18, 1 To 2)
    ReDim arrBid19DataTmp(1 To counterbid19, 1 To 2)
    ReDim arrOff19DataTmp(1 To counterOff19, 1 To 2)
    ReDim arrBid20DataTmp(1 To counterbid20, 1 To 2)
    ReDim arrOff20DataTmp(1 To counterOff20, 1 To 2)
    ReDim arrBid21DataTmp(1 To counterbid21, 1 To 2)
    ReDim arrOff21DataTmp(1 To counterOff21, 1 To 2)
    ReDim arrBid22DataTmp(1 To counterbid22, 1 To 2)
    ReDim arrOff22DataTmp(1 To counterOff22, 1 To 2)
    ReDim arrBid23DataTmp(1 To counterbid23, 1 To 2)
    ReDim arrOff23DataTmp(1 To counterOff23, 1 To 2)
    ReDim arrBid24DataTmp(1 To counterbid24, 1 To 2)
    ReDim arrOff24DataTmp(1 To counterOff24, 1 To 2)

counterbid1 = 0: counterbid2 = 0: counterbid3 = 0: counterbid4 = 0: counterbid5 = 0: counterbid6 = 0
counterbid7 = 0: counterbid8 = 0: counterbid9 = 0: counterbid10 = 0: counterbid11 = 0: counterbid12 = 0
counterbid13 = 0: counterbid14 = 0: counterbid15 = 0: counterbid16 = 0: counterbid17 = 0: counterbid18 = 0
counterbid19 = 0: counterbid20 = 0: counterbid21 = 0: counterbid22 = 0: counterbid23 = 0: counterbid24 = 0

counterOff1 = 0: counterOff2 = 0: counterOff3 = 0: counterOff4 = 0: counterOff5 = 0: counterOff6 = 0
counterOff7 = 0: counterOff8 = 0: counterOff9 = 0: counterOff10 = 0: counterOff11 = 0: counterOff12 = 0
counterOff13 = 0: counterOff14 = 0: counterOff15 = 0: counterOff16 = 0: counterOff17 = 0: counterOff18 = 0
counterOff19 = 0: counterOff20 = 0: counterOff21 = 0: counterOff22 = 0: counterOff23 = 0: counterOff24 = 0

'Insert values in the array of Bids and Offers. Convert it to DoubleValues
'introws-1 since you start from i=0 instead of i=1
Dim j As Long
For i = 0 To introws - 1

Select Case arrBO(i + 1, 1) 'Loop for Hours
    
    Case 1
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid1DataTmp(counterbid1 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid1 = counterbid1 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff1DataTmp(counterOff1 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff1 = counterOff1 + 1
    End If
    
    Case 2
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid2DataTmp(counterbid2 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid2 = counterbid2 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff2DataTmp(counterOff2 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff2 = counterOff2 + 1
    End If
    
    Case 3
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid3DataTmp(counterbid3 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid3 = counterbid3 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff3DataTmp(counterOff3 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff3 = counterOff3 + 1
    End If
    
    Case 4
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid4DataTmp(counterbid4 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid4 = counterbid4 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff4DataTmp(counterOff4 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff4 = counterOff4 + 1
    End If
    
    Case 5
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid5DataTmp(counterbid5 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid5 = counterbid5 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff5DataTmp(counterOff5 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff5 = counterOff5 + 1
    End If
    
    Case 6
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid6DataTmp(counterbid6 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid6 = counterbid6 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff6DataTmp(counterOff6 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff6 = counterOff6 + 1
    End If
    
    Case 7
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid7DataTmp(counterbid7 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid7 = counterbid7 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff7DataTmp(counterOff7 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff7 = counterOff7 + 1
    End If
    
    Case 8
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid8DataTmp(counterbid8 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid8 = counterbid8 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff8DataTmp(counterOff8 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff8 = counterOff8 + 1
    End If
    
    Case 9
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid9DataTmp(counterbid9 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid9 = counterbid9 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff9DataTmp(counterOff9 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff9 = counterOff9 + 1
    End If
    
    Case 10
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid10DataTmp(counterbid10 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid10 = counterbid10 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff10DataTmp(counterOff10 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff10 = counterOff10 + 1
    End If
    
    Case 11
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid11DataTmp(counterbid11 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid11 = counterbid11 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff11DataTmp(counterOff11 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff11 = counterOff11 + 1
    End If
    
    Case 12
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid12DataTmp(counterbid12 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid12 = counterbid12 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff12DataTmp(counterOff12 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff12 = counterOff12 + 1
    End If
    
    Case 13
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid13DataTmp(counterbid13 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid13 = counterbid13 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff13DataTmp(counterOff13 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff13 = counterOff13 + 1
    End If
    
    Case 14
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid14DataTmp(counterbid14 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid14 = counterbid14 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff14DataTmp(counterOff14 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff14 = counterOff14 + 1
    End If
    
    Case 15
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid15DataTmp(counterbid15 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid15 = counterbid15 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff15DataTmp(counterOff15 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff15 = counterOff15 + 1
    End If
    
    Case 16
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid16DataTmp(counterbid16 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid16 = counterbid16 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff16DataTmp(counterOff16 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff16 = counterOff16 + 1
    End If
    
    Case 17
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid17DataTmp(counterbid17 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid17 = counterbid17 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff17DataTmp(counterOff17 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff17 = counterOff17 + 1
    End If
    
    Case 18
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid18DataTmp(counterbid18 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid18 = counterbid18 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff18DataTmp(counterOff18 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff18 = counterOff18 + 1
    End If
    
    Case 19
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid19DataTmp(counterbid19 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid19 = counterbid19 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff19DataTmp(counterOff19 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff19 = counterOff19 + 1
    End If
    
    Case 20
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid20DataTmp(counterbid20 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid20 = counterbid20 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff20DataTmp(counterOff20 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff20 = counterOff20 + 1
    End If
    
    Case 21
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid21DataTmp(counterbid21 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid21 = counterbid21 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff21DataTmp(counterOff21 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff21 = counterOff21 + 1
    End If
    
    Case 22
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid22DataTmp(counterbid22 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid22 = counterbid22 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff22DataTmp(counterOff22 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff22 = counterOff22 + 1
    End If
    
    Case 23
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid23DataTmp(counterbid23 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid23 = counterbid23 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff23DataTmp(counterOff23 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff23 = counterOff23 + 1
    End If
    
    Case 24
    If arrBO(i + 1, 5) = "C" Then
    For j = 0 To 1
    arrBid24DataTmp(counterbid24 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterbid24 = counterbid24 + 1
    ElseIf arrBO(i + 1, 5) = "V" Then
    For j = 0 To 1
    arrOff24DataTmp(counterOff24 + 1, j + 1) = CDbl(arrBO(i + 1, j + 6))
    Next j
    counterOff24 = counterOff24 + 1
    End If
    
Case Else
End Select

Next i
'EraseTotalData
Erase arrBO()


'****Sort Bid and Offer Values in a descending order****'
'Step 4
'MySortDescending PublicBids
    arrBid1DataTmp = MySortDesc(arrBid1DataTmp)
    arrBid2DataTmp = MySortDesc(arrBid2DataTmp)
    arrBid3DataTmp = MySortDesc(arrBid3DataTmp)
    arrBid4DataTmp = MySortDesc(arrBid4DataTmp)
    arrBid5DataTmp = MySortDesc(arrBid5DataTmp)
    arrBid6DataTmp = MySortDesc(arrBid6DataTmp)
    arrBid7DataTmp = MySortDesc(arrBid7DataTmp)
    arrBid8DataTmp = MySortDesc(arrBid8DataTmp)
    arrBid9DataTmp = MySortDesc(arrBid9DataTmp)
    arrBid10DataTmp = MySortDesc(arrBid10DataTmp)
    arrBid11DataTmp = MySortDesc(arrBid11DataTmp)
    arrBid12DataTmp = MySortDesc(arrBid12DataTmp)
    arrBid13DataTmp = MySortDesc(arrBid13DataTmp)
    arrBid14DataTmp = MySortDesc(arrBid14DataTmp)
    arrBid15DataTmp = MySortDesc(arrBid15DataTmp)
    arrBid16DataTmp = MySortDesc(arrBid16DataTmp)
    arrBid17DataTmp = MySortDesc(arrBid17DataTmp)
    arrBid18DataTmp = MySortDesc(arrBid18DataTmp)
    arrBid19DataTmp = MySortDesc(arrBid19DataTmp)
    arrBid20DataTmp = MySortDesc(arrBid20DataTmp)
    arrBid21DataTmp = MySortDesc(arrBid21DataTmp)
    arrBid22DataTmp = MySortDesc(arrBid22DataTmp)
    arrBid23DataTmp = MySortDesc(arrBid23DataTmp)
    arrBid24DataTmp = MySortDesc(arrBid24DataTmp)

'Step 5
'PB Layout
    PBLayout MyHub:=MyHub, arrBid1DataTmp:=arrBid1DataTmp, arrBid2DataTmp:=arrBid2DataTmp, arrBid3DataTmp:=arrBid3DataTmp _
    , arrBid4DataTmp:=arrBid4DataTmp, arrBid5DataTmp:=arrBid5DataTmp, arrBid6DataTmp:=arrBid6DataTmp _
    , arrBid7DataTmp:=arrBid7DataTmp, arrBid8DataTmp:=arrBid8DataTmp, arrBid9DataTmp:=arrBid9DataTmp _
    , arrBid10DataTmp:=arrBid10DataTmp, arrBid11DataTmp:=arrBid11DataTmp, arrBid12DataTmp:=arrBid12DataTmp _
    , arrBid13DataTmp:=arrBid13DataTmp, arrBid14DataTmp:=arrBid14DataTmp, arrBid15DataTmp:=arrBid15DataTmp _
    , arrBid16DataTmp:=arrBid16DataTmp, arrBid17DataTmp:=arrBid17DataTmp, arrBid18DataTmp:=arrBid18DataTmp _
    , arrBid19DataTmp:=arrBid19DataTmp, arrBid20DataTmp:=arrBid20DataTmp, arrBid21DataTmp:=arrBid21DataTmp _
    , arrBid22DataTmp:=arrBid22DataTmp, arrBid23DataTmp:=arrBid23DataTmp, arrBid24DataTmp:=arrBid24DataTmp

'Step 6
'MySortDescending PublicOffers
    arrOff1DataTmp = MySortDesc(arrOff1DataTmp)
    arrOff2DataTmp = MySortDesc(arrOff2DataTmp)
    arrOff3DataTmp = MySortDesc(arrOff3DataTmp)
    arrOff4DataTmp = MySortDesc(arrOff4DataTmp)
    arrOff5DataTmp = MySortDesc(arrOff5DataTmp)
    arrOff6DataTmp = MySortDesc(arrOff6DataTmp)
    arrOff7DataTmp = MySortDesc(arrOff7DataTmp)
    arrOff8DataTmp = MySortDesc(arrOff8DataTmp)
    arrOff9DataTmp = MySortDesc(arrOff9DataTmp)
    arrOff10DataTmp = MySortDesc(arrOff10DataTmp)
    arrOff11DataTmp = MySortDesc(arrOff11DataTmp)
    arrOff12DataTmp = MySortDesc(arrOff12DataTmp)
    arrOff13DataTmp = MySortDesc(arrOff13DataTmp)
    arrOff14DataTmp = MySortDesc(arrOff14DataTmp)
    arrOff15DataTmp = MySortDesc(arrOff15DataTmp)
    arrOff16DataTmp = MySortDesc(arrOff16DataTmp)
    arrOff17DataTmp = MySortDesc(arrOff17DataTmp)
    arrOff18DataTmp = MySortDesc(arrOff18DataTmp)
    arrOff19DataTmp = MySortDesc(arrOff19DataTmp)
    arrOff20DataTmp = MySortDesc(arrOff20DataTmp)
    arrOff21DataTmp = MySortDesc(arrOff21DataTmp)
    arrOff22DataTmp = MySortDesc(arrOff22DataTmp)
    arrOff23DataTmp = MySortDesc(arrOff23DataTmp)
    arrOff24DataTmp = MySortDesc(arrOff24DataTmp)
    
'Step 7
'PO Layout
     POLayout MyHub:=MyHub, arrOff1DataTmp:=arrOff1DataTmp, arrOff2DataTmp:=arrOff2DataTmp, arrOff3DataTmp:=arrOff3DataTmp _
    , arrOff4DataTmp:=arrOff4DataTmp, arrOff5DataTmp:=arrOff5DataTmp, arrOff6DataTmp:=arrOff6DataTmp _
    , arrOff7DataTmp:=arrOff7DataTmp, arrOff8DataTmp:=arrOff8DataTmp, arrOff9DataTmp:=arrOff9DataTmp _
    , arrOff10DataTmp:=arrOff10DataTmp, arrOff11DataTmp:=arrOff11DataTmp, arrOff12DataTmp:=arrOff12DataTmp _
    , arrOff13DataTmp:=arrOff13DataTmp, arrOff14DataTmp:=arrOff14DataTmp, arrOff15DataTmp:=arrOff15DataTmp _
    , arrOff16DataTmp:=arrOff16DataTmp, arrOff17DataTmp:=arrOff17DataTmp, arrOff18DataTmp:=arrOff18DataTmp _
    , arrOff19DataTmp:=arrOff19DataTmp, arrOff20DataTmp:=arrOff20DataTmp, arrOff21DataTmp:=arrOff21DataTmp _
    , arrOff22DataTmp:=arrOff22DataTmp, arrOff23DataTmp:=arrOff23DataTmp, arrOff24DataTmp:=arrOff24DataTmp


End Sub
