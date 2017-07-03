Attribute VB_Name = "modACPrep"
Option Explicit
Option Base 1

Sub PrepareZonalSheets()
Dim arrMyZones() As Variant
Dim arrMyZoneCodes() As Variant
Dim arrMyZonalSales() As Variant
Dim arrMyZonalBids() As Variant
Dim arrMyZonalTotal() As Variant
Dim arrMyZonalPrices() As Variant
Dim arrMyZonalOrigQ() As Variant
Dim arrMyZonalOAQBS() As Variant
Dim WB As Workbook
Dim WS As Worksheet, ZoneWS As Worksheet
Dim j As Long, i As Long
Dim ZoneWSName As String

Set WB = ThisWorkbook
Set WS = WB.Worksheets("ImportedData")

arrMyZonalTotal = WS.Range("O32:P55")
arrMyZonalSales = WS.Range("O60:P83")
arrMyZonalBids = WS.Range("O90:P113")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set WS = Nothing
Set WS = WB.Worksheets("OMEL")
arrMyZonalPrices = WS.Range("E2:F25")
ReDim arrMyZonalOrigQ(1 To 24, 1 To 1)
ReDim arrMyZonalOAQBS(1 To 24, 1 To 4)
For i = 1 To 24
'ES + MI + PT
If WS.Range("B39").Offset(0, i).Value >= WS.Range("B39").Offset(1, i).Value Then
arrMyZonalOrigQ(i, 1) = WS.Range("B39").Offset(0, i).Value
Else
arrMyZonalOrigQ(i, 1) = WS.Range("B39").Offset(1, i).Value
End If

'OAQBS
'Compras
'ES
arrMyZonalOAQBS(i, 1) = WS.Range("B31").Offset(0, i).Value
'PT
arrMyZonalOAQBS(i, 2) = WS.Range("B35").Offset(0, i).Value
'Ventas
'ES
arrMyZonalOAQBS(i, 3) = WS.Range("B31").Offset(1, i).Value
'PT
arrMyZonalOAQBS(i, 4) = WS.Range("B35").Offset(1, i).Value
Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

ReDim arrMyZones(1 To 2)
ReDim arrMyZoneCodes(1 To 2)
arrMyZones = Array("ES", "PT")
arrMyZoneCodes = Array(1, 2)

For i = 1 To 2

ZoneWSName = CStr(arrMyZones(i))
Set ZoneWS = WB.Worksheets(ZoneWSName)
ZoneWS.Cells.Clear
ZoneWS.Cells(1, 5).Value = "Hour"
ZoneWS.Cells(2, 5).Value = "Bids"
ZoneWS.Cells(3, 5).Value = "Offers"
ZoneWS.Cells(4, 5).Value = "BidsOffers"
ZoneWS.Cells(5, 5).Value = "GuessPoint"
ZoneWS.Cells(6, 5).Value = "Direction"
ZoneWS.Cells(7, 5).Value = "OPrice"
ZoneWS.Cells(8, 5).Value = "OQuantity"
ZoneWS.Cells(11, 5).Value = "GuessPoint2"
ZoneWS.Cells(13, 5).Value = "OAQ Buy"
ZoneWS.Cells(14, 5).Value = "OAQ Sell"
For j = 1 To 24
ZoneWS.Cells(1, (6 * j)).Value = j
ZoneWS.Cells(2, (6 * j)).Value = (arrMyZonalBids(j, i) * 2) + 1
ZoneWS.Cells(3, (6 * j)).Value = arrMyZonalSales(j, i) * 2
ZoneWS.Cells(4, (6 * j)).Value = (arrMyZonalTotal(j, i) * 2) + 1
ZoneWS.Cells(7, (6 * j)).Value = arrMyZonalPrices(j, i)
ZoneWS.Cells(8, (6 * j)).Value = arrMyZonalOrigQ(j, 1)
'ZoneWS.Cells(8, (6 * j)).Value = arrMyZonalOrigQ(j, i)
ZoneWS.Cells(13, (6 * j)).Value = arrMyZonalOAQBS(j, i)
ZoneWS.Cells(14, (6 * j)).Value = arrMyZonalOAQBS(j, 2 + i)

Next j
Set ZoneWS = Nothing
Next i

Erase arrMyZonalTotal
Erase arrMyZonalSales
Erase arrMyZonalBids
Erase arrMyZonalPrices
Erase arrMyZonalOrigQ
Erase arrMyZonalOAQBS
Set WB = Nothing
Set WS = Nothing

End Sub

Sub CreateACCurves()
Dim WB As Workbook
Dim WS As Worksheet
Dim i As Long, j As Long
Set WB = ThisWorkbook
Set WS = WB.Worksheets("ImportedData")

'ES and PT
For i = 1 To 2
For j = 1 To 24

MyAFFunction MyHour:=j, MyZone:=i, MyType:="V"
CreateACResults MyHour:=j, MyZone:=i, MyType:="V"

MyAFFunction MyHour:=j, MyZone:=i, MyType:="C"
CreateACResults MyHour:=j, MyZone:=i, MyType:="C"

Next j
Next i

ThisWorkbook.Worksheets("ImportedData").AutoFilterMode = False

If ThisWorkbook.Worksheets("PT").Range("F4").Value <= 1 Then
ThisWorkbook.Worksheets("PT").Range("A1").Value = "Original Volume"
End If
End Sub
