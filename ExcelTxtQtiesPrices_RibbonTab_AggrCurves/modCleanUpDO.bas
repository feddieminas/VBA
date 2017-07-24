Attribute VB_Name = "modCleanUpDO"
Option Explicit
Option Base 1

Sub ReconfigureZones()
Dim arrBOTable() As Variant
Dim Lastrow As Long
Dim WB As Workbook
Dim WS As Worksheet
Dim Z As Long

Set WB = ThisWorkbook
Set WS = WB.Worksheets("ImportedData")

With WS
Lastrow = .Cells(.Rows.Count, "C").End(xlUp).Row
arrBOTable = .Range("C4:C" & Lastrow)
End With

'loop for no of the table rows retreive from AccessD
For Z = 1 To UBound(arrBOTable)

'Already looping for the number of Hours
Select Case arrBOTable(Z, 1)

Case "MI"
arrBOTable(Z, 1) = 1
Case "ES"
arrBOTable(Z, 1) = 1
Case "PT"
arrBOTable(Z, 1) = 2

End Select

Next Z

WS.Range("C4").Offset(0, 0).Resize(UBound(arrBOTable), 1).Value = arrBOTable

Erase arrBOTable
Set WB = Nothing
Set WS = Nothing

End Sub
