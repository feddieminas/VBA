Attribute VB_Name = "modXMLImport"
Option Explicit
Option Compare Text
Option Base 1

Sub ClrConts()
QtiesClrConts
PricesClrConts
End Sub

Sub QtiesClrConts()
ThisWorkbook.Worksheets("QQties").Range("B8:CW373").ClearContents
ThisWorkbook.Worksheets("QQties").Range("DA8:GV373").ClearContents
End Sub

Sub RetQties()
Application.ScreenUpdating = False

'Dim MyarrQtoH() As Double
Dim MyRetOccur As Boolean

'Real
ThisWorkbook.Worksheets("QQties").Range("B8:CW373").ClearContents
MyRetOccur = RetValuesXML(ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Value & _
"\" & "FloatQties_" & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value & "_" & _
ThisWorkbook.Worksheets("Dashboard").Range("B13").Value & ".xml", _
"Real", 100, 8, 1, "QQties", "B8")

'With ThisWorkbook
'.Worksheets("HQties").Range("B8:Z373").ClearContents

'MyarrQtoH = QuarterToHour(.Worksheets("QQties").Range("B8:CW373"), _
'.Worksheets("HQties").Range("$B$7:$Z$7"))
'.Worksheets("HQties").Range("B8:Z373").Value = MyarrQtoH
'End With

'Forecast
ThisWorkbook.Worksheets("QQties").Range("DA8:GV373").ClearContents
MyRetOccur = RetValuesXML(ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Value & _
"\" & "FloatQties_" & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value & "_" & _
ThisWorkbook.Worksheets("Dashboard").Range("B13").Value & ".xml", _
"Forecast", 100, 8, 1, "QQties", "DA8")

'With ThisWorkbook
'.Worksheets("HQties").Range("AD8:BB373").ClearContents

'MyarrQtoH = QuarterToHour(.Worksheets("QQties").Range("DA8:GV373"), _
'.Worksheets("HQties").Range("$AD$7:$BB$7"))
'.Worksheets("HQties").Range("AD8:BB373").Value = MyarrQtoH
'End With

'Erase MyarrQtoH()

'Conclusion Msgbox
If MyRetOccur = True Then
MsgBox "Qties Retrieved"
Else
MsgBox "Qties Not Retrieved", vbCritical
End If

Application.ScreenUpdating = True
End Sub

Sub PricesClrConts()
ThisWorkbook.Worksheets("Prices").Range("B6:Z371").ClearContents
ThisWorkbook.Worksheets("Prices").Range("AD6:BB371").ClearContents
End Sub

Sub RetPrices()
Dim MyRetOccur As Boolean

'PUN
ThisWorkbook.Worksheets("Prices").Range("B6:Z371").ClearContents
MyRetOccur = RetValuesXML(ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Value & _
"\" & "FloatPrices_" & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value & ".xml", _
"PUNIndex", 25, 5, 0, "Prices", "B6")

'NORD
ThisWorkbook.Worksheets("Prices").Range("AD6:BB371").ClearContents
MyRetOccur = RetValuesXML(ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Value & _
"\" & "FloatPrices_" & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value & ".xml", _
"NORDIndex", 25, 5, 0, "Prices", "AD6")

'Conclusion Msgbox
If MyRetOccur = True Then
MsgBox "Market Prices Retrieved"
Else
MsgBox "Market Prices Not Retrieved", vbCritical
End If
End Sub

