Attribute VB_Name = "modWaterfallbar"
Option Explicit
Option Base 1

' Worked in Windows 10 and Excel 2016

Function MyWaterfallBar(sMyPrefixShape As String, rChartRange As Range, rDataRange As Range, rMaxNominator, rMinDenominator As Range) As Double
Dim YHeight As Single, MyStartLeft As Single, MyStartTop As Single, MyCellWidth As Single
Dim MyBalanceHeight As Single, MyBlankOffset As Single
Dim MyBlank As Double, MyPercent As Double, MyFull As Double
Dim MyShape As Shape, MyWS As String
Dim arrData() As Variant
Dim MyStartColumn As Long, MyDataPoint As Long
Dim MyScale As Double, MyBaseStart As Double, MyTopEnd As Double
Dim lDataPoints As Long, MyColour As Long
MyWS = rChartRange.Worksheet.Name
MyScale = rMaxNominator.Value - rMinDenominator.Value

MyStartColumn = rChartRange.Column
MyStartLeft = rChartRange.Left
MyStartTop = rChartRange.Top
YHeight = rChartRange.Cells.Height
MyCellWidth = rChartRange.Cells(1, 1).Width

arrData = rDataRange.Value
lDataPoints = UBound(arrData(), 2)
MyDataPoint = Application.Caller.Column - MyStartColumn + 1

Select Case MyDataPoint

Case Is = 1

MyPercent = arrData(1, 1) / MyScale
MyBlank = 1 - MyPercent
MyFull = MyPercent

MyBlankOffset = CSng(MyBlank * YHeight)
MyBalanceHeight = CSng(MyFull * YHeight)

Case Is > 1

If CDbl(arrData(1, MyDataPoint)) > CDbl(arrData(1, MyDataPoint - 1)) Then

MyBaseStart = CDbl(arrData(1, MyDataPoint - 1)) / MyScale
MyTopEnd = CDbl(arrData(1, MyDataPoint) / MyScale)

MyPercent = MyTopEnd
MyBlank = 1 - MyPercent
MyFull = MyTopEnd - MyBaseStart
MyBlankOffset = CSng(MyBlank * YHeight)
MyBalanceHeight = CSng(MyFull * YHeight)
MyColour = 12566463
Else
MyTopEnd = CDbl(arrData(1, MyDataPoint - 1)) / MyScale
MyBaseStart = CDbl(arrData(1, MyDataPoint) / MyScale)
MyPercent = MyTopEnd
MyBlank = 1 - MyPercent
MyFull = MyTopEnd - MyBaseStart
MyBlankOffset = CSng(MyBlank * YHeight)
MyBalanceHeight = CSng(MyFull * YHeight)
MyColour = 192
End If

End Select

'''' If rDataRange.Value2 = (-9999.9999 / 1000) Then MyPercent = 0 ''''

Set MyShape = ThisWorkbook.Worksheets(MyWS).Shapes(sMyPrefixShape & CStr(MyDataPoint))
With MyShape
If MyDataPoint = 1 Then
.Top = MyStartTop + MyBlankOffset
.Height = MyBalanceHeight
.Width = MyCellWidth * 0.7
.Left = MyStartLeft + ((MyCellWidth * 0.3) / 2)
.Fill.ForeColor.RGB = (MyColour)

Else
.Top = MyStartTop + MyBlankOffset
.Height = MyBalanceHeight
.Width = MyCellWidth * 0.7
.Left = MyStartLeft + ((MyCellWidth * 0.3) / 2) + (MyCellWidth * (MyDataPoint - 1))
.Fill.ForeColor.RGB = MyColour
End If
End With

MyWaterfallBar = MyPercent

Set MyShape = Nothing


End Function

