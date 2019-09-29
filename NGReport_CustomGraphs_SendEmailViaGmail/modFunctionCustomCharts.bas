Attribute VB_Name = "modFunctionCustomCharts"
Option Explicit

' Worked in Windows 10 and Excel 2016

Function MyBar(sMyShape As String, rRange As Range, rNominator As Range, rDenominator As Range) As Double
Dim YHeight As Single, MyStartLeft As Single, MyStartTop As Single, MyCellWidth As Single
Dim MyBalanceHeight As Single, MyBlankOffset As Single
Dim MyBlank As Double, MyPercent As Double, MyFull As Double
Dim MyShape As Shape
Dim MyWS As String
MyWS = rRange.Worksheet.Name
MyStartLeft = rRange.Left
MyStartTop = rRange.Top
YHeight = rRange.Cells.Height
MyCellWidth = rRange.Cells(1, 1).Width
MyPercent = rNominator.Value2 / rDenominator.Value2
If MyPercent > 1 Then MyPercent = 1
If rNominator.Value2 = (-9999.9999 / 1000) Then MyPercent = 0
MyBlank = 1 - MyPercent
MyFull = MyPercent
MyBlankOffset = CSng(MyBlank * YHeight)
MyBalanceHeight = CSng(MyFull * YHeight)
Set MyShape = ThisWorkbook.Worksheets(MyWS).Shapes(sMyShape)
With MyShape
.Top = MyStartTop + MyBlankOffset
.Height = MyBalanceHeight
.Width = MyCellWidth * 0.7
.Left = MyStartLeft + ((MyCellWidth * 0.3) / 2)
End With
MyBar = MyPercent
Set MyShape = Nothing
End Function

Function MyDiffLine(MyType As Long, rRange As Range, rSNominator As Range, rSDenominator As Range, rENominator As Range, rEDenominator As Range, rMax As Range) As Double
Dim YHeight As Single, MyStartLeft As Single, MyStartTop As Single, MyCellWidth As Single, MyEndLeft As Single, MyNextCellWidth As Single
Dim MyBalanceHeight As Single, MyBlankOffset As Single, MyNameLeft As Single, MyNameTop As Single, MyEndTop As Single
Dim MyStartBlank As Double, MyStartPercent As Double, MyEndPercent As Double, MyEndBlank As Double, MyAxisScale As Double
Dim MyShape As Shape, SPP As Shape
Dim MyWS As String, MyConName As String

'MyType represents which time series for colour coding of line
'rRange reflects the range of the line - vertical axis
'rSNominator is the diff for the start of the line
'rsDenominator is the max of the vertical scale so axis range in cell ai31
'rENominator is the diff for the end of the line
'rEDenominator is the max of the vertical scale so axis range in cell ai31
'rMax is the axis range value

MyWS = rRange.Worksheet.Name
MyNameLeft = rRange.Left
MyNameTop = rRange.Top
MyAxisScale = (rMax.Value2) / 2

Select Case MyType
Case Is = 1
MyConName = "ConD1_" & MyNameLeft & "_" & MyNameTop
Case Is = 2
MyConName = "ConD2_" & MyNameLeft & "_" & MyNameTop
Case Is = 3
MyConName = "ConD3_" & MyNameLeft & "_" & MyNameTop
Case Is = 4
MyConName = "ConD4_" & MyNameLeft & "_" & MyNameTop
Case Is = 5
MyConName = "ConD5_" & MyNameLeft & "_" & MyNameTop
End Select

On Error Resume Next
Set MyShape = ThisWorkbook.Worksheets(MyWS).Shapes(MyConName)
MyShape.Delete

MyStartLeft = rRange.Left
MyStartTop = rRange.Top
MyEndTop = rRange.Top

YHeight = rRange.Cells.Height
'Calculate Cell Widths - we want midpoint of cell to midpoint of next cell to right
MyCellWidth = rRange.Cells(1, 1).Width
MyNextCellWidth = rRange.Offset(0, 1).Cells(1, 1).Width

MyStartLeft = MyStartLeft + ((MyCellWidth) / 2)
MyEndLeft = MyStartLeft + (MyCellWidth / 2) + (MyNextCellWidth / 2)

'Calculate Height
MyStartPercent = (rSNominator.Value2 + MyAxisScale) / rSDenominator.Value2
MyEndPercent = (rENominator.Value2 + MyAxisScale) / rEDenominator.Value2

If rSNominator.Value2 = (-9999.9999 / 1000) Then
MyStartPercent = 0.5
MyEndPercent = 0.5
End If

If rSNominator.Value2 <> (-9999.9999 / 1000) And rENominator.Value2 = (-9999.9999 / 1000) Then
MyEndPercent = MyStartPercent
End If


MyStartBlank = 1 - MyStartPercent
MyEndBlank = 1 - MyEndPercent

MyBlankOffset = CSng(MyStartBlank * YHeight)
MyStartTop = MyStartTop + MyBlankOffset

MyBlankOffset = CSng(MyEndBlank * YHeight)
MyEndTop = MyEndTop + MyBlankOffset

Set SPP = ThisWorkbook.Worksheets(MyWS).Shapes.AddConnector(msoConnectorStraight, MyStartLeft, MyStartTop, MyEndLeft, MyEndTop)
With SPP
  .Name = MyConName
If MyStartPercent = 0 And MyEndPercent = 0 Then
  .Line.ForeColor.RGB = RGB(0, 0, 0)
  .Line.Weight = 0.25
Else
Select Case MyType
Case Is = 1
  .Line.ForeColor.RGB = RGB(255, 0, 0)
Case Is = 2
  .Line.ForeColor.RGB = RGB(153, 204, 255)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
Case Is = 3
  .Line.ForeColor.RGB = RGB(128, 0, 128)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
Case Is = 4
  .Line.ForeColor.RGB = RGB(255, 153, 0)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
Case Is = 5
  .Line.ForeColor.RGB = RGB(0, 0, 0)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
End Select
End If
End With
MyDiffLine = MyStartPercent
Set SPP = Nothing
Set MyShape = Nothing
End Function

Function MyLine(MyType As Long, rRange As Range, rSNominator As Range, rSDenominator As Range, rENominator As Range, rEDenominator As Range) As Double
Dim YHeight As Single, MyStartLeft As Single, MyStartTop As Single, MyCellWidth As Single, MyEndLeft As Single, MyNextCellWidth As Single
Dim MyBalanceHeight As Single, MyBlankOffset As Single, MyNameLeft As Single, MyNameTop As Single, MyEndTop As Single
Dim MyStartBlank As Double, MyStartPercent As Double, MyEndPercent As Double, MyEndBlank As Double
Dim MyShape As Shape, SPP As Shape
Dim MyWS As String, MyConName As String

MyWS = rRange.Worksheet.Name
MyNameLeft = rRange.Left
MyNameTop = rRange.Top

Select Case MyType
Case Is = 1
MyConName = "Con1_" & MyNameLeft & "_" & MyNameTop
Case Is = 2
MyConName = "Con2_" & MyNameLeft & "_" & MyNameTop
Case Is = 3
MyConName = "Con3_" & MyNameLeft & "_" & MyNameTop
Case Is = 4
MyConName = "Con4_" & MyNameLeft & "_" & MyNameTop
Case Is = 5
MyConName = "Con5_" & MyNameLeft & "_" & MyNameTop
End Select

On Error Resume Next
Set MyShape = ThisWorkbook.Worksheets(MyWS).Shapes(MyConName)
MyShape.Delete

MyStartLeft = rRange.Left
MyStartTop = rRange.Top
MyEndTop = rRange.Top

YHeight = rRange.Cells.Height
'Calculate Cell Widths - we want midpoint of cell to midpoint of next cell to right
MyCellWidth = rRange.Cells(1, 1).Width
MyNextCellWidth = rRange.Offset(0, 1).Cells(1, 1).Width

MyStartLeft = MyStartLeft + ((MyCellWidth) / 2)
MyEndLeft = MyStartLeft + (MyCellWidth / 2) + (MyNextCellWidth / 2)

'Calculate Height
MyStartPercent = rSNominator.Value2 / rSDenominator.Value2
MyEndPercent = rENominator.Value2 / rEDenominator.Value2

If rSNominator.Value2 = (-9999.9999 / 1000) Then
MyStartPercent = 0
MyEndPercent = 0
End If

If rSNominator.Value <> (-9999.9999 / 1000) And rENominator.Value2 = (-9999.9999 / 1000) Then
MyEndPercent = MyStartPercent
End If
MyStartBlank = 1 - MyStartPercent
MyEndBlank = 1 - MyEndPercent
MyBlankOffset = CSng(MyStartBlank * YHeight)
MyStartTop = MyStartTop + MyBlankOffset
MyBlankOffset = CSng(MyEndBlank * YHeight)
MyEndTop = MyEndTop + MyBlankOffset
Set SPP = ThisWorkbook.Worksheets(MyWS).Shapes.AddConnector(msoConnectorStraight, MyStartLeft, MyStartTop, MyEndLeft, MyEndTop)
With SPP
  .Name = MyConName
If MyStartPercent = 0 And MyEndPercent = 0 Then
  .Line.ForeColor.RGB = RGB(0, 0, 0)
  .Line.Weight = 0.25
Else
Select Case MyType
Case Is = 1
  .Line.ForeColor.RGB = RGB(255, 0, 0)
Case Is = 2
  .Line.ForeColor.RGB = RGB(153, 204, 255)
  .Line.Weight = 0.01
  .Line.Transparency = 0.7
Case Is = 3
  .Line.ForeColor.RGB = RGB(128, 0, 128)
  .Line.Weight = 0.01
  .Line.Transparency = 0.7
Case Is = 4
  .Line.ForeColor.RGB = RGB(255, 153, 0)
  .Line.Weight = 0.01
  .Line.Transparency = 0.7
Case Is = 5
  .Line.ForeColor.RGB = RGB(0, 0, 0)
  .Line.Weight = 0.01
  .Line.Transparency = 0.7
End Select
End If
End With
MyLine = MyStartPercent
Set SPP = Nothing
Set MyShape = Nothing
End Function

Function MyLineNew(MyFunctionRange As Range, MyDataRange As Range, MyChartRange As Range, MyDenominator As Long) As Double
Dim YHeight As Single, MyStartLeft As Single, MyStartTop As Single, MyCellWidth As Single, MyEndLeft As Single, MyNextCellWidth As Single
Dim MyBalanceHeight As Single, MyBlankOffset As Single, MyNameLeft As Single, MyNameTop As Single, MyEndTop As Single
Dim MyStartBlank As Double, MyStartPercent As Double, MyEndPercent As Double, MyEndBlank As Double
Dim MyShape As Shape, SPP As Shape
Dim MyWS As String, MyShapeName As String
Dim arrData() As Variant
Dim MyX As Long, MyY As Long
Dim MyXStart As Long, MyYStart As Long
Dim SNominator As Double, ENominator As Double
Dim MyType As Long


MyWS = Application.Caller.Worksheet.Name
MyNameLeft = MyChartRange.Left
MyNameTop = MyChartRange.Top
MyXStart = MyFunctionRange.Column
MyYStart = MyFunctionRange.Cells(1, 1).Row
MyX = Application.Caller.Column - MyXStart + 1
MyY = Application.Caller.Row - MyYStart + 1

arrData = MyDataRange.Value
SNominator = CDbl(arrData(MyY, MyX))
ENominator = CDbl(arrData(MyY, MyX + 1))
MyType = MyY

MyShapeName = CStr(arrData(MyY, UBound(arrData(), 2))) & CStr(MyY) & "_" & CStr(MyX)

On Error Resume Next
Set MyShape = ThisWorkbook.Worksheets(MyWS).Shapes(MyShapeName)
MyShape.Delete

MyStartLeft = MyChartRange.Left
MyStartTop = MyChartRange.Top
MyEndTop = MyChartRange.Top

YHeight = MyChartRange.Cells.Height
'Calculate Cell Widths - we want midpoint of cell to midpoint of next cell to right
MyCellWidth = MyChartRange.Cells(1, 1).Width
MyNextCellWidth = MyChartRange.Offset(0, 1).Cells(1, 1).Width

MyStartLeft = MyStartLeft + ((MyCellWidth) * (MyX - 1))
MyEndLeft = MyStartLeft + (MyCellWidth)

'Calculate Height
MyStartPercent = SNominator
MyEndPercent = ENominator

If SNominator = (-9999.9999) Then
MyStartPercent = 0
MyEndPercent = 0
End If

If SNominator <> (-9999.9999) And ENominator = (-9999.9999) Then
MyEndPercent = MyStartPercent
End If

MyStartBlank = 1 - MyStartPercent
MyEndBlank = 1 - MyEndPercent

MyBlankOffset = CSng(MyStartBlank * YHeight)
MyStartTop = MyStartTop + MyBlankOffset

MyBlankOffset = CSng(MyEndBlank * YHeight)
MyEndTop = MyEndTop + MyBlankOffset

Set SPP = ThisWorkbook.Worksheets(MyWS).Shapes.AddConnector(msoConnectorStraight, MyStartLeft, MyStartTop, MyEndLeft, MyEndTop)
With SPP
  .Name = MyShapeName
If MyStartPercent = 0 And MyEndPercent = 0 Then
  .Line.ForeColor.RGB = RGB(0, 0, 0)
  .Line.Weight = 0.25
Else
Select Case MyType

Case Is = 6
  .Line.ForeColor.RGB = RGB(255, 0, 0)
  .Line.Weight = 1.5
  
Case Is = 5
  .Line.ForeColor.RGB = 6174487 'RGB(153, 204, 255)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
Case Is = 4
  .Line.ForeColor.RGB = RGB(128, 0, 128)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
Case Is = 3
  .Line.ForeColor.RGB = RGB(255, 153, 0)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
Case Is = 2
  .Line.ForeColor.RGB = RGB(0, 0, 0)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2
Case Is = 1
  .Line.ForeColor.RGB = 3969911 'RGB(204, 204, 255)
  .Line.Weight = 0.5
  .Line.Transparency = 0.2

End Select
End If
End With
MyLineNew = MyStartPercent
Set SPP = Nothing
Set MyShape = Nothing
End Function


