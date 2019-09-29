Attribute VB_Name = "modCopyRangetoPNG"
Option Explicit

' Worked in Windows 10 and Excel 2016

Sub CopyRangeToPNG()
' Sources
' https://software-solutions-online.com/excel-vba-save-table-as-jpeg/

Dim SourceWS As Worksheet
Dim DestWS As Worksheet
Dim strPath As String
Dim i As Integer, j As Integer
Dim rng As Range
Dim intCount As Integer
Dim objPic As Shape
Dim objChart As Chart

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

' Create imgs path
strPath = ThisWorkbook.Path & IIf(Right(ThisWorkbook.Path, 1) = "\", "", "\") & "imgs\"
If Not FileFolderExists(strPath) Then MkDir strPath

Set SourceWS = ThisWorkbook.Worksheets("Sheet1")
Set DestWS = ThisWorkbook.Worksheets("Sheet2")

Dim myImgsCounter As Integer
myImgsCounter = 12

For i = 1 To myImgsCounter
Select Case i
Case Is = 1
Set rng = SourceWS.Range("A1:AM23") 'Balance Situation
Case Is = 2
Set rng = SourceWS.Range("A31:AM55") 'Total Consumption
Case Is = 3
Set rng = SourceWS.Range("A72:AM95") 'Industry Consumption
Case Is = 4
Set rng = SourceWS.Range("A108:AM131") 'Power Generation Consumption
Case Is = 5
Set rng = SourceWS.Range("A139:AM158") 'LDZ consumptoin
Case Is = 6
Set rng = SourceWS.Range("A161:AM177") 'AU Imp
Case Is = 7
Set rng = SourceWS.Range("A178:AM194") 'CH Imp
Case Is = 8
Set rng = SourceWS.Range("A195:AM211") 'Libya Imp
Case Is = 9
Set rng = SourceWS.Range("A212:AM228") 'Tunisia
Case Is = 10
Set rng = SourceWS.Range("A229:AM245") 'Cavarzere
Case Is = 11
Set rng = SourceWS.Range("A246:AM262") 'Panigaglia
Case Is = 12
Set rng = SourceWS.Range("A263:AM279") 'Livorno
End Select

' Select source workbook
SourceWS.Select

' Copy the range as an image
Call rng.CopyPicture(xlScreen, xlPicture)

' Remove all previous shapes in destws
intCount = DestWS.Shapes.Count
For j = 1 To intCount
DestWS.Shapes.Item(1).Delete
Next j

' Create an empty chart in destws
DestWS.Shapes.AddChart

' Select dest workbook
DestWS.Select

' Select the shape in destws
DestWS.Shapes.Item(1).Select
Set objChart = ActiveChart

DestWS.Shapes.Item(1).Width = rng.Width
DestWS.Shapes.Item(1).Height = rng.Height

' Paste the range into the chart
objChart.Paste

On Error Resume Next

Kill strPath & "mytestfile" & i & ".png"

' Save the chart as a PNG
objChart.Export strPath & "mytestfile" & i & ".png"

On Error GoTo 0

' Delete shape
DestWS.Shapes.Item(1).Delete

Next i

' Go Back to Buttons Sheet
ThisWorkbook.Worksheets("Buttons").Select

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

Set rng = Nothing
Set DestWS = Nothing
Set SourceWS = Nothing
End Sub

