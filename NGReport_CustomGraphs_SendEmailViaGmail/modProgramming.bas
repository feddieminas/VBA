Attribute VB_Name = "modProgramming"
Option Compare Text
Option Explicit

' Worked in Windows 10 and Excel 2016

Sub calcSheet1()
' One launches this macro in case the Workbook is opening and
' cells from the custom UDF functions have not been recalculated

ThisWorkbook.Worksheets("Sheet1").Calculate
Application.CalculateFull

End Sub

'''' Archived Macros ''''

Sub ChangeShapeName()
Dim MyShape As Shape
Dim sName As String

For Each MyShape In ThisWorkbook.Worksheets("Sheet1").Shapes
sName = MyShape.Name
If Left(sName, 4) = "SEAS" Then
sName = Replace(sName, "SEAS", "IMP_STRAT")
MyShape.Name = sName
End If
Next

End Sub

Sub MyLinesDelete()
Dim FullRange As Range, MyConName As String
Dim MyLiveShape As Shape
    
    For Each MyLiveShape In ThisWorkbook.Worksheets("Sheet1").Shapes
    If Left(MyLiveShape.Name, 4) = "Con1" Then MyLiveShape.Delete
    Next MyLiveShape

End Sub
Sub CheckMyColour()
Dim sShape As Shape

Set sShape = ThisWorkbook.Worksheets("Sheet1").Shapes("IMP_CAV_19")
MsgBox sShape.Fill.ForeColor.RGB

End Sub

