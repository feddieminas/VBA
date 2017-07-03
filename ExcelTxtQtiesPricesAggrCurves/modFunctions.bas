Attribute VB_Name = "modFunctions"
Option Explicit
Option Base 1

'Need to change the Filter Column Number criteria to Sort Descending
'''''Function MySortDesc(Array_Values() As Object) = If it referes for ex. to ExcelNameRange
Function MySortDesc(Array_Values() As Double) As Double()
'''''Application.Calculation = xlCalculationManual
Application.Volatile (False)
'''''Application.ScreenUpdating = False

Dim size As Long, changed As Long, mycols As Long
Dim tmp As Double
Dim nums() As Double
Dim myrows As Long, i As Long, j As Long, cycle As Long
mycols = UBound(Array_Values, 2)
myrows = UBound(Array_Values, 1)
'''''mycols = Array_Values.Columns.Count
'''''myrows = Application.WorksheetFunction.CountA(Array_Values.Rows) / mycols
'''''myrows = ThisWorkbook.Worksheets("Notepad").Range("FD1").Value

ReDim Preserve nums(1 To myrows, 1 To mycols)
For j = 1 To mycols
For i = 1 To myrows
nums(i, j) = Array_Values(i, j)
Next i
Next j

Do
cycle = cycle + 1
changed = 0

For i = 1 To myrows - 1

'2 is the Filter Column Number criteria (PricesBO)
If nums(i, 2) < nums(i + 1, 2) Then

For j = 1 To mycols
tmp = nums(i + 1, j)
nums(i + 1, j) = nums(i, j)
nums(i, j) = tmp
Next j

changed = 1
End If

Next i
Loop Until changed = 0

MySortDesc = nums
'''''MySortDesc = WorksheetFunction.Transpose(nums)
Application.Volatile (True)
'''''Application.Calculation = xlCalculationAutomatic
'''''Application.ScreenUpdating = True
End Function

Public Function FileFolderExists(strFullPath As String) As Boolean
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
EarlyExit:
    On Error GoTo 0
End Function

