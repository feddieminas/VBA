Attribute VB_Name = "MyFunctions"

'***********************************
'EPEX_PowerBidOffer_Template.xlsm

'To avoid having any recalculation issues and UDF functions are not getting recognised (NAME shown in cells),
'I recommend you do the following step :
'Go to Excel Options --> Trust Center --> Trust Center Settings --> Trusted locations --> Click below on
'Allow Trusted Locations on my network (not recommended)
'Then press on Add new location..., browse your filepath and click on 'Subfolders of this locator are also trusted'.
'Press ok, ok to Trusted Locations and finally ok on Excel Options.
'Save the file, close Microsoft Excel Application and open again the file.
'***********************************

Function MyDAMSuffix(cell As Range)
Select Case cell.Value
Case "Austria"
Damsfx = "AU"
Case "France"
Damsfx = "FR"
Case "Germany"
Damsfx = "DE-AMP"
Case "Switzerland"
Damsfx = "CH"
End Select

MyDAMSuffix = Damsfx  'ex. MyDAMSuffix(ThisWorkbook.Worksheets("MyTemplate").Range("B2"))
End Function

Function MergeRanges(ParamArray arguments() As Variant) As Variant()
Dim cell As Range, temp() As Variant, argument As Variant
Dim iRows As Integer, i As Integer
ReDim temp(0)
For Each argument In arguments
  For Each cell In argument
    If cell <> "" Then
      temp(UBound(temp)) = cell
      ReDim Preserve temp(UBound(temp) + 1)
    End If
  Next cell
Next argument
ReDim Preserve temp(UBound(temp) - 1)
iRows = Range(Application.Caller.Address).Rows.Count
For i = UBound(temp) To iRows
  ReDim Preserve temp(UBound(temp) + 1)
  temp(UBound(temp)) = ""
Next i
MergeRanges = Application.Transpose(temp)
End Function

Function MySort(Array_Values As Object)
Application.Calculation = xlCalculationManual
Application.Volatile (False)
Application.ScreenUpdating = False
Dim size As Long, changed As Long
Dim tmp As Double, nums() As Double
Dim limit As Long, i As Long, cycle As Long
'limit = Array_Values.Rows.Count
limit = ThisWorkbook.Worksheets("MyTemplate").Range("G1").Value

ReDim Preserve nums(1 To limit)
For i = 1 To limit
nums(i) = Array_Values(i)
Next i
timer_val = Timer
Do
cycle = cycle + 1
changed = 0
For i = 1 To limit - 1
If nums(i) > nums(i + 1) Then
tmp = nums(i)
nums(i) = nums(i + 1)
nums(i + 1) = tmp
changed = 1
End If
Next i
Loop Until changed = 0
MySort = WorksheetFunction.Transpose(nums)
Application.Volatile (True)
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Function

