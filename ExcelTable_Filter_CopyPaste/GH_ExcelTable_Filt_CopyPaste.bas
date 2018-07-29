Attribute VB_Name = "GH_ExcelTable_Filt_CopyPaste"
'Create an Excel Table, Filter it and Paste the desired rows and columns to a destination worksheet range
Option Explicit
Option Base 1

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Worked the project in Windows 7 and Excel 2007, 2010
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'******************************************************************
'Example of an Excel Table Filtering, Copying and Paste values
'******************************************************************

'Instructions :
'Run CreateFileTemplate and MyExcelTable macro in order
'From Line Ln147 to Ln250 I demonstrate various ways of filtering your data according to the type of variable. One can
'comment and uncomment to try

Const StartCol = "A"   'Manually adjust your Const Values. INSERT or keep it
Const StartRow = 1
Const EndCol = "F"
Const Wst = "AllDeals"


Sub CreateTableTemplate() 'Need to Run only once to create your Template table
Dim i As Long

Dim arrchr() As String
ReDim arrchr(1 To 6)
arrchr(1) = StartCol
arrchr(2) = Chr(Asc(StartCol) + 1)
arrchr(3) = Chr(Asc(StartCol) + 2)
arrchr(4) = Chr(Asc(StartCol) + 3)
arrchr(5) = Chr(Asc(StartCol) + 4)
arrchr(6) = EndCol
'Notes: Asc function above converts a given Char to Integer
'Chr function above converts a given Integer to Char

With ThisWorkbook.Worksheets(1)
.Cells.Clear

.Name = Wst
.Range(arrchr(1) & StartRow) = "Book"
.Range(arrchr(2) & StartRow) = "Counterparty"
.Range(arrchr(3) & StartRow) = "StartDate"
.Range(arrchr(4) & StartRow) = "EndDate"
.Range(arrchr(5) & StartRow) = "Price"
.Range(arrchr(6) & StartRow) = "MWh"

For i = 2 To 100

If i < 50 Then

If i Mod 2 = 0 Then
.Range(arrchr(1) & i) = "Germany"
.Range(arrchr(2) & i) = "EEX FUTURES"
.Range(arrchr(3) & i) = Format("01/01/2014", "DD/MM/YYYY")
.Range(arrchr(4) & i) = Format("31/12/2014", "DD/MM/YYYY")
.Range(arrchr(5) & i) = CDbl(Format(i + 2.22, "0.00"))
.Range(arrchr(6) & i) = 1
Else
.Range(arrchr(1) & i) = "Italy"
.Range(arrchr(2) & i) = "TEI Energy"
.Range(arrchr(3) & i) = Format("01/04/2014", "MM/DD/YYYY")
'.Range(arrchr(3) & i) = Format("01/04/2014", "DD/MM/YYYY") 'Issue with local settings. See comment on Ln 150
.Range(arrchr(4) & i) = Format("30/06/2014", "DD/MM/YYYY")
.Range(arrchr(5) & i) = CDbl(Format(i + 1.01, "0.00"))
.Range(arrchr(6) & i) = 2
End If

Else

If i Mod 2 = 0 Then
.Range(arrchr(1) & i) = "France"
.Range(arrchr(2) & i) = "EEX FR FUTURES"
.Range(arrchr(3) & i) = Format("01/01/2015", "DD/MM/YYYY")
.Range(arrchr(4) & i) = Format("31/03/2015", "DD/MM/YYYY")
.Range(arrchr(5) & i) = CDbl(Format(i + 2.42, "0.00"))
.Range(arrchr(6) & i) = 3
Else
.Range(arrchr(1) & i) = "Switzerland"
.Range(arrchr(2) & i) = "EEX CH FUTURES"
.Range(arrchr(3) & i) = Format("01/01/2014", "DD/MM/YYYY")
.Range(arrchr(4) & i) = Format("31/03/2014", "DD/MM/YYYY")
.Range(arrchr(5) & i) = CDbl(Format(i + 1.11, "0.00"))
.Range(arrchr(6) & i) = 4
End If

End If

Next i

'Some Duplicate Price Values
.Range(arrchr(5) & 7) = CDbl(8.22)
.Range(arrchr(5) & 14) = CDbl(14.01)
.Range(arrchr(5) & 31) = CDbl(30.22)
End With

Erase arrchr()
End Sub

Sub AddTable(EndRow As Long)
With ThisWorkbook.Worksheets(Wst)
     '.Select
     .ListObjects.Add(xlSrcRange, .Range(StartCol & StartRow & ":" & EndCol & EndRow), , xlYes).Name = _
        "Table1"
End With
End Sub

Sub DeleteTable()
Dim objListObj As ListObject

With ThisWorkbook.Worksheets(Wst)
    Set objListObj = .ListObjects("Table1")
    objListObj.Unlist
End With

Set objListObj = Nothing

End Sub

Sub MyExcelTable() 'Assume there are Headers to your Data
Application.ScreenUpdating = False

Dim EndRow As Long
With ThisWorkbook.Sheets(Wst)
'.AutoFilterMode = False

EndRow = .Cells(.Rows.Count, "A").End(xlUp).Row

'If .Range(StartCol & StartRow).ListObject.Name = "Table1" Then AddTable EndRow
If .ListObjects.Count = 0 Then AddTable EndRow
End With

Dim rng As Range
Set rng = ThisWorkbook.Worksheets(Wst).Range(StartCol & StartRow)

Dim My_Table As ListObject
Set My_Table = rng.ListObject
If My_Table Is Nothing Then GoTo NoTableCreated

'**********************************Filtering***********************************************'
Dim i As Long

'DATES
Dim MyDateFilter As String
Dim MyDateFilter2 As String

'Filter with Criteria
rng.AutoFilter Field:=3 'Field 3 is Excel Column C
MyDateFilter = Format("01/01/2014", "dd/mm/yyyy")
'VBA due to its settings might recognise your "DD/MM/YYYY" as "MM/DD/YYYY". If so modify your MyDateFilter in that format
'ex. MyDateFilter = Format("01/01/2014", "mm/dd/yyyy")
'Other :
'Format for Date and Time would be MyDateFilter = Format("01/01/2014", "dd/mm/yyyy hh:mm:ss")
'MyDateFilter = Format("01/01/2014 13:00:00", "dd/mm/yyyy hh:mm:ss")
If Val(Application.Version) = 12 Then  'Excel 2007
rng.AutoFilter Field:=3, Criteria1:=CDate(MyDateFilter)
'OR within a range
'rng.AutoFilter Field:=3, Criteria1:=">=" & CDate(MyDateFilter)
ElseIf Val(Application.Version) >= 14 Then 'Excel 2010 and on
rng.AutoFilter Field:=3, Criteria1:=CStr(MyDateFilter) 'Excel 2010 and assume and on
'OR within a range
'rng.AutoFilter Field:=3, Criteria1:=">=" & CStr(MyDateFilter)
End If

'Filter with two Criteria
rng.AutoFilter Field:=3
MyDateFilter2 = Format("01/04/2014", "mm/dd/yyyy")
'VBA due to its settings might recognise your "DD/MM/YYYY" as "MM/DD/YYYY". If so modify your MyDateFilter in that format
'ex. MyDateFilter2 = Format("01/04/2014", "mm/dd/yyyy")
If Val(Application.Version) = 12 Then  'Excel 2007
rng.AutoFilter Field:=3, Criteria1:=">=" & CDate(MyDateFilter), Operator:=xlAnd, Criteria2:="<=" & CDate(MyDateFilter2)
ElseIf Val(Application.Version) >= 14 Then 'Excel 2010 and assume and on
rng.AutoFilter Field:=3, Criteria1:=">=" & CStr(MyDateFilter), Operator:=xlAnd, Criteria2:="<=" & CStr(MyDateFilter2)
End If
'The or operation is Operator:=xlOr

rng.AutoFilter Field:=3
'*******


'STRING
Dim MyStringFilter As String
Dim strarr(3) As String

'Filter with Criteria
rng.AutoFilter Field:=2 'Field 2 is Excel Column B
MyStringFilter = "EEX FUTURES"
rng.AutoFilter Field:=2, Criteria1:="=" & MyStringFilter

'Filter with Criteria using InputBox
'rng.AutoFilter Field:=2
'MyStringFilter = InputBox("Type your filter item.")
'If MyStringFilter = "" Then GoTo EndTable:
'rng.AutoFilter Field:=2, Criteria1:="=" & MyStringFilter

'Filter with Criteria Multiple Strings as an Array of Strings
rng.AutoFilter Field:=2
rng.AutoFilter Field:=2, Criteria1:=Array( _
    "EEX FUTURES", "EEX FR FUTURES", "TEI Energy"), Operator:=xlFilterValues

'Filter with Criteria Multiple Strings with initialising an Array of Strings
rng.AutoFilter Field:=2
strarr(1) = "EEX FUTURES"
strarr(2) = "EEX FR FUTURES"
strarr(3) = "TEI Energy"
rng.AutoFilter Field:=2, Criteria1:=strarr, Operator:=xlFilterValues

'Filter with Criteria Multiple Strings with initialising an Array of Strings on an InputBox
'rng.AutoFilter Field:=2
'For i = 1 To UBound(strarr)
'strarr(i) = InputBox("Type your " & i & " filter item.")
'Next i
'rng.AutoFilter Field:=2, Criteria1:=strarr, Operator:=xlFilterValues
'Erase strarr()

rng.AutoFilter Field:=2
'*******


'NUMBER
Dim MyNumFilter As Variant
Dim numarr(3) As Variant

'Filter with Criteria Single Number
rng.AutoFilter Field:=5 'Field 5 is Excel Column E
MyNumFilter = CStr("14.01")
rng.AutoFilter Field:=5, Criteria1:=">=" & MyNumFilter 'If MyNumFilter declared as integer could change to CInt or CLng if Long

'Filter with Criteria Multiple Numbers with initialising an Array of Numbers
rng.AutoFilter Field:=5
numarr(1) = CStr("8.22")
numarr(2) = CStr("14.01")
numarr(3) = CStr("30.22")
rng.AutoFilter Field:=5, Criteria1:=numarr, Operator:=xlFilterValues

'Filter with Criteria Multiple Numbers with initialising an Array of Numbers on an InputBox
'rng.AutoFilter Field:=5
'For i = 1 To UBound(numarr)
'numarr(i) = CStr(InputBox("Type your " & i & " filter item."))
'Next i
'rng.AutoFilter Field:=5, Criteria1:=numarr, Operator:=xlFilterValues
'Erase numarr()

'rng.AutoFilter Field:=5
'*******

'***************************Copying and Pasting***********************************************'

On Error GoTo EndTable

Dim arethererows As Long
arethererows = My_Table.Range.SpecialCells(xlCellTypeVisible).Offset(1, 0).Cells.Count
 
Dim FilteredStartRow As Long, FilteredEndRow As Long
Dim FilteredStartCol As String, FilteredEndCol As String

If arethererows > My_Table.Range.Columns.Count Then

FilteredStartRow = My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, 1).Row - StartRow + 1
FilteredEndRow = My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Cells(1, 1).End(xlDown).Row

'*******Static Copy and Paste Values. From next line to line 283*******'

FilteredStartCol = "A"                      'Manually Adjust your Start Column    INSERT or keep it
FilteredEndCol = "A"                        'Manually Adjust your End Column      INSERT or keep it
'Make Sure your End Column Character is higher or equal to your Start Column Character

ThisWorkbook.Sheets(Wst).Select
My_Table.Range.SpecialCells(xlCellTypeVisible).Range(FilteredStartCol & FilteredStartRow & ":" & _
FilteredEndCol & FilteredEndRow).Copy

Dim MyDestWsht As String, FilteredDestRow As Long, FilteredDestCol As Long
MyDestWsht = "Sheet2"                       'Manually Adjust your Destination Worksheet INSERT or keep it

If ThisWorkbook.Worksheets.Count = 1 Then   'Check if there a second Sheet and name it
ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = MyDestWsht
Else
ThisWorkbook.Worksheets(2).Name = MyDestWsht
End If

FilteredDestRow = 3                         'Manually Adjust your Destination Row         INSERT or keep it
FilteredDestCol = 2                         'Manually Adjust your Destination Column      INSERT or keep it

With ThisWorkbook.Worksheets(MyDestWsht)
.Cells(FilteredDestRow, FilteredDestCol).PasteSpecial xlPasteColumnWidths
'.Cells(FilteredDestRow, FilteredDestCol).PasteSpecial xlPasteValues
.Cells(FilteredDestRow, FilteredDestCol).PasteSpecial xlPasteValuesAndNumberFormats '.... MyDestWsht Column B
End With
Application.CutCopyMode = False

'*******

'UNIQUE Records of a Single Table Column Filtered  .... MyDestWsht Column F
Dim MyFiltDupl() As Variant
MyFiltDupl = MyCopyFiltDupl(My_Table, 5) '5 is Price Column
With ThisWorkbook.Worksheets(MyDestWsht)
.Cells(FilteredDestRow, FilteredDestCol + Asc(FilteredEndCol) - _
Asc(FilteredStartCol) + 4).Resize(UBound(MyFiltDupl, 1), UBound(MyFiltDupl, 2)) = MyFiltDupl
End With
Erase MyFiltDupl()
'UNIQUE Records of a Single Table Column UnFiltered .... MyDestWsht Column D
With ThisWorkbook.Worksheets(MyDestWsht) '5 is Price Column
My_Table.ListColumns(5).Range.AdvancedFilter Action:=xlFilterCopy, _
CopyToRange:=.Cells(FilteredDestRow - 1, FilteredDestCol + Asc(FilteredEndCol) - Asc(FilteredStartCol) + 2), Unique:=True
End With
'Notes: Asc function above converts a given Char to Integer

'**********************************************************************************************'

End If

EndTable:
'Dates
rng.AutoFilter Field:=3
'String
rng.AutoFilter Field:=2
'Number
rng.AutoFilter Field:=5

ThisWorkbook.Sheets(MyDestWsht).Select

Set My_Table = Nothing
If ThisWorkbook.Sheets(Wst).ListObjects.Count > 0 Then DeleteTable

NoTableCreated:
Set rng = Nothing
Application.ScreenUpdating = True
End Sub

Function MyCopyFiltDupl(My_Table As ListObject, ListCol As Integer) As Variant() 'Function which copies VisibleRows and not allow Duplicate Records
Dim temparr() As Variant
Dim VisibleRows As Long
VisibleRows = (My_Table.Range.SpecialCells(xlCellTypeVisible).Offset(1, 0).Cells.Count / My_Table.Range.Columns.Count) - 1
ReDim temparr(VisibleRows, 1)
'OR
'You can as well consider the whole list of rows
'ReDim temparr(My_Table.ListRows.Count, 1)

'******* 'Begin My Visible ListRows of the Selected ListCol to an array
'The below loop which inserts your visible excel rows to an array can also
'be used as an alternate to the above static copy and paste method - Ln269 to Ln288

Dim MyTableVisibleRng As Range, i As Long, j As Long, k As Long
Set MyTableVisibleRng = My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible)
k = 1
For i = 1 To MyTableVisibleRng.Areas.Count
For j = 1 To MyTableVisibleRng.Areas(i).Rows.Count
temparr(k, 1) = ThisWorkbook.Worksheets(Wst).Cells(MyTableVisibleRng.Areas(i).Cells(j * My_Table.Range.Columns.Count).Row, ListCol).Value
'temparr(k, 1) = ThisWorkbook.Worksheets(Wst).Cells(MyTableVisibleRng.Areas(i).Row, ListCol).Value
k = k + 1
Next j
Next i
Set MyTableVisibleRng = Nothing
'******* 'End
                            'Below is the Operation to remove the duplicate values from the temparr created

k = k - 1
For i = 1 To k             'Set Duplicate Values to a value of 9999
For j = i + 1 To k
If temparr(i, 1) = temparr(j, 1) Then temparr(j, 1) = 9999
Next j
Next i

Dim temparr2() As Variant  'Create the new array with the unique values
ReDim temparr2(k, 1)

i = 1
j = 1
While (k > 0)
If temparr(j, 1) <> 9999 Then
temparr2(i, 1) = temparr(j, 1)
i = i + 1
End If
j = j + 1
k = k - 1
Wend

Erase temparr()
MyCopyFiltDupl = temparr2   'return my array
Erase temparr2()
End Function
