Attribute VB_Name = "SQLRetrieve"
Option Explicit
'Option Base 1

''' are the Comments
'   are the Comments to uncomment depending on your selection

Sub MySQLRetQuery()  '''Retrieve My Selected Records into the Excel Worksheet

Dim rsData As ADODB.Recordset
Dim sConnect As String
Dim sSQL As String
Dim WB As Workbook
Dim WS As Worksheet
Dim start As Double, finish As Double, totaltime As Double

'''Turn off Screen Updating and Automatic Calculation
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

start = Timer

'''Set Workbook and Worksheet Objects
Set WB = ThisWorkbook
Set WS = WB.Worksheets("SelectDati")

'''Clear Contents of fields
WS.Range("A:G").ClearContents

MyDatabaseLogin

On Error GoTo myend

'''ConnectionString to MySQL Database
sConnect = "Driver={" & SQLODBCDriver & "};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"
'sConnect = "Driver={MySQL ODBC 5.3 ANSI Driver};SERVER=" & ServerName & ";DATABASE=" & DatabaseName & _
";USER=" & UserName & ";PASSWORD=" & Password & ";Option=3"

''''**************************************************************''''''''''''''''''''''''''''''

'Dim idate As String

'''1)
'''DDate declared as Number to MySQL Table

'idate = InputBox("Select Date YYYYMMDD")
'If idate = "" Or Len(idate) <> 8 Then idate = Format(Now, "YYYYMMDD")

sSQL = "SELECT * " & _
"FROM `test`.`CrossBorder` ORDER BY IDIndex, Hour" '& _
"where DDate=" & CLng(idate)

'''MsgBox sSQL


'''2)
'''DDate declared as Date to MySQL Table

'idate = InputBox("Select Date YYYYMMDD")
'If idate = "" Or Len(idate) <> 8 Then idate = CLng(Format(Now, "YYYYMMDD"))
'If idate <> "" And Len(idate) = 8 Then
'idate = CLng(idate)
'idate = CVar(Left(idate, 4) & "-" & Mid(idate, 5, 2) & "-" & Right(idate, 2))
'End If

'sSQL = "SELECT * " & _
"FROM `test`.`CrossBorder` " & _
"where DDate=DATE_FORMAT('" & (Left(idate, 4) & "/" & Mid(idate, 5, 2) & "/" & Right(idate, 2)) & "', '%Y-%m-%d')"

'''Comments
'''MsgBox CDate(Left(idate, 4) & "/" & Mid(idate, 5, 2) & "/" & Right(idate, 2)) 'Date format tries
'''MsgBox DateSerial(CInt(Left(idate, 4)), CInt(Mid(idate, 5, 2)), CInt(Right(idate, 2))) 'Date format tries
'''if campus on table is declared as datetime one could apply also the follow format
'''SELECT DATE_FORMAT('2011-10-10 19:46:00', '%M %d, %Y');

'''MsgBox sSQL
      
''''**************************************************************''''''''''''''''''''''''''''''
      
'''Create the Recordset object and run the query.
Set rsData = New ADODB.Recordset

'''Open the Connection
rsData.CursorLocation = adUseServer

'''Depends on the Excel Application Version you have, the way you will connect
'''to the Database differs
If Val(Application.Version) >= 12 Then
rsData.Open sSQL, sConnect, adOpenDynamic, adLockOptimistic, adCmdText
Else
rsData.Open sSQL, sConnect, adOpenForwardOnly, adLockReadOnly, adCmdText
End If
    
'''Make Sure we got records back.EOF means End of File
If Not rsData.EOF Then

'''Two Methods possible

'''1)
'''***************************************
Dim r As Long, c As Long, intRows As Long, intCols As Long
Dim arrFinalData() As Variant, arrData() As Variant

    arrData = rsData.GetRows             '''Assign to an array the Records Retrieve from the sSQL
    intRows = UBound(arrData, 2)
    intCols = UBound(arrData, 1)
    
    '''Transpose the array data
    ReDim arrFinalData(1 To intRows + 1, 1 To intCols + 1)
        For r = 0 To intRows
            For c = 0 To intCols
                If IsNumeric(arrData(c, r)) Then
                
                If Int(arrData(c, r)) <> arrData(c, r) Then
                arrFinalData(r + 1, c + 1) = CDbl(Format(CDbl(arrData(c, r)), "0.00"))
                Else
                arrFinalData(r + 1, c + 1) = CDbl(arrData(c, r))
                End If
                
                Else
                arrFinalData(r + 1, c + 1) = arrData(c, r)
                End If
            Next c
        Next r

If IsArrayEmpty(arrFinalData) = False Then
WS.Cells.Clear
WS.Range("A2").Resize(intRows + 1, intCols + 1).Value = arrFinalData
End If
Erase arrData(): Erase arrFinalData()
'''***************************************

'''OR

'''2)
'''***************************************
'''Use the copyfromrecordset to retrieve values on your worksheet
'''Just the Body Range, No Headers have been taken into account
'WS.Range("A2").CopyFromRecordset rsData
'''***************************************
    
    '''Close the Recordset object
    rsData.Close
    '''Add headers to the worksheet
    With WS.Range("A1:G1")
        .Value = Array("IDIndex", "DDate", "Hour", "Border", "Purpose", "Qty", "Price")
        .Font.Bold = True
    End With
    
    '''Fit the column widths to the data
    WS.UsedRange.EntireColumn.AutoFit
Else
    '''Close the Recordset object
    rsData.Close
    MsgBox "Error: No records returned.", vbCritical
End If

myend:

'''Close the Connection with the Database
If rsData.State <> adStateClosed Then
rsData.Close
End If

'''Destroy the Recordset Object
If Not rsData Is Nothing Then Set rsData = Nothing

'''Timer Ends
finish = Timer
totaltime = Format(finish - start, "0.00")

'''Destroy the WorkSheet Object
Set WS = Nothing

If Err.Number = 0 Then
MsgBox "MySQL " & intRows + 1 & " Records Retrieved. Stored on Range L12." & vbLf & vbLf _
& "Go to SelectDati Worksheet"
'''Insert number of seconds
WB.Worksheets("Dashboard").Range("I10").Value = totaltime
'''SQLRecords
WB.Worksheets("Dashboard").Range("L12").Value = intRows + 1
'''MyLastUpdate
WB.Worksheets("Dashboard").Range("L10").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
If DateSerial(Year(Now), Month(Now), Day(Now)) <> WB.Worksheets("Dashboard").Range("L10").Value Then
WB.Worksheets("Dashboard").Range("L10").Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
End If

Else
MsgBox "MySQL Records Not Retrieved", vbCritical
WB.Worksheets("Dashboard").Range("I10").Value = ""
End If

'''Destroy the Workbook Object
Set WB = Nothing

'''Turn on Screen Updating and Automatic Calculation
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

Private Function IsArrayEmpty(MyArray As Variant) As Boolean '''Check if an array created is empty
IsArrayEmpty = False
Err.Number = 0
On Error GoTo ErrHandler:

    Dim Element As Variant
    For Each Element In MyArray
        If IsEmpty(Element) Or IsNull(Element) Then
           IsArrayEmpty = True
           Exit Function
        End If
    Next Element
    
ErrHandler:
If Err.Number <> 0 Then IsArrayEmpty = True
End Function
