VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ccbsImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit
''''''''''''''''''''''
' Windows API Declares
''''''''''''''''''''''
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetTickCount Lib "kernel32" () As Long


Private pSourceWBNameA As String
Private pSourceWBNameB As String                                'Private Properties are what are used within the Class
Private pSourceWSName As String                                 'as opposed to the Public Properties which are used to
Private pSourceWSRangeA As String                               'to interact between the standard module and the class module
Private pSourceWSRangeB As String
Private pDestWBName As String                                   'there are two reasons for this. First to allow for any attribute
Private pDestWSName As String                                   'assigned by the user in the Standard Module, to be fully checked for
Private pDestWSRange As String                                  'errors by error handling that needs to be inserted here.
Private pboolHorizontal As Boolean                              'second to ensure that a normal user has no possibility to screw up the code.
Private pboolChngeNeg As Boolean                                'as you can see we now have an opportunity to insert a lot of clearly defined
Public Property Let bboolHorizontal(Value As Boolean)           'error handling here.
pboolHorizontal = Value
End Property
Public Property Let bboolChngeNeg(Value As Boolean)
pboolChngeNeg = Value
End Property
Public Property Let sSourceWSName(Value As String)
pSourceWSName = Value
End Property
Public Property Let sSourceWBNameA(Value As String)
pSourceWBNameA = Value
End Property
Public Property Let sSourceWBNameB(Value As String)
pSourceWBNameB = Value
End Property
Public Property Let sSourceWSRangeA(Value As String)
pSourceWSRangeA = "$" & Value
End Property
Public Property Let sSourceWSRangeB(Value As String)
pSourceWSRangeB = "$" & Value
End Property
Public Property Let sDestWSName(Value As String)
pDestWSName = Value
End Property
Public Property Let sDestWBName(Value As String)
pDestWBName = Value
End Property
Public Property Let sDestWSRange(Value As String)
pDestWSRange = Value
End Property

Sub fImpRangeWB1()
Dim rsData As ADODB.Recordset
Dim sConnect As String, sSQL As String, MyRng As String, MyOTECheck As String
Dim r As Integer, C As Integer
Dim intRows As Long, intCols As Long, i As Long
Dim arrFinalDataA() As Variant, arrData() As Variant
Dim Wb As Workbook, WS As Worksheet
Dim bTranspose As Boolean, bNegative As Boolean, IsClosed As Boolean

If Len(pDestWBName) = 0 Then
Set Wb = Selection.Parent.Parent
Set WS = Wb.Worksheets(pDestWSName)
Else
Set Wb = Workbooks(pDestWBName)
Set WS = Wb.Worksheets(pDestWSName)
End If
On Error GoTo ErrHandler100:
                                                                
IsClosed = WaitForFileClose(FileName:=pSourceWBNameA, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
If IsClosed = False Then GoTo MyNextLine:
                                                                
                                                                
                                                                
                                                                'Need to set HDR equal to No
sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & pSourceWBNameA & ";" & _
"Extended Properties=" & """Excel 12.0 Macro;HDR=NO""" & ";"

MyRng = "[" & pSourceWSName & pSourceWSRangeA & "];"

sSQL = "Select* from " & MyRng
Set rsData = New ADODB.Recordset
rsData.Open sSQL, sConnect, adOpenForwardOnly, _
adLockReadOnly, adCmdText

If Not rsData.EOF Then                                          'Make Sure we got records back
    arrData = rsData.GetRows                                    'Assign to an array the Records Retrieve from the sSQL
    intRows = UBound(arrData, 2)
    intCols = UBound(arrData, 1)
    
MyOTECheck = left(pSourceWSName, 3)

Select Case MyOTECheck                                          'This is because OTE data is stored originally horizontally in the source wb
    
    Case Is = "OTE"
     ReDim arrFinalDataA(1 To intCols + 1, 1 To intRows + 1)
        For r = 0 To intCols
            For C = 0 To intRows
                If IsNumeric(arrData(r, C)) Then
                arrFinalDataA(r + 1, C + 1) = CDbl(arrData(r, C))
                Else
                arrFinalDataA(r + 1, C + 1) = arrData(r, C)
                End If
            Next C
        Next r
    intRows = UBound(arrFinalDataA, 1) - 1  'this is to reset these variables in order to make sure that the correct paste into the worksheet happens
    intCols = UBound(arrFinalDataA, 2) - 1
    
    Case Else
    ReDim arrFinalDataA(1 To intRows + 1, 1 To intCols + 1)
        For r = 0 To intRows
            For C = 0 To intCols
                If IsNumeric(arrData(C, r)) Then
                arrFinalDataA(r + 1, C + 1) = CDbl(arrData(C, r))
                Else
                arrFinalDataA(r + 1, C + 1) = arrData(C, r)
                End If
            Next C
        Next r
End Select
Erase arrData()
If rsData.State <> adStateClosed Then
rsData.Close
End If
If Not rsData Is Nothing Then Set rsData = Nothing

If IsArrayEmpty(arrFinalDataA()) Then GoTo MyNextLine:

If Len(pDestWBName) = 0 Then
        bTranspose = False
        bNegative = False
        With WS.Range(pDestWSRange)
            If .Columns.Count = 2 And .Rows.Count = 1 Then bTranspose = True
            If .Columns.Count = 3 And .Rows.Count = 1 Then
                bTranspose = True
                bNegative = True
            End If

        If .Columns.Count = 1 And .Rows.Count = 3 Then
            bTranspose = False
            bNegative = True
        End If
        
        If bNegative = True Then
            For i = 1 To UBound(arrFinalDataA)
                arrFinalDataA(i, 1) = -arrFinalDataA(i, 1)
            Next i
        End If

        Select Case bTranspose
            Case Is = True
                .Cells(1, 1).Resize(intCols + 1, intRows + 1).Value = Application.WorksheetFunction.Transpose(arrFinalDataA)
                .Cells(1, 1).Resize(intCols + 1, intRows + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
            Case Is = False
                .Cells(1, 1).Resize(intRows + 1, intCols + 1).Value = arrFinalDataA
                .Cells(1, 1).Resize(intRows + 1, intCols + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
        End Select

        End With

Else

        bTranspose = pboolHorizontal
        bNegative = pboolChngeNeg
        
        With WS.Range(pDestWSRange)
        
        If bNegative = True Then
            For i = 1 To UBound(arrFinalDataA)
                arrFinalDataA(i, 1) = -arrFinalDataA(i, 1)
            Next i
        End If

        Select Case bTranspose
            Case Is = True
                .Cells(1, 1).Resize(intCols + 1, intRows + 1).Value = Application.WorksheetFunction.Transpose(arrFinalDataA)
                .Cells(1, 1).Resize(intCols + 1, intRows + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
                .Cells(1, 1).Resize(intCols + 1, intRows + 1).Interior.ColorIndex = 0
                .Cells(1, 1).Resize(intCols + 1, intRows + 1).Font.ColorIndex = 1
                gcbTagsUpdated = gcbTagsUpdated + 1
            Case Is = False
                .Cells(1, 1).Resize(intRows + 1, intCols + 1).Value = arrFinalDataA
                .Cells(1, 1).Resize(intRows + 1, intCols + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
                .Cells(1, 1).Resize(intRows + 1, intCols + 1).Interior.ColorIndex = 0
                .Cells(1, 1).Resize(intRows + 1, intCols + 1).Font.ColorIndex = 1
                gcbTagsUpdated = gcbTagsUpdated + 1
        End Select

        End With

End If
End If

Erase arrFinalDataA()
Set WS = Nothing: Set Wb = Nothing
Exit Sub
ErrHandler100:
Resume MyNextLine:
MyNextLine:
With WS.Range(pDestWSRange)
    If bTranspose Then
        .Resize(1, 24).Interior.ColorIndex = 3
        .Resize(1, 24).Font.ColorIndex = 2
        .Resize(1, 24).Font.Bold = True
    Else
        .Resize(24, 1).Interior.ColorIndex = 3
        .Resize(24, 1).Font.ColorIndex = 2
        .Resize(24, 1).Font.Bold = True
    End If
End With

gcbTagsProblems = gcbTagsProblems + 1

Set WS = Nothing: Set Wb = Nothing

End Sub

Sub fImpRangeWB2()
Dim rsData As ADODB.Recordset
Dim sConnect As String, sSQL As String, MyRng As String
Dim r As Integer, C As Integer
Dim intRows As Long, intCols As Long, i As Long
Dim arrFinalDataA() As Variant, arrData() As Variant
Dim Wb As Workbook, WS As Worksheet
Dim bTranspose As Boolean, bNegative As Boolean, IsClosed As Boolean

If Len(pDestWBName) = 0 Then
Set Wb = Selection.Parent.Parent
Set WS = Wb.Worksheets(pDestWSName)
Else
Set Wb = Workbooks(pDestWBName)
Set WS = Wb.Worksheets(pDestWSName)
End If
                                                                
On Error GoTo ErrHandler100:
                                                                
IsClosed = WaitForFileClose(FileName:=pSourceWBNameA, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
If IsClosed = False Then GoTo MyNextLine:
                                                                
                                                                
                                                                'Need to set HDR equal to No
sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & pSourceWBNameA & ";" & _
"Extended Properties=" & """Excel 12.0 Macro;HDR=NO""" & ";"

MyRng = "[" & pSourceWSName & pSourceWSRangeA & "];"

sSQL = "Select* from " & MyRng
Set rsData = New ADODB.Recordset
rsData.Open sSQL, sConnect, adOpenForwardOnly, _
adLockReadOnly, adCmdText

If Not rsData.EOF Then                                          'Make Sure we got records back
    arrData = rsData.GetRows                                    'Assign to an array the Records Retrieve from the sSQL
    intRows = UBound(arrData, 2)
    intCols = UBound(arrData, 1)
    ReDim arrFinalDataA(1 To intRows + 1, 1 To intCols + 1)
        For r = 0 To intRows
            For C = 0 To intCols
                If IsNumeric(arrData(C, r)) Then
                arrFinalDataA(r + 1, C + 1) = CDbl(arrData(C, r))
                Else
                arrFinalDataA(r + 1, C + 1) = arrData(C, r)
                End If
            Next C
        Next r

If rsData.State <> adStateClosed Then
rsData.Close
End If
If Not rsData Is Nothing Then Set rsData = Nothing
Erase arrData()

If IsArrayEmpty(arrFinalDataA()) Then GoTo MyNextLine:

IsClosed = WaitForFileClose(FileName:=pSourceWBNameB, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
If IsClosed = False Then GoTo MyNextLine:


sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & pSourceWBNameB & ";" & _
"Extended Properties=" & """Excel 12.0 Macro;HDR=NO""" & ";"

MyRng = "[" & pSourceWSName & pSourceWSRangeB & "];"

sSQL = "Select* from " & MyRng
Set rsData = New ADODB.Recordset
rsData.Open sSQL, sConnect, adOpenForwardOnly, _
adLockReadOnly, adCmdText

If Not rsData.EOF Then                                          'Make Sure we got records back
    arrData = rsData.GetRows                                    'Assign to an array the Records Retrieve from the sSQL
    intRows = UBound(arrData, 2)
    intCols = UBound(arrData, 1)
    ReDim arrFinalDataB(1 To intRows + 1, 1 To intCols + 1)
        For r = 0 To intRows
            For C = 0 To intCols
                If IsNumeric(arrData(C, r)) Then
                arrFinalDataB(r + 1, C + 1) = CDbl(arrData(C, r))
                Else
                arrFinalDataB(r + 1, C + 1) = arrData(C, r)
                End If
            Next C
        Next r
End If
If rsData.State <> adStateClosed Then
rsData.Close
End If
If Not rsData Is Nothing Then Set rsData = Nothing
Erase arrData()

If IsArrayEmpty(arrFinalDataB()) Then GoTo MyNextLine:
                                                                    'to create delta take data from primary wb and deduct data from secondary wb
For i = 1 To UBound(arrFinalDataA())
arrFinalDataA(i, 1) = arrFinalDataA(i, 1) - arrFinalDataB(i, 1)
Next i

Erase arrFinalDataB()

If Len(pDestWBName) = 0 Then
        bTranspose = False
        bNegative = False
        With WS.Range(pDestWSRange)
            If .Columns.Count = 2 And .Rows.Count = 1 Then bTranspose = True
            If .Columns.Count = 3 And .Rows.Count = 1 Then
                bTranspose = True
                bNegative = True
            End If

            If .Columns.Count = 1 And .Rows.Count = 3 Then
                bTranspose = False
                bNegative = True
            End If
    
            If bNegative = True Then
                For i = 1 To UBound(arrFinalDataA)
                    arrFinalDataA(i, 1) = -arrFinalDataA(i, 1)
                Next i
            End If

            Select Case bTranspose
                Case Is = True
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).Value = Application.WorksheetFunction.Transpose(arrFinalDataA)
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).Interior.ColorIndex = 0
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).Font.ColorIndex = 1
                    gcbTagsUpdated = gcbTagsUpdated + 1
                Case Is = False
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).Value = arrFinalDataA
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).Interior.ColorIndex = 0
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).Font.ColorIndex = 1
                    gcbTagsUpdated = gcbTagsUpdated + 1
            End Select
        End With
Else

        bTranspose = pboolHorizontal
        bNegative = pboolChngeNeg
        With WS.Range(pDestWSRange)
            If bNegative = True Then
                For i = 1 To UBound(arrFinalDataA)
                    arrFinalDataA(i, 1) = -arrFinalDataA(i, 1)
                Next i
            End If

            Select Case bTranspose
                Case Is = True
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).Value = Application.WorksheetFunction.Transpose(arrFinalDataA)
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).Interior.ColorIndex = 0
                    .Cells(1, 1).Resize(intCols + 1, intRows + 1).Font.ColorIndex = 1
                    gcbTagsUpdated = gcbTagsUpdated + 1
                    
                Case Is = False
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).Value = arrFinalDataA
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).NumberFormat = "#,##0.00;[Red](#,##0.00)"
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).Interior.ColorIndex = 0
                    .Cells(1, 1).Resize(intRows + 1, intCols + 1).Font.ColorIndex = 1
                    gcbTagsUpdated = gcbTagsUpdated + 1
            End Select
        End With
End If
End If


Erase arrData(): Erase arrFinalDataA()
Set WS = Nothing: Set Wb = Nothing
Exit Sub

ErrHandler100:
Resume MyNextLine:
MyNextLine:
With WS.Range(pDestWSRange)
    If bTranspose Then
        .Resize(1, 24).Interior.ColorIndex = 3
        .Resize(1, 24).Font.ColorIndex = 2
        .Resize(1, 24).Font.Bold = True
    Else
        .Resize(24, 1).Interior.ColorIndex = 3
        .Resize(24, 1).Font.ColorIndex = 2
        .Resize(24, 1).Font.Bold = True
    End If
End With
gcbTagsProblems = gcbTagsProblems + 1
Set WS = Nothing: Set Wb = Nothing

End Sub

Private Function IsArrayEmpty(MyArray As Variant) As Boolean
Dim Element As Variant
    
    For Each Element In MyArray
        If IsEmpty(Element) Or IsNull(Element) Then
           IsArrayEmpty = True
           Exit Function
        End If
    Next Element
End Function

'******************************************************************************************************
'Chip Pearson code for checking if file is open

Private Function WaitForFileClose(FileName As String, ByVal TestIntervalMilliseconds As Long, _
    ByVal TimeOutMilliseconds As Long) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WaitForFileClose
' This function tests to see if a specified file is open. If the file is not
' open, it returns a value of True and exists immediately. If FileName is
' open, the code goes into a wait loop, testing whether the is still open
' every TestIntervalMilliSeconds. If the is closed while the function is
' waiting, the function exists with a result of True. If TimeOutMilliSeconds
' is reached and file remains open, the function exits with a result of
' False. The function will return True is FileName does not exist.
' If TimeOutMilliSeconds is reached and the file remains open, the function
' returns False.
' If FileName refers to a workbook that is open Shared, the function returns
' True and exits immediately.
' This function requires the IsFileOpen function and the Sleep and GetTickCount
' API functions.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim StartTickCount As Long
Dim EndTickCount As Long
Dim TickCountNow As Long
Dim FileIsOpen As Boolean
Dim Done As Boolean
Dim CancelKeyState As Long

'''''''''''''''''''''''''''''''''''''''''''''''
' Before we do anything, first test if the file
' is open. If it is not, get out immediately.
'''''''''''''''''''''''''''''''''''''''''''''''
FileIsOpen = IsFileOpen(FileName:=FileName)
If FileIsOpen = False Then
    WaitForFileClose = True
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If TestIntervalMilliseconds <= 0, use a default value of 500.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If TestIntervalMilliseconds <= 0 Then
    TestIntervalMilliseconds = 500
End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Here, we save the state of EnableCancelKey, and set it to
' xlErrorHandler. This will cause an error 18 to raised if the
' user press CTLR+BREAK. In this case, we'll abort the wait
' procedure and return False.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CancelKeyState = Application.EnableCancelKey
Application.EnableCancelKey = xlErrorHandler
On Error GoTo ErrHandler:

'''''''''''''''''''''''''''''''
' Get the current tick count.
'''''''''''''''''''''''''''''''
StartTickCount = GetTickCount()
If TimeOutMilliseconds <= 0 Then
    ''''''''''''''''''''''''''''''''''''''''
    ' If TimeOutMilliSeconds is negative,
    ' we'll wait forever.
    ''''''''''''''''''''''''''''''''''''''''
    EndTickCount = -1
Else
    ''''''''''''''''''''''''''''''''''''''''
    ' If TimeOutMilliseconds > 0, get the
    ' tick count value at which we will
    ' give up on the wait and return
    ' false.
    ''''''''''''''''''''''''''''''''''''''''
    EndTickCount = StartTickCount + TimeOutMilliseconds
End If

Done = False
Do Until Done
    ''''''''''''''''''''''''''''''''''''''''''''''''
    ' Test if the file is open. If it is closed,
    ' exit with a result of True.
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If IsFileOpen(FileName:=FileName) = False Then
        WaitForFileClose = True
        Application.EnableCancelKey = CancelKeyState
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''
    ' Go to sleep for TestIntervalMilliSeconds
    ' milliseconds.
    '''''''''''''''''''''''''''''''''''''''''
    Sleep dwMilliseconds:=TestIntervalMilliseconds
    TickCountNow = GetTickCount()
    If EndTickCount > 0 Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' If EndTickCount > 0, a specified timeout
        ' value was provided. Test if we have
        ' exceeded the time. Do one last test for
        ' FileOpen, and exit.
        '''''''''''''''''''''''''''''''''''''''''''
        If TickCountNow >= EndTickCount Then
            WaitForFileClose = Not (IsFileOpen(FileName))
            Application.EnableCancelKey = CancelKeyState
            Exit Function
        Else
            '''''''''''''''''''''''''''''''''''''''''
            ' TickCountNow is less than EndTickCount,
            ' so continue to wait.
            '''''''''''''''''''''''''''''''''''''''''
        End If
    Else
        ''''''''''''''''''''''''''''''''
        ' EndTickCount < 0, meaning wait
        ' forever. Test if the file
        ' is open. If the file is not
        ' open, exit with a TRUE result.
        ''''''''''''''''''''''''''''''''
        If IsFileOpen(FileName:=FileName) = False Then
            WaitForFileClose = True
            Application.EnableCancelKey = CancelKeyState
            Exit Function
        End If
        
    End If
    DoEvents
Loop

'''''''''''''''''''''''''''''''''''''''''''''''''''
' The following Exit Function line will never be
' executed, but it included for logical consistency.
' The return code is set and the function is
' terminated within the Loop above.
'''''''''''''''''''''''''''''''''''''''''''''''''''
Exit Function

ErrHandler:
'''''''''''''''''''''''''''''''''''
' This is the error handler block.
' For any error, return False.
'''''''''''''''''''''''''''''''''''
Application.EnableCancelKey = CancelKeyState
WaitForFileClose = False

End Function


Private Function IsFileOpen(FileName As String) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsFileOpen
' By Chip Pearson www.cpearson.com/excel chip@cpearson.com
' This function determines whether a file is open by any program. Returns TRUE or FALSE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim FileNum As Integer
Dim ErrNum As Integer

On Error Resume Next   ' Turn error checking off.

'''''''''''''''''''''''''''''''''''''''''''
' If we were passed in an empty string,
' there is no file to test so return FALSE.
'''''''''''''''''''''''''''''''''''''''''''
If FileName = vbNullString Then
    IsFileOpen = False
    Exit Function
End If

'''''''''''''''''''''''''''''''
' If the file doesn't exist,
' it isn't open so get out now.
'''''''''''''''''''''''''''''''
If Dir(FileName) = vbNullString Then
    IsFileOpen = False
    Exit Function
End If
''''''''''''''''''''''''''
' Get a free file number.
''''''''''''''''''''''''''
FileNum = FreeFile()
'''''''''''''''''''''''''''
' Attempt to open the file
' and lock it.
'''''''''''''''''''''''''''
Err.Clear
Open FileName For Input Lock Read As #FileNum
''''''''''''''''''''''''''''''''''''''
' Save the error number that occurred.
''''''''''''''''''''''''''''''''''''''
ErrNum = Err.Number
On Error GoTo 0        ' Turn error checking back on.
Close #FileNum       ' Close the file.
''''''''''''''''''''''''''''''''''''
' Check to see which error occurred.
''''''''''''''''''''''''''''''''''''
Select Case ErrNum
    Case 0
    '''''''''''''''''''''''''''''''''''''''''''
    ' No error occurred.
    ' File is NOT already open by another user.
    '''''''''''''''''''''''''''''''''''''''''''
        IsFileOpen = False

    Case 70
    '''''''''''''''''''''''''''''''''''''''''''
    ' Error number for "Permission Denied."
    ' File is already opened by another user.
    '''''''''''''''''''''''''''''''''''''''''''
        IsFileOpen = True

    '''''''''''''''''''''''''''''''''''''''''''
    ' Another error occurred. Assume the file
    ' cannot be accessed.
    '''''''''''''''''''''''''''''''''''''''''''
    Case Else
        IsFileOpen = True
        
End Select

End Function


