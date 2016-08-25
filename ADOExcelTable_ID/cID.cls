VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSourceWBName As String
Private pSourceWSName As String
Private pSourceWSRange As String
Private pDestWBName As String
Private pDestWSName As String
Private pDestWSRange As String
Private pDDate As String
Private pID As String
                                                                                                                     
Public Property Let sSourceWBName(Value As String)
pSourceWBName = Value
End Property
Public Property Let sSourceWSName(Value As String)
pSourceWSName = Value
End Property
Public Property Let sSourceWSRange(Value As String)
pSourceWSRange = "$" & Value
End Property
Public Property Let sDestWBName(Value As String)
pDestWBName = Value
End Property
Public Property Let sDestWSName(Value As String)
pDestWSName = Value
End Property
Public Property Let sDestWSRange(Value As String)
pDestWSRange = Value
End Property
Public Property Let sDDate(Value As String)
pDDate = Value
End Property
Public Property Let sID(Value As String)
pID = Value
End Property

Function fUpdWS() As Boolean
Dim sConnect As String
Dim objCommand As ADODB.Command
Dim lRecordsAffected As Long

fUpdWS = False

sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & pDestWBName & ";" & _
           "Extended Properties=""Excel 8.0;HDR=Yes;"";"
           
'Below in ''' comments is a function called WaitForFileClose written by CPearson
'Below is the link with its code. You can copy it on a separate module
'http://www.cpearson.com/excel/WaitForFileClose.htm
    
'WaitForFileClose is an excellent function to specify whether an Excel file
'has been closed or currently used by another person
           
'''Dim IsClosed As Boolean
'''IsClosed = WaitForFileClose(FileName:=pDestWBName, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
'''If IsClosed = True Then

On Error GoTo ErrHandler
           
Set objCommand = New ADODB.Command
objCommand.ActiveConnection = sConnect

objCommand.CommandText = "SELECT ID FROM [" & pDestWSName & pDestWSRange & "] " & _
"WHERE DDATE=" & pDDate

Dim Rs1 As ADODB.Recordset     'Check ID value that exists on the Database
Set Rs1 = objCommand.Execute()
Dim IDRet As Long

IDRet = CLng(Rs1.Fields.Item(0).Value)  'Create an array of values for ppl
Set Rs1 = Nothing

'Below :
'a)Assume that a Negative ID is a Deleted Record
'b)Same ID which already exists on Excel Database, no need to be updated

If IDRet >= 0 And IDRet <> CLng(pID) Then
objCommand.CommandText = "UPDATE [" & pDestWSName & pDestWSRange & "] " & _
"SET [" & "ID" & "]=" & pID & " WHERE DDATE=" & pDDate

'**************************************************************************************************
'Various : Example to update multiple columns
'Assume you have an excel database with 7 columns (Date, Hour, Zone, Sign, SellPrice, BuyPrice, Quantity)
'Date is in format YYYYMMDD
'Zone and Sign are String columns
'SellPrice, BuyPrice and Quantity are columns with Numbers
'You can update your Query in format similar as below in comments

'objCommand.CommandText = "UPDATE [" & pDestWSName & pDestWSRange & "] " & _
"SET [" & "Zone" & "]='" & "South" & _
"' ," & "[" & "Sign" & "]='" & "+" & _
"' ," & "[" & "SellPrice" & "]=" & 50.5 & _
" ," & "[" & "BuyPrice" & "]=" & 45 & _
" ," & "[" & "Quantity" & "]=" & 10 & _
" WHERE [Date]=" & pDDate & " and [Hour]=" & 1

'p.s. If you experience issues with local settings having comma as decimal separatos
'and values do not get updated use the Replace function
'Assume that 50.5 value is a value on your Excel Workbook (Not Excel Database) at Range E2 then
'Replace(Format(Range("E2").Value, "##0.0"), ",", ".")
'**************************************************************************************************

On Error Resume Next
objCommand.Execute RecordsAffected:=lRecordsAffected, Options:=adCmdText Or adExecuteNoRecords
If Err.Number = 0 Then fUpdWS = True
On Error GoTo 0
End If

If IDRet = CLng(pID) Then fUpdWS = True 'A record with same value already exists

'Reset the Recordset Object
ErrHandler:
If CBool(objCommand.ActiveConnection.State And adStateOpen) Then objCommand.ActiveConnection.Close
Set objCommand = Nothing

'''Else
'''End If

End Function

Function fRetWS() As String
Dim rsData As ADODB.Recordset
Dim sConnect As String, sSQL As String, MyRng As String
Dim r As Integer, c As Integer
Dim intRows As Long, intCols As Long
Dim arrFinalData() As Variant, arrData() As Variant
                                                                
sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & pSourceWBName & ";" & _
"Extended Properties=" & """Excel 8.0;HDR=YES""" & ";"

MyRng = "[" & pSourceWSName & pSourceWSRange & "]"

sSQL = "Select* from " & MyRng & " WHERE DDATE=" & pDDate & ";"
 
fRetWS = "NULL"

'Below in ''' comments is a function called WaitForFileClose written by CPearson
'Below is the link with its code. You can copy it on a separate module
'http://www.cpearson.com/excel/WaitForFileClose.htm
    
'WaitForFileClose is an excellent function to specify whether an Excel file
'has been closed or currently used by another person

'''Dim IsClosed As Boolean
'''IsClosed = WaitForFileClose(FileName:=pSourceWBName, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
'''If IsClosed = True Then

On Error GoTo ErrHandler

Set rsData = New ADODB.Recordset
rsData.Open sSQL, sConnect, adOpenForwardOnly, _
adLockReadOnly, adCmdText

If Not rsData.EOF Then                                          'Make Sure we got records back
    arrData = rsData.GetRows                                    'Assign to an array the Records Retrieve from the sSQL
    intRows = UBound(arrData, 2)
    intCols = UBound(arrData, 1)
    ReDim arrFinalData(1 To intRows + 1, 1 To intCols + 1)
        For r = 0 To intRows
            For c = 0 To intCols
                If IsNumeric(arrData(c, r)) Then
                arrFinalData(r + 1, c + 1) = CDbl(arrData(c, r))
                Else
                arrFinalData(r + 1, c + 1) = arrData(c, r)
                End If
            Next c
        Next r

End If

ErrHandler:
If IsArrayEmpty(arrFinalData) Then
fRetWS = "NULL"
Else
'DDATE arrFinalData(1, 1)
'ID    arrFinalData(1, 2)
fRetWS = CStr(arrFinalData(1, 2))
End If

'Reset the Recordset Object
If rsData.State <> adStateClosed Then
rsData.Close
End If

If Not rsData Is Nothing Then Set rsData = Nothing

'''End If

Erase arrData(): Erase arrFinalData()
End Function

Function fInsWS() As Boolean
Dim sConnect As String
Dim objCommand As ADODB.Command
Dim lRecordsAffected As Long

fInsWS = False

sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & pSourceWBName & ";" & _
           "Extended Properties=""Excel 8.0;HDR=Yes;"";"
           
'Below in ''' comments is a function called WaitForFileClose written by CPearson
'Below is the link with its code. You can copy it on a separate module
'http://www.cpearson.com/excel/WaitForFileClose.htm
    
'WaitForFileClose is an excellent function to specify whether an Excel file
'has been closed or currently used by another person
           
'''Dim IsClosed As Boolean
'''IsClosed = WaitForFileClose(FileName:=pDestWBName, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
'''If IsClosed = True Then

On Error GoTo ErrHandler
           
Set objCommand = New ADODB.Command
objCommand.ActiveConnection = sConnect

objCommand.CommandText = "INSERT INTO [" & pDestWSName & "$] IN '' [Excel 8.0;Database=" & pDestWBName & "] SELECT * FROM [" & pSourceWSName & pSourceWSRange & "]"
On Error Resume Next
objCommand.Execute RecordsAffected:=lRecordsAffected, Options:=adCmdText Or adExecuteNoRecords
If Err.Number = 0 Then fInsWS = True
On Error GoTo 0

'Reset the Recordset Object
ErrHandler:
If CBool(objCommand.ActiveConnection.State And adStateOpen) Then objCommand.ActiveConnection.Close
Set objCommand = Nothing

'''End If
End Function

Function fDelWS() As Boolean
Dim sConnect As String, sSQL As String, MyRng As String
Dim rsData As ADODB.Recordset
Dim lRecordsAffected As Long

fDelWS = False
           
sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & pDestWBName & ";" & _
"Extended Properties=" & """Excel 8.0;HDR=YES""" & ";"
           
'Below in ''' comments is a function called WaitForFileClose written by CPearson
'Below is the link with its code. You can copy it on a separate module
'http://www.cpearson.com/excel/WaitForFileClose.htm
    
'WaitForFileClose is an excellent function to specify whether an Excel file
'has been closed or currently used by another person
           
'''Dim IsClosed As Boolean
'''IsClosed = WaitForFileClose(FileName:=pDestWBName, _
                TestIntervalMilliseconds:=500, TimeOutMilliseconds:=10000)
'''If IsClosed = True Then

On Error GoTo ErrHandler
           
MyRng = "[" & pDestWSName & pDestWSRange & "]"

sSQL = "Select* from " & MyRng & " WHERE DDATE=" & pDDate & ";"
           
Set rsData = New ADODB.Recordset
rsData.Open sSQL, sConnect, adOpenDynamic, _
adLockPessimistic, adCmdText

rsData.MoveFirst
 Do Until rsData.EOF
 If rsData.Fields("DDATE").Value = pDDate Then
    rsData.Fields("ID").Value = -1925  'Set Delete Record to a value of -1925
    ''''rsData.Delete Doesnt Work
    rsData.Update
    If rsData.State = 1 Then fDelWS = True
 End If
    rsData.MoveNext
 Loop

'Reset the Recordset Object
ErrHandler:
If rsData.State <> adStateClosed Then
rsData.Close
End If

If Not rsData Is Nothing Then Set rsData = Nothing

'''Else
'''fDelWS = False
'''End If

End Function

Private Function IsArrayEmpty(MyArray As Variant) As Boolean
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

