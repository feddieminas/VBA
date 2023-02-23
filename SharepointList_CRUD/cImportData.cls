VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSourceSharePointURL As String
Private pSourceSharePointListGUID As String
Private pSourceSharePointListName As String
Private pDestWSName As String
Private pCollConns As Collection
                                                                                                                               
Public Property Let sSourceSharePointURL(Value As String)
pSourceSharePointURL = Value
End Property
Public Property Let sSourceSharePointListGUID(Value As String)
pSourceSharePointListGUID = Replace(Replace(Replace(Value, "%7B", ""), "%7D", ""), "%2D", "-")
End Property
Public Property Let sSourceSharePointListName(Value As String)
pSourceSharePointListName = Value
End Property
Public Property Let sDestWSName(Value As String)
pDestWSName = Value
End Property

Property Get CollConns() As Collection
    Set CollConns = pCollConns
End Property

Property Set CollConns(CollConns As Collection)
    Set pCollConns = CollConns
End Property

Public Function SetCollConns() As Collection
Dim tmpColl As Collection
Set tmpColl = New Collection

With tmpColl
    .Add CStr("Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=2;RetrieveIds=Yes;" & _
         "DATABASE=" & pSourceSharePointURL & ";LIST={" & _
          pSourceSharePointListGUID & "};")
    .Add CStr("SELECT * FROM [" & pSourceSharePointListName & "];")
End With

Set SetCollConns = tmpColl
Set tmpColl = Nothing
End Function

'********************
'INSERT
'********************
Sub fInsLst()

Dim rsData As ADODB.Recordset
Dim fldCol As ADODB.Field

Set pCollConns = SetCollConns()

Set rsData = New ADODB.Recordset
'''rsData.CursorLocation = adUseServer

'''adOpenStatic and AdLockOptimistic or adOpenDynamic and AdLockOptimistic
rsData.Open Me.CollConns(2), Me.CollConns(1), adOpenDynamic, _
adLockOptimistic, adCmdText

If rsData.Supports(adAddNew) Then

 rsData.AddNew
 rsData.Fields("Title") = "test email3"
 rsData.Fields("Person") = 19 'id of the person in the organization.
 'One can capture its id by inserting manually a row in the Sharepoint List
 'at the browser and pressing btn READ LIST on Sheet1. Column E on Person will result to a number rather than a person itself.
 'then capturing the id, one can create a dictionary to map the actual person's name.
 rsData.Fields("Task Date") = Format(Now, "DD/MM/YYYY")
 rsData.Fields("Task Assigned") = "test INSERT Record"
 rsData.Update

End If

Set pCollConns = Nothing

If rsData.State <> adStateClosed Then
rsData.Close
End If

If Not rsData Is Nothing Then Set rsData = Nothing

End Sub


'********************
'UPDATE
'********************
Sub fUpdLst()

Dim rsData As ADODB.Recordset
'Dim sConnect As String, sSQL As String
Dim fldCol As ADODB.Field

Set pCollConns = SetCollConns()

Set rsData = New ADODB.Recordset
'rsData.CursorLocation = adUseServer

'adOpenStatic and AdLockOptimistic or adOpenDynamic and AdLockOptimistic
rsData.Open Me.CollConns(2), Me.CollConns(1), adOpenDynamic, _
adLockOptimistic, adCmdText

If rsData.Supports(adUpdate) Then

 rsData.MoveFirst
 While Not rsData.EOF

 For Each fldCol In rsData.Fields
 If fldCol.Name = "Title" And fldCol.Value = "test email3" Then
 
  rsData("Task Assigned").Value = "test UPDATE Record"
  rsData.Update
 
 End If
 Next
 rsData.MoveNext
 
 Wend

End If

Set pCollConns = Nothing

If rsData.State <> adStateClosed Then
rsData.Close
End If

If Not rsData Is Nothing Then Set rsData = Nothing

End Sub

'********************
'DELETE
'********************
Sub fDelLst()

Dim rsData As ADODB.Recordset

Set pCollConns = SetCollConns()

Set rsData = New ADODB.Recordset
'rsData.CursorLocation = adUseServer

'adOpenStatic and AdLockOptimistic or adOpenDynamic and AdLockOptimistic
rsData.Open Me.CollConns(2), Me.CollConns(1), adOpenDynamic, _
adLockOptimistic, adCmdText

If rsData.Supports(adDelete) Then

 rsData.MoveFirst
 Do Until rsData.EOF
 Debug.Print rsData.Fields("ID"), rsData.Fields("Title"), rsData.Fields("Compliance Asset Id")
 Debug.Print rsData.Fields("Task Date"), rsData.Fields("Person"), rsData.Fields("Task Assigned")
 Debug.Print rsData.Fields("Content Type"), rsData.Fields("Attachments"), rsData.Fields("App Created By")
 Debug.Print rsData.Fields("App Modified By"), rsData.Fields("Workflow Instance ID"), rsData.Fields("File Type")
 Debug.Print rsData.Fields("Modified"), rsData.Fields("Created"), rsData.Fields("Created By")
 Debug.Print rsData.Fields("Modified By"), rsData.Fields("URL Path"), rsData.Fields("Path")
 Debug.Print rsData.Fields("Item Type"), rsData.Fields("Encoded Absolute URL")
 
 If rsData.Fields("Title").Value = "test email3" Then
    rsData.Delete
    rsData.Update
 End If
    rsData.MoveNext
 Loop
 
End If

Set pCollConns = Nothing

If rsData.State <> adStateClosed Then
rsData.Close
End If

If Not rsData Is Nothing Then Set rsData = Nothing

End Sub

'********************
'READ - Import
'********************
Sub fImpLst(Optional ClearTheCells As Boolean = True)

Dim rsData As ADODB.Recordset
Dim r As Integer, c As Integer
Dim WB As Workbook, WS As Worksheet
Dim arrData() As Variant, arrFinalData() As Variant
Dim intRows As Long, intCols As Long

Set WB = ThisWorkbook
Set WS = ThisWorkbook.Worksheets(pDestWSName)
If ClearTheCells Then WS.Cells.Clear

Set pCollConns = SetCollConns()

Set rsData = New ADODB.Recordset

'adOpenStatic and AdLockOptimistic or adOpenDynamic and AdLockOptimistic
rsData.Open Me.CollConns(2), Me.CollConns(1), adOpenDynamic, _
adLockOptimistic, adCmdText

'Make Sure we got records back
If Not rsData.EOF Then

arrData = rsData.GetRows
intRows = UBound(arrData, 2)
intCols = UBound(arrData, 1)

ReDim arrFinalData(1 To intRows + 2, 1 To intCols + 1)

'Insert Headers
Dim fldCol As ADODB.Field
c = 1
For Each fldCol In rsData.Fields
arrFinalData(1, c) = fldCol.Name
c = c + 1
Next

For r = 0 To intRows
For c = 0 To intCols

'Convert the Numbers to Double
If IsNumeric(arrData(c, r)) Then
arrFinalData(r + 2, c + 1) = CDbl(arrData(c, r))
Else
arrFinalData(r + 2, c + 1) = arrData(c, r)
End If

'Convert to Date the Datenumber
If IsNumeric(arrFinalData(r + 1, 1)) Then
arrFinalData(r + 2, 1) = CDate(arrFinalData(r + 1, 1))
End If

'For Negative Exponential Numbers
If (arrData(c, r) > -0.00000000001) And (arrData(c, r) < 0.00000000001) Then
arrFinalData(r + 2, c + 1) = 0
End If

Next c
Next r

WS.Range("A1").Resize(intRows + 2, intCols + 1).Value = arrFinalData
End If

Set pCollConns = Nothing

If rsData.State <> adStateClosed Then
rsData.Close
End If

If Not rsData Is Nothing Then Set rsData = Nothing

Erase arrData(): Erase arrFinalData()
Set WS = Nothing: Set WB = Nothing

End Sub

