Attribute VB_Name = "mod_1_ExchResults"
Option Explicit

Sub GetDatafromExchITSpot() 'Retrieves data from the IPEX xlsx file you downloaded from the Exchange and saved on a folder of your choice. Default is C:\

Dim oCn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cmd As ADODB.Command
Dim strSQL As String
Dim MySourceFile As String

Application.ScreenUpdating = False

ThisWorkbook.Worksheets("ExchRes").Cells.Clear

MySourceFile = ThisWorkbook.Worksheets("Settings").Range("G6").Value & ThisWorkbook.Worksheets("Settings").Range("F8").Value

Set oCn = New ADODB.Connection
With oCn
    .Provider = "Microsoft.ACE.OLEDB.12.0;"
    .ConnectionString = "Data Source=" & MySourceFile & ";" & "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    .ConnectionTimeout = 40
    On Error Resume Next
    .Open
End With
If oCn = "" Then GoTo myEnd

'1) Retrieve all
'strSQL = "Select* from " _
& "[rptOffers$A14:X1013]"
'OR
'2) Retrieve only the Accepted Offers
strSQL = "Select* from " _
& "[rptOffers$A14:X1013]" & " where Stato='Accettato';"

Set cmd = New ADODB.Command
With cmd
    .ActiveConnection = oCn
    .CommandText = strSQL
End With

Set rs = New ADODB.Recordset
With rs
    .ActiveConnection = oCn
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open cmd
End With

Dim fldField As ADODB.Field
Dim fld As Variant, i As Integer
i = 1

For Each fld In rs.Fields
        ThisWorkbook.Worksheets("ExchRes").Cells(1, i).Value = fld.Name  'One can use also an array to store the Headers (then retrieve the array to Excel row 1) but it's ok to do this procedure for few cells
        i = i + 1
Next fld

'MsgBox rs.Fields.Item(0).Name

ThisWorkbook.Worksheets("ExchRes").Range("A2").CopyFromRecordset rs

ThisWorkbook.Worksheets("ExchRes").Columns("A:X").AutoFit

'*******************
'Filtering. Insert the values to the specified columns on your destination sheet MIQty. With...End With
'MI is the Intraday Market Exchange of the Power Market in Italy
'*******************

With ThisWorkbook
    .Worksheets("MIQty").Range("C2:C1000").Value = .Worksheets("ExchRes").Range("C2:C1000").Value 'data
    .Worksheets("MIQty").Range("E2:E1000").Value = .Worksheets("ExchRes").Range("G2:G1000").Value 'stato
    .Worksheets("MIQty").Range("G2:G1000").Value = .Worksheets("ExchRes").Range("H2:H1000").Value 'scopo
    .Worksheets("MIQty").Range("I2:I1000").Value = .Worksheets("ExchRes").Range("E2:E1000").Value 'ora
    .Worksheets("MIQty").Range("D2:D1000").Value = .Worksheets("ExchRes").Range("D2:D1000").Value 'mercato
    .Worksheets("MIQty").Range("A2:A1000").Value = .Worksheets("ExchRes").Range("A2:A1000").Value 'unità
    .Worksheets("MIQty").Range("O2:O1000").Value = .Worksheets("ExchRes").Range("K2:K1000").Value 'MIacc.
    .Worksheets("MIQty").Range("R2:R1000").Value = .Worksheets("ExchRes").Range("L2:L1000").Value 'Prezzo
End With
    
'''Kill MySourceFile

myEnd:

On Error GoTo 0
Set cmd = Nothing
Set rs = Nothing
Set oCn = Nothing

Application.ScreenUpdating = True
    
End Sub

'Notes
'https://stackoverflow.com/questions/18144838/vba-ado-connection-to-xlsx-file
'https://stackoverflow.com/questions/22920607/iterating-through-adodb-fields-inside-a-recordset-loop

