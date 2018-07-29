Attribute VB_Name = "mod_2a_Oracle"
Option Explicit

'Worksheet : Oracle (Sheet 5)
'Takes the values from the sheet Database, inserts them on Oracle Sheet and uploads them on the SQL database.

'MI is the Intraday Market Exchange of the Power Market in Italy

'Used the Oracle InProc Server 5.0 Type Library on C:\Oracle\Ora92\product\11.2.0\client_1\bin\oip11.tlb. Need to import your Oracle version lib
'as one might have a different version. On the VBA editor, one can do it through Tab Tools-->References-->Browse for the library or select from the list

Sub MIBordersTable()

Dim arethererows As Long, OracleStartRow As Long, DatabaseLastRow As Long, DatabaseListRows As Long, anno As Long
Dim FieldNum As Long, ErrNum As Long, CCount As Long, DatesC As Long, MIsC As Long
Dim My_Table As ListObject
Dim ActiveCellInTable As Boolean
Dim market As String
Dim Rng As Range
Dim m As Integer, DatesR As Integer
Dim DataStart As String, DataFinish As String

Application.ScreenUpdating = False

    ThisWorkbook.Worksheets("Oracle").Range("C6:G65536").ClearContents
    
''''******************************************************************************''''
    
    'Sort Date Values in Ascending Order to determine the lowest and Highest Date
    With ThisWorkbook.Worksheets("MIQty")
    
    'find StartRow of DeliveryDates
    DatesR = .Range("Dates").Cells(1, 1).Row
    
    'Count Delivery Days
    DatesC = .Range("AK5").Value
    
    'Count MIMarkets
    MIsC = .Range("AK6").Value
    
    'If is onyl one Day then dont sort out the values
    If DatesC = 1 Then GoTo mynext
      
    .Range("AN1:AN51").AutoFilter
    .AutoFilter.Sort.SortFields.Add _
    Key:=.Range("AN1:AN51"), SortOn:=xlSortOnValues, Order:=xlAscending
    .AutoFilter.Sort.Header = xlYes
    .AutoFilter.Sort.Orientation = xlTopToBottom
    .AutoFilter.Sort.SortMethod = xlPinYin
    .AutoFilter.Sort.Apply
    .AutoFilter.Sort.SortFields.Clear
    .AutoFilterMode = False
    
mynext:
    
    'Format(DateSerial(2001, 1, 1), "dddd dd/mm/yyyy")
    
    'VBA due to its settings might recognise your date as either "MM/DD/YYYY" or "DD/MM/YYYY".
    'Default below is at format "MM/DD/YYYY". In case it shows a wrong date to u, then modify your DataStart and DataFinish in the format :
    'ex. DataStart = Format(ThisWorkbook.Worksheets("Oracle").Range("DeliveryDay").Value, "dd/mm/yyyy")
    'ex. DataFinish = Format(.Range("AN" & DatesR - 1).Offset(DatesC, 0).Value, "dd/mm/yyyy")
    
    'DeliveryDay to Start
    ThisWorkbook.Worksheets("Oracle").Range("DeliveryDay").Value = .Range("Dates").Cells(1, 1).Value
    DataStart = Format(ThisWorkbook.Worksheets("Oracle").Range("DeliveryDay").Value, "mm/dd/yyyy") 'If not correct date appears, then change your format to "dd/mm/yyyy"
    
    'search again for StartRow of DeliveryDates
    DatesR = .Range("Dates").Cells(1, 1).Row
    
    'DeliveryDay to Finish
    DataFinish = Format(.Range("AN" & DatesR - 1).Offset(DatesC, 0).Value, "mm/dd/yyyy") 'If not correct date appears, then change your format to "dd/mm/yyyy"
    
    'First row to start input values for Oracle
    OracleStartRow = 6
    
    End With
    
''''***********************************************************************************************''''
    
    'Create a Table on the Database
    With ThisWorkbook.Worksheets("Database")
    
        DatabaseLastRow = .Cells(.Rows.count, "A").End(xlUp).Row
        
       On Error GoTo continue
       Application.GoTo Reference:=.Range("A2")
       
         .ListObjects.Add(xlSrcRange, Range("$A$1:$E$" & DatabaseLastRow), , xlYes).Name = _
        "MIListDtbs"
       
continue:
       .ListObjects("MIListDtbs").Resize Range("$A$1:$E$" & DatabaseLastRow)
       
       .ListObjects("MIListDtbs").TableStyle = "TableStyleMedium11"
        
    Set Rng = .Range("A" & DatabaseLastRow)
    
    End With
    
''''***********************************************************************************************''''
    
    For m = 1 To MIsC
    
    'Market
    ThisWorkbook.Worksheets("Oracle").Range("Market").Value = _
    ThisWorkbook.Worksheets("MIQty").Range("MIs").Cells(m, 1).Value
    market = ThisWorkbook.Worksheets("Oracle").Range("Market").Value
    
''''************************************************************************************************''''

    'Test if rng is in a a list or Table
    On Error Resume Next
    ActiveCellInTable = (Rng.ListObject.Name <> "")
    On Error GoTo 0

    'If the cell is in a List or Table run the code
    If ActiveCellInTable = True Then

        Set My_Table = Rng.ListObject
        FieldNum = Rng.Column - My_Table.Range.Cells(1).Column + 1
        
    'Filter the range Date and Market
    My_Table.Range.AutoFilter Field:=FieldNum, Criteria1:=">=" & DataStart, _
    Operator:=xlAnd, Criteria2:="<=" & DataFinish
    
    My_Table.Range.AutoFilter Field:=FieldNum + 2, Criteria1:=market
    
    arethererows = My_Table.Range.SpecialCells(xlCellTypeVisible).Offset(1, 0).Cells.count
    
''''*************************************************************************************************''''
    
    If arethererows > 5 And DatabaseLastRow > 1 Then '5 is the number of header (columns) of the SQL table
    
    On Error GoTo fine
    
    '-1 is the HeaderCount
    DatabaseListRows = (arethererows / 5) - 1
    
   ' DatabaseListRows = My_Table.Range.SpecialCells(xlCellTypeVisible).Areas(3).Rows.Count
    'DatabaseListRows = My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Rows.Count
    
    'Copy Data and Ora
    My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Resize(DatabaseListRows, 2).Copy
    
    With ThisWorkbook.Worksheets("Oracle").Range("C" & OracleStartRow)
    .PasteSpecial xlPasteColumnWidths
    .PasteSpecial xlPasteValues
   ' .PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    End With
    
    'Copy Unita
    My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Offset(0, 3).Resize(DatabaseListRows, 1).Copy
    
    With ThisWorkbook.Worksheets("Oracle").Range("E" & OracleStartRow)
    .PasteSpecial xlPasteColumnWidths
    .PasteSpecial xlPasteValues
   ' .PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    End With
    
    'Copy Mercato
    My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Offset(0, 2).Resize(DatabaseListRows, 1).Copy
    
    With ThisWorkbook.Worksheets("Oracle").Range("F" & OracleStartRow)
    .PasteSpecial xlPasteColumnWidths
    .PasteSpecial xlPasteValues
   ' .PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    End With
    
    'Copy MWhaccetatto
    My_Table.DataBodyRange.SpecialCells(xlCellTypeVisible).Offset(0, 4).Resize(DatabaseListRows, 1).Copy
    
    With ThisWorkbook.Worksheets("Oracle").Range("G" & OracleStartRow)
    .PasteSpecial xlPasteColumnWidths
    .PasteSpecial xlPasteValues
   ' .PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    End With
    
    End If
    
''''*************************************************************************************************''''
    
    'If there will be more than one MI Market to upload then find the next start row for following market
    With ThisWorkbook.Worksheets("Oracle")
    OracleStartRow = .Cells(.Rows.count, "C").End(xlUp).Row + 1
    End With
    
''''*************************************************************************************************''''
    
    End If
    
    Next m
    
fine:

    'Unlist the table
    Application.CutCopyMode = False
    My_Table.Range.AutoFilter Field:=FieldNum
    My_Table.Range.AutoFilter Field:=FieldNum + 2
    My_Table.Unlist
  
  Set Rng = Nothing
  Set My_Table = Nothing
  
Application.GoTo Reference:=ThisWorkbook.Worksheets("Oracle").Range("C6")

Application.ScreenUpdating = True

End Sub


'SQL Table

'CREATE TABLE MEREUC_RW.IPEXUNITS
'(
'  Day     DATE NOT NULL,
'  ORA     NUMBER(2) NOT NULL,
'  UNITA   VARCHAR2(200) NOT NULL,
'  MERCATO VARCHAR2(200) NOT NULL,
'  MWHACC Number(6, 3)
')
'/

'Goto Module modSQL to view also its procedure and view

Sub OracleUpload()

Dim OraSession As OraSession
Dim OraDatabase As OraDatabase
Dim sqlStatement As OraSqlStmt
Dim SID As String, Utente As String, Password As String, Procedura As String
Dim DatabaseRows As Long
Dim i As Integer

Application.ScreenUpdating = False

SID = Range("SID").Value
Utente = Range("UTENTE").Value
Password = Range("PASSWORD").Value
Procedura = Range("PROCEDURA").Value

DatabaseRows = Cells(Rows.count, "C").End(xlUp).Row - 5

If DatabaseRows = 0 Then GoTo myfinish

Set OraSession = CreateObject("OracleInProcServer.XOraSession")

On Error GoTo wrongORAsettings

Set OraDatabase = OraSession.OpenDatabase(SID, Utente & "/" & Password, 0&)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Set ORAPARM_INPUT or ORAPARM_OUTPUT Parameters
'Set the Declaration of the Parameter .servertype

OraDatabase.Parameters.Add "DataP", Range("C6").Value, ORAPARM_INPUT
OraDatabase.Parameters("DataP").serverType = 12  'ORATYPE_DATE

OraDatabase.Parameters.Add "oraP", Range("D6").Value, ORAPARM_INPUT
OraDatabase.Parameters("oraP").serverType = 2  'ORATYPE_NUMBER

OraDatabase.Parameters.Add "UnitaP", Range("E6").Value, ORAPARM_INPUT
OraDatabase.Parameters("UnitaP").serverType = 1  'ORATYPE_VARCHAR2

OraDatabase.Parameters.Add "MercatoP", Range("F6").Value, ORAPARM_INPUT
OraDatabase.Parameters("MercatoP").serverType = 1  'ORATYPE_VARCHAR2

'Below the MWh accepted is at a float format. Thus a quantity might be partially accepted with
'decimal points. Thus if decimal on your PC is a comma then on your SQL database it will be
'entered as a 0,151 value. If your decimal is dot, then it will be 0.151
'I advise the value on the database to be on decimal dot rather than comma (i.e. 0.151).

OraDatabase.Parameters.Add "MwhaccP", Range("G6").Value, ORAPARM_INPUT
OraDatabase.Parameters("MwhaccP").serverType = 4  'ORATYPE_FLOAT

'Create the SQL Statement. Can be any

Set sqlStatement = OraDatabase.CreateSql("Begin " & Utente & "." & Procedura & _
"(:DataP, :oraP,:UnitaP,:MercatoP,:MwhaccP); end;", 0&)

'MsgBox sqlStatement.Sql

Range("C3:G3").Value = Range("C6:G6").Value

If DatabaseRows = 1 Then GoTo myfinishUpload

i = 1

For i = 1 To DatabaseRows - 1

OraDatabase.Parameters("DataP").Value = Range("C" & 6 + i).Value

OraDatabase.Parameters("oraP").Value = Range("D" & 6 + i).Value

OraDatabase.Parameters("UnitaP").Value = Range("E" & 6 + i).Value

OraDatabase.Parameters("MercatoP").Value = Range("F" & 6 + i).Value

OraDatabase.Parameters("MwhaccP").Value = Range("G" & 6 + i).Value

'Commit

sqlStatement.Refresh

Range("C3:G3").Value = Range("C" & 6 + i & ":G" & 6 + i).Value

Next i

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

myfinishUpload:

'Remove your Parameters

OraDatabase.Parameters.Remove "DataP"

OraDatabase.Parameters.Remove "oraP"

OraDatabase.Parameters.Remove "UnitaP"

OraDatabase.Parameters.Remove "MercatoP"

OraDatabase.Parameters.Remove "MwhaccP"

'Take out the Oracle Object from Oracle

Set sqlStatement = Nothing

OraDatabase.Close

Set OraDatabase = Nothing

wrongORAsettingsfin:
Set OraSession = Nothing

myfinish:

Application.ScreenUpdating = True

Exit Sub
wrongORAsettings:
MsgBox "Check your ORA Settings inserted on Sheet Settings. Have a look also that your SQL table and procedure has been inserted", vbExclamation
Resume wrongORAsettingsfin
End Sub

Sub clrcontents()

ThisWorkbook.Worksheets("Oracle").Range("C6:G65536").ClearContents

End Sub




