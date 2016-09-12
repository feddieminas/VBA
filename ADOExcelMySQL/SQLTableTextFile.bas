Attribute VB_Name = "SQLTableTextFile"
Option Explicit
Option Base 1

''' are the Comments
'   are the Comments to uncomment depending on your selection

Sub MySQLTxtFileRowData() '''Create your Text file with Records to insert on the Database
                          '''Each column separated by Tab
Dim sLine As String
Dim sFName As String
Dim iFNumber As Integer
Dim MyPathway As String
Dim i As Long, j As Long, h As Long, arrBorders() As Variant, arrBuyPrices() As Variant, arrSellPrices() As Variant
Dim arrTblData() As Variant
Dim start As Double, finish As Double, totaltime As Double

start = Timer

Const sTab As String = vbTab

MyPathway = ThisWorkbook.Path

sFName = MyPathway & "\" & "MyTableData" & ".txt"

If FolderFileExists(sFName) Then Kill sFName

iFNumber = FreeFile

'''Create a text file with Random Table Data

ReDim arrTblData(100, 7)
j = 1: h = 1
arrBorders = Array("DECH", "CHDE", "FRCH", "CHFR")

arrSellPrices = Array(29.93, 24.32, 23.6, 23, 23.04, 24.93, 32.79, 38.68, 40.34, 38.96, 34.76, 36.11, 33.97, 35.93, 32.93, 32.91, 30.98, 35.77, 39.73, 42.03, 40.82, 38.01, 32.6, 32.52, 0)
arrBuyPrices = Array(22.92, 17.31, 16.59, 15.99, 16.03, 17.92, 25.78, 31.67, 33.33, 31.95, 27.75, 29.1, 26.96, 28.92, 25.92, 25.9, 23.97, 28.76, 32.72, 35.02, 33.81, 31, 25.59, 25.51, 0)

On Error GoTo myend

For i = 1 To UBound(arrTblData)
arrTblData(i, 1) = j                                            '''IDIndex

'''1)
'''DDate declared as Number to MySQL Table

arrTblData(i, 2) = CLng(Format(Now + j - 1, "YYYYMMDD"))        '''Date

'OR

'''2)
'''DDate declared as Date to MySQL Table

'arrTblData(i, 2) = CLng(Format(Now + j - 1, "YYYYMMDD"))        '''Date
'arrTblData(i, 2) = CVar(Left(arrTblData(i, 2), 4) & "-" & Mid(arrTblData(i, 2), 5, 2) & "-" & Right(arrTblData(i, 2), 2))


arrTblData(i, 3) = h                                            '''HOUR
arrTblData(i, 4) = arrBorders(j)                                '''BORDER

If j = 1 Xor j = 4 Then
arrTblData(i, 5) = "SELL"                                       '''PURPOSE
arrTblData(i, 6) = IIf(h = 25, 0, 10)                           '''QTY
arrTblData(i, 7) = IIf(i > 50 And h <> 25, CDbl(arrSellPrices(h) + 0.21), CDbl(arrSellPrices(h)))   '''PRICE

Else
arrTblData(i, 5) = "BUY"                                                                       '''PURPOSE
arrTblData(i, 6) = IIf(h = 25, 0, 50)                                                          '''QTY
arrTblData(i, 7) = IIf(i > 50 And h <> 25, CDbl(arrBuyPrices(h) - 0.21), CDbl(arrBuyPrices(h)))      '''PRICE

End If

h = h + 1
If h = 26 Then
h = 1
j = j + 1
End If
Next i
Erase arrBorders(): Erase arrSellPrices(): Erase arrBuyPrices()

Open sFName For Append As #iFNumber
                                                                                                        
    For i = 1 To UBound(arrTblData)
                                                                                                        
    '''IDINDEX DDATE HOUR
    sLine = arrTblData(i, 1) & sTab & arrTblData(i, 2) & sTab & arrTblData(i, 3) & sTab
    '''BORDER PURPOSE
    sLine = sLine & arrTblData(i, 4) & sTab & arrTblData(i, 5) & sTab
    
    '''QTY
    If Application.DecimalSeparator = "," And Int(arrTblData(i, 6)) <> arrTblData(i, 6) Then
    sLine = sLine & Replace(Format(arrTblData(i, 6), "0.00"), ",", ".") & sTab
    Else
    sLine = sLine & arrTblData(i, 6) & sTab
    End If
    
    '''PRICE
    If Application.DecimalSeparator = "," And Int(arrTblData(i, 7)) <> arrTblData(i, 7) Then
    sLine = sLine & Replace(Format(arrTblData(i, 7), "0.00"), ",", ".") & sTab
    Else
    sLine = sLine & arrTblData(i, 7) & sTab
    End If

    Print #iFNumber, sLine

    Next i
    
Close #iFNumber

MsgBox "MySQL Text File Created on " & MyPathway

finish = Timer
totaltime = Format(finish - start, "0.00")

'''Insert the Number of Seconds
ThisWorkbook.Worksheets("Dashboard").Range("E7").Value = totaltime

Exit Sub

myend:
On Error Resume Next '''MyErrorHandling
Close #iFNumber
On Error GoTo 0
MsgBox "MySQL Text File Not Created on " & MyPathway, vbCritical
ThisWorkbook.Worksheets("Dashboard").Range("E7").Value = ""
End Sub

Private Function FolderFileExists(FFName As String) As Boolean
Dim FFNameThere As String
FFNameThere = Dir(FFName, vbDirectory)
If FFNameThere = "" Then FFNameThere = Dir(FFName)
If FFNameThere = "" Then
FolderFileExists = False
Else
FolderFileExists = True
End If
End Function

