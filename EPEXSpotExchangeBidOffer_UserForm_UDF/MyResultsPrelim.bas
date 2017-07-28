Attribute VB_Name = "MyResultsPrelim"
Option Explicit

'***********************************
'EPEX_PowerBidOffer_Template.xlsm

'20170727_MarketResult_Comp_CH.xls is the file that one downloads it from the european power exchange (EPEX)
'on behalf of his company. Originally it's a csv file. It's being converted for simplicity to an excel file.
'It is assumed on this workbook that one has downloaded it already and placed it on the folder indicated on MyLists sheet.
'The file you retrieve it on MarketResults spreadsheet via MyDAMResultFile procedure below.
'***********************************

Sub MyResultsCalc()

Dim j, i As Integer 'where j is considered to be hour
Dim ResultWS As Worksheet
Dim MyWS As Worksheet
Dim MyPrice As Double, MyQuantity As Double, MyBidHour As Double, MyBidQuantity As Double, MyBidPrice As Double
Dim MyBidBook As String
Dim lastrow As Long
Dim BookA As Double, BookB As Double, BookC As Double
Dim MyTotal As Double
Dim MyMarket As String
Dim MyAPPP As Double
Dim counter As Integer, checkcounter As Integer, MyRounding As Integer, MyListRow As Integer
Dim MyLinearConstant As Double
Dim MyProSubTotal As Integer, MyFinalCounterLimit As Integer, MyFinalCounter As Integer, finalcheckcounter As Integer
Dim finalcheckcounterx As Integer, finalcheckcountery As Integer


Set ResultWS = ThisWorkbook.Worksheets("MarketResults")

With ResultWS
    .Range(Cells(40, 2), Cells(63, 5)).ClearContents
    .Range(Cells(40, 2), Cells(63, 5)).Interior.ColorIndex = 0
    .Range(Cells(67, 1), Cells(500, 5)).ClearContents
    .Range(Cells(67, 1), Cells(500, 5)).Interior.ColorIndex = 0
    .Cells(38, 1).Value = "Preliminary Iteration Results - Do Not Use - for comparison purposes only"
End With


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual


Set MyWS = ThisWorkbook.Worksheets("MyTemplate")

MyMarket = MyWS.Cells(2, 2).Value

If MyMarket = "France" Then
    'MyAPPP = 0.01
    MyAPPP = ThisWorkbook.Worksheets("MyLists").Range("F2").Value '0.1
    MyRounding = 2
    MyLinearConstant = 100
Else
    MyAPPP = ThisWorkbook.Worksheets("MyLists").Range("F2").Value '0.1
    MyRounding = 1
    MyLinearConstant = 10
End If

With MyWS
    lastrow = .Cells(.Rows.Count, "B").End(xlUp).Row
End With

For j = 1 To 24

BookA = 0
BookB = 0
BookC = 0
MyTotal = 0

MyPrice = ResultWS.Cells(3 + j, 2).Value
MyQuantity = ResultWS.Cells(3 + j, 4).Value

For i = 6 To lastrow

    MyBidPrice = MyWS.Cells(i, 3).Value
    MyBidQuantity = MyWS.Cells(i, 2).Value
    MyBidBook = MyWS.Cells(i, 4).Value
    MyBidHour = MyWS.Cells(i, 1).Value

If MyBidHour = j Then

Select Case MyBidBook

Case "Scheduling"
    If MyBidQuantity > 0 And MyBidPrice >= MyPrice Then
        BookA = MyBidQuantity + BookA
'    ElseIf MyBidQuantity < 0 And MyBidPrice <= MyPrice + MyAPPP Then
    ElseIf MyBidQuantity < 0 And MyBidPrice + MyAPPP <= MyPrice Then
        BookA = MyBidQuantity + BookA
    Else
    End If

Case "Italian"
    If MyBidQuantity > 0 And MyBidPrice >= MyPrice Then
        BookB = MyBidQuantity + BookB
'    ElseIf MyBidQuantity < 0 And MyBidPrice <= MyPrice + MyAPPP Then
    ElseIf MyBidQuantity < 0 And MyBidPrice + MyAPPP <= MyPrice Then
        BookB = MyBidQuantity + BookB
    Else
    End If
    
Case "Continental"
    If MyBidQuantity > 0 And MyBidPrice >= MyPrice Then
        BookC = MyBidQuantity + BookC
'    ElseIf MyBidQuantity < 0 And MyBidPrice <= MyPrice + MyAPPP Then
    ElseIf MyBidQuantity < 0 And MyBidPrice + MyAPPP <= MyPrice Then
        BookC = MyBidQuantity + BookC
    Else
    End If

End Select

MyTotal = BookA + BookB + BookC

Else
End If

Next i

With ResultWS
    .Cells(39 + j, 2).Value = BookA
    .Cells(39 + j, 3).Value = BookB
    .Cells(39 + j, 4).Value = BookC
    .Cells(39 + j, 5).Value = MyTotal
End With

Next j

'Check we won what we expected or have we been partially acccepted
MyListRow = 0

For counter = 1 To 24

If ResultWS.Cells(39 + counter, 5).Value <> ResultWS.Cells(3 + counter, 4).Value Then
'were there interpolated bids or just firm bids

For checkcounter = 1 To lastrow - 5
' so we loop thru the bids, if hours match, and my bid price and actual bid price rounded down with buy and rounded up with sale
' match then this was prorated.

'BUY
If MyWS.Cells(5 + checkcounter, 1).Value = counter And MyWS.Cells(5 + checkcounter, 2).Value > 0 And MyWS.Cells(5 + checkcounter, 3).Value = WorksheetFunction.RoundDown(ResultWS.Cells(3 + counter, 2), MyRounding) Then
'If MyWS.Cells(5 + checkcounter, 1).Value = counter And MyWS.Cells(5 + checkcounter, 2).Value > 0 And MyWS.Cells(5 + checkcounter, 3).Value = WorksheetFunction.Round(ResultWS.Cells(3 + counter, 2), MyRounding) Then
MyListRow = MyListRow + 1
With ResultWS
    .Cells(66 + MyListRow, 1).Value = checkcounter 'bid
    .Cells(66 + MyListRow, 2).Value = counter 'hour
    .Cells(66 + MyListRow, 3).Value = MyWS.Cells(5 + checkcounter, 4).Value 'book
    .Cells(66 + MyListRow, 4).Value = MyWS.Cells(5 + checkcounter, 2).Value 'quantity
    .Cells(66 + MyListRow, 5).Value = MyWS.Cells(5 + checkcounter, 3).Value 'price
    .Cells(66 + MyListRow, 6).Value = MyWS.Cells(5 + checkcounter, 6).Value 'APP
    .Cells(66 + MyListRow, 7).Value = .Cells(3 + counter, 2).Value 'DAM
 '   .Cells(66 + MyListRow, 10).Value = (ResultWS.Cells(3 + counter, 2).Value - MyWS.Cells(5 + checkcounter, 3).Value) * MyLinearConstant * MyWS.Cells(5 + checkcounter, 2).Value
    .Cells(66 + MyListRow, 10).Value = MyWS.Cells(5 + checkcounter, 2).Value - (.Cells(3 + counter, 2).Value - MyWS.Cells(5 + checkcounter, 3).Value) * MyLinearConstant * MyWS.Cells(5 + checkcounter, 2).Value
    .Cells(66 + MyListRow, 11).Value = WorksheetFunction.Round(.Cells(66 + MyListRow, 10), 0)
    MyProSubTotal = ResultWS.Cells(66 + MyListRow, 11) + MyProSubTotal
End With

'SELL
ElseIf MyWS.Cells(5 + checkcounter, 1).Value = counter And MyWS.Cells(5 + checkcounter, 2).Value < 0 And MyWS.Cells(5 + checkcounter, 3).Value = WorksheetFunction.RoundDown(ResultWS.Cells(3 + counter, 2), MyRounding) Then
MyListRow = MyListRow + 1

With ResultWS
    .Cells(66 + MyListRow, 1).Value = checkcounter 'bid
    .Cells(66 + MyListRow, 2).Value = counter 'hour
    .Cells(66 + MyListRow, 3).Value = MyWS.Cells(5 + checkcounter, 4).Value 'book
    .Cells(66 + MyListRow, 4).Value = MyWS.Cells(5 + checkcounter, 2).Value 'quantity
    .Cells(66 + MyListRow, 5).Value = MyWS.Cells(5 + checkcounter, 3).Value 'price
    .Cells(66 + MyListRow, 6).Value = MyWS.Cells(5 + checkcounter, 6).Value 'APP
    .Cells(66 + MyListRow, 7).Value = .Cells(3 + counter, 2).Value 'DAM
   ' .Cells(66 + MyListRow, 10).Value = (MyWS.Cells(5 + checkcounter, 6).Value - .Cells(3 + counter, 2).Value) * MyLinearConstant * MyWS.Cells(5 + checkcounter, 2).Value
    .Cells(66 + MyListRow, 10).Value = MyWS.Cells(5 + checkcounter, 2).Value - (MyWS.Cells(5 + checkcounter, 6).Value - .Cells(3 + counter, 2).Value) * MyLinearConstant * MyWS.Cells(5 + checkcounter, 2).Value
    .Cells(66 + MyListRow, 11).Value = WorksheetFunction.Round(.Cells(66 + MyListRow, 10), 0)
    MyProSubTotal = .Cells(66 + MyListRow, 11) + MyProSubTotal
End With

Else
End If

Next checkcounter

'MsgBox MyProSubTotal

ResultWS.Cells(66 + MyListRow, 12).Value = MyProSubTotal
ResultWS.Cells(66 + MyListRow, 12).Interior.ColorIndex = 6

Else
'Nothing
End If
'Loop

Next counter

'so now go through final check to update preliminary table and convert to final results table

MyFinalCounterLimit = MyListRow

'MsgBox MyFinalCounterLimit

ResultWS.Range(Cells(66 + MyListRow + 3, 1), Cells(66 + MyListRow + 3 + 25, 5)).Value = ResultWS.Range(Cells(38, 1), Cells(63, 5)).Value
ResultWS.Cells(38, 1).Value = "Final Book Allocation Results"

For MyFinalCounter = 1 To MyFinalCounterLimit

Select Case ResultWS.Cells(66 + MyFinalCounter, 3).Value

'Why plus 39.does it need any paste special values add or subtract?
Case "Italian"
With ResultWS
'    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 3).Value = .Cells(66 + MyFinalCounter, 11).Value
    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 3).Value = .Cells(66 + MyFinalCounter, 11).Value + .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 3).Value
    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 3).Interior.ColorIndex = 6
    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 5).Value = .Cells(66 + MyFinalCounter, 12).Value + .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 5).Value
'    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 5).Value = .Cells(66 + MyFinalCounter, 12).Value
End With


Case "Continental"
With ResultWS
    '.Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 4).Value = .Cells(66 + MyFinalCounter, 11).Value
    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 4).Value = .Cells(66 + MyFinalCounter, 11).Value + .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 4).Value
    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 4).Interior.ColorIndex = 6
    .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 5).Value = .Cells(66 + MyFinalCounter, 12).Value + .Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 5).Value
    '.Cells(.Cells(66 + MyFinalCounter, 2).Value + 39, 5).Value = .Cells(66 + MyFinalCounter, 12).Value
End With

End Select


Next MyFinalCounter

'Final Check of Totals

For finalcheckcounter = 1 To 24
    If ResultWS.Cells(39 + finalcheckcounter, 5).Value = ResultWS.Cells(3 + finalcheckcounter, 4).Value Then
    ResultWS.Cells(39 + finalcheckcounter, 5).Interior.ColorIndex = 4
Else
    ResultWS.Cells(39 + finalcheckcounter, 5).Interior.ColorIndex = 2
End If
Next finalcheckcounter


With ResultWS
    lastrow = .Cells(.Rows.Count, "B").End(xlUp).Row - 24
End With

'Highlight where preliminary iteration is wrong


'For finalcheckcounterx = 1 To 14
For finalcheckcounterx = 1 To 4
For finalcheckcountery = 1 To 24
    If ResultWS.Cells(39 + finalcheckcountery, finalcheckcounterx + 1).Value <> ResultWS.Cells(lastrow + finalcheckcountery, finalcheckcounterx + 1).Value Then
    ResultWS.Cells(lastrow + finalcheckcountery, finalcheckcounterx + 1).Interior.ColorIndex = 3
    Else
    End If
Next finalcheckcountery
Next finalcheckcounterx

'Insert a Formula for the Totals
ResultWS.Range("E40:E63").FormulaR1C1 = "=RC2+RC3+RC4"

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Set ResultWS = Nothing
Set MyWS = Nothing

End Sub

Sub MyDAMResultFile()

Dim Yeardata As Long
Dim MyLink, isitthere, DeliveryDate As String
Dim WB As Workbook
Dim DAMLastRow As Integer

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
End With

'Find the year
Yeardata = Year(ThisWorkbook.Worksheets("MyTemplate").Range("B3").Value)

'Delivery Date in YYYYMMDD Format
DeliveryDate = Format(ThisWorkbook.Worksheets("MyTemplate").Range("B3").Value, "YYYYMMDD")

Dim Damsfx As String, Company As String
Damsfx = MyDAMSuffix(ThisWorkbook.Worksheets("MyTemplate").Range("B2"))
Company = "Comp"

'IL Filepath concatenate with Filename
MyLink = ThisWorkbook.Worksheets("MyLists").Range("FolderPathtoUse").Value & _
IIf(Right(ThisWorkbook.Worksheets("MyLists").Range("FolderPathtoUse").Value, 1) = "\", "", "\") _
& DeliveryDate & "_MarketResult_" & Company & "_" & Damsfx & ".xls"

'Is there any Filename as the one you stated above. Dir is the Directory
isitthere = Dir(MyLink)
    
'if there is then open the specified workbook
If isitthere <> "" Then
      
'Set the Object as readonly and updatelinks
Set WB = Workbooks.Open(MyLink, True, True)

With WB.Worksheets(1)

'find the last row on the worksheet
'''''DAMLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

'Copy paste the Range to Your Worksheet
ThisWorkbook.Worksheets("MarketResults").Range("A1:L34").Value = .Range("A30:L63").Value

'''''ThisWorkbook.Worksheets("MarketResults").Range("A1:L" & DAMLastRow).Value = _
'''''.Range("A1:L" & DAMLastRow).Value 'Do not uncomment it on MarketResults worksheet

End With

WB.Close
Set WB = Nothing

Else

MsgBox "NO DAM file has been found for the Delivery Day " & DeliveryDate
ThisWorkbook.Worksheets("MarketResults").Range("A1:L34").ClearContents

End If

With Application
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
End With

End Sub

