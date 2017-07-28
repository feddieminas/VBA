Attribute VB_Name = "MyTry"
Option Explicit

'***********************************
'EPEX_PowerBidOffer_Template.xlsm
'***********************************

Sub EnableFunctionality()

Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

If Not Application.CalculationState = xlDone Then
    DoEvents
End If

Application.CalculateFull

End Sub

Sub MyTriggerInput()

Dim Sourcewb As Workbook
Dim Destwb As Worksheet
Dim anno As Integer, startrow As Integer, lastrowtrigger As Integer, i As Integer, TotalTriggerRows As Integer
Dim isitthere As String, MyLink As String
Dim Filename(3) As String

With Application
.ScreenUpdating = False
.EnableEvents = False
.DisplayAlerts = False
.Calculation = xlCalculationManual
End With

Set Destwb = ThisWorkbook.Worksheets("MyTemplate")

Destwb.Range("TriggerHourlyTemplate").ClearContents

Filename(0) = "BuyRange Italian_" & Format(Destwb.Range("B3").Value, "YYYYMMDD")
Filename(1) = "SellRange Italian_" & Format(Destwb.Range("B3").Value, "YYYYMMDD")
Filename(2) = "BuyRange Continental_" & Format(Destwb.Range("B3").Value, "YYYYMMDD")
Filename(3) = "SellRange Continental_" & Format(Destwb.Range("B3").Value, "YYYYMMDD")

TotalTriggerRows = 0

For i = 0 To 3

On Error GoTo nextfile

startrow = Destwb.Cells(Destwb.Rows.Count, 1).End(xlUp).Row + 1

anno = Year(Date + 1)

MyLink = ThisWorkbook.Worksheets("MyLists").Range("FolderPathtoUse").Value & _
IIf(Right(ThisWorkbook.Worksheets("MyLists").Range("FolderPathtoUse").Value, 1) = "\", "", "\") _
& Filename(i) & ".xls"

isitthere = Dir(MyLink)
    
If isitthere <> "" Then
    
Set Sourcewb = Workbooks.Open(MyLink, True, True)

With Sourcewb.Sheets("Output")
lastrowtrigger = .Cells(.Rows.Count, 1).End(xlUp).Row
'MsgBox lastrowtrigger
TotalTriggerRows = TotalTriggerRows + lastrowtrigger - 1
'MsgBox TotalTriggerRows
If TotalTriggerRows > 971 Then lastrowtrigger = lastrowtrigger - (TotalTriggerRows - 971)
'Trigger rows 1000
'GoTo Sourcewbclose
If lastrowtrigger > 1 Then
ThisWorkbook.Worksheets("MyTemplate").Range("A" & startrow & ":D" & startrow + lastrowtrigger - 2).Value = _
.Range("A2:D" & lastrowtrigger).Value
End If
End With

Sourcewbclose:

Sourcewb.Close
Set Sourcewb = Nothing

If lastrowtrigger = 1 Then GoTo notriggers

Else

notriggers:

nextfile:

MsgBox "No Triggers for " & Filename(i)

End If

Next i

Erase Filename
Set Destwb = Nothing

With Application
.ScreenUpdating = True
.EnableEvents = True
.DisplayAlerts = True
.Calculation = xlCalculationAutomatic
End With

End Sub

Sub Letstry()
Dim i As Integer
'Dim j As Integer 'where j is the equivalent of looping through hour
Dim k As Integer ' where k is the equivalent of looping through distinct price point
Dim StartPoint As Range
Dim MyHour As Range
Dim MyQuantity As Range
Dim MyMatrixPosition As Range
Dim MyNoDistinctValues As Integer
Dim MySubTotal1 As Double, MySubTotal2 As Double, MySubTotal3 As Double, MySubTotal4 As Double, MySubTotal5 As Double
Dim MySubTotal6 As Double, MySubTotal7 As Double, MySubTotal8 As Double, MySubTotal9 As Double, MySubTotal10 As Double
Dim MySubTotal11 As Double, MySubTotal12 As Double, MySubTotal13 As Double, MySubTotal14 As Double, MySubTotal15 As Double
Dim MySubTotal16 As Double, MySubTotal17 As Double, MySubTotal18 As Double, MySubTotal19 As Double, MySubTotal20 As Double
Dim MySubTotal21 As Double, MySubTotal22 As Double, MySubTotal23 As Double, MySubTotal24 As Double, MyFirmTriggerRows As Double
Dim MyWS As Worksheet

Set MyWS = ThisWorkbook.Worksheets("MyTemplate")
MyWS.Range(Cells(5, 11), Cells(29, 1962)).ClearContents

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
'Application.Volatile (False)
'MyWS.Cells(32, 22).Value = Format(Now, "HH:MM:ss")

MyNoDistinctValues = MyWS.Cells(1, 9).Value
MyFirmTriggerRows = MyWS.Cells(1, 8).Value

For k = 1 To MyNoDistinctValues

'For j = 1 To 24

For i = 1 To MyFirmTriggerRows
'293
'50

Set StartPoint = MyWS.Cells(4 + i, 1)
Set MyHour = MyWS.Cells(4 + i, 1)
Set MyMatrixPosition = MyWS.Cells(4 + i, 5)
Set MyQuantity = MyWS.Cells(4 + i, 2)

Select Case MyHour.Value
Case 1
'3000 and buy
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal1 = MyQuantity.Value + MySubTotal1
'sell interpolation
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal1 = MyQuantity.Value + MySubTotal1
'-3000
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal1 = MyQuantity.Value + MySubTotal1
Else
End If

Case 2
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal2 = MyQuantity.Value + MySubTotal2
'If MyQuantity.Value >= 0 And MyMatrixPosition.Value >= k Then
'MySubTotal2 = MyQuantity.Value + MySubTotal2
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal2 = MyQuantity.Value + MySubTotal2
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal2 = MyQuantity.Value + MySubTotal2

Else
End If

Case 3
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal3 = MyQuantity.Value + MySubTotal3
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal3 = MyQuantity.Value + MySubTotal3
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal3 = MyQuantity.Value + MySubTotal3

Else
End If

Case 4
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal4 = MyQuantity.Value + MySubTotal4
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal4 = MyQuantity.Value + MySubTotal4
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal4 = MyQuantity.Value + MySubTotal4
Else
End If


Case 5
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal5 = MyQuantity.Value + MySubTotal5
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal5 = MyQuantity.Value + MySubTotal5
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal5 = MyQuantity.Value + MySubTotal5

Else
End If


Case 6
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal6 = MyQuantity.Value + MySubTotal6
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal6 = MyQuantity.Value + MySubTotal6
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal6 = MyQuantity.Value + MySubTotal6

Else
End If


Case 7
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal7 = MyQuantity.Value + MySubTotal7
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal7 = MyQuantity.Value + MySubTotal7
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal7 = MyQuantity.Value + MySubTotal7

Else
End If


Case 8
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal8 = MyQuantity.Value + MySubTotal8
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal8 = MyQuantity.Value + MySubTotal8
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal8 = MyQuantity.Value + MySubTotal8

Else
End If


Case 9
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal9 = MyQuantity.Value + MySubTotal9
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal9 = MyQuantity.Value + MySubTotal9
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal9 = MyQuantity.Value + MySubTotal9

Else
End If



Case 10
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal10 = MyQuantity.Value + MySubTotal10
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal10 = MyQuantity.Value + MySubTotal10
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal10 = MyQuantity.Value + MySubTotal10

Else
End If


Case 11
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal11 = MyQuantity.Value + MySubTotal11
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal11 = MyQuantity.Value + MySubTotal11
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal11 = MyQuantity.Value + MySubTotal11

Else
End If

Case 12
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal12 = MyQuantity.Value + MySubTotal12
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal12 = MyQuantity.Value + MySubTotal12
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal12 = MyQuantity.Value + MySubTotal12

Else
End If

Case 13
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal13 = MyQuantity.Value + MySubTotal13
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal13 = MyQuantity.Value + MySubTotal13
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal13 = MyQuantity.Value + MySubTotal13

Else
End If

Case 14
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal14 = MyQuantity.Value + MySubTotal14
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal14 = MyQuantity.Value + MySubTotal14
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal14 = MyQuantity.Value + MySubTotal14

Else
End If


Case 15
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal15 = MyQuantity.Value + MySubTotal15
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal15 = MyQuantity.Value + MySubTotal15
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal15 = MyQuantity.Value + MySubTotal15

Else
End If


Case 16
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal16 = MyQuantity.Value + MySubTotal16
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal16 = MyQuantity.Value + MySubTotal16
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal16 = MyQuantity.Value + MySubTotal16

Else
End If


Case 17
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal17 = MyQuantity.Value + MySubTotal17
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal17 = MyQuantity.Value + MySubTotal17
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal17 = MyQuantity.Value + MySubTotal17
Else
End If


Case 18
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal18 = MyQuantity.Value + MySubTotal18
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal18 = MyQuantity.Value + MySubTotal18
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal18 = MyQuantity.Value + MySubTotal18
Else
End If


Case 19
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal19 = MyQuantity.Value + MySubTotal19
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal19 = MyQuantity.Value + MySubTotal19
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal19 = MyQuantity.Value + MySubTotal19

Else
End If



Case 20
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal20 = MyQuantity.Value + MySubTotal20
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal20 = MyQuantity.Value + MySubTotal20
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal20 = MyQuantity.Value + MySubTotal20

Else
End If


Case 21
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal21 = MyQuantity.Value + MySubTotal21
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal21 = MyQuantity.Value + MySubTotal21
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal21 = MyQuantity.Value + MySubTotal21
Else
End If

Case 22
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal22 = MyQuantity.Value + MySubTotal22
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal22 = MyQuantity.Value + MySubTotal22
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal22 = MyQuantity.Value + MySubTotal22

Else
End If

Case 23
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal23 = MyQuantity.Value + MySubTotal23
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal23 = MyQuantity.Value + MySubTotal23
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal23 = MyQuantity.Value + MySubTotal23

Else
End If

Case 24
If MyQuantity.Value >= 0 And MyMatrixPosition.Value > k Then
MySubTotal24 = MyQuantity.Value + MySubTotal24
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value < k Then
MySubTotal24 = MyQuantity.Value + MySubTotal24
ElseIf MyQuantity.Value <= 0 And MyMatrixPosition.Value = 1 Then
MySubTotal24 = MyQuantity.Value + MySubTotal24

Else
End If
End Select

Next i

MyWS.Cells(6, 10 + k).Value = MySubTotal1
MyWS.Cells(7, 10 + k).Value = MySubTotal2
MyWS.Cells(8, 10 + k).Value = MySubTotal3
MyWS.Cells(9, 10 + k).Value = MySubTotal4
MyWS.Cells(10, 10 + k).Value = MySubTotal5
MyWS.Cells(11, 10 + k).Value = MySubTotal6
MyWS.Cells(12, 10 + k).Value = MySubTotal7
MyWS.Cells(13, 10 + k).Value = MySubTotal8
MyWS.Cells(14, 10 + k).Value = MySubTotal9
MyWS.Cells(15, 10 + k).Value = MySubTotal10
MyWS.Cells(16, 10 + k).Value = MySubTotal11
MyWS.Cells(17, 10 + k).Value = MySubTotal12
MyWS.Cells(18, 10 + k).Value = MySubTotal13
MyWS.Cells(19, 10 + k).Value = MySubTotal14
MyWS.Cells(20, 10 + k).Value = MySubTotal15
MyWS.Cells(21, 10 + k).Value = MySubTotal16
MyWS.Cells(22, 10 + k).Value = MySubTotal17
MyWS.Cells(23, 10 + k).Value = MySubTotal18
MyWS.Cells(24, 10 + k).Value = MySubTotal19
MyWS.Cells(25, 10 + k).Value = MySubTotal20
MyWS.Cells(26, 10 + k).Value = MySubTotal21
MyWS.Cells(27, 10 + k).Value = MySubTotal22
MyWS.Cells(28, 10 + k).Value = MySubTotal23
MyWS.Cells(29, 10 + k).Value = MySubTotal24
MySubTotal1 = 0
MySubTotal2 = 0
MySubTotal3 = 0
MySubTotal4 = 0
MySubTotal5 = 0
MySubTotal6 = 0
MySubTotal7 = 0
MySubTotal8 = 0
MySubTotal9 = 0
MySubTotal10 = 0
MySubTotal11 = 0
MySubTotal12 = 0
MySubTotal13 = 0
MySubTotal14 = 0
MySubTotal15 = 0
MySubTotal16 = 0
MySubTotal17 = 0
MySubTotal18 = 0
MySubTotal19 = 0
MySubTotal20 = 0
MySubTotal21 = 0
MySubTotal22 = 0
MySubTotal23 = 0
MySubTotal24 = 0

Set StartPoint = Nothing
Set MyHour = Nothing
Set MyMatrixPosition = Nothing
Set MyQuantity = Nothing


'Next j
Next k

'***************************************************************************************************************

'''''If MyWS.Range("MyLBoundC").Value = 0 Then
'''''If MyWS.Range("MyUBoundC").Offset(1, 0).Value = 24 Then 'Only Buy

'''''ElseIf MyWS.Range("MyLBoundC").Offset(1, 0).Value = 24 Then 'Only Sell
'''''MyWS.Range(Cells(5, 11), Cells(5, 11)).Value = MyWS.Range("MyLBoundC").Offset(-1, 0).Value
'''''MyWS.Range(Cells(6, 11), Cells(29, 11)).Value = 0
'''''Else  'Both Buy and Sell

'''''End If
'''''End If

'''''If MyWS.Range("MyUBoundC").Value = 0 Then

'''''MyWS.Range(Cells(5, 11), Cells(5, 11)).Value = MyWS.Range("MyUBoundC").Offset(-1, 0).Value
'''''MyWS.Range(Cells(6, 11), Cells(29, 11)).Value = 0
'''''End If

MyWS.Range(Cells(5, 11), Cells(5, 11 + MyNoDistinctValues - 1)).Value = WorksheetFunction.Transpose(MyWS.Range(Cells(6, 9), Cells(6 + MyNoDistinctValues - 1, 9)).Value)

For i = 1 To 24
Select Case MyWS.Range(Cells(5, 12), Cells(5, 12)).Cells(i + 1, 1).Value
Case Is >= 0  'Buy --> 3000 a zero
'MyWS.Range(Cells(5, 11), Cells(5, 11)).Cells(i + 1, 1).Value = 0
MyWS.Range(Cells(5, 11), Cells(5, 11)).Cells(i + 1, 1).Value = _
MyWS.Range(Cells(5, 12), Cells(5, 12)).Cells(i + 1, 1).Value
Case Is < 0
MyWS.Range(Cells(5, 11), Cells(5, 11)).Cells(i + 1, 1).Value = 0
Case Else
End Select

Select Case MyWS.Range(Cells(5, 11 + MyNoDistinctValues - 2), Cells(5, 11 + MyNoDistinctValues - 2)).Cells(i + 1, 1).Value
Case Is < 0   'Sell --> -500 a zero
MyWS.Range(Cells(5, 11 + MyNoDistinctValues - 1), Cells(5, 11 + MyNoDistinctValues - 1)).Cells(i + 1, 1).Value = _
MyWS.Range(Cells(5, 11 + MyNoDistinctValues - 2), Cells(5, 11 + MyNoDistinctValues - 2)).Cells(i + 1, 1).Value
Case Is > 0
MyWS.Range(Cells(5, 11 + MyNoDistinctValues - 2), Cells(5, 11 + MyNoDistinctValues - 2)).Cells(i + 1, 1).Value = 0
'MyWS.Range(Cells(5, 11 + MyNoDistinctValues - 1), Cells(5, 11 + MyNoDistinctValues - 1)).Cells(i + 1, 1).Value = 0
Case Else
End Select
Next i

'MyWS.Cells(33, 22).Value = Format(Now, "HH:MM:ss")
Set MyWS = Nothing
'Application.Volatile (True)
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

'EPEX bids copy
ThisWorkbook.Worksheets("MyTemplate").Range("Interpolationbid").Copy

'***************************************************************************************************************

End Sub
