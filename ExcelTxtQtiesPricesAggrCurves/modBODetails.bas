Attribute VB_Name = "modBODetails"

Sub SummariseBO()
Dim arrBOTable() As Variant
Dim Lastrow As Long
Dim WB As Workbook
Dim WS As Worksheet
Dim j As Long
Dim Q As Long
Dim t As Long
Dim Z As Long
Dim FilterHr As Long
Dim FilterZo As Long
Dim FilterBO As String

ESHr1 = 0
PTHr1 = 0
ESHr2 = 0
PTHr2 = 0
ESHr3 = 0
PTHr3 = 0
ESHr4 = 0
PTHr4 = 0
ESHr5 = 0
PTHr5 = 0
ESHr6 = 0
PTHr6 = 0
ESHr7 = 0
PTHr7 = 0
ESHr8 = 0
PTHr8 = 0
ESHr9 = 0
PTHr9 = 0
ESHr10 = 0
PTHr10 = 0
ESHr11 = 0
PTHr11 = 0
ESHr12 = 0
PTHr12 = 0
ESHr13 = 0
PTHr13 = 0
ESHr14 = 0
PTHr14 = 0
ESHr15 = 0
PTHr15 = 0
ESHr16 = 0
PTHr16 = 0
ESHr17 = 0
PTHr17 = 0
ESHr18 = 0
PTHr18 = 0
ESHr19 = 0
PTHr19 = 0
ESHr20 = 0
PTHr20 = 0
ESHr21 = 0
PTHr21 = 0
ESHr22 = 0
PTHr22 = 0
ESHr23 = 0
PTHr23 = 0
ESHr24 = 0
PTHr24 = 0

ESHr1OFF = 0
PTHr1OFF = 0
ESHr2OFF = 0
PTHr2OFF = 0
ESHr3OFF = 0
PTHr3OFF = 0
ESHr4OFF = 0
PTHr4OFF = 0
ESHr5OFF = 0
PTHr5OFF = 0
ESHr6OFF = 0
PTHr6OFF = 0
ESHr7OFF = 0
PTHr7OFF = 0
ESHr8OFF = 0
PTHr8OFF = 0
ESHr9OFF = 0
PTHr9OFF = 0
ESHr10OFF = 0
PTHr10OFF = 0
ESHr11OFF = 0
PTHr11OFF = 0
ESHr12OFF = 0
PTHr12OFF = 0
PTHr13OFF = 0
ESHr13OFF = 0
PTHr14OFF = 0
ESHr14OFF = 0
PTHr15OFF = 0
ESHr15OFF = 0
PTHr16OFF = 0
ESHr16OFF = 0
PTHr17OFF = 0
ESHr17OFF = 0
PTHr18OFF = 0
ESHr18OFF = 0
PTHr19OFF = 0
ESHr19OFF = 0
PTHr20OFF = 0
ESHr22OFF = 0
PTHr21OFF = 0
ESHr22OFF = 0
PTHr22OFF = 0
ESHr23OFFOFF = 0
PTHr23OFFOFF = 0
ESHr24OFFOFF = 0
PTHr24OFFOFF = 0

ESHr1BID = 0
PTHr1BID = 0
ESHr2BID = 0
PTHr2BID = 0
ESHr3BID = 0
PTHr3BID = 0
ESHr4BID = 0
PTHr4BID = 0
ESHr5BID = 0
PTHr5BID = 0
ESHr6BID = 0
PTHr6BID = 0
ESHr7BID = 0
PTHr7BID = 0
ESHr8BID = 0
PTHr8BID = 0
ESHr9BID = 0
PTHr9BID = 0
ESHr10BID = 0
PTHr10BID = 0
ESHr11BID = 0
PTHr11BID = 0
ESHr12BID = 0
PTHr12BID = 0
PTHr13BID = 0
ESHr13BID = 0
PTHr14BID = 0
ESHr14BID = 0
PTHr15BID = 0
ESHr15BID = 0
PTHr16BID = 0
ESHr16BID = 0
PTHr17BID = 0
ESHr17BID = 0
PTHr18BID = 0
ESHr18BID = 0
PTHr19BID = 0
ESHr19BID = 0
PTHr20BID = 0
ESHr22BID = 0
PTHr21BID = 0
ESHr22BID = 0
PTHr22BID = 0
ESHr23BIDBID = 0
PTHr23BIDBID = 0
ESHr24BIDBID = 0
PTHr24BIDBID = 0



Set WB = ThisWorkbook
Set WS = WB.Worksheets("ImportedData")

With WS
Lastrow = .Cells(.Rows.Count, "A").End(xlUp).Row
arrBOTable = .Range("A4:H" & Lastrow + 3)
End With

For j = 1 To UBound(arrBOTable)

FilterHr = arrBOTable(j, 1)
FilterZo = arrBOTable(j, 3)
FilterBO = arrBOTable(j, 5)

Select Case FilterHr

Case 1
If FilterZo = 1 Then ESHr1 = ESHr1 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr1OFF = ESHr1OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr1BID = ESHr1BID + 1
If FilterZo = 2 Then PTHr1 = PTHr1 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr1OFF = PTHr1OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr1BID = PTHr1BID + 1

Case 2
If FilterZo = 1 Then ESHr2 = ESHr2 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr2OFF = ESHr2OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr2BID = ESHr2BID + 1
If FilterZo = 2 Then PTHr2 = PTHr2 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr2OFF = PTHr2OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr2BID = PTHr2BID + 1

Case 3
If FilterZo = 1 Then ESHr3 = ESHr3 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr3OFF = ESHr3OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr3BID = ESHr3BID + 1
If FilterZo = 2 Then PTHr3 = PTHr3 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr3OFF = PTHr3OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr3BID = PTHr3BID + 1

Case 4
If FilterZo = 1 Then ESHr4 = ESHr4 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr4OFF = ESHr4OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr4BID = ESHr4BID + 1
If FilterZo = 2 Then PTHr4 = PTHr4 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr4OFF = PTHr4OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr4BID = PTHr4BID + 1

Case 5
If FilterZo = 1 Then ESHr5 = ESHr5 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr5OFF = ESHr5OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr5BID = ESHr5BID + 1
If FilterZo = 2 Then PTHr5 = PTHr5 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr5OFF = PTHr5OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr5BID = PTHr5BID + 1

Case 6
If FilterZo = 1 Then ESHr6 = ESHr6 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr6OFF = ESHr6OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr6BID = ESHr6BID + 1
If FilterZo = 2 Then PTHr6 = PTHr6 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr6OFF = PTHr6OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr6BID = PTHr6BID + 1

Case 7
If FilterZo = 1 Then ESHr7 = ESHr7 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr7OFF = ESHr7OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr7BID = ESHr7BID + 1
If FilterZo = 2 Then PTHr7 = PTHr7 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr7OFF = PTHr7OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr7BID = PTHr7BID + 1

Case 8
If FilterZo = 1 Then ESHr8 = ESHr8 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr8OFF = ESHr8OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr8BID = ESHr8BID + 1
If FilterZo = 2 Then PTHr8 = PTHr8 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr8OFF = PTHr8OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr8BID = PTHr8BID + 1

Case 9
If FilterZo = 1 Then ESHr9 = ESHr9 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr9OFF = ESHr9OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr9BID = ESHr9BID + 1
If FilterZo = 2 Then PTHr9 = PTHr9 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr9OFF = PTHr9OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr9BID = PTHr9BID + 1

Case 10
If FilterZo = 1 Then ESHr10 = ESHr10 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr10OFF = ESHr10OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr10BID = ESHr10BID + 1
If FilterZo = 2 Then PTHr10 = PTHr10 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr10OFF = PTHr10OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr10BID = PTHr10BID + 1

Case 11
If FilterZo = 1 Then ESHr11 = ESHr11 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr11OFF = ESHr11OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr11BID = ESHr11BID + 1
If FilterZo = 2 Then PTHr11 = PTHr11 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr11OFF = PTHr11OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr11BID = PTHr11BID + 1

Case 12
If FilterZo = 1 Then ESHr12 = ESHr12 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr12OFF = ESHr12OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr12BID = ESHr12BID + 1
If FilterZo = 2 Then PTHr12 = PTHr12 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr12OFF = PTHr12OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr12BID = PTHr12BID + 1

Case 13
If FilterZo = 1 Then ESHr13 = ESHr13 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr13OFF = ESHr13OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr13BID = ESHr13BID + 1
If FilterZo = 2 Then PTHr13 = PTHr13 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr13OFF = PTHr13OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr13BID = PTHr13BID + 1

Case 14
If FilterZo = 1 Then ESHr14 = ESHr14 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr14OFF = ESHr14OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr14BID = ESHr14BID + 1
If FilterZo = 2 Then PTHr14 = PTHr14 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr14OFF = PTHr14OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr14BID = PTHr14BID + 1

Case 15
If FilterZo = 1 Then ESHr15 = ESHr15 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr15OFF = ESHr15OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr15BID = ESHr15BID + 1
If FilterZo = 2 Then PTHr15 = PTHr15 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr15OFF = PTHr15OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr15BID = PTHr15BID + 1

Case 16
If FilterZo = 1 Then ESHr16 = ESHr16 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr16OFF = ESHr16OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr16BID = ESHr16BID + 1
If FilterZo = 2 Then PTHr16 = PTHr16 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr16OFF = PTHr16OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr16BID = PTHr16BID + 1

Case 17
If FilterZo = 1 Then ESHr17 = ESHr17 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr17OFF = ESHr17OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr17BID = ESHr17BID + 1
If FilterZo = 2 Then PTHr17 = PTHr17 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr17OFF = PTHr17OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr17BID = PTHr17BID + 1

Case 18
If FilterZo = 1 Then ESHr18 = ESHr18 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr18OFF = ESHr18OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr18BID = ESHr18BID + 1
If FilterZo = 2 Then PTHr18 = PTHr18 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr18OFF = PTHr18OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr18BID = PTHr18BID + 1

Case 19
If FilterZo = 1 Then ESHr19 = ESHr19 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr19OFF = ESHr19OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr19BID = ESHr19BID + 1
If FilterZo = 2 Then PTHr19 = PTHr19 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr19OFF = PTHr19OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr19BID = PTHr19BID + 1

Case 20
If FilterZo = 1 Then ESHr20 = ESHr20 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr20OFF = ESHr20OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr20BID = ESHr20BID + 1
If FilterZo = 2 Then PTHr20 = PTHr20 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr20OFF = PTHr20OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr20BID = PTHr20BID + 1

Case 21
If FilterZo = 1 Then ESHr21 = ESHr21 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr21OFF = ESHr21OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr21BID = ESHr21BID + 1
If FilterZo = 2 Then PTHr21 = PTHr21 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr21OFF = PTHr21OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr21BID = PTHr21BID + 1

Case 22
If FilterZo = 1 Then ESHr22 = ESHr22 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr22OFF = ESHr22OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr22BID = ESHr22BID + 1
If FilterZo = 2 Then PTHr22 = PTHr22 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr22OFF = PTHr22OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr22BID = PTHr22BID + 1

Case 23
If FilterZo = 1 Then ESHr23 = ESHr23 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr23OFF = ESHr23OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr23BID = ESHr23BID + 1
If FilterZo = 2 Then PTHr23 = PTHr23 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr23OFF = PTHr23OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr23BID = PTHr23BID + 1

Case 24
If FilterZo = 1 Then ESHr24 = ESHr24 + 1
If FilterZo = 1 And FilterBO = "V" Then ESHr24OFF = ESHr24OFF + 1
If FilterZo = 1 And FilterBO = "C" Then ESHr24BID = ESHr24BID + 1
If FilterZo = 2 Then PTHr24 = PTHr24 + 1
If FilterZo = 2 And FilterBO = "V" Then PTHr24OFF = PTHr24OFF + 1
If FilterZo = 2 And FilterBO = "C" Then PTHr24BID = PTHr24BID + 1
End Select

Next j

Erase arrBOTable

Dim ESChoice As String, PTChoice As String


ESChoice = 0
PTChoice = 0



Dim arrResTotal() As Long

ReDim arrResTotal(1 To 24, 1 To 2)
For i = 1 To 24
For j = 1 To 2
arrResTotal(i, j) = 0
Next j
Next i
For i = 1 To 24





Select Case i
Case 1
ESChoice = ESHr1
PTChoice = PTHr1

Case 2
ESChoice = ESHr2
PTChoice = PTHr2

Case 3
ESChoice = ESHr3
PTChoice = PTHr3

Case 4
ESChoice = ESHr4
PTChoice = PTHr4

Case 5
ESChoice = ESHr5
PTChoice = PTHr5

Case 6
ESChoice = ESHr6
PTChoice = PTHr6

Case 7
ESChoice = ESHr7
PTChoice = PTHr7

Case 8
ESChoice = ESHr8
PTChoice = PTHr8

Case 9
ESChoice = ESHr9
PTChoice = PTHr9

Case 10
ESChoice = ESHr10
PTChoice = PTHr10

Case 11
ESChoice = ESHr11
PTChoice = PTHr11

Case 12
ESChoice = ESHr12
PTChoice = PTHr12

Case 13
ESChoice = ESHr13
PTChoice = PTHr13

Case 14
ESChoice = ESHr14
PTChoice = PTHr14

Case 15
ESChoice = ESHr15
PTChoice = PTHr15

Case 16
ESChoice = ESHr16
PTChoice = PTHr16

Case 17
ESChoice = ESHr17
PTChoice = PTHr17

Case 18
ESChoice = ESHr18
PTChoice = PTHr18

Case 19
ESChoice = ESHr19
PTChoice = PTHr19

Case 20
ESChoice = ESHr20
PTChoice = PTHr20

Case 21
ESChoice = ESHr21
PTChoice = PTHr21

Case 22
ESChoice = ESHr22
PTChoice = PTHr22

Case 23
ESChoice = ESHr23
PTChoice = PTHr23

Case 24
ESChoice = ESHr24
PTChoice = PTHr24

End Select

If IsNumeric(ESChoice) = False Then ESChoice = 0
If IsNumeric(PTChoice) = False Then PTChoice = 0


arrResTotal(i, 1) = ESChoice
arrResTotal(i, 2) = PTChoice
Next i


WS.Range("O32").Offset(0, 0).Resize(24, 2).Value = arrResTotal


ReDim arrResTotal(1 To 24, 1 To 2)

For i = 1 To 24
For j = 1 To 2
arrResTotal(i, j) = 0
Next j
Next i


For i = 1 To 24
ESChoice = 0
PTChoice = 0

Select Case i
Case 1
ESChoice = ESHr1OFF
PTChoice = PTHr1OFF

Case 2
ESChoice = ESHr2OFF
PTChoice = PTHr2OFF

Case 3
ESChoice = ESHr3OFF
PTChoice = PTHr3OFF

Case 4
ESChoice = ESHr4OFF
PTChoice = PTHr4OFF

Case 5
ESChoice = ESHr5OFF
PTChoice = PTHr5OFF

Case 6
ESChoice = ESHr6OFF
PTChoice = PTHr6OFF

Case 7
ESChoice = ESHr7OFF
PTChoice = PTHr7OFF

Case 8
ESChoice = ESHr8OFF
PTChoice = PTHr8OFF

Case 9
ESChoice = ESHr9OFF
PTChoice = PTHr9OFF

Case 10
ESChoice = ESHr10OFF
PTChoice = PTHr10OFF

Case 11
ESChoice = ESHr11OFF
PTChoice = PTHr11OFF

Case 12
ESChoice = ESHr12OFF
PTChoice = PTHr12OFF

Case 13
ESChoice = ESHr13OFF
PTChoice = PTHr13OFF

Case 14
ESChoice = ESHr14OFF
PTChoice = PTHr14OFF

Case 15
ESChoice = ESHr15OFF
PTChoice = PTHr15OFF

Case 16
ESChoice = ESHr16OFF
PTChoice = PTHr16OFF

Case 17
ESChoice = ESHr17OFF
PTChoice = PTHr17OFF

Case 18
ESChoice = ESHr18OFF
PTChoice = PTHr18OFF

Case 19
ESChoice = ESHr19OFF
PTChoice = PTHr19OFF

Case 20
ESChoice = ESHr20OFF
PTChoice = PTHr20OFF

Case 21
ESChoice = ESHr21OFF
PTChoice = PTHr21OFF

Case 22
ESChoice = ESHr22OFF
PTChoice = PTHr22OFF

Case 23
ESChoice = ESHr23OFF
PTChoice = PTHr23OFF

Case 24
ESChoice = ESHr24OFF
PTChoice = PTHr24OFF

End Select

If IsNumeric(ESChoice) = False Then ESChoice = 0
If IsNumeric(PTChoice) = False Then PTChoice = 0

arrResTotal(i, 1) = ESChoice
arrResTotal(i, 2) = PTChoice
Next i


WS.Range("O60").Offset(0, 0).Resize(24, 2).Value = arrResTotal


ReDim arrResTotal(1 To 24, 1 To 2)

For i = 1 To 24
For j = 1 To 2
arrResTotal(i, j) = 0
Next j
Next i

For i = 1 To 24
ESChoice = 0
PTChoice = 0

Select Case i
Case 1
ESChoice = ESHr1BID
PTChoice = PTHr1BID

Case 2
ESChoice = ESHr2BID
PTChoice = PTHr2BID

Case 3
ESChoice = ESHr3BID
PTChoice = PTHr3BID

Case 4
ESChoice = ESHr4BID
PTChoice = PTHr4BID

Case 5
ESChoice = ESHr5BID
PTChoice = PTHr5BID

Case 6
ESChoice = ESHr6BID
PTChoice = PTHr6BID

Case 7
ESChoice = ESHr7BID
PTChoice = PTHr7BID

Case 8
ESChoice = ESHr8BID
PTChoice = PTHr8BID

Case 9
ESChoice = ESHr9BID
PTChoice = PTHr9BID

Case 10
ESChoice = ESHr10BID
PTChoice = PTHr10BID

Case 11
ESChoice = ESHr11BID
PTChoice = PTHr11BID

Case 12
ESChoice = ESHr12BID
PTChoice = PTHr12BID

Case 13
ESChoice = ESHr13BID
PTChoice = PTHr13BID

Case 14
ESChoice = ESHr14BID
PTChoice = PTHr14BID

Case 15
ESChoice = ESHr15BID
PTChoice = PTHr15BID

Case 16
ESChoice = ESHr16BID
PTChoice = PTHr16BID

Case 17
ESChoice = ESHr17BID
PTChoice = PTHr17BID

Case 18
ESChoice = ESHr18BID
PTChoice = PTHr18BID

Case 19
ESChoice = ESHr19BID
PTChoice = PTHr19BID

Case 20
ESChoice = ESHr20BID
PTChoice = PTHr20BID

Case 21
ESChoice = ESHr21BID
PTChoice = PTHr21BID

Case 22
ESChoice = ESHr22BID
PTChoice = PTHr22BID

Case 23
ESChoice = ESHr23BID
PTChoice = PTHr23BID

Case 24
ESChoice = ESHr24BID
PTChoice = PTHr24BID

End Select

If IsNumeric(ESChoice) = False Then ESChoice = 0
If IsNumeric(PTChoice) = False Then PTChoice = 0



arrResTotal(i, 1) = ESChoice
arrResTotal(i, 2) = PTChoice
Next i


WS.Range("O90").Offset(0, 0).Resize(24, 2).Value = arrResTotal

Erase arrResTotal

End Sub

