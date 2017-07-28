Attribute VB_Name = "modTrigInputRange"

'***********************************
'BuyRange Continental_20170727.xls
'***********************************

Public MyOutputWS As Worksheet
Public MySourceWS As Worksheet
Public ModifiedRange As Range

Sub SelectRange()

Dim OriginalRange As Range
Dim NoofColumns As Integer
Dim NoofRows As Integer
Dim ModifiedColumns As Integer
Dim ModifiedRows As Integer
Dim MyHour As Integer
Dim MyPrice As Double
Dim MyPrompt As String
Dim sumtotalrows As Long
Dim sumtotalrange As Long
Dim Sumpopup As VbMsgBoxResult

Set MySourceWS = Worksheets("TraderTables")
Set MyOutputWS = Worksheets("Output")
MyOutputWS.Columns("A:D").ClearContents

MyPrompt = "Please select the full range of the trigger price & quantity table." & vbNewLine & _
        "Make sure that you include the farleft column with the hours, and the top row with the trigger prices!"
    
On Error Resume Next

Application.DisplayAlerts = False

Set OriginalRange = Application.InputBox(Prompt:=MyPrompt, _
    Title:="SPECIFY RANGE OF TRIGGER PRICE TABLE", Left:=300, Top:=400, Type:=8)

On Error GoTo 0

Application.DisplayAlerts = True
        

If OriginalRange Is Nothing Then

Exit Sub

Else

NoofColumns = OriginalRange.Columns.Count
NoofRows = OriginalRange.Rows.Count
End If
           
ModifiedColumns = NoofColumns - 1
ModifiedRows = NoofRows - 1
           
Set ModifiedRange = OriginalRange.Offset(1, 1).Resize(ModifiedRows, ModifiedColumns)

With MySourceWS
sumtotalrange = Application.WorksheetFunction.Sum(.Range(.Cells(1, 1), .Cells(NoofRows, NoofColumns)))
sumtotalrows = Application.WorksheetFunction.Sum(.Range("A1:A25").EntireRow)
End With

If sumtotalrange <> sumtotalrows Then
Sumpopup = MsgBox("Select your Input Range again", vbExclamation)
Exit Sub
End If

With frmTASA
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With
        

End Sub

Sub MyPartTwo()
Dim MyCounter As Integer
Dim MyChosenBook As String
Dim MyPurchaseSale As Double
Dim NegativeValues As VbMsgBoxResult

If ThisWorkbook.Worksheets("MyUserForm").Range("B6").Value = 1 Then MyChosenBook = "Continental"

If ThisWorkbook.Worksheets("MyUserForm").Range("B3").Value = 1 Then MyPurchaseSale = 1

With MyOutputWS
     .Columns("A:D").ClearContents
     .Range("A1").Value = "Hour"
     .Range("B1").Value = "Quantity"
     .Range("C1").Value = "Price"
     .Range("D1").Value = "Book"
End With
              
        
MyCounter = 2
        
For Each c In ModifiedRange.Cells
If IsEmpty(c) Then GoTo MyBlankCell:
If c.Value < 0 Then
    NegativeValues = MsgBox("Insert Positive Value on Range " & c.Address(RowAbsolute:=False, ColumnAbsolute:=False), vbCritical)
    ThisWorkbook.Worksheets("Output").Columns("A:D").ClearContents
Exit Sub
End If
    MyOutputWS.Cells(MyCounter, 1).Value = MySourceWS.Cells(c.Row, 1).Value
    MyOutputWS.Cells(MyCounter, 2).Value = c.Value * MyPurchaseSale
    MyOutputWS.Cells(MyCounter, 3).Value = MySourceWS.Cells(1, c.Column).Value
    MyOutputWS.Cells(MyCounter, 4).Value = MyChosenBook
    MyCounter = MyCounter + 1
MyBlankCell:
Next c
MyOutputWS.Select

'MsgBox "Yeeehaaaa"

End Sub
