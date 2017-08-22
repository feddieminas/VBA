Attribute VB_Name = "modRibbon"

Public ribbon As IRibbonUI
Public arrMenuDates(1 To 2) As String
Public arrMenuMarkets(1 To 7) As String

Public Sub RibbonLoad(rb As IRibbonUI)
'
' Code for onLoad callback. Ribbon control customUI
'
Set ribbon = rb
ThisWorkbook.Worksheets("Settings").Range("F1").Value = ObjPtr(ribbon)
FillarrMenuDates
FillarrMenuMarkets
End Sub


'*****************************
'Dates
'*****************************

Sub FillarrMenuDates()
Dim i As Long
For i = 0 To 1
arrMenuDates(i + 1) = Format(Date + i, "DD/MM/YYYY") 'If not correct date appears on your Ribbon's Combo box, then change your format to "MM/DD/YYYY"
Next i
End Sub

Sub GetItemCountDates(control As IRibbonControl, ByRef count)
'Combo box callback.
    count = UBound(arrMenuDates)
End Sub

Sub GetTextDates(control As IRibbonControl, ByRef text)
'Combo box callback.
    text = arrMenuDates(2)
ThisWorkbook.Worksheets("Settings").Range("F2").Value = CDate(text)
 'reset  array
ribbon.InvalidateControl ("ComboDates")
End Sub

Sub GetItemLabelDates(control As IRibbonControl, index As Integer, _
           ByRef label)
    label = arrMenuDates(index + 1)
End Sub

Public Sub OnChangeDates(control As IRibbonControl, text As String)
'
' Code for onChange callback. Ribbon control editBox
'
ThisWorkbook.Worksheets("Settings").Range("F2").Value = CDate(text)
'reset  array
'ribbon.InvalidateControl ("ComboDates")
End Sub


'*****************************
'Markets
'*****************************

Sub FillarrMenuMarkets()
arrMenuMarkets(1) = "MI1"
arrMenuMarkets(2) = "MI2"
arrMenuMarkets(3) = "MI3"
arrMenuMarkets(4) = "MI4"
arrMenuMarkets(5) = "MI5"
arrMenuMarkets(6) = "MI6"
arrMenuMarkets(7) = "MI7"
End Sub

Sub GetItemCountMarkets(control As IRibbonControl, ByRef count)
'Combo box callback.
    count = UBound(arrMenuMarkets)
End Sub

Sub GetTextMarkets(control As IRibbonControl, ByRef text)
'Combo box callback.
    text = arrMenuMarkets(1)
ThisWorkbook.Worksheets("Settings").Range("F4").Value = text
 'reset  array
ribbon.InvalidateControl ("ComboMarkets")
End Sub

Sub GetItemLabelMarkets(control As IRibbonControl, index As Integer, _
           ByRef label)
    label = arrMenuMarkets(index + 1)
End Sub

Public Sub OnChangeMarkets(control As IRibbonControl, text As String)
'
' Code for onChange callback. Ribbon control editBox
'
ThisWorkbook.Worksheets("Settings").Range("F4").Value = text 'default is MI1 market, the first Intraday session of the Italian Exchange
'ThisWorkbook.Worksheets("MIQty").Select
 'reset  array
'ribbon.InvalidateControl ("ComboDates")
End Sub



