VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'***********************************
'EPEX_PowerBidOffer_Template.xlsm
'***********************************

'Semi-Manual

Sub FixedPosPlace()

Worksheets("MyTemplate").Range("B6:B29").Value = Range("J4:J27").Value

End Sub

Sub TotalPosPlace()

Worksheets("MyTemplate").Range("B6:B29").Value = Range("N4:N27").Value

End Sub

'Auto

'One can uncomment the following macro to trigger the events.
'This would eventually mean that whilst you are modifying your exchange bid/offer,
'those values will automatically go to your bid spreadsheet (MyTemplate Sheet)
'and then ready to proceed to place them on the exchange.

'Private Sub WorkSheet_Change(ByVal Target As Range)

'Dim rng As Range
'Set rng = Range("J4:J27,L4:M27,B4:D27")

'If Intersect(Target, rng) Is Nothing Then GoTo myEnd

'Application.EnableEvents = False
'Application.ScreenUpdating = False

'If Range("H1") = Date + 1 Then

'Select Case Range("S1").Value

'Case Is <> 0
   
'Worksheets("MyTemplate").Range("B6:B29").Value = Range("N4:N27").Value
  
'Case Else

'Worksheets("MyTemplate").Range("B6:B29").Value = Range("J4:J27").Value

'End Select

'End If

'Application.EnableEvents = True
'Application.ScreenUpdating = True

'myEnd:

'Set rng = Nothing

'End Sub

