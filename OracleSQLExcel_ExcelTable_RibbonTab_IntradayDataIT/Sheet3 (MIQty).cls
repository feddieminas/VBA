VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Application.EnableEvents = False

Select Case Target.Address

Case Is = Range("O1").Address 'Q.tà acc.

GetDatafromExchITSpot
Application.Wait Now + TimeValue("00:00:02")
'Sleep 1000
MIunits

End Select

Application.EnableEvents = True

End Sub

