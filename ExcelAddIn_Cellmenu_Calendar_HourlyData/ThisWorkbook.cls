VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Private Sub Workbook_BeforeClose(Cancel As Boolean) 'Actual, Forecast Ops, Forecast Ops Shift
On Error Resume Next

DeletePopUpMenu

On Error GoTo 0

End Sub

Private Sub Workbook_Open()
On Error Resume Next
SetUpMenus

On Error GoTo 0
'Call AddToCellMenu 'need to uncomment it for the excel add-inn

End Sub

Private Sub Workbook_Activate()
    Call AddToCellMenu
End Sub

Private Sub Workbook_Deactivate()
    Call DeleteFromCellMenu
    Call DeletePopUpMenu
End Sub
