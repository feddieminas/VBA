VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Worked the project in Windows 7 and Excel 2007, 2010
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

'Libraries Used :
'Visual Basic for Applications
'Microsoft Excel 15.0 Object Library
'OLE Automation
'Microsoft Office 15.0 Object Library
'Microsoft ActiveX Data Objects 2.8 Library
'Microsoft Forms 2.0 Object Library

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
