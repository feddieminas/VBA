VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar01 
   Caption         =   "Select Target Date"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2865
   OleObjectBlob   =   "frmCalendar01.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCalendar01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit

Private Sub cmdClose_Click()
    gcombarDate011 = 0
    Unload Me
End Sub

Private Sub MonthView011_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
        gcombarDate011 = CDbl(DateClicked)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
        Me.MonthView011.Value = Date
        gcombarDate011 = 0
End Sub
