VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar02 
   Caption         =   "Select Target Date &  Delta Date..."
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5910
   OleObjectBlob   =   "frmCalendar02.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCalendar02"
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
    gcombarDate021 = 0
    gcombarDate022 = 0
    Unload Me
End Sub

Private Sub MonthView021_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
        gcombarDate021 = CDbl(DateClicked)
End Sub

Private Sub MonthView022_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    If gcombarDate021 = 0 Then
    MsgBox "No Target Delivery Date has been selected. Please select one before selecting the Delta Date."
    gcombarDate022 = 0
    Else
    gcombarDate022 = CDbl(DateClicked)
    Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
        Me.MonthView021.Value = Date
        Me.MonthView022.Value = Date
        gcombarDate021 = 0
        gcombarDate022 = 0
End Sub
