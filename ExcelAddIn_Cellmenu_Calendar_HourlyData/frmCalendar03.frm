VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar03 
   Caption         =   "Select Target Date, Delta Date & Publication Date..."
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8925
   OleObjectBlob   =   "frmCalendar03.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCalendar03"
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

    gcombarDate031 = 0
    gcombarDate032 = 0
    gcombarDate033 = 0
    Unload Me

End Sub

Private Sub MonthView031_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
        gcombarDate031 = CDbl(DateClicked)
End Sub

Private Sub MonthView032_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    If gcombarDate031 = 0 Then
    MsgBox "No Target Delivery Date has been selected. Please select one before selecting the Delta Date."
    gcombarDate032 = 0
    Else
    gcombarDate032 = CDbl(DateClicked)
    End If
End Sub
Private Sub MonthView033_DateClick(ByVal DateClicked As Date)
    On Error Resume Next
    If gcombarDate032 = 0 Then
    MsgBox "No Delta Date has been selected. Please select one before selecting the Publication Date."
    gcombarDate033 = 0
    Else
    gcombarDate033 = CDbl(DateClicked)
    Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
        Me.MonthView031.Value = Date
        Me.MonthView032.Value = Date
        Me.MonthView033.Value = Date
        
        gcombarDate031 = 0
        gcombarDate032 = 0
        gcombarDate033 = 0
End Sub
