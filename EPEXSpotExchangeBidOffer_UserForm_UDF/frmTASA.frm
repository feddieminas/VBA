VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTASA 
   Caption         =   "Company Automatic Scheduling Assistant"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5220
   OleObjectBlob   =   "frmTASA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTASA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***********************************
'BuyRange Continental_20170727.xls
'***********************************

Private Sub btnCaC_Click()
Dim MyMessage As String
Dim BidType As Integer
Dim BidBook As Integer
Dim ChoiceA As String
Dim ChoiceB As String
Dim Response As VbMsgBoxResult

If ThisWorkbook.Worksheets("MyUserForm").Range("B3").Value = 1 Then
ChoiceA = "Purchases"
End If

If ThisWorkbook.Worksheets("MyUserForm").Range("B6").Value = 1 Then
ChoiceB = "Continental"
End If

MyMessage = "Pls confirm that the following selection is correct:" & vbNewLine & _
        "These trigger prices are for:" & vbNewLine & _
        ChoiceA & vbNewLine & _
        "Book: " & vbNewLine & ChoiceB
        
Unload Me

Response = MsgBox(MyMessage, vbYesNo, "TEI Automatic Scheduling Assistant")

If Response = vbYes Then Call MyPartTwo
If Response = vbNo Then Exit Sub

End Sub

Private Sub btnContinental_Click()
ThisWorkbook.Worksheets("MyUserForm").Range("B6").Value = 1
ThisWorkbook.Worksheets("MyUserForm").Range("B7").Value = 0
End Sub

Private Sub btnEC_Click()
Unload Me
End Sub

Private Sub btnPurchases_Click()
ThisWorkbook.Worksheets("MyUserForm").Range("B3").Value = 1
ThisWorkbook.Worksheets("MyUserForm").Range("B4").Value = 0
End Sub

Private Sub UserForm_Click()

End Sub
