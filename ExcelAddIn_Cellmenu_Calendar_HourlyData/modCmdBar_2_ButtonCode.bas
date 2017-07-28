Attribute VB_Name = "modCmdBar_2_ButtonCode"

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit
Option Base 1
Option Compare Text

Public gcombarDate011 As Double
Public gcombarDate021 As Double
Public gcombarDate022 As Double
Public gcombarDate031 As Double
Public gcombarDate032 As Double
Public gcombarDate033 As Double
Public gcombarDate041 As Double
Public gcombarDate042 As Double

Private Sub cbSwitchOff()
With Application
.EnableEvents = False
.Calculation = xlCalculationManual
.ScreenUpdating = False
.DisplayStatusBar = False
End With
End Sub

Private Sub cbSwitchOn()
With Application
.EnableEvents = True
.Calculation = xlCalculationAutomatic
.ScreenUpdating = True
.DisplayStatusBar = True
End With
End Sub

Sub myCreateButtons(myCurrentMenuItem As CommandBarPopup, MyButtonName As String, MySubName As String)
Dim myitem As CommandBarButton
Dim i As Long
Dim MyZone As String * 7, MyType As String * 3

Select Case MyButtonName 'myCurrentMenuItem.Caption

Case Is = "Actual"
    MyZone = left(MySubName, 7)
    MyType = Mid(MySubName, 9, 3)
    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)

    With myitem
      .Caption = "Actual"
      .OnAction = "'" & ThisWorkbook.Name & "'!" & "'ACT_UAL " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
      .FaceId = 217
    End With

Case Is = "Delivered"
    MyZone = left(MySubName, 7)
    MyType = Mid(MySubName, 9, 3)
    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)

    With myitem
      .Caption = "Actual"
      .OnAction = "'" & ThisWorkbook.Name & "'!" & "'ACT_UAL " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
      .FaceId = 217
    End With


Case Is = "Forecast EC00 Ens Actual Deltas"
   MyZone = left(MySubName, 7)
    MyType = Mid(MySubName, 9, 3)

    
    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
    With myitem
        .Caption = "Forecast EC00 Ens Actual Deltas"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FAD_ES " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
    End With

Case Is = "Forecast EC00 Op Actual Deltas"
   MyZone = left(MySubName, 7)
   MyType = Mid(MySubName, 9, 3)
   Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
   With myitem
        .Caption = "Forecast EC00 Op Actual Deltas"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FAD_OP " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
   End With


Case Is = "Forecast EC00 Op"
   MyZone = left(MySubName, 7)
   MyType = Mid(MySubName, 9, 3)
    
    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
    With myitem
        .Caption = "Forecast EC00 Op"
       .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FOR_OP " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
    End With


Case Is = "Forecast EC00 Ens"
   MyZone = left(MySubName, 7)
   MyType = Mid(MySubName, 9, 3)

    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
    With myitem
        .Caption = "Forecast EC00 Ens"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FOR_ES " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
    End With


Case Is = "Forecast EC00 Op Deltas"
   MyZone = left(MySubName, 7)
   MyType = Mid(MySubName, 9, 3)


    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
    With myitem
        .Caption = "Forecast EC00 Op Deltas"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FDE_OP " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
    End With


Case Is = "Forecast EC00 Ens Deltas"
   MyZone = left(MySubName, 7)
   MyType = Mid(MySubName, 9, 3)

    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
    With myitem
        .Caption = "Forecast EC00 Ens Deltas"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FDE_ES " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
    End With


Case Is = "Forecast EC00 Op Shift"
   MyZone = left(MySubName, 7)
   MyType = Mid(MySubName, 9, 3)

    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
    With myitem
        .Caption = "Forecast EC00 Op Shift"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FSH_OP " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
    End With


Case Is = "Forecast EC00 Ens Shift"
   MyZone = left(MySubName, 7)
   MyType = Mid(MySubName, 9, 3)

    Set myitem = myCurrentMenuItem.Controls.Add(Type:=msoControlButton)
    With myitem
        .Caption = "Forecast EC00 Ens Shift"
        .OnAction = "'" & ThisWorkbook.Name & "'!" & "'FSH_ES " & """" & MyType & """" & ", " & """" & MyZone & """" & "'"
        .FaceId = 217
    End With


'''Case Is = "Prices" 'cbRetrieveCompData Procedure not show it on a .bas file. One can view it through the xlsm file.
    '''cbRetrieveCompData cbsMyType:="PRI", cbsMyData:="ACT", cbsMyZone:=MySubName, cbsMyDate1:=gcombarDate011

End Select
Set myitem = Nothing

'Set myCurrentMenuItem = Nothing:

End Sub

'On Action Subs:
Sub ACT_UAL(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar01
If gcombarDate011 = 0 Then Exit Sub
Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False


cbRetrieveData cbsMyData:="ACT", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate011, cbsMyDate2:=0    'cbsMyDate2=0 as there is no 2nd wb

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    Actual" & Chr(10) _
                & "Target  Date: " & CStr(Format(gcombarDate011, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new
    'Debug.Print sCommentTT
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&ACT&" & MyType & "&" & MyZone & "&" & gcombarDate011 & "&" & "0" & "&CompTAG>>" 'new
    'Debug.Print sCommentTT
    '"0" in the End is the DeltaDate
    
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
      
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
    .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
       .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 125
    End If
End With

cbSwitchOn

End Sub

Sub FAD_ES(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar02
If gcombarDate021 = 0 Then Exit Sub
If gcombarDate022 = 0 Then Exit Sub
Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False


cbRetrieveData cbsMyData:="FAD_ES", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate021, cbsMyDate2:=gcombarDate022

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    FAD_ES" & Chr(10) _
                & "Target Date:  " & CStr(Format(gcombarDate021, "DD/MM/YYYY")) & Chr(10) & "Delta Date:    " & CStr(Format(gcombarDate022, "DD/MM/YYYY")) & Chr(10) _
                 & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta Date:   __/__/____"  'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FAD_ES&" & MyType & "&" & MyZone & "&" & gcombarDate021 & "&" & gcombarDate022 & "&CompTAG>>"
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
    .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
      .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 135
    End If
End With

cbSwitchOn

'MsgBox "It works. FAD_ES. MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub

Sub FAD_OP(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar02
If gcombarDate021 = 0 Then Exit Sub
If gcombarDate022 = 0 Then Exit Sub

Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False


cbRetrieveData cbsMyData:="FAD_OP", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate021, cbsMyDate2:=gcombarDate022

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    FAD_OP" & Chr(10) _
                & "Target Date: " & CStr(Format(gcombarDate021, "DD/MM/YYYY")) & Chr(10) & "Delta Date:   " & CStr(Format(gcombarDate022, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta    Date: __/__/____"  'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FAD_OP&" & MyType & "&" & MyZone & "&" & gcombarDate021 & "&" & gcombarDate022 & "&CompTAG>>"
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
    .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
       .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 135
    End If
End With

cbSwitchOn


'MsgBox "It works. FAD_OP. MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub

Sub FOR_OP(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar01
If gcombarDate011 = 0 Then Exit Sub

Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False

cbRetrieveData cbsMyData:="FOR_OP", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate011, cbsMyDate2:=0

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    FOR_OP" & Chr(10) _
                & "Target Date:  " & CStr(Format(gcombarDate011, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FOR_OP&" & MyType & "&" & MyZone & "&" & gcombarDate011 & "&" & "0" & "&CompTAG>>"
    '"0" is the DeltaDate
    
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
        .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
      .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 125
    End If

    
End With

cbSwitchOn

'MsgBox "It works. FOR_OP. MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub

Sub FOR_ES(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar01
If gcombarDate011 = 0 Then Exit Sub
Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False

cbRetrieveData cbsMyData:="FOR_ES", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate011, cbsMyDate2:=0

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    FOR_ES" & Chr(10) _
                & "Target Date:  " & CStr(Format(gcombarDate011, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new

    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FOR_ES&" & MyType & "&" & MyZone & "&" & gcombarDate011 & "&" & "0" & "&CompTAG>>"
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
       .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
      .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 125
    End If

    
End With

cbSwitchOn

'MsgBox "It works. FOR_ES. MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub

Sub FDE_OP(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar02
If gcombarDate021 = 0 Then Exit Sub
If gcombarDate022 = 0 Then Exit Sub
Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False

cbRetrieveData cbsMyData:="FDE_OP", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate021, cbsMyDate2:=gcombarDate022

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    FDE_OP" & Chr(10) _
                & "Target Date:  " & CStr(Format(gcombarDate021, "DD/MM/YYYY")) & Chr(10) & "Delta  Date:   " & CStr(Format(gcombarDate022, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta    Date: __/__/____"  'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FDE_OP&" & MyType & "&" & MyZone & "&" & gcombarDate021 & "&" & gcombarDate022 & "&CompTAG>>"
    'gcombarDate022 is the DeltaDate
    
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
       .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
       .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 135
    End If
    
End With

cbSwitchOn

'MsgBox "It works. FDE_OP. MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub
Sub FDE_ES(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar02
If gcombarDate021 = 0 Then Exit Sub
If gcombarDate022 = 0 Then Exit Sub
Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False

cbRetrieveData cbsMyData:="FDE_ES", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate021, cbsMyDate2:=gcombarDate022

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:        " & MyType & Chr(10) & "Data Type:    FDE_ES" & Chr(10) _
                & "Target Date:  " & CStr(Format(gcombarDate021, "DD/MM/YYYY")) & Chr(10) & "Delta  Date:   " & CStr(Format(gcombarDate022, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta    Date: __/__/____"  'new
                
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FDE_ES&" & MyType & "&" & MyZone & "&" & gcombarDate021 & "&" & gcombarDate022 & "&CompTAG>>"
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
       .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
          .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 135

    End If
    
    
End With

cbSwitchOn

'MsgBox "It works. FDE_ES.MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub
Sub FSH_OP(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar01
If gcombarDate011 = 0 Then Exit Sub
Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False

cbRetrieveData cbsMyData:="FSH_OP", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate011, cbsMyDate2:=(gcombarDate011 - 1)

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    FSH_OP" & Chr(10) _
                & "Target Date:  " & CStr(Format(gcombarDate011, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new

    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FSH_OP&" & MyType & "&" & MyZone & "&" & gcombarDate011 & "&" & (gcombarDate011 - 1) & "&CompTAG>>"
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
       .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
          .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 125

    End If
    
End With

cbSwitchOn

'MsgBox "It works. FSH_OP.MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub

Sub FSH_ES(MyType As String, MyZone As String)
Dim sCommentTT As String
Dim bolTranspose As Boolean, bolNegative As Boolean
Dim cbMySelectedCells As String, cbMySelectedWS As String, cbMySelectedWB As String
Dim lArea As Long
OpenCalendar01
If gcombarDate011 = 0 Then Exit Sub


Call cbSwitchOff

cbMySelectedCells = Selection.Address
cbMySelectedWS = Selection.Worksheet.Name
cbMySelectedWB = Selection.Parent.Parent.Name
bolTranspose = False
bolNegative = False

cbRetrieveData cbsMyData:="FSH_ES", cbsMyType:=MyType, cbsMyZone:=MyZone, cbsMyDate1:=gcombarDate011, cbsMyDate2:=(gcombarDate011 - 1)

With Workbooks(cbMySelectedWB).Worksheets(cbMySelectedWS).Range(cbMySelectedCells)
    If .Columns.Count = 2 And .Rows.Count = 1 Then bolTranspose = True
    If .Columns.Count = 3 And .Rows.Count = 1 Then
        bolTranspose = True
        bolNegative = True
    End If
    If .Columns.Count = 1 And .Rows.Count = 3 Then
        bolTranspose = False
        bolNegative = True
    End If
    
    sCommentTT = "Tag Date:     " & CStr(Date) & Chr(10) & "Publisher:      MK" & Chr(10) & "Zone:           " & MyZone & Chr(10) & "Product:       " & MyType & Chr(10) & "Data Type:    FSH_ES" & Chr(10) _
                & "Target Date:  " & CStr(Format(gcombarDate011, "DD/MM/YYYY")) & Chr(10) _
                & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new

    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(Date)) & "&" & bolTranspose & "&" & bolNegative & "&MK" & "&FSH_ES&" & MyType & "&" & MyZone & "&" & gcombarDate011 & "&" & (gcombarDate011 - 1) & "&CompTAG>>"
    If .Cells(1, 1).Comment Is Nothing Then
       .Cells(1, 1).AddComment.Text sCommentTT
    Else
        .Cells(1, 1).Comment.Delete
        .Cells(1, 1).AddComment.Text sCommentTT
    End If
       .Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If .Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = .Cells(1, 1).Comment.Shape.Width * .Cells(1, 1).Comment.Shape.Height
       .Cells(1, 1).Comment.Shape.Width = 120
       .Cells(1, 1).Comment.Shape.Height = 125

    End If
   
    
    
End With

cbSwitchOn

'MsgBox "It works. FSH_ES .MyType is " & MyType & " MyZone is " & MyZone & "  " & MyDate1 & " / " & MyDate2

End Sub

Private Sub OpenCalendar01()
   With frmCalendar01
   .StartUpPosition = 0
   .left = convertMouseToForm.left
   .top = convertMouseToForm.top
   .Show
   End With
End Sub

Private Sub OpenCalendar02()
   With frmCalendar02
   .StartUpPosition = 0
   .left = convertMouseToForm.left
   .top = convertMouseToForm.top
   .Show
   End With
End Sub

Private Sub OpenCalendar03()
   With frmCalendar03
   .StartUpPosition = 0
   .left = convertMouseToForm.left
   .top = convertMouseToForm.top
   .Show
   End With
End Sub

Private Sub OpenCalendar04()
   With frmCalendar04
   .StartUpPosition = 0
   .left = convertMouseToForm.left
   .top = convertMouseToForm.top
   .Show
   End With
End Sub

