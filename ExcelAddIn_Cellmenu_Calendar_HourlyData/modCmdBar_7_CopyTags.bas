Attribute VB_Name = "modCmdBar_7_CopyTags"

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit
Option Base 1

Private Type cbPasteUserSel
        Range As Range
        Anchor As Range
        Count As Long
        ColCount As Long
        RowCount As Long
        StartCol As Long
        StartRow As Long
        Vertical As Boolean
        Row As Long
        Col As Long
End Type

Private Type cbCopyUserSel
    Range As Range
    Anchor As Range
    Count As Long
    ColCount As Long
    RowCount As Long
    StartCol As Long
    StartRow As Long
    Vertical As Boolean
    Row As Long
    Col As Long
End Type

Private Type tCompTAG
    ActualTag As Boolean
    PublicationDate As Double
    Horizontal As Boolean
    ChngNegative As Boolean
    Publisher As String
    MyData As String
    MyType As String
    MyZone As String
    TargetDate As Double
    DeltaDate As Double
End Type

Sub cbCopyPasteTags()
Dim cbMyPrompt As String, arrMyComments() As String
Dim cbMyIndex As Long, cbCommentCount As Long, CPDiff As Long, cbAnchDiff As Long, cbTAGAdjust As Long
Dim cbMyComment As Comment, sCommentTT As String
Dim cbCopyArea As cbCopyUserSel, cbPasteArea As cbPasteUserSel
Dim C As Range
Dim arrCurrentTag() As String, MyCurrentTag As tCompTAG
Dim MyTagStart As Long, MyTagEnd As Long, i As Long, lArea As Long
Dim OneCellCheck As Boolean
On Error Resume Next

cbMyPrompt = "Please select the full range of Comp Tags" & vbNewLine & "that you wish to copy:"

Application.DisplayAlerts = False
MyCopyRetry:


With cbCopyArea
    Set .Range = Application.InputBox(Prompt:=cbMyPrompt, _
        Title:="Copy Tags", left:=300, top:=400, Type:=8)

    On Error GoTo 0

    Application.DisplayAlerts = True
        
        If .Range Is Nothing Then
            Exit Sub
        Else
            .ColCount = .Range.Columns.Count
            .RowCount = .Range.Rows.Count
            Set .Anchor = .Range.Cells(1, 1)
            .Row = .Anchor.Row
            .Col = .Anchor.Column
            .Count = .Range.Cells.Count
        End If
          
            If .ColCount > 1 And .RowCount > 1 Then
                MsgBox "The selected range must be either one column wide" & vbNewLine & "or one row wide. Please reselect copy range"
                GoTo MyCopyRetry:
            End If
          
            If .ColCount > 1 Then
                .Vertical = True
            Else
                If .Count > 1 Then
                .Vertical = False   ' this is the problem when only one cell
                Else
                OneCellCheck = True ' we need to double check exactly what to do
                End If
            End If

          
            cbCommentCount = 0
            
            For Each C In .Range.Cells                      'Count the comments in the selected range
                If Not (C.Comment Is Nothing) Then          'if there are none then exit
                    cbCommentCount = cbCommentCount + 1
                End If
            Next C
          
            If cbCommentCount = 0 Then
                MsgBox "No cells containing Comp TAGs were selected. Please retry!"
                GoTo MyCopyRetry:
            Else
                ReDim arrMyComments(1 To cbCommentCount)    'since comments exist create an array to store them in.
                cbCommentCount = 0
                   
               For Each C In .Range.Cells                  'copy comments into the array
                    If Not (C.Comment Is Nothing) Then
                        cbCommentCount = cbCommentCount + 1
                        MyTagStart = InStr(1, C.Comment.Text, "<<CompTAG")
                        MyTagEnd = InStrRev(C.Comment.Text, "CompTAG>>")
                        Debug.Print C.Comment.Text
                        'arrMyComments(cbCommentCount) = CStr(c.Comment.Text)
                        arrMyComments(cbCommentCount) = Mid(C.Comment.Text, MyTagStart, MyTagEnd - MyTagStart + 1 + 8)
                        'Debug.Print arrMyComments(cbCommentCount)
                    End If
                Next C
            End If
          
            If OneCellCheck = True Then
            arrCurrentTag() = Split(arrMyComments(1), "&")
            .Vertical = Not (CBool(arrCurrentTag(2)))
            Erase arrCurrentTag()
            End If

End With
          
          
With cbPasteArea

            cbMyPrompt = "Please select the full range" & vbNewLine & "you wish to paste Comp Tags to:"
MyPasteRetry:

            Set .Range = Application.InputBox(Prompt:=cbMyPrompt, _
            Title:="Paste Tags", left:=300, top:=400, Type:=8)
           
            If .Range Is Nothing Then
                Exit Sub
            Else
                .ColCount = .Range.Columns.Count
                .RowCount = .Range.Rows.Count
                Set .Anchor = .Range.Cells(1, 1)
                .Row = .Anchor.Row
                .Col = .Anchor.Column
                .Count = .Range.Cells.Count
            End If

            If .ColCount > 1 And .RowCount > 1 Then
                MsgBox "The selected range must be either one column wide" _
                & vbNewLine & "or one row wide. Please reselect paste range"
                GoTo MyPasteRetry:
            End If

            If .ColCount > 1 Then
                .Vertical = True
            Else
                If .Count > 1 Then
                .Vertical = False
                Else
                .Vertical = cbCopyArea.Vertical
                End If
                
            End If

            If .Vertical = False And cbCopyArea.Vertical = False Then         '
                cbAnchDiff = cbPasteArea.Row - cbCopyArea.Row
                cbTAGAdjust = cbAnchDiff - cbCommentCount + 1
                Debug.Print "Combo 1"
            ElseIf .Vertical = True And cbCopyArea.Vertical = True Then
                cbAnchDiff = cbPasteArea.Col - cbCopyArea.Col
                cbTAGAdjust = cbAnchDiff - cbCommentCount + 1 'if we are going backwards does that mean it should be -1?????
                Debug.Print "Combo 2"
            ElseIf .Vertical = True And cbCopyArea.Vertical = False Then
                cbAnchDiff = cbPasteArea.Row - cbCopyArea.Row
                cbTAGAdjust = cbAnchDiff - cbCommentCount + 1 'if we are going backwards does that mean it should be -1?????
                Debug.Print "Combo 3"
            ElseIf .Vertical = False And cbCopyArea.Vertical = True Then
                cbAnchDiff = cbPasteArea.Col - cbCopyArea.Col
                cbTAGAdjust = cbAnchDiff - cbCommentCount + 1 'if we are going backwards does that mean it should be -1?????
                Debug.Print "Combo 4"
            End If
            Debug.Print "cbTAGAdjust= " & cbTAGAdjust
            Debug.Print "cbAnchDiff= " & cbAnchDiff
            
            cbCommentCount = 0
            cbMyIndex = 0
            
            For Each C In .Range.Cells

                If cbCommentCount = UBound(arrMyComments()) Then
                cbCommentCount = 0             'looping thru the comments
                cbMyIndex = 0                       'correction prob cbMyIndex not necessary and cbTAGAdjust should just be adjusted by 1
                cbMyIndex = cbMyIndex + 1
                cbTAGAdjust = cbTAGAdjust + cbMyIndex
                End If
                Debug.Print "1) cbMyIndex= " & cbMyIndex
                Debug.Print "2) cbTAGAdjust= " & cbTAGAdjust
                cbCommentCount = cbCommentCount + 1
                Debug.Print "3) cbCommentCount= " & cbCommentCount
                'insert tag adjustment
                
                 arrCurrentTag() = Split(arrMyComments(cbCommentCount), "&")        'this could be wrong **************************************

With MyCurrentTag
                    .PublicationDate = CDbl(arrCurrentTag(1))
                    If cbPasteArea.Count = 1 Then
                    .Horizontal = CBool(arrCurrentTag(2))
                    Else
                    .Horizontal = Not (cbPasteArea.Vertical)
                    End If
                    .ChngNegative = CBool(arrCurrentTag(3))
                    .Publisher = CStr(arrCurrentTag(4))
                    .MyData = CStr(arrCurrentTag(5))
                    .MyType = CStr(arrCurrentTag(6))
                    .MyZone = CStr(arrCurrentTag(7))
                    .TargetDate = CDbl(arrCurrentTag(8))
                    .DeltaDate = CDbl(arrCurrentTag(9))
                    'Now adjust the Target Date - leaving Delta Date unchanged but can be changed here if revision required
                    .TargetDate = .TargetDate + cbTAGAdjust
Select Case .MyData
                
Case Is = "ACT"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    Actual" & Chr(10) _
    & "Target  Date: " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&ACT&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & "0" & "&CompTAG>>"  'new

Case Is = "FOR_OP"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    FOR_OP" & Chr(10) _
    & "Target Date:  " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FOR_OP&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & "0" & "&CompTAG>>"

Case Is = "FAD_OP"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    FAD_OP" & Chr(10) _
    & "Target Date: " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) & "Delta Date:   " & CStr(Format(.DeltaDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta    Date: __/__/____"  'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FAD_OP&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & .DeltaDate & "&CompTAG>>"

Case Is = "FDE_OP"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    FDE_OP" & Chr(10) _
    & "Target Date:  " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) & "Delta  Date:   " & CStr(Format(.DeltaDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta    Date: __/__/____"  'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FDE_OP&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & .DeltaDate & "&CompTAG>>"

Case Is = "FSH_OP"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    FSH_OP" & Chr(10) _
    & "Target Date:  " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new

    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FSH_OP&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & (.TargetDate - 1) & "&CompTAG>>"

Case Is = "FOR_ES"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    FOR_ES" & Chr(10) _
    & "Target Date:  " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new

    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FOR_ES&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & "0" & "&CompTAG>>"
      
Case Is = "FAD_ES"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    FAD_ES" & Chr(10) _
    & "Target Date:  " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) & "Delta Date:    " & CStr(Format(.DeltaDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta Date:   __/__/____"  'new
    
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FAD_ES&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & .DeltaDate & "&CompTAG>>"

Case Is = "FDE_ES"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:        " & .MyType & Chr(10) & "Data Type:    FDE_ES" & Chr(10) _
    & "Target Date:  " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) & "Delta  Date:   " & CStr(Format(.DeltaDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" & Chr(10) & "Delta    Date: __/__/____"  'new
                
    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FDE_ES&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & .DeltaDate & "&CompTAG>>"
 
Case Is = "FSH_ES"
    sCommentTT = "Tag Date:     " & CStr(Format(.PublicationDate, "DD/MM/YYYY")) & Chr(10) & "Publisher:      " & .Publisher & Chr(10) & "Zone:           " & .MyZone & Chr(10) & "Product:       " & .MyType & Chr(10) & "Data Type:    FSH_ES" & Chr(10) _
    & "Target Date:  " & CStr(Format(.TargetDate, "DD/MM/YYYY")) & Chr(10) _
    & Chr(10) & "Refresh Details:" & Chr(10) & "Refresh Date: __/__/____" & Chr(10) & "Target  Date: __/__/____" 'new

    sCommentTT = sCommentTT & Chr(10) & Chr(10) & Chr(10) & "<<CompTAG:&" & CStr(CDbl(.PublicationDate)) & "&" & .Horizontal & "&" & .ChngNegative & "&" & .Publisher & "&FSH_ES&" & .MyType & "&" & .MyZone & "&" & .TargetDate & "&" & (.TargetDate - 1) & "&CompTAG>>"
                            
End Select
     End With
                
C.Cells(1, 1).AddComment.Text CStr(sCommentTT)
   If C.Cells(1, 1).Comment Is Nothing Then
       C.Cells(1, 1).AddComment.Text sCommentTT
    Else
        C.Cells(1, 1).Comment.Delete
        C.Cells(1, 1).AddComment.Text sCommentTT
    End If
       C.Cells(1, 1).Comment.Shape.TextFrame.AutoSize = True
    If C.Cells(1, 1).Comment.Shape.Width > 125 Then
       lArea = C.Cells(1, 1).Comment.Shape.Width * C.Cells(1, 1).Comment.Shape.Height
       C.Cells(1, 1).Comment.Shape.Width = 120
       C.Cells(1, 1).Comment.Shape.Height = 135

    End If
            Next C

End With


Set cbPasteArea.Range = Nothing: Set cbCopyArea.Range = Nothing: Erase arrMyComments()

End Sub

