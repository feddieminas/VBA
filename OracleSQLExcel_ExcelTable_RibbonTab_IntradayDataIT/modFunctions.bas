Attribute VB_Name = "modFunctions"

#If Mac Then

#If VBA7 Then
Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As LongPtr
#Else
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#End If

#ElseIf VBA7 Then
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#Else
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
#End If


Function isworkbookopen(ByRef checkbook As String)

    Dim wbook As Workbook

    On Error Resume Next
    Set wbook = Workbooks(checkbook)

    If wbook Is Nothing Then
    isworkbookopen = False
    Else
    isworkbookopen = True
    End If

    Set wbook = Nothing
    On Error GoTo 0

End Function

Function Cl_G_H(Target_Sheet, Flag)

    Dim Graph_R As Variant
    Dim H_links As Variant
    
    For Each Graph_R In ThisWorkbook.Sheets(Target_Sheet).Shapes
        Graph_R.Delete
    Next Graph_R
    
    For Each H_links In ThisWorkbook.Sheets(Target_Sheet).Hyperlinks
        H_links.Delete
    Next H_links
    
    If Flag = 1 Then ThisWorkbook.Sheets(Target_Sheet).Cells.Clear
    

End Function
