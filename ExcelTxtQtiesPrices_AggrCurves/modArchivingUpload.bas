Attribute VB_Name = "modArchivingUpload"

Sub ArchiveACUpload()

    Dim Sourcewb As Workbook
    Dim DestWB As Workbook
    Dim MyFilePathName As String, MyDate As String

    Set Sourcewb = ThisWorkbook

    For Each WS In Sourcewb.Worksheets
    
    If WS.Name = "ES" Or WS.Name = "PT" Then
    MyDate = Sourcewb.Worksheets("Dashboard").Range("O9").Value
    MyFileName = "IB" & MyDate
    Else
    GoTo MySkipover:
    End If
    
    Sourcewb.Sheets(Array("ES", "PT")).Copy

    Set DestWB = ActiveWorkbook

    MyFilePathName = ThisWorkbook.Worksheets("Dashboard").Range("J16").Value

If Val(Application.Version) > 11 Then

    With DestWB
        .SaveAs MyFilePathName, FileFormat:=56
        .Close savechanges:=False
    End With

Else
    With DestWB
        .SaveAs MyFilePathName & ".xls"
        .Close savechanges:=False
    End With


End If

MySkipover:
    
    Next WS

    Set Sourcewb = Nothing
    Set DestWB = Nothing

  
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub


