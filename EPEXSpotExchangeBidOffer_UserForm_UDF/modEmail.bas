Attribute VB_Name = "modEmail"
Option Explicit

'***********************************
'EPEX_PowerBidOffer_Template.xlsm

'CH_Trigger_Results_20170727.xls is the output of the email (attached file created)
'20170727 is the Delivery Day (D+1)
'***********************************

Sub FinalBookResults_Mail()

' Works in Excel 2000, Excel 2002, Excel 2003, Excel 2007, Excel 2010, Outlook 2000, Outlook 2002, Outlook 2003, Outlook 2007, Outlook 2010 etc.
    
    Dim Source As Range
    Dim Dest As Workbook
    Dim WB As Workbook
    Dim TempFilePath As String
    Dim TempFileName, TempFile As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String
    Dim fso As Object
    Dim ts As Object

    Set Source = Nothing
    On Error Resume Next
    Set Source = ThisWorkbook.Worksheets("MarketResults").Range("A1:M63")
    On Error GoTo 0

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    Set WB = ActiveWorkbook
    Set Dest = Workbooks.Add(xlWBATWorksheet)
    Source.Copy
    With Dest.Worksheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial Paste:=xlPasteValues
        .Cells(1).PasteSpecial Paste:=xlPasteFormats
        .Cells(1).PasteSpecial Paste:=xlPasteFormulas
        .Cells(1).Select
       ' .Columns("A:M").EntireColumn.AutoFit
        Application.CutCopyMode = False
       .Range("A65:D74").Value = ThisWorkbook.Worksheets("MarketResults").Range("N13:Q22").Value
       .Range("H39:M63").ClearContents
    End With

    'TempFilePath = Environ$("temp") & "\"
    TempFilePath = ThisWorkbook.Worksheets("MyLists").Range("FolderPathtoUse").Value & _
    IIf(Right(ThisWorkbook.Worksheets("MyLists").Range("FolderPathtoUse").Value, 1) = "\", "", "\")
    
    Dim Damsfx As String
    Damsfx = MyDAMSuffix(ThisWorkbook.Worksheets("MyTemplate").Range("B2"))
    
    TempFileName = Damsfx & "_Trigger_Results_" & Format(Range("B2"), "YYYYMMDD")
    
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    FileExtStr = ".xls": FileFormatNum = -4143
    
    'print in an Html Format the Email Body that would like to be shown
    With Dest.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=Dest.Worksheets(1).Name, _
         Source:=Dest.Worksheets(1).Range("A38:G75").Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    strbody = ts.ReadAll
    ts.Close
    strbody = Replace(strbody, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Dest
        .Worksheets(1).Range("A65:D74").ClearContents
        .SaveAs TempFilePath & TempFileName & FileExtStr, _
                FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = "faidon.dermesonoglou@gmail.com" 'using ; one can send to multiple recipients
            .CC = ""
            .BCC = ""
            .Subject = TempFileName
            .HTMLBody = strbody
            '.Body = ThisWorkbook.Worksheets("MarketResults").Range("N13:S22")
            .Attachments.Add Dest.FullName
            ' You can add other files by uncommenting the following statement.
            '.Attachments.Add ("C:\test.txt")
            ' In place of the following statement, you can use ".Display" to
            ' display the e-mail message.
            .Send
        End With
        On Error GoTo 0
        .Close SaveChanges:=False
    End With

    'Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing
    
    Set fso = Nothing
    Set ts = Nothing
    
    Set WB = Nothing
    Set Dest = Nothing
    
    Set Source = Nothing

    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
End Sub


