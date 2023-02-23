Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Sub Intro()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False

End Sub

Sub Conclusion()

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

'********************
'INSERT
'********************
Sub InsertList()

Call Intro

Dim IData As cImportData
Set IData = New cImportData
With IData ' Import the Data on your Destination Sheet
'URL LINK
.sSourceSharePointURL = ThisWorkbook.Worksheets("Settings").Range("B1").Value
'LIST
'https://nickgrattan.wordpress.com/2008/04/29/finding-the-id-guid-for-a-sharepoint-list/
.sSourceSharePointListGUID = ThisWorkbook.Worksheets("Settings").Range("B2").Value
.sSourceSharePointListName = ThisWorkbook.Worksheets("Settings").Range("B3").Value
.fInsLst
End With
Set IData = Nothing

Call Conclusion

End Sub

'********************
'UPDATE
'********************
Sub UpdateList()

Call Intro

Dim IData As cImportData
Set IData = New cImportData
With IData ' Import the Data on your Destination Sheet
'URL LINK
.sSourceSharePointURL = ThisWorkbook.Worksheets("Settings").Range("B1").Value
'LIST
'https://nickgrattan.wordpress.com/2008/04/29/finding-the-id-guid-for-a-sharepoint-list/
.sSourceSharePointListGUID = ThisWorkbook.Worksheets("Settings").Range("B2").Value
.sSourceSharePointListName = ThisWorkbook.Worksheets("Settings").Range("B3").Value
.fUpdLst
End With
Set IData = Nothing

Call Conclusion

End Sub

'********************
'DELETE
'********************
Sub DeleteList()

Call Intro

Dim IData As cImportData
Set IData = New cImportData
With IData ' Import the Data on your Destination Sheet
'URL LINK
.sSourceSharePointURL = ThisWorkbook.Worksheets("Settings").Range("B1").Value
'LIST
'https://nickgrattan.wordpress.com/2008/04/29/finding-the-id-guid-for-a-sharepoint-list/
.sSourceSharePointListGUID = ThisWorkbook.Worksheets("Settings").Range("B2").Value
.sSourceSharePointListName = ThisWorkbook.Worksheets("Settings").Range("B3").Value
.fDelLst
End With
Set IData = Nothing

Call Conclusion

End Sub

'********************
'READ
'********************
Sub ReadImportList()

Call Intro

Dim IData As cImportData
Set IData = New cImportData
With IData ' Import the Data on your Destination Sheet
'URL LINK
.sSourceSharePointURL = ThisWorkbook.Worksheets("Settings").Range("B1").Value
'LIST
'https://nickgrattan.wordpress.com/2008/04/29/finding-the-id-guid-for-a-sharepoint-list/
.sSourceSharePointListGUID = ThisWorkbook.Worksheets("Settings").Range("B2").Value
.sSourceSharePointListName = ThisWorkbook.Worksheets("Settings").Range("B3").Value
.sDestWSName = "Sheet1"
.fImpLst True
End With
Set IData = Nothing

Call Conclusion

End Sub
