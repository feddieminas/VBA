Attribute VB_Name = "modMain"

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Worked the project in Windows 7 and Excel 2007, 2010
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'Need to download 7Zip
'Need to create manually a directory to insert your zip file ABN Amro Clearing - Sample XML files.zip
'Default folder is C:\ABNAmbroXML
'If needed one can modify the folder directory on Sheet Dashboard at Range E15.
'If one would modify it on spreadsheet, then it would need to create manually a directory on PC
'to insert the zip file.

'Libraries Used
'Visual Basic for Applications
'Microsoft Excel 12.0 Object Library
'OLE Automation
'Microsoft Office 12.0 Object Library
'Microsoft XML v6.0
'Microsoft Forms 12.0 Object Library

Sub ClearingWB()

MyClearWS
MyClearDash

End Sub

Sub Unzipping()

MyUnzip

End Sub

Sub ABNAmbroFiles()

ABNFilesLoop

End Sub

Sub Archiving()

MyArchive

End Sub
