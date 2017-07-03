Attribute VB_Name = "modDownload"
Option Explicit
Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias _
    "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Function MyTextFileQuery(MyDirSourceWb As String, MyDestWS As String) As Boolean
Dim connectionName As String, qt As QueryTable

connectionName = "TEXT;" + MyDirSourceWb

On Error GoTo MyError:

With Worksheets(MyDestWS).QueryTables.Add(Connection:=connectionName, Destination:=Worksheets(MyDestWS).Range("A1"))
.Name = MyDirSourceWb
.FieldNames = True
.RowNumbers = False
.FillAdjacentFormulas = False
.PreserveFormatting = True
.RefreshOnFileOpen = False
.RefreshStyle = xlOverwriteCells
.SavePassword = False
.SaveData = True
.AdjustColumnWidth = True
.RefreshPeriod = 0
.TextFilePromptOnRefresh = False
.TextFilePlatform = 1252                                             'Might need modification
.TextFileDecimalSeparator = ","                                      'for anglosaxon this would be changed to "."
.TextFileThousandsSeparator = "."                                    'Might need modification
.TextFileStartRow = 1                                              'you could change this if you wanted to import starting from a different row
.TextFileParseType = xlDelimited
.TextFileTextQualifier = xlTextQualifierDoubleQuote
.TextFileConsecutiveDelimiter = False
.TextFileTabDelimiter = False
.TextFileSemicolonDelimiter = True                                   'False typical european separator...then should be set to true
.TextFileCommaDelimiter = False                                      'thiss would be the normal anglosaxon comma separator, and would then be set to true
.TextFileSpaceDelimiter = False
'.TextFileOtherDelimiter =  True                                     '"-" 'specifically for EEX Transparency Website
.Refresh BackgroundQuery:=False
End With
MyTextFileQuery = True

For Each qt In ThisWorkbook.Worksheets(MyDestWS).QueryTables                        'this is vital in order to destroy the connection to the text file.
qt.Delete                                                                           'hence why it appears also in the event of error handling
Next qt

Exit Function

MyError:

For Each qt In ThisWorkbook.Worksheets(MyDestWS).QueryTables
qt.Delete
Next qt
MyTextFileQuery = False
End Function

Public Sub OMELPriceDownload()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim MyFDestination As String, MySource As String
Dim l As Long

'********************************* Prices ***************************************'
With ThisWorkbook.Worksheets("Dashboard")
MySource = .Range("J5").Value
MyFDestination = "OMEL"
End With
ThisWorkbook.Worksheets("Notepad").Cells.Clear

l = DeleteUrlCacheEntry(MySource)
Application.Wait Now() + TimeValue("00:00:03")
MyTextFileQuery MySource, MyFDestination
Application.Wait Now() + TimeValue("00:00:03")

'Adjustments
'Insert Headers Year Month Day Hour ES PT
With ThisWorkbook.Worksheets("OMEL")
.Range("A1:F1").Value = Array("YEAR", "MONTH", "DAY", "HOUR", "ES", "PT")
End With

'********************************* Quantities ************************************'
MySource = ThisWorkbook.Worksheets("Dashboard").Range("J6").Value
MyFDestination = "Notepad"
l = DeleteUrlCacheEntry(MySource)
Application.Wait Now() + TimeValue("00:00:03")
MyTextFileQuery MySource, MyFDestination
Application.Wait Now() + TimeValue("00:00:03")

'Adjustments
'put alltogether to the OMEL Worksheet
With ThisWorkbook
.Worksheets("OMEL").Range("A29:AB43").Value = .Worksheets("Notepad").Range("A3:AB17").Value
.Worksheets("OMEL").UsedRange.EntireColumn.AutoFit
.Worksheets("Notepad").Cells.Clear
End With

End Sub

Public Sub MyOMELPriceDownload()
OMELPriceDownload

Dim MyFFDestination As String
'********************************* Archive *****************************************'
MyFFDestination = ThisWorkbook.Worksheets("Dashboard").Range("J14").Value
ThisWorkbook.Worksheets("OMEL").Copy

Dim DestWB As Workbook
Set DestWB = ActiveWorkbook
DestWB.SaveAs Filename:=MyFFDestination, FileFormat:=56
DestWB.Close
Set DestWB = Nothing
End Sub


'*************************** AC Downloads************************************
Public Sub MyOMELACDownload()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Dim MyFDestination As String, MySource As String
With ThisWorkbook.Worksheets("Dashboard")
MySource = .Range("J7").Value
MyFDestination = "ImportedData"
End With
ThisWorkbook.Worksheets("ImportedData").Cells.Clear

'Dim L As Long
'L = DeleteUrlCacheEntry(MySource)
MyTextFileQuery MySource, MyFDestination

Dim MyFFDestination As String

'********************************* Archive *****************************************'

MyFFDestination = ThisWorkbook.Worksheets("Dashboard").Range("J15").Value
ThisWorkbook.Worksheets("ImportedData").Copy

Dim DestWB As Workbook
Set DestWB = ActiveWorkbook
DestWB.Worksheets(1).Name = "MD"
DestWB.SaveAs Filename:=MyFFDestination, FileFormat:=56
DestWB.Close
Set DestWB = Nothing

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub



