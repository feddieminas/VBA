VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'Worked the project in Windows 7 and Excel 2007, 2010
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'Create Aggregate Curves for the Iberia Power Market

Private Sub Workbook_Open()
Exit Sub 'to avoid going the file in automatic whilst Open. Can use the ribbon button
         'If one wants automatically the file to run, comment the exit sub

Application.DisplayAlerts = False
Application.ScreenUpdating = False
        
MyOMELPriceDownload
ThisWorkbook.Worksheets("Dashboard").Range("B17").Value = "Prices"
'************************************************************************'

MyOMELACDownload
ThisWorkbook.Worksheets("Dashboard").Range("B17").Value = "AC Curves"
'************************************************************************'

ReconfigureZones
MySourceDataResort

SummariseBO
PrepareZonalSheets

CreateACCurves
ArchiveACUpload
'Application.Wait (Now + TimeValue("0:00:01"))
Application.GoTo ThisWorkbook.Worksheets("Dashboard").Range("B15")
ThisWorkbook.Worksheets("Dashboard").Range("B17").Value = "AC Curves"
'************************************************************************'
      
'Application.Quit

    'With ThisWorkbook
          '.Close savechanges:=True
    'End With
'Application.Quit

End Sub
