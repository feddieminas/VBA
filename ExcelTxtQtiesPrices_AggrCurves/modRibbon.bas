Attribute VB_Name = "modRibbon"
Option Explicit

Sub MyMacro1(iribbon As IRibbonControl)
MyOMELPriceDownload

ThisWorkbook.Worksheets("Dashboard").Range("B17").Value = "Prices"
End Sub


Sub MyMacro2(iribbon As IRibbonControl)
MyOMELACDownload

ThisWorkbook.Worksheets("Dashboard").Range("B17").Value = "Create AC Curves"
End Sub


Sub MyMacro3(iribbon As IRibbonControl)

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With

ReconfigureZones
MySourceDataResort

SummariseBO
PrepareZonalSheets

CreateACCurves
ArchiveACUpload
'Application.Wait (Now + TimeValue("0:00:01"))
Application.GoTo ThisWorkbook.Worksheets("Dashboard").Range("B15")
ThisWorkbook.Worksheets("Dashboard").Range("B17").Value = "AC Curves"

With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub


