VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'***********************************
'EPEX_PowerBidOffer_Template.xlsm

'To avoid having any recalculation issues and UDF functions are not getting recognised (NAME shown in cells),
'I recommend you do the following step :
'Go to Excel Options --> Trust Center --> Trust Center Settings --> Trusted locations --> Click below on
'Allow Trusted Locations on my network (not recommended)
'Then press on Add new location..., browse your filepath and click on 'Subfolders of this locator are also trusted'.
'Press ok, ok to Trusted Locations and finally ok on Excel Options.
'Save the file, close Microsoft Excel Application and open again the file.
'***********************************

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean) 'Double Click

Application.EnableEvents = False
Application.ScreenUpdating = False

Select Case Target.Address

Case Is = Range("B5").Address

Range("E1").Calculate
MsgBox "Book Scheduling = " & Range("E1").Value

Case Is = Range("M3").Address, Range("N3").Address

MyTriggerInput

Case Is = Range("S34").Address, Range("T34").Address, Range("U34").Address, Range("V34").Address

Range("P56:AN75").ClearContents 'Block Bids 20 rows

Case Else
End Select

Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range) 'Single Click

Application.EnableEvents = False

Dim rng As Range
Set rng = Range("Q3,R3,S3,T3")
If Intersect(Target, rng) Is Nothing Then GoTo myEnd

On Error GoTo myEnd

Letstry

Range("N1").Value = Format(Now(), "yyyy-MM-dd hh:mm:ss")
Range("M1").FormulaR1C1 = "Exchange Offer last created on " & (WeekdayName(Weekday(Range("N1"), vbMonday)))

myEnd:

Set rng = Nothing
        
Application.EnableEvents = True

End Sub

Sub BlockBidsCopy()

On Error Resume Next
Range("Blockbids").Copy
On Error GoTo 0

End Sub
