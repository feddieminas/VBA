Attribute VB_Name = "modXMLExportFloatQ"
Option Explicit
Option Compare Text
Option Base 1

'Used Microsoft XML v3.0 library

Sub CreateXMLFloatsQ() 'XML Export (i.e. File Creation) of Qties... Values have three decimal values as default

Dim xmlDoc As DOMDocument, objIntro As IXMLDOMProcessingInstruction
Dim objRoot As IXMLDOMElement, objRecord As IXMLDOMElement, objNameF As IXMLDOMElement, objNameFF As IXMLDOMElement, objNameFFF As IXMLDOMElement

Dim row As Long, col As Long, NumberofRecords As Long, NumberofCategories As Long
Dim MyCurrentFolder As String, MyDataSample() As Variant
Dim HourLoop As Long, colQ As Long, colH As Long, i As Long
Dim c As Range

ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Calculate
MyCurrentFolder = ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Value 'ThisWorkbook Saved Folder

Application.ScreenUpdating = False
'''Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Range("Hour23").Calculate
Range("Hour25").Calculate

If ThisWorkbook.Worksheets("QQties").Range("B1").Value = "" Then GoTo NoCreation
On Error GoTo NoCreation

MyDataSample = Array(Range("FloatQQ"), Range("FloatHQ"))

NumberofRecords = Range("FloatQQ").Rows.Count        'Rows
NumberofCategories = Range("FloatQQ").Columns.Count  'Cols

'*************************Headers********************************

Set xmlDoc = CreateObject("Microsoft.XMLDOM")  'or Set it as New Dom Document
Set objIntro = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
xmlDoc.InsertBefore objIntro, xmlDoc.ChildNodes(0)

Set objRoot = xmlDoc.createElement("Data"): xmlDoc.appendChild objRoot

With objRoot 'Attributes
    .setAttribute "Purpose", MyDataSample(1)(1, 1)
    .setAttribute "Year", ThisWorkbook.Worksheets("Dashboard").Range("Year").Value
    '.setAttribute "Version", "1"
    .setAttribute "LastUpdate", Format(Now(), "YYYYMMDD hh:mm")
End With

    Set objRecord = xmlDoc.createElement(MyDataSample(1)(1, 1)): objRoot.appendChild objRecord
                                                                                                
    colH = 1
    For col = 1 To NumberofCategories Step 103 'Step 103 since it loops over the right in case you have more than one yearly profile. Ex. Real and Forecast
    
    If col > 1 Then colH = 29
    
    For row = 4 To NumberofRecords
    HourLoop = 0
    
    If MyDataSample(1)(2, col) <> "" Then
    Set objNameF = xmlDoc.createElement(MyDataSample(1)(2, col)): objRecord.appendChild objNameF 'Real and Forecast
    Else
    GoTo NextCol
    End If
    
    Set objNameFF = xmlDoc.createElement(ThisWorkbook.Worksheets("QQties").Cells(1, 1).Value): objNameF.appendChild objNameFF
    objNameFF.Text = CStr(ThisWorkbook.Worksheets("QQties").Cells(1, 2).Value) 'Client
    Set objNameFF = xmlDoc.createElement(ThisWorkbook.Worksheets("QQties").Cells(2, 1).Value): objNameF.appendChild objNameFF
    objNameFF.Text = CStr(ThisWorkbook.Worksheets("QQties").Cells(2, 2).Value) 'POD
    Set objNameFF = xmlDoc.createElement(ThisWorkbook.Worksheets("QQties").Cells(3, 1).Value): objNameF.appendChild objNameFF
    objNameFF.Text = CStr(ThisWorkbook.Worksheets("QQties").Cells(3, 2).Value) 'Zone
    
    Set objNameFF = xmlDoc.createElement("Month"): objNameF.appendChild objNameFF 'Month
    objNameFF.Text = CStr(Month(MyDataSample(1)(row, col)))
    Set objNameFF = xmlDoc.createElement("Day"): objNameF.appendChild objNameFF 'Day
    objNameFF.Text = CStr(Day(MyDataSample(1)(row, col)))
    Set objNameFF = xmlDoc.createElement("Weekday"): objNameF.appendChild objNameFF 'Weekday
    objNameFF.Text = CStr(Weekday(MyDataSample(1)(row, col), 2))
    Set objNameFF = xmlDoc.createElement("PublicHoliday"): objNameF.appendChild objNameFF 'PublicHoliday
    objNameFF.Text = CStr("NonH")
    For Each c In ThisWorkbook.Worksheets("Dashboard").Range("NatHolidays").Cells
    If CDate(c.Value) = MyDataSample(1)(row, col) Then objNameFF.Text = CStr("Hol")
    Next c
    Set objNameFF = xmlDoc.createElement("Status"): objNameF.appendChild objNameFF 'Status
    objNameFF.Text = CStr(MyDataSample(1)(2, col))
    

    For colQ = 1 To 100 Step 4 '100 since are 25 hours multiplied the 4 quarterly values
                                                                                                                                                                                
    HourLoop = HourLoop + 1

    Set objNameFF = xmlDoc.createElement("HQ" & CStr(HourLoop)): objNameF.appendChild objNameFF   'Node of Quarterly and Hourly values
    
    Set objNameFFF = xmlDoc.createElement("H" & CStr(HourLoop)): objNameFF.appendChild objNameFFF 'Hourly values

    If IsNumeric(MyDataSample(2)(row, colH + HourLoop)) Then
    
    If Application.DecimalSeparator = "," And Int(MyDataSample(2)(row, colH + HourLoop)) <> MyDataSample(2)(row, colH + HourLoop) Then
    objNameFFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyDataSample(2)(row, colH + HourLoop))), "#0.000"), ",", "."))
    'or
    'objNameFFF.Text = Replace(Format(MyDataSample(2)(row, colH + HourLoop), "#0.000"), ",", ".")
    
    Else
    objNameFFF.Text = CStr(MyDataSample(2)(row, colH + HourLoop))
    End If
    
    Else
    objNameFFF.Text = CStr(MyDataSample(2)(row, colH + HourLoop))
    End If
    
    If HourLoop >= 24 And Range("Hour23").Value = MyDataSample(2)(row, colH) Then 'SummerTime
    objNameFFF.Text = vbNullString
    End If
    
    If HourLoop > 24 And Range("Hour25").Value <> MyDataSample(2)(row, colH) Then 'WinterTime
    objNameFFF.Text = vbNullString
    End If
    
        For i = 1 To 4  'Quarterly values
        Set objNameFFF = xmlDoc.createElement("Q" & CStr(i)): objNameFF.appendChild objNameFFF
    
        If IsNumeric(MyDataSample(1)(row, col + colQ + i - 1)) Then
    
        If Application.DecimalSeparator = "," And Int(MyDataSample(1)(row, col + colQ + i - 1)) <> MyDataSample(1)(row, col + colQ + i - 1) Then
        objNameFFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyDataSample(1)(row, col + colQ + i - 1))), "#0.000"), ",", "."))
        'or
        'objNameFFF.Text = Replace(Format(MyDataSample(1)(row, col + colQ + i - 1), "#0.000"), ",", ".") 'Works also
    
        Else
        objNameFFF.Text = CStr(MyDataSample(1)(row, col + colQ + i - 1))
        End If
    
        Else
        objNameFFF.Text = CStr(MyDataSample(1)(row, col + colQ + i - 1))
        End If
    
        If colQ >= 92 And Range("Hour23").Value = MyDataSample(1)(row, col) Then 'SummerTime
        objNameFFF.Text = vbNullString
        End If
    
        If colQ > 96 And Range("Hour25").Value <> MyDataSample(1)(row, col) Then 'WinterTime
        objNameFFF.Text = vbNullString
        End If
    
        Next i
    
    Next colQ
    
NextRow:
    Next row
    
NextCol:
    Next col
            

xmlDoc.Save MyCurrentFolder & "/" & MyDataSample(1)(1, 1) & "_" & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value & _
"_" & ThisWorkbook.Worksheets("Dashboard").Range("B13").Value & ".xml"

NoCreation:
Set xmlDoc = Nothing: Set objRoot = Nothing: Set objRecord = Nothing: Set objNameF = Nothing: Set objNameFF = Nothing
Set objIntro = Nothing: Erase MyDataSample() 'Clean Up

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
'''Application.EnableEvents = True

'Conclusion Msgbox
If Err.Number <> 0 Then
MsgBox "Qties not created for Client " & ThisWorkbook.Worksheets("Dashboard").Range("B10").Value & " Year " _
& ThisWorkbook.Worksheets("Dashboard").Range("Year").Value, vbCritical
Exit Sub
End If

MsgBox "Qties created for Client " & ThisWorkbook.Worksheets("Dashboard").Range("B10").Value & " Year " _
& ThisWorkbook.Worksheets("Dashboard").Range("Year").Value

ThisWorkbook.Worksheets("Dashboard").Range("M10").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
If DateSerial(Year(Now), Month(Now), Day(Now)) <> ThisWorkbook.Worksheets("Dashboard").Range("M10").Value Then
ThisWorkbook.Worksheets("Dashboard").Range("M10").Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
End If
End Sub


