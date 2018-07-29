Attribute VB_Name = "modXMLExportFloatP"
Option Explicit
Option Compare Text
Option Base 1

'Used Microsoft XML v3.0 library

Sub CreateXMLFloatsP() 'XML Export (i.e. File Creation) of Prices... Values have two decimal values as default

Dim xmlDoc As DOMDocument, objIntro As IXMLDOMProcessingInstruction
Dim objRoot As IXMLDOMElement, objRecord As IXMLDOMElement, objNameF As IXMLDOMElement, objNameFF As IXMLDOMElement

Dim row As Long, col As Long, NumberofRecords As Long, NumberofCategories As Long
Dim MyCurrentFolder As String, MyDataSample() As Variant
Dim HourLoop As Long, colH As Long
Dim c As Range
Dim MyAvg() As Double

ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Calculate
MyCurrentFolder = ThisWorkbook.Worksheets("Dashboard").Range("XMLFolder").Value 'ThisWorkbook Saved Folder
                                                                                                            
Range("Hour23").Calculate
Range("Hour25").Calculate

On Error GoTo NoCreation
                                                                                                            
MyDataSample = Range("FloatP")

NumberofRecords = UBound(MyDataSample(), 1)       'Rows
NumberofCategories = UBound(MyDataSample(), 2)    'Cols

'*************************Headers********************************

Set xmlDoc = CreateObject("Microsoft.XMLDOM")  'or Set it as New Dom Document
Set objIntro = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
xmlDoc.InsertBefore objIntro, xmlDoc.ChildNodes(0)

Set objRoot = xmlDoc.createElement("Data"): xmlDoc.appendChild objRoot

With objRoot 'Attributes
    .setAttribute "Purpose", MyDataSample(1, 1)
    .setAttribute "Year", ThisWorkbook.Worksheets("Dashboard").Range("Year").Value
    '.setAttribute "Version", "1"
    .setAttribute "LastUpdate", Format(Now(), "YYYYMMDD hh:mm")
End With
    
    Set objRecord = xmlDoc.createElement(MyDataSample(1, 1)): objRoot.appendChild objRecord
                                                                                                                                                                                
    For col = 1 To NumberofCategories Step 28

    For row = 4 To NumberofRecords
    HourLoop = 0
    
    If MyDataSample(2, col) <> "" Then
    Set objNameF = xmlDoc.createElement(MyDataSample(2, col) & "Index"): objRecord.appendChild objNameF
    Else
    GoTo NextCol
    End If
    
    Set objNameFF = xmlDoc.createElement("Month"): objNameF.appendChild objNameFF   'Month
    objNameFF.Text = CStr(Month(MyDataSample(row, col)))
    Set objNameFF = xmlDoc.createElement("Day"): objNameF.appendChild objNameFF     'Day
    objNameFF.Text = CStr(Day(MyDataSample(row, col)))
    Set objNameFF = xmlDoc.createElement("Weekday"): objNameF.appendChild objNameFF 'Weekday
    objNameFF.Text = CStr(Weekday(MyDataSample(row, col), 2))
    Set objNameFF = xmlDoc.createElement("PublicHoliday"): objNameF.appendChild objNameFF 'PublicHoliday
    objNameFF.Text = CStr("NonH")
    For Each c In ThisWorkbook.Worksheets("Dashboard").Range("NatHolidays").Cells
    If CDate(c.Value) = MyDataSample(row, col) Then objNameFF.Text = CStr("Hol")
    Next c
    Set objNameFF = xmlDoc.createElement("Status"): objNameF.appendChild objNameFF 'Status
    objNameFF.Text = CStr(MyDataSample(2, col))
    
    For colH = 1 To 24 '24 are the hours and internally loop for h25
                                                                                                                                                                                
NextHour25:
    HourLoop = HourLoop + 1
    
    Set objNameFF = xmlDoc.createElement("H" & CStr(HourLoop)): objNameF.appendChild objNameFF 'Hourly values

    If IsNumeric(MyDataSample(row, col + colH)) Then
    
    If Application.DecimalSeparator = "," And Int(MyDataSample(row, col + colH)) <> MyDataSample(row, col + colH) Then
    objNameFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyDataSample(row, col + colH))), "#0.00"), ",", "."))
    'or
    'objNameFF.Text = CStr(Replace(Format(MyDataSample(row, col + colH), "#0.00"), ",", "."))

    Else
    objNameFF.Text = CStr(MyDataSample(row, col + colH))
    End If
    
    If colH >= 24 And Range("Hour23").Value = MyDataSample(row, col) Then 'SummerTime
    objNameFF.Text = vbNullString
    End If
    
    If colH > 24 And Range("Hour25").Value <> MyDataSample(row, col) Then 'WinterTime
    objNameFF.Text = vbNullString
    End If
    
    Else
    
    objNameFF.Text = vbNullString
    'objNameFF.Text = CStr(MyDataSample(row, col + colH))
    'objNameFF.Text = CStr(0)
    End If
    
    If colH = 24 Then
    colH = colH + 1
    GoTo NextHour25
    End If
    
    Next colH
    
    MyAvg = MyAverages(MyDataSample, row, col) 'Bload, Peak, Offpeak, OffP1, OffP2
    
    Set objNameFF = xmlDoc.createElement("Bload"): objNameF.appendChild objNameFF
    If MyAvg(1) + MyAvg(2) + MyAvg(3) + MyAvg(4) + MyAvg(5) = 0 Then
    objNameFF.Text = vbNullString
    Else
    
    If Application.DecimalSeparator = "," Then
    objNameFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyAvg(1))), "#0.00"), ",", "."))
    'or
    'objNameFF.Text = Replace(Format(MyAvg(1), "#0.00"), ",", ".")
    
    Else
    objNameFF.Text = CStr(MyAvg(1))
    End If
    
    End If
    
    Set objNameFF = xmlDoc.createElement("Peak"): objNameF.appendChild objNameFF
    If MyAvg(1) + MyAvg(2) + MyAvg(3) + MyAvg(4) + MyAvg(5) = 0 Then
    objNameFF.Text = vbNullString
    Else
    
    If Application.DecimalSeparator = "," Then
    objNameFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyAvg(2))), "#0.00"), ",", "."))
    'or
    'objNameFF.Text = Replace(Format(MyAvg(2), "#0.00"), ",", ".")
    
    Else
    objNameFF.Text = CStr(MyAvg(2))
    End If
    
    End If
    
    Set objNameFF = xmlDoc.createElement("Offpeak"): objNameF.appendChild objNameFF
    If MyAvg(1) + MyAvg(2) + MyAvg(3) + MyAvg(4) + MyAvg(5) = 0 Then
    objNameFF.Text = vbNullString
    Else
    
    If Application.DecimalSeparator = "," Then
    objNameFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyAvg(3))), "#0.00"), ",", "."))
    'or
    'objNameFF.Text = Replace(Format(MyAvg(3), "#0.00"), ",", ".")
    
    Else
    objNameFF.Text = CStr(MyAvg(3))
    End If
    
    End If
    
    Set objNameFF = xmlDoc.createElement("OffP1"): objNameF.appendChild objNameFF
    If MyAvg(1) + MyAvg(2) + MyAvg(3) + MyAvg(4) + MyAvg(5) = 0 Then
    objNameFF.Text = vbNullString
    Else
    
    If Application.DecimalSeparator = "," Then
    objNameFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyAvg(4))), "#0.00"), ",", "."))
    'or
    'objNameFF.Text = Replace(Format(MyAvg(4), "#0.00"), ",", ".")
    
    Else
    objNameFF.Text = CStr(MyAvg(4))
    End If
    
    End If
    
    Set objNameFF = xmlDoc.createElement("OffP2"): objNameF.appendChild objNameFF
    If MyAvg(1) + MyAvg(2) + MyAvg(3) + MyAvg(4) + MyAvg(5) = 0 Then
    objNameFF.Text = vbNullString
    Else
    
    If Application.DecimalSeparator = "," Then
    objNameFF.Text = CStr(Replace(Format(XMLStringtoVal(CStr(MyAvg(5))), "#0.00"), ",", "."))
    'or
    'objNameFF.Text = Replace(Format(MyAvg(5), "#0.00"), ",", ".")
    
    Else
    objNameFF.Text = CStr(MyAvg(5))
    End If
    
    End If
    
    Erase MyAvg()
    
NextRow:
    Next row
    
NextCol:
    Next col
            

xmlDoc.Save MyCurrentFolder & "/" & MyDataSample(1, 1) & "_" & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value & _
".xml"

NoCreation:
Set xmlDoc = Nothing: Set objRoot = Nothing: Set objRecord = Nothing: Set objNameF = Nothing: Set objNameFF = Nothing
Set objIntro = Nothing: Erase MyDataSample() 'Clean Up

'Conclusion Msgbox
If Err.Number <> 0 Then
MsgBox "Market prices created for " & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value, vbCritical
Exit Sub
End If

MsgBox "Market prices created for " & ThisWorkbook.Worksheets("Dashboard").Range("Year").Value

ThisWorkbook.Worksheets("Dashboard").Range("M17").Value = Format(Now, "dd/mm/yyyy hh:mm:ss")
If DateSerial(Year(Now), Month(Now), Day(Now)) <> ThisWorkbook.Worksheets("Dashboard").Range("M17").Value Then
ThisWorkbook.Worksheets("Dashboard").Range("M17").Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
End If
End Sub
