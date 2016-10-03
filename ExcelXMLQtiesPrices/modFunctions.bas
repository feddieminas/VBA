Attribute VB_Name = "modFunctions"
Option Explicit
Option Base 1

'Function Retrieves the filepath folder of this Workbook used
Function MyFilepath() As String
Application.Volatile (False)
MyFilepath = ThisWorkbook.Path
Application.Volatile (True)
End Function

'Function Retrieves the March Day of the 23 hours and the October Day of the 25 hours
Public Function MySummertimeAdjustment(StartDate As Double, EndDate As Double, Optional Profile As String = "Baseload") As Date
Application.Volatile (False)

Dim arrSumWin(1 To 2) As Double
Dim arrWinSum(1 To 2) As Double
Dim MyAdjustment As Double
Dim i As Long, YCount As Long

'Convert the Start and End Date as Dates to find your Year
Dim DStartDate As Date, DEndDate As Date
DStartDate = CDate(StartDate)
DEndDate = CDate(EndDate)
Dim YSDeal As Long, YEDeal As Long
YSDeal = Year(DStartDate)
YEDeal = Year(DEndDate)

'Your Summer time is on October. Loop through October Dates to find the Last Sunday of the Month
Dim arrSummerFDS As Date, arrSummerDS(1 To 31) As Date, WDSummerDS As Long
Dim arrSummerFDE As Date, arrSummerDE(1 To 31) As Date, WDSummerDE As Long

'DealStartDate
arrSummerFDS = DateSerial(YSDeal, 10, 1)
For i = 1 To 31
arrSummerDS(i) = arrSummerFDS - 1 + i
Next i
Do
WDSummerDS = Weekday(arrSummerDS(i - 1), 2)
i = i - 1
Loop Until WDSummerDS = 7
arrSumWin(1) = CDbl(arrSummerDS(i))

'DealEndDate
arrSummerFDE = DateSerial(YEDeal, 10, 1)
For i = 1 To 31
arrSummerDE(i) = arrSummerFDE - 1 + i
Next i
Do
WDSummerDE = Weekday(arrSummerDE(i - 1), 2)
i = i - 1
Loop Until WDSummerDE = 7
arrSumWin(2) = CDbl(arrSummerDE(i))

If Month(StartDate) = 10 And Month(EndDate) = 10 Then
MySummertimeAdjustment = arrSumWin(2)
End If

'Your Winter time is on March. Loop through March Dates to find the Last Sunday of the Month
Dim arrWinterFDS As Date, arrWinterDS(1 To 31) As Date, WDWinterDS As Long
Dim arrWinterFDE As Date, arrWinterDE(1 To 31) As Date, WDWinterDE As Long

'DealStartDate
arrWinterFDS = DateSerial(YSDeal, 3, 1)
For i = 1 To 31
arrWinterDS(i) = arrWinterFDS - 1 + i
Next i
Do
WDWinterDS = Weekday(arrWinterDS(i - 1), 2)
i = i - 1
Loop Until WDWinterDS = 7
arrWinSum(1) = CDbl(arrWinterDS(i))

'DealEndDate
arrWinterFDE = DateSerial(YEDeal, 3, 1)
For i = 1 To 31
arrWinterDE(i) = arrWinterFDE - 1 + i
Next i
Do
WDWinterDE = Weekday(arrWinterDE(i - 1), 2)
i = i - 1
Loop Until WDWinterDE = 7
arrWinSum(2) = CDbl(arrWinterDE(i))

If Month(StartDate) = 3 And Month(EndDate) = 3 Then
MySummertimeAdjustment = arrWinSum(2)
End If

'Clean Up
Erase arrWinSum: Erase arrSumWin: Erase arrSummerDS: Erase arrSummerDE: Erase arrWinterDS: Erase arrWinterDE
Exit Function


MyErrorHandling:
MySummertimeAdjustment = 0
Exit Function

'Profile not recognised - provide error code
ErrorHandling1:
Resume MyErrorHandling:
Application.Volatile (True)
End Function

Sub DatesAnnualTry() 'from function below
Dim arrd() As Date

arrd = DatesAnnual(2017, "h")

ThisWorkbook.Worksheets("MyClients").Range("L1").Value = "DatesH"
ThisWorkbook.Worksheets("MyClients").Range("L2").Resize(UBound(arrd, 1), UBound(arrd, 2)).Value = arrd

ThisWorkbook.Worksheets("MyClients").Range("K1").Value = "DatesD"
ThisWorkbook.Worksheets("MyClients").Range("K2").Resize(UBound(arrd, 1), UBound(arrd, 2)).Value = arrd
End Sub

'Function returns an array of Dates of the year whether daily (ex. 365 values) or on hourly basis (ex. 365 *25)
Function DatesAnnual(MyYear As Integer, MyTypeDorH As String) As Date()
Dim arrDates() As Date
If MyYear Mod 4 = 0 Then
If MyTypeDorH = "d" Then
ReDim arrDates(366, 1)
ElseIf MyTypeDorH = "h" Then
ReDim arrDates(366 * 25, 1)
End If
Else
If MyTypeDorH = "d" Then
ReDim arrDates(365, 1)
ElseIf MyTypeDorH = "h" Then
ReDim arrDates(365 * 25, 1)
End If
End If

Dim i As Integer, dh As Integer, StartDate As Date

'd
dh = 1
StartDate = CDate(DateSerial(MyYear, 1, dh))
arrDates(1, 1) = StartDate

For i = 2 To UBound(arrDates)
If MyTypeDorH = "d" Then
dh = dh + 1
StartDate = StartDate + 1
End If

If Month(StartDate) <> Month(StartDate - 1) Then dh = 1

arrDates(i, 1) = CDate(DateSerial(MyYear, Month(StartDate), dh))

If MyTypeDorH = "h" And i Mod 25 = 0 Then
dh = dh + 1
StartDate = StartDate + 1
End If
Next i

DatesAnnual = arrDates
Erase arrDates()
End Function

'*******************************************
'FLOAT QUANTITIES
'*******************************************

'Cell B8 of HQties  =QuarterToHour(QQties!B8:CW373;HQties!$B$7:$Z$7)
Sub RunMyQuarterToHourFuncVBA() 'from function below
Dim MyarrQtoH() As Double

With ThisWorkbook
MyarrQtoH = QuarterToHour(Worksheets("QQties").Range("B8:CW373"), _
Worksheets("HQties").Range("$B$7:$Z$7"))

.Worksheets("HQties").Range("B8:Z373").Value = MyarrQtoH
End With

Erase MyarrQtoH()
End Sub

'Function converts quarterly to hourly SUM values
'Excel Formula
'Cell B8 of HQties =IF(ISERROR(SUM(OFFSET(QQties!$A8;0;((B$7-1)*4)+1;1;4)));"";SUM(OFFSET(QQties!$A8;0;((B$7-1)*4)+1;1;4)))
'VBA Formula
Function QuarterToHour(MyRange As Range, MyHours As Range) As Double()
On Error GoTo myend:

Dim rtransposed As Boolean, RRows As Long, RCols As Long
'''''rtransposed True'''''
rtransposed = True
RRows = MyRange.Columns.Count
RCols = MyRange.Rows.Count
'''''rtransposed True'''''

'''''rtransposed false'''''
If MyRange.Rows.Count > 1 And Int(MyRange.Rows.Count / 4) = MyHours.Cells.Count Then
rtransposed = False
If Int(RCols / 4) <> MyHours.Cells.Count Then Exit Function
End If
'''''rtransposed false'''''

'''''rtransposed True'''''
If rtransposed = True Then
If Int(RRows / 4) <> MyHours.Cells.Count Then Exit Function
End If
'''''rtransposed True'''''

'''''htransposed'''''
Dim htransposed As Boolean
htransposed = True
If MyHours.Rows.Count > 1 And MyHours.Rows.Count > MyHours.Columns.Count Then htransposed = False
Dim MySumRng() As Double 'Variant
If htransposed = True Then
ReDim MySumRng(RCols, MyHours.Columns.Count) 'transpose true RCols RRows
ElseIf htransposed = False And rtransposed = True Then
ReDim MySumRng(RCols, MyHours.Rows.Count)
Else
ReDim MySumRng(MyHours.Rows.Count, RRows)
End If
'''''htransposed'''''

Dim i As Long, j As Long, h As Long
h = 1
For i = 1 To RCols
For j = 1 To RRows

Select Case rtransposed
Case True
If IsNumeric(MyRange.Cells(i, j).Value) Then
MySumRng(i, h) = MySumRng(i, h) + MyRange.Cells(i, j).Value
Else
MySumRng(i, h) = MySumRng(i, h) + 0
End If
If j Mod 4 = 0 Then h = h + 1
If j = RRows Then h = 1
Case False
If IsNumeric(MyRange.Cells(i, j).Value) Then
MySumRng(h, j) = MySumRng(h, j) + MyRange.Cells(i, j).Value 'Transpose False
Else
MySumRng(h, j) = MySumRng(h, j) + 0
End If
End Select

Next j

If rtransposed = False Then
If i Mod 4 = 0 Then h = h + 1
If i = RCols Then h = 1
End If

Next i

QuarterToHour = MySumRng

myend:
Erase MySumRng()
Set MyHours = Nothing
Set MyRange = Nothing
End Function

Sub RunMyHourToQuarterFuncVBA() 'from function below
Dim MyarrHtoQ() As Double

With ThisWorkbook
MyarrHtoQ = HourToQuarter(Worksheets("HQties").Range("B8:Z373"), _
Worksheets("QQties").Range("$B$7:$CW$7"))

.Worksheets("QQties").Range("B8:CW373").Value = MyarrQtoH
End With

Erase MyarrHtoQ()
End Sub

'Function converts hourly to quarterly SUM values
Function HourToQuarter(MyRange As Range, MyHours As Range) As Double()
On Error GoTo myend:

Dim rtransposed As Boolean, RRows As Long, RCols As Long
'''''rtransposed True'''''
rtransposed = True
RRows = MyRange.Columns.Count
RCols = MyRange.Rows.Count
'''''rtransposed True'''''

'''''rtransposed false'''''
If MyRange.Rows.Count > 1 And Int(MyRange.Rows.Count * 4) = MyHours.Cells.Count Then
rtransposed = False
If Int(RCols * 4) <> MyHours.Cells.Count Then Exit Function
RCols = RCols * 4
End If
'''''rtransposed false'''''

'''''rtransposed True'''''
If rtransposed = True Then
If Int(RRows * 4) <> MyHours.Cells.Count Then Exit Function
RRows = RRows * 4
End If
'''''rtransposed True'''''

'''''htransposed'''''
Dim htransposed As Boolean
htransposed = True
If MyHours.Rows.Count > 1 And MyHours.Rows.Count > MyHours.Columns.Count Then htransposed = False
Dim MySumRng() As Double
If htransposed = True Then
ReDim MySumRng(RCols, MyHours.Columns.Count) 'transpose true RCols RRows
ElseIf htransposed = False And rtransposed = True Then
ReDim MySumRng(RCols, MyHours.Rows.Count)
Else
ReDim MySumRng(MyHours.Rows.Count, RRows)
End If
'''''htransposed'''''

Dim i As Long, j As Long, h As Long
h = 1
For i = 1 To RCols
For j = 1 To RRows

Select Case rtransposed
Case True
If IsNumeric(MyRange.Cells(i, Int((j - 1) / 4) + 1).Value) Then
MySumRng(i, h) = MySumRng(i, h) + MyRange.Cells(i, Int((j - 1) / 4) + 1).Value
Else
MySumRng(i, h) = MySumRng(i, h) + 0
End If
h = h + 1
If j = RRows Then h = 1
Case False
If IsNumeric(MyRange.Cells(Int((i - 1) / 4) + 1, j).Value) Then
MySumRng(h, j) = MySumRng(h, j) + MyRange.Cells(Int((i - 1) / 4) + 1, j).Value 'Transpose False
Else
MySumRng(h, j) = MySumRng(h, j) + 0
End If
End Select

Next j

If rtransposed = False Then
h = h + 1
If i = RCols Then h = 1
End If

Next i

HourToQuarter = MySumRng

myend:
Erase MySumRng()
Set MyHours = Nothing
Set MyRange = Nothing
End Function

'*************
'FLOAT PRICES
'*************

'Function retieves the Averages of selected array values
Function MyAverages(MyDataSample() As Variant, row As Long, col As Long) As Double()

Dim MyAveragesTmp() As Double
ReDim MyAveragesTmp(1 To 5)

Dim i As Integer
MyAveragesTmp(1) = 0 '"Bload"
MyAveragesTmp(2) = 0 '"Peak"
MyAveragesTmp(3) = 0 '"Offpeak"
MyAveragesTmp(4) = 0 '"OffP1"
MyAveragesTmp(5) = 0 '"OffP2"

Dim Hours As Long, PlusMinus As Integer
Hours = 24
PlusMinus = 0
If MyDataSample(row, col) = Range("Hour25").Value Then
Hours = 25
PlusMinus = 1
End If
If MyDataSample(row, col) = Range("Hour23").Value Then
Hours = 23
PlusMinus = -1
End If

For i = 1 To Hours
MyAveragesTmp(1) = MyAveragesTmp(1) + MyDataSample(row, col + i) '"Bload"

If i <= 8 + PlusMinus Then
MyAveragesTmp(3) = MyAveragesTmp(3) + MyDataSample(row, col + i) '"Offpeak"
MyAveragesTmp(4) = MyAveragesTmp(4) + MyDataSample(row, col + i) '"OffP1"
End If

If i >= 9 + PlusMinus And i <= 20 + PlusMinus Then
MyAveragesTmp(2) = MyAveragesTmp(2) + MyDataSample(row, col + i) '"Peak"
End If

If i >= 21 + PlusMinus Then
MyAveragesTmp(3) = MyAveragesTmp(3) + MyDataSample(row, col + i) '"Offpeak"
MyAveragesTmp(5) = MyAveragesTmp(5) + MyDataSample(row, col + i) '"OffP2"
End If

Next i

MyAveragesTmp(1) = MyAveragesTmp(1) / Hours          '"Bload"
MyAveragesTmp(2) = MyAveragesTmp(2) / 12             '"Peak"
MyAveragesTmp(3) = MyAveragesTmp(3) / 12 + PlusMinus '"Offpeak"
MyAveragesTmp(4) = MyAveragesTmp(4) / 8 + PlusMinus  '"OffP1"
MyAveragesTmp(5) = MyAveragesTmp(5) / 4              '"OffP2"


MyAverages = MyAveragesTmp
Erase MyAveragesTmp()
End Function


'*********************************************************************************************************
'XML
'*********************************************************************************************************

Sub XMLStringtoValTry()
Dim MystringVal As String
MystringVal = "12"
MsgBox XMLStringtoVal(MystringVal)
MystringVal = "12.01"
MsgBox XMLStringtoVal(MystringVal)
MystringVal = "34.12"
MsgBox XMLStringtoVal(MystringVal)
End Sub

Function XMLStringtoVal(mystr As String) As Double 'function which takes string and converts to number
If mystr = "" Then Exit Function

Dim dotpos As Integer, commapos As Integer
dotpos = InStr(mystr, ".")
commapos = InStr(mystr, ",")

If dotpos + commapos = 0 Then 'If no Decimals
XMLStringtoVal = CDbl(mystr)
Exit Function
End If

'MyNumber
Dim NumLeft As Double
NumLeft = Left(mystr, dotpos + commapos - 1)

'MyDecimals
Dim NumRight As Double
NumRight = Mid(mystr, dotpos + commapos + 1, Len(mystr) - dotpos - commapos)
NumRight = NumRight / (Application.WorksheetFunction.Power(10, Len(mystr) - dotpos - commapos))

XMLStringtoVal = NumLeft + NumRight
End Function


'Function retrieves values from an XML file. Works well for time-series xmls
'Parameteres Explained
'MyFileName = Complete filepath plus filename and extension
'parNode = Name of the Parent Node of your time series values (ex. "Real" on Qties)
'NumberofCategories = Number of columns array will be created (ex. 100 on Qties for quarterly values)
'FirstChildNodeIndex = parNode childindex where your looping for your values is commencing (Note: Index starts with 0. If insert 8 then is the ninth node)
'SecondChildNodeIndex = FirstChildNodeIndex childindex if there exists a SubNode (ex. on Qties not Prices) (Note: Index starts with 0. If insert 8 then is the ninth node)
'DestWS = Worksheet you want to paste your values (ex. "QQties")
'DestRng = Start Range you want your array to filter your values (ex. "B8")
Function RetValuesXML(MyFileName As String, parNode As String, NumberofCategories As Long, _
FirstChildNodeIndex As Long, SecondChildNodeIndex As Long, DestWS As String, DestRng As String) As Boolean

Dim xml As New DOMDocument, oXMLNode As IXMLDOMNode
Dim oXMLNodeList As IXMLDOMNodeList

Dim i As Long, q As Long
Dim MyResults() As Variant

'xml.async = False
Call xml.Load(MyFileName)

If xml.ChildNodes.Length = 0 Then
RetValuesXML = False
GoTo endfunction:
End If

Set oXMLNodeList = xml.DocumentElement.SelectNodes("//" & xml.ChildNodes.Item(1).ChildNodes.Item(0).BaseName & "/" & parNode)

ReDim MyResults(1 To oXMLNodeList.Length, 1 To NumberofCategories)

Dim tmpFirstChildNodeIndex As Long, tmpSecondChildNodeIndex As Long
tmpFirstChildNodeIndex = FirstChildNodeIndex
tmpSecondChildNodeIndex = SecondChildNodeIndex

For i = 1 To oXMLNodeList.Length
Set oXMLNode = oXMLNodeList.Item(i - 1)
FirstChildNodeIndex = tmpFirstChildNodeIndex: SecondChildNodeIndex = tmpSecondChildNodeIndex

    For q = 1 To NumberofCategories
    
        If FirstChildNodeIndex = oXMLNode.ChildNodes.Length Then Exit For
        
        If oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex) Is Nothing Then
        MyResults(i, q) = Null
        
        Else
        
        If oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text = vbNullString Then
            MyResults(i, q) = oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text
        Else
            If Application.DecimalSeparator = "," Then
            
            Select Case parNode
            Case "Fascie"
            MyResults(i, q) = "F" & XMLStringtoVal(oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text)
            'OR
            'MyResults(i, q) = "F" & CDbl(Replace(oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text, ".", ","))
            
            Case Else
            MyResults(i, q) = XMLStringtoVal(oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text)
            'OR
            'MyResults(i, q) = CDbl(Replace(oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text, ".", ","))
            End Select
            
            Else
            
            Select Case parNode
            Case "Fascie"
            MyResults(i, q) = "F" & CDbl(oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text)
            
            Case Else
            MyResults(i, q) = CDbl(oXMLNode.ChildNodes(FirstChildNodeIndex).ChildNodes(SecondChildNodeIndex).Text)
            End Select
            
            End If
        End If
        
        End If
        
        If SecondChildNodeIndex = 0 Then
        FirstChildNodeIndex = FirstChildNodeIndex + 1
        
        Else
        If q Mod 4 = 0 Then FirstChildNodeIndex = FirstChildNodeIndex + 1  'Quarterly Values
        SecondChildNodeIndex = SecondChildNodeIndex + 1
        If SecondChildNodeIndex = 5 Then SecondChildNodeIndex = 1
        End If
    
    Next q
    
    
Next i

With ThisWorkbook.Worksheets(DestWS)
.Range(DestRng).Resize((UBound(MyResults(), 1)), (UBound(MyResults(), 2))).Value = MyResults
RetValuesXML = True
End With

endfunction:
Set oXMLNodeList = Nothing: Set oXMLNode = Nothing: Set xml = Nothing
Erase MyResults
End Function
