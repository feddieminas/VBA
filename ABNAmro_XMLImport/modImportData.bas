Attribute VB_Name = "modImportData"
Option Explicit
Option Base 1
Option Compare Text

'SubMacro can do
'Delete Sheets
'insertSheets
'NodiY
'Attributes
'NodiX Headers
'NodiX values

'Alternative way to parse XML is using SAX
'http://sax.sourceforge.net/event.html          Explanation of the concept
'http://stackoverflow.com/questions/5626653/how-can-i-improve-the-speed-of-xml-parsing-in-vba

Sub ABNFilesLoop()
Dim Filepath As String
Dim MyFileFinder() As String, MyFileSuffix As String
Dim i As Integer

Application.ScreenUpdating = False

For i = 1 To ThisWorkbook.Worksheets("Dashboard").Range("C2").Value

Filepath = Left(ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i).Value, 8)
Filepath = ThisWorkbook.Worksheets("Dashboard").Range("E15").Value & _
IIf(Right(ThisWorkbook.Worksheets("Dashboard").Range("E15").Value, 1) = "\", "", "\") & "Unzipped\" & Filepath & "\" & _
ThisWorkbook.Worksheets("Dashboard").Range("C" & 2 + i).Value
MyFileFinder() = Split(Filepath, "-")
MyFileFinder() = Split(MyFileFinder(3))
MyFileSuffix = MyFileFinder(0)              'Codify your Sheet Purpose
'ThisWorkbook.Worksheets("Dashboard").Range("D" & 2 + i).Value = MyFileSuffix

XMLDOMDynamic Filepath, MyFileSuffix
Next i

Erase MyFileFinder

Application.ScreenUpdating = True

End Sub

Sub XMLDOMDynamic(Filepath As String, MyFileSuffix As String)

'Dim Filepath As String
'Dim MyFileFinder() As String, MyFileSuffix As String
Dim WB As Workbook
Dim WS() As String
Dim S As Long
Dim XMLLoad As Boolean
Dim Obj As DOMDocument
Dim Nodi As IXMLDOMNodeList
Dim Nodo As IXMLDOMNode
Dim NodiYL As Long, NodiXL As Long, NodiZL As Long, NodiFL As Long  'L as Lengths, Element Counter NodiY=Rows, NodiX=Columns
Dim NodiYBN As String, NodiZBN As String, Noditmp As String
Dim NodiYC As Long, NodiXC As Long, NodiZC As Long, NodiFC As Long   'C as a Node Counter
Dim NodiYList() As String, NodiZList() As String
Dim StartRow() As Long, StartRowC As Long
Dim y As Long, x As Long, z As Long, f As Long, n As Long, ExtraCol As Long
Dim NodiHasChilds As Boolean

'Step 1 FilePath and Suffix Name of Sheet

'Step 2 a
Set WB = ThisWorkbook
If SheetExists(MyFileSuffix & 1, ThisWorkbook) Then   'Create at least your Min 1 Sheet for a MyFileSuffix
Else
With WB
.Sheets.Add After:=.Sheets(.Sheets.Count)
S = .Sheets.Count
.Worksheets(S).Name = MyFileSuffix & 1
End With
End If

'Step 2 b
For x = 2 To 10                                        'If other existing sheets of MyFileSuffix Delete as default
If SheetExists(MyFileSuffix & x, ThisWorkbook) Then    'Assume you have 10 different MyFileSuffix
Application.DisplayAlerts = False
ThisWorkbook.Worksheets(MyFileSuffix & x).Delete
End If
Next x
Application.DisplayAlerts = True


'Step 3
Set Obj = New DOMDocument
        Obj.async = False
        XMLLoad = Obj.Load(Filepath)                   'Load MyXML object

If XMLLoad = False Then GoTo fine                      'If an XML file correctly structured exists, proceed


'Step 4
NodiYC = 0  'NODIY  All my Topic elements. I will then need for each one to loop and find myrows

'Step 4 a
'MainPurpose
NodiYL = Obj.ChildNodes.Length
For y = 1 To NodiYL
NodiYBN = Obj.ChildNodes.Item(y - 1).BaseName

If NodiYBN <> "xml" Then                                'Not take into account the very first line of code
NodiYC = NodiYC + 1
ReDim Preserve NodiYList(1 To NodiYC)
NodiYList(NodiYC) = NodiYBN
End If

Next y

'Step 4 b
y = NodiYC                                              'SubPosition Accounts of MainPurpose
For y = y To UBound(NodiYList)
Set Nodi = Obj.DocumentElement.SelectNodes("/" & NodiYList(y))
NodiZL = Nodi.Item(0).ChildNodes.Length
For z = 1 To NodiZL
NodiZBN = Nodi.Item(0).ChildNodes.Item(z - 1).BaseName

If NodiZBN <> Noditmp Then                              'We want distinct values of NodiY SubPosition Accounts
NodiYC = NodiYC + 1
ReDim Preserve NodiYList(1 To NodiYC)
NodiYList(NodiYC) = NodiZBN
End If

Noditmp = NodiZBN
Next z
Next y
Set Nodi = Nothing


'Step 4 c
'Evaluations of SubPosition Accounts
y = NodiYC
On Error GoTo fine                                      'If no Contents exist within the XML, then End Macro

For y = y To UBound(NodiYList)
Set Nodi = Obj.DocumentElement.SelectNodes("/" & NodiYList(y - 1) & "/" & NodiYList(y))
NodiZL = Nodi.Item(0).ChildNodes.Length                  'ASSUME the very first Item of each Selected SubPosition Account Node
For z = 1 To NodiZL
NodiZBN = Nodi.Item(0).ChildNodes.Item(z - 1).BaseName

If NodiZBN <> Noditmp Then                               'We want distinct values of NodiY
NodiYC = NodiYC + 1
ReDim Preserve NodiYList(1 To NodiYC)
NodiYList(NodiYC) = NodiZBN
End If

Noditmp = NodiZBN
Next z
Next y
Set Nodi = Nothing

'Step 4 d
NodiYC = 0                                               'Summary of Node Y. My Final List
For y = 1 To UBound(NodiYList)                           'VERIFY. We want distinct values of NodiYList
Noditmp = NodiYList(y)
For z = y To UBound(NodiYList) - 1
If NodiYList(z + 1) = Noditmp Then
NodiYC = NodiYC + 1
NodiYList(z + 1) = "NULL"
End If
Next z
Next y
NodiYC = NodiYC / 2                                     'it will be at least multiply by 2 a single combination
ReDim Preserve NodiYList(1 To UBound(NodiYList) - NodiYC + 1)
NodiYList(UBound(NodiYList)) = "NULL"                   'Finish NodiYList with a NULL character as your end

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Step 5
'NODIX Excel Setups

'Step 5 a                                   'Attributes of each Node to see how many Sheets do I need
StartRowC = 0: y = 1
For y = 1 To UBound(NodiYList)
If NodiYList(y) <> "NULL" Then              'NULL is my final value array and in any case any duplicate NodiY
Set Nodi = Obj.DocumentElement.SelectNodes("//" & NodiYList(y))   '// AllElements of Selected Node

NodiZL = Nodi.Item(0).Attributes.Length     'ASSUMPTION that if there are attributes, that would be only Item 0 as default
If NodiZL > 0 Then
NodiXL = Nodi.Length

'Step 5 b
StartRowC = StartRowC + NodiXL              'MyTotalSheets of MainPurpose
End If
Set Nodi = Nothing
End If
Next y

'Step 5 c
If StartRowC = 0 Then GoTo fine             'Equivalent to MyTotalSheets, MyStartRow array
ReDim StartRow(1 To StartRowC)
For S = 1 To StartRowC
StartRow(S) = 0
Next S


'Step 5 d                                    MyExtraSheets to 1 ???
For y = 1 To StartRowC                      'Could also write it as UBound(StartRowC)
If SheetExists(MyFileSuffix & y, ThisWorkbook) Then
Else
With WB
.Sheets.Add After:=.Sheets(.Sheets.Count)
S = .Sheets.Count
.Worksheets(S).Name = MyFileSuffix & y
End With
End If
Next y
'Step 5 e
ReDim WS(1 To StartRowC)                    'MyTotalSheets array
For S = 1 To StartRowC
WS(S) = MyFileSuffix & S
WB.Worksheets(WS(S)).Cells.Clear
Next S


'Step 6
'NODIX MyRows from NodiY

y = 1
For y = 1 To UBound(NodiYList)
If NodiYList(y) <> "NULL" Then  'NULL is my final value array and in any case any duplicate NodiY

'Step 6 a
Set Nodi = Obj.DocumentElement.SelectNodes("//" & NodiYList(y))    '// AllElements of Selected Node
For S = 1 To StartRowC                                             'Loop for Worksheets and StartRows

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Step 6 b                                                           My NodiY StartRow
With WB.Worksheets(WS(S))
StartRow(S) = .Cells(.Rows.Count, "A").End(xlUp).Row
End With
If StartRow(S) = 1 And WB.Worksheets(WS(S)).Cells(StartRow(S), 1).Value = "" Then
WB.Worksheets(WS(S)).Cells(StartRow(S), 1).Value = NodiYList(y)    'My NodiY First Title
Else
StartRow(S) = StartRow(S) + 3
WB.Worksheets(WS(S)).Cells(StartRow(S), 1).Value = NodiYList(y)    'My NodiY Other Titles
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Step 6 c                                 See if Attributes to fill. If there are Attributes, these are your portfolio accounts
NodiZL = Nodi.Item(0).Attributes.Length  'ASSUMPTION that if there are attributes, that would be only Item 0 as default
If NodiZL > 0 Then

z = 0
StartRow(S) = StartRow(S) + 2
NodiZL = Nodi.Item(S - 1).Attributes.Length   'Item Attributes
For z = 1 To NodiZL                           'Fill your Attributes of SubPosition Accounts of MainPurpose
WB.Worksheets(WS(S)).Cells(StartRow(S) + z - 1, 1).Value = Nodi.Item(S - 1).Attributes(z - 1).BaseName   'My NodiX Title
WB.Worksheets(WS(S)).Cells(StartRow(S) + z - 1, 2).Value = Nodi.Item(S - 1).Attributes(z - 1).Text       'MyNodiX Text
Next z
StartRow(S) = StartRow(S) + NodiZL + 2        'Amend MyStartRow after filling its Attributes

                                              'Evaluations of SubPosition Accounts Counter per Sheet.
'Step 6 d                                     'How many Items with no Attributes (Third Node theoretically) afterwards will we loop
z = 1
For x = y + 1 To UBound(NodiYList) - 1                     'No Need the last array NULL or -1
WB.Worksheets(WS(S)).Cells(z, 27).Value = NodiYList(x)     'Name of Evaluations of SubPosition Accounts Nodes
z = z + 1
Next x
NodiXL = Nodi.Item(S - 1).ChildNodes.Length
NodiFL = z - 1
'Step 6 e
For f = 1 To NodiFL                                        'NodiY Names of Evaluations of SubPosition Accounts Nodes
WB.Worksheets(WS(S)).Cells(f, 28).Value = 0
For x = 1 To NodiXL                                        'NodeY Counters.How many Items per Selected Node
If WB.Worksheets(WS(S)).Cells(f, 27).Value = Nodi.Item(S - 1).ChildNodes.Item(x - 1).BaseName Then
WB.Worksheets(WS(S)).Cells(f, 28).Value = WB.Worksheets(WS(S)).Cells(f, 28).Value + 1
End If
Next x
Next f

Else
StartRow(S) = StartRow(S) + 3                              'Amend MyStartRow if no Attributes
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Step 7                                              'Very First NodiY element or Attributes and NextBasename if same as Next array then no Values. U then goto Next Loop
NodiYBN = Nodi.Item(0).ChildNodes.Item(0).BaseName   'CrossCheck that NextValueArray has a different value
If (y = 1 Or NodiZL > 0) And NodiYBN = NodiYList(y + 1) Then
GoTo MyNextNode

Else

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Step 8
'NODIX MyHeaderRows

'Step 8 a
If WB.Worksheets(WS(S)).Cells(y - 2, 28).Value = 0 Then    'Check if NodeY Counters exist for Selected Element for the Worksheet S. if Not, no need to loop
WB.Worksheets(WS(S)).Cells(StartRow(S) - 3, 1).Value = ""
GoTo MyNextNode
End If

'Step 8 b                                                  'HEADERS FIRST
If S = 1 Then                                              'If first SubPosition Account
x = WB.Worksheets(WS(S)).Cells(y - 2, 28).Value - 1        'Use -1 because first Item is consider as Item(0) lower bound
Else
For z = 2 To S                                             'Second SubPosition Account and after
x = WB.Worksheets(WS(z - 1)).Cells(y - 2, 28).Value
x = x + WB.Worksheets(WS(z)).Cells(y - 2, 28).Value - 1    'Use -1 because first Item is consider as Item(0) lower bound
Next z
End If
If x = -1 Then x = 0

'Step 8 c
NodiZL = Nodi.Item(x).ChildNodes.Length
If x > 0 Then                             'CrossCheck is the max ChildNodes length if more than 1 Item. If not Amend
For z = 0 To x - 1
If Nodi.Item(z).ChildNodes.Length > NodiZL Then
NodiZL = Nodi.Item(z).ChildNodes.Length   'Amend Selected item length
x = z                                     'Amend Selected item
End If
Next z
End If
ReDim NodiZList(1 To NodiZL)

'Step 8 d1
NodiFC = 0                                    'HEADERS Loop
For z = 1 To NodiZL
NodiZList(z) = Nodi.Item(x).ChildNodes.Item(z - 1).BaseName
WB.Worksheets(WS(S)).Cells(StartRow(S), z + NodiFC).Value = Nodi.Item(x).ChildNodes.Item(z - 1).BaseName                                    'NodiZList(z)

NodiFL = Nodi.Item(x).ChildNodes.Item(z - 1).ChildNodes.Length
If NodiFL = 0 Then GoTo MyNextColH
NodiHasChilds = Nodi.Item(x).ChildNodes.Item(z - 1).ChildNodes.Item(0).HasChildNodes

'Step 8 d2                                     'If HEADERS have another SubHeaders
If NodiFL > 1 And NodiHasChilds = True Then    'ASSUMPTION that only SubHeaders exist, no any further Subs
For f = 1 To NodiFL                            'When we use 1, we should loop from f-1 to retrieve first Item
WB.Worksheets(WS(S)).Cells(StartRow(S) - 1, z + NodiFC + f - 1).Value = NodiZList(z)   'HEADERS ReLoop
WB.Worksheets(WS(S)).Cells(StartRow(S), z + NodiFC + f - 1).Value = _
Nodi.Item(x).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).BaseName                    'SUBHEADERS Loop
Next f
NodiFC = NodiFC + NodiFL - 1 'For the Column Offset of Further Loop
End If

MyNextColH:
Next z

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Step 9
'NODIX FILL Row Values

'Step 9 a  Start and EndLoop of Looping Items
If WB.Worksheets(WS(S)).Cells(y - 2, 28).Value > 0 Then
'1
If S = 1 Then                                           'If first SubPosition Account
NodiXL = WB.Worksheets(WS(S)).Cells(y - 2, 28).Value    'MyTotalRows of first SubPosition Account
NodiYC = 1                                              'StartLoop Items
Else
'2
For z = 2 To S                                          'Second SubPosition Account and after
NodiXL = WB.Worksheets(WS(z - 1)).Cells(y - 2, 28).Value
NodiXL = NodiXL + WB.Worksheets(WS(z)).Cells(y - 2, 28).Value
NodiYC = NodiXL - WB.Worksheets(WS(z)).Cells(y - 2, 28).Value + 1    '+1 because then we loop Item(x) with Item(x - 1)
Next z
End If
Else
'3
NodiXL = Nodi.Length                                    'If 1 SubPosition Account
NodiYC = 1
End If

'Step 9 b                                               'Fill ROWValues Loop
NodiXC = 0
NodiZL = 0
For x = NodiYC To NodiXL                                'X Loop per each Item of SubPosition Account
NodiXC = NodiXC + 1
NodiFC = 0

NodiZL = Nodi.Item(x - 1).ChildNodes.Length             'In case there are less columns or need to insert an extra column
For z = 1 To NodiZL                                     'Z Loop per each SubItem of SubPosition Account
ExtraCol = 0

'Step 9 b1                           'CrossCheck Same Header either on StartRow or SubStartRow to fill value
If WB.Worksheets(WS(S)).Cells(StartRow(S) - 1, z + NodiFC).Value = _
Nodi.Item(x - 1).ChildNodes.Item(z - 1).BaseName Or _
WB.Worksheets(WS(S)).Cells(StartRow(S), z + NodiFC).Value = _
Nodi.Item(x - 1).ChildNodes.Item(z - 1).BaseName Then
WB.Worksheets(WS(S)).Cells(StartRow(S) + NodiXC, z + NodiFC).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).Text

Else
With WB.Worksheets(WS(S))
ExtraCol = .Cells(StartRow(S), .Columns.Count).End(xlToLeft).Column + 1

'Step 9 b2                           'If not before having same Header on that Loop, loop thru each Item to find whether there is same Header on another Row
For n = 1 To ExtraCol
If .Cells(StartRow(S) - 1, n).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).BaseName _
Or .Cells(StartRow(S), n).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).BaseName Then
.Cells(StartRow(S) + NodiXC, n).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).Text
Exit For
End If
Next n

'Step 9 b3                           'if not matched Header then we need to add an extra column, a New Header
If n = ExtraCol + 1 Then
.Cells(StartRow(S), ExtraCol).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).BaseName    'NodiXName ExtraCol
.Cells(StartRow(S) + NodiXC, ExtraCol).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).Text 'NodiXText
End If
End With

End If

'Step 9 c                            'If there are other childs with numbers to fill in
NodiFL = Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Length
If NodiFL = 0 Then GoTo MyNextColV
NodiHasChilds = Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(0).HasChildNodes
If NodiHasChilds = True Then                             'If ROWValues have another SubROWValues

For f = 1 To NodiFL                                      'ASSUMPTION that only SubROWValues exist, no any further Subs
'Step 9 c1
If ExtraCol = 0 Then

If WB.Worksheets(WS(S)).Cells(StartRow(S), z + NodiFC + f - 1).Value = _
Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).BaseName Then
WB.Worksheets(WS(S)).Cells(StartRow(S) + NodiXC, z + NodiFC + f - 1).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).Text
Else
For n = z - 1 To NodiZL
If WB.Worksheets(WS(S)).Cells(StartRow(S), n).Value = _
Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).BaseName Then
WB.Worksheets(WS(S)).Cells(StartRow(S) + NodiXC, n).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).Text
End If
Next n
End If

'Step 9 c2
Else
If n = ExtraCol + 1 Then
WB.Worksheets(WS(S)).Cells(StartRow(S) - 1, ExtraCol - 2 + NodiFC + f - 1).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).BaseName
WB.Worksheets(WS(S)).Cells(StartRow(S), ExtraCol - 2 + NodiFC + f - 1).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).BaseName
WB.Worksheets(WS(S)).Cells(StartRow(S) + NodiXC, ExtraCol - 2 + NodiFC + f - 1).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).Text
Else
WB.Worksheets(WS(S)).Cells(StartRow(S) + NodiXC, n + f - 1).Value = Nodi.Item(x - 1).ChildNodes.Item(z - 1).ChildNodes.Item(f - 1).Text            'n + NodiFC + f - 1
End If
End If
Next f

NodiFC = NodiFC + NodiFL - 1
End If


MyNextColV:
Next z

Next x


MyNextNode:
End If  'NodiYL = 1 And NodiYBN = NodiYList(y + 1)

'Step 10 AutoColimn Fit
WB.Worksheets(WS(S)).Columns("A:AD").AutoFit             'Autofit of Columns
Next S
End If  'NodiYList(y)
Set Nodi = Nothing
Next y

fine:
'Step 11 Erase Objects
Erase NodiYList: Erase NodiZList
Erase WS: Erase StartRow
Set Obj = Nothing
Set WB = Nothing

End Sub

'Functions
Public Function SheetExists(SheetName As String, MyWorkbook As Workbook) As Boolean
    Dim WS As Worksheet
    SheetExists = False
    For Each WS In MyWorkbook.Worksheets
        If WS.Name = SheetName Then SheetExists = True
    Next WS
End Function
