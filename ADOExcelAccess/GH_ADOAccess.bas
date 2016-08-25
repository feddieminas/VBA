Attribute VB_Name = "GH_ADOAccess"
Option Explicit
Option Base 1

'Tried in Windows 7 and Excel 2007, 2010

'**************************************************************
'ADO Microsoft Access TABLE DELETE, INSERT, UPDATE, RETRIEVE
'**************************************************************

'1)Install Microsoft Access
'2)Need to check in your ODBC Data Source Administrator
'on User DSN tab that you have :
'User Data Sources Name = MS Access Database
'User Data Sources Driver = Microsoft Excel Driver(*mdb,*accdb)
'3)Add VBA Library reference : VBA Editor --> Tools --> References --> Microsoft XML, v3.0
'4)Add VBA Library reference :
'Need to add the MDAC Library named Microsoft ActiveX Data Objects 2.8 Library
'VBA Editor --> Tools --> References --> Microsoft ActiveX Data Objects 2.8 Library
'Alternatively you can find it VBA Editor --> Tools --> References --> Browse
'C:\Program Files (x86)\Common Files\System\ado\msado28.tlb

'Instructions
'Insert file extension on Line 25. Then run CreateAnXMLFile, ADcreation and ADDelInsUpd macro in order

Const dbExt = "accdb"  'can be accdb or mdb 'INSERT or keep it
                       'accdb is the updated file version of mdb since office pack 2007

Sub CreateAnXMLFile()
'A record XML Sample
'<?xml version="1.0" encoding="UTF-8" standalone="true"?>
'-<MyPriceCurve xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
'-<MyMktCurve>
'<DeliveryDate>20160302</DeliveryDate>
'<Hour>1</Hour>
'<Prc>0</Prc>
'<Qty>27</Qty>
'<Purpose>Buy</Purpose>
'</MyMktCurve>
'</MyPriceCurve>

Dim xmlDoc As DOMDocument, objIntro As IXMLDOMProcessingInstruction
Dim objRoot As IXMLDOMElement, objRecord As IXMLDOMElement, objNameF As IXMLDOMElement

Dim MyFolder As String
Dim NumberofMktCurves As Integer, NumberofCategories As Integer, r As Integer, c As Integer

MyFolder = ThisWorkbook.Path 'INSERT or keep it

'**************************
'Create your XML head

Set xmlDoc = CreateObject("Microsoft.XMLDOM")
Set objIntro = xmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8' standalone='yes'")
xmlDoc.InsertBefore objIntro, xmlDoc.ChildNodes(0)

Set objRoot = xmlDoc.createElement("MyPriceCurve"): xmlDoc.appendChild objRoot
With objRoot
    .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
End With

'***************************************************************
'Create my Header and array of Data to insert on the XML file

NumberofMktCurves = 50

Dim Headers(5) As String
Headers(1) = "DeliveryDate"
Headers(2) = "Hour"
Headers(3) = "Prc"
Headers(4) = "Qty"
Headers(5) = "Purpose"

NumberofCategories = UBound(Headers)

Dim MyDataSample() As Variant
ReDim MyDataSample(1 To NumberofMktCurves, 1 To NumberofCategories)

For r = 1 To NumberofMktCurves                              '25 Hours
MyDataSample(r, 1) = Format(Now, "YYYYMMDD")                'DeliveryDate
MyDataSample(r, 2) = r Mod 25                               'Hour

If r <= 24 Then
MyDataSample(r, 3) = CDbl((40 - 1.01 + 1) * Rnd + 1)           'Prc
MyDataSample(r, 4) = Int((10 - 1 + 1) * Rnd + 1)               'Qty
MyDataSample(r, 5) = "BID"                                     'Purpose
ElseIf r <= 49 Then
MyDataSample(r, 3) = CDbl((55 - 1.01 + 1) * Rnd + 1)           'Prc
MyDataSample(r, 4) = Int((15 - 1 + 1) * Rnd + 1)               'Qty
MyDataSample(r, 5) = "OFFER"                                   'Purpose
Else
MyDataSample(r, 3) = 0                                          'Prc Hour 25
MyDataSample(r, 4) = 0                                          'Qty Hour 25
End If
Next r
MyDataSample(25, 2) = 25                                        'Hour Hour 25
MyDataSample(25, 5) = "BID"                                     'Purpose Hour 25
MyDataSample(50, 2) = 25                                        'Hour Hour 25
MyDataSample(50, 5) = "OFFER"                                   'Purpose Hour 25


'**************************
'Create your XML body

For r = 1 To NumberofMktCurves
Set objRecord = xmlDoc.createElement("MyMktCurve"): objRoot.appendChild objRecord

    For c = 1 To NumberofCategories
    Set objNameF = xmlDoc.createElement(Headers(c)): objRecord.appendChild objNameF
    If CStr(MyDataSample(r, c)) <> "" Then
    If IsNumeric(MyDataSample(r, c)) And c > 1 Then
    
    If Application.DecimalSeparator = "," And Int(MyDataSample(r, c)) <> MyDataSample(r, c) Then
    MyDataSample(r, c) = Replace(Format(MyDataSample(r, c), "#0.00"), ",", ".")
    objNameF.Text = CStr(MyDataSample(r, c))
    ElseIf Int(MyDataSample(r, c)) <> MyDataSample(r, c) Then 'Case with Application.DecimalSeparator = "."
    objNameF.Text = CStr(Format(MyDataSample(r, c), "#0.00"))
    Else
    objNameF.Text = CStr(CDbl(MyDataSample(r, c)))
    End If
    
    Else
    objNameF.Text = CStr(MyDataSample(r, c))
    End If
    Else
    objNameF.Text = vbNullString
    End If
    Set objNameF = Nothing
    Next c

Set objRecord = Nothing
Next r

xmlDoc.Save MyFolder & "\" & "RandPriceCurve.xml" 'INSERT or keep it

Set xmlDoc = Nothing: Set objRoot = Nothing: Set objRecord = Nothing: Set objNameF = Nothing
Set objIntro = Nothing: Erase MyDataSample(): Erase Headers()
End Sub

Sub ADcreation()
Dim appAccess As Object
Dim dbfile As String
Dim xmlfile As String

Application.DisplayAlerts = False

Set appAccess = CreateObject("Access.Application")

'Create Access File
dbfile = ThisWorkbook.Path & "\" & "RandPriceCurve." & dbExt  'Access file mdb or accdb  'INSERT or keep it
xmlfile = ThisWorkbook.Path & "\" & "RandPriceCurve.xml"      'xml file to import in Access  'INSERT or keep it
If FolderFileExists(dbfile) Then Kill dbfile 'GoTo myend

'appAccess.Visible = True
appAccess.NewCurrentDatabase dbfile
appAccess.ImportXml DataSource:=xmlfile
appAccess.CloseCurrentDatabase

'Quitting Microsoft Access Application
'Option 1
appAccess.DoCmd.Quit
'Option 2
'appAccess.Application.Quit 'Those two commands in comments are syntatically correct but were hanging the Access Application
'appAccess.Application.Quit acQuitSaveNone

'myend:

Set appAccess = Nothing

Application.DisplayAlerts = True

End Sub

Sub ADDelInsUpd()

Dim dbfile As String, dbtable As String
Dim xmlfile As String
Dim fldCol As ADODB.Field 'for scanning through update

dbfile = ThisWorkbook.Path & "\" & "RandPriceCurve." & dbExt 'INSERT or keep it
xmlfile = ThisWorkbook.Path & "\" & "RandPriceCurve.xml"     'INSERT or keep it

dbtable = MyTableName(xmlfile)

Dim sConnect As String
Dim sSQL As String
Dim rsData As ADODB.Recordset

If Val(Application.Version) >= 12 Then
sConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
"Data Source=" & dbfile & ";"
Else
sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & dbfile & ";"
End If

Set rsData = New ADODB.Recordset
rsData.CursorLocation = adUseServer

If Val(Application.Version) >= 12 Then
rsData.Open "[" & dbtable & "]", sConnect, adOpenDynamic, adLockOptimistic, adCmdTable
Else
rsData.Open "[" & dbtable & "]", sConnect, adOpenForwardOnly, adLockOptimistic, adCmdTable
End If

'********************
'DELETE
'********************

'Delete existing records in the Database     'Delete Hour 25 on OFFER
If rsData.Supports(adDelete) Then

 rsData.MoveFirst
 Do Until rsData.EOF
 If rsData.Fields("Hour").Value = 25 And rsData.Fields("Purpose").Value = "OFFER" Then
    rsData.Delete
    rsData.Update
 End If
    rsData.MoveNext
 Loop
 
End If

'********************
'INSERT
'********************

'Insert records to the Database
If rsData.Supports(adAddNew) Then           'Insert Hour 25 on Offer

 rsData.AddNew
 rsData.Fields("DeliveryDate").Value = Format(Now, "YYYYMMDD")
 rsData.Fields("Hour").Value = 25
 On Error Resume Next
 If Application.DecimalSeparator = "," Then
 rsData.Fields("Prc").Value = Replace(Format(CDbl(19.08), "##0.00"), ",", ".")
 Else
 rsData.Fields("Prc").Value = Format(CDbl(19.08), "##0.00")
 End If
 On Error GoTo 0
 rsData.Fields("Qty").Value = 100
 rsData.Fields("Purpose").Value = "OFFER"
 rsData.Update

End If

'********************
'UPDATE
'********************

'Update existing records on Database        'Update Hour 25 Prc and Qty to zero on OFFER
If rsData.Supports(adUpdate) Then

 rsData.MoveFirst 'can be in comments
 While Not rsData.EOF

 For Each fldCol In rsData.Fields
 If fldCol.Name = "Hour" Then               'Show RETRIEVE Column Name              (i.e. fldCol.Name)
 If fldCol.Value = 25 Then                  'Show RETRIEVE Hour Column Row Value    (i.e. fldCol.Value)
 
  If rsData("Purpose").Value = "OFFER" Then 'Show RETRIEVE Purpose Column Row Value (i.e. rsData("Purpose").Value)
  rsData("Prc").Value = 0
  rsData("Qty").Value = 0
  rsData.Update
  End If
  
  If rsData("Purpose").Value = "BID" Then 'Show RETRIEVE Purpose Column Row Value (i.e. rsData("Purpose").Value)
  rsData("Prc").Value = 0                 'Set also Hour25 BID originally created to zero values
  rsData("Qty").Value = 0
  rsData.Update
  End If
 
 End If
 End If
 Next
 rsData.MoveNext
 
 Wend

End If

'********************

'Reset the Recordset Object
If rsData.State <> adStateClosed Then rsData.Close
If Not rsData Is Nothing Then Set rsData = Nothing

End Sub

Private Function FolderFileExists(FFName As String) As Boolean
Dim FFNameThere As String
FFNameThere = Dir(FFName, vbDirectory)
If FFNameThere = "" Then FFNameThere = Dir(FFName)
If FFNameThere = "" Then
FolderFileExists = False
Else
FolderFileExists = True
End If
End Function

Private Function MyTableName(file As String) As String
Dim xml As New DOMDocument
xml.async = False
Call xml.Load(file)
If xml.ChildNodes.Length = 0 Then GoTo myend
MyTableName = xml.ChildNodes.Item(1).ChildNodes.Item(0).nodeName
myend:
Set xml = Nothing
End Function
