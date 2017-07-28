Attribute VB_Name = "modCmdBar_1_Code"

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit
Option Base 1
Option Compare Text

Sub SetUpMenus()
Dim cbrAllBars As CommandBars
Dim menuItem0 As CommandBarPopup, menuItem1 As CommandBarPopup, menuItem2 As CommandBarPopup, menuItem3 As CommandBarPopup, menuItem4 As CommandBarPopup
Dim myitem As CommandBarButton, currentmenuItem As CommandBarPopup
Dim myBar As CommandBar
Dim i As Long
Dim vTable() As Variant

On Error Resume Next

Set cbrAllBars = Application.CommandBars

vTable() = ThisWorkbook.Worksheets("Sheet1").Range("nCompShortCut").Value

For i = 1 To UBound(vTable(), 1)

    If Len(CStr(Trim(vTable(i, 1)))) > 0 And vTable(i, 2) = "msoBarPopup" Then
    Set myBar = cbrAllBars.Item(CStr(vTable(i, 1))) 'does it exist already if it does set it - otherwise create an error and carry on
    myBar.Delete
    Set myBar = cbrAllBars.Add(vTable(i, 1), Position:=msoBarPopup, menuBar:=False, Temporary:=True)
    End If


Select Case vTable(i, 2)

    Case Is = "msoControlPopup"
    
        Select Case vTable(i, 5)
    
        Case Is = 0
        Set menuItem0 = myBar.Controls.Add(Type:=msoControlPopup)
        menuItem0.Caption = vTable(i, 3)
        Set currentmenuItem = menuItem0
        
        Case Is = 1
        Set menuItem1 = menuItem0.Controls.Add(Type:=msoControlPopup)
        menuItem1.Caption = vTable(i, 3)
        If menuItem1.Caption = "Prices" Then menuItem1.BeginGroup = True
        Set currentmenuItem = menuItem1
        
        Case Is = 2
        Set menuItem2 = menuItem1.Controls.Add(Type:=msoControlPopup)
        menuItem2.Caption = vTable(i, 3)
        If menuItem2.Caption = "Prices" Then menuItem2.BeginGroup = True
        Set currentmenuItem = menuItem2
        
        Case Is = 3
        Set menuItem3 = menuItem2.Controls.Add(Type:=msoControlPopup)
        menuItem3.Caption = vTable(i, 3)
        Set currentmenuItem = menuItem3
        
        Case Is = 4
        Set menuItem4 = menuItem3.Controls.Add(Type:=msoControlPopup)
        menuItem4.Caption = vTable(i, 3)
        Set currentmenuItem = menuItem4
        
        End Select

    Case Is = "msoControlButton"

        myCreateButtons myCurrentMenuItem:=currentmenuItem, MyButtonName:=CStr(vTable(i, 3)), MySubName:=CStr(vTable(i, 4))
    
End Select


Next i
    Erase vTable()
    
Set myitem = Nothing:  Set myBar = Nothing
Set menuItem0 = Nothing: Set menuItem1 = Nothing: Set menuItem2 = Nothing: Set menuItem3 = Nothing: Set menuItem4 = Nothing
Set currentmenuItem = Nothing

       
End Sub
Sub DeletePopUpMenu()
    Dim vTable() As Variant
    
    Dim i As Long
    'Delete PopUp menu if it exist
    On Error Resume Next
    vTable() = ThisWorkbook.Worksheets("Sheet1").Range("nCompShortCut")
    For i = 1 To UBound(vTable(), 1)
    If Len(CStr(Trim(vTable(i, 1)))) > 0 And vTable(i, 2) = "msoBarPopup" Then
    Application.CommandBars(CStr(vTable(i, 1))).Delete
    
    End If
    Next i
    
    Erase vTable()
    
    On Error GoTo 0
End Sub

Sub DisplayPopUpMenu()
Dim Mname As String

    On Error Resume Next
    Mname = "CompShortCut"
    Application.CommandBars(Mname).ShowPopup
  '  Cancel = True

    On Error GoTo 0
End Sub

