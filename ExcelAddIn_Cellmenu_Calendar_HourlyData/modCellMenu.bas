Attribute VB_Name = "modCellMenu"

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit

Sub AddToCellMenu()
Dim ContextMenu As CommandBar
Dim MySubMenu As CommandBarControl
Dim MyRefreshSubMenu1 As CommandBarPopup, MyRefreshSubMenu2 As CommandBarPopup, MyRefreshSubMenu As CommandBarPopup
'Dim MyRedundancySubMenu As CommandBarPopup
'Delete the controls first to avoid duplicates
Call DeleteFromCellMenu

'Set ContextMenu to the Cell menu
Set ContextMenu = Application.CommandBars("Cell")

'Add one custom button to the Cell menu
With ContextMenu.Controls.Add(Type:=msoControlButton, before:=1)
    .OnAction = "'" & ThisWorkbook.Name & "'!" & "DisplayPopUpMenu"
    .FaceId = 925 '109
    .Caption = "Comp Retrieve MK Data"
    .Tag = "My_Cell_Control_Tag"
End With

''copytags
With ContextMenu.Controls.Add(Type:=msoControlButton, before:=2)
    .OnAction = "'" & ThisWorkbook.Name & "'!" & "cbCopyPasteTags"
    .FaceId = 19
    .Caption = "Comp Tags Copy and Paste"
    .Tag = "My_Cell_Control_Tag"
End With
''copytags


'Add custom menu with three buttons
Set MyRefreshSubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)
        
With MyRefreshSubMenu
    .Caption = "Comp Refresh Menu"
    .Tag = "My_Cell_Control_Tag"
      
            Set MyRefreshSubMenu1 = .Controls.Add(Type:=msoControlPopup)
              
                With MyRefreshSubMenu1
                    .Caption = "Refresh Absolute"
                    .Tag = "My_Cell_Control_Tag"
                End With
          
            Set MyRefreshSubMenu2 = .Controls.Add(Type:=msoControlPopup)
         
                With MyRefreshSubMenu2
                    .Caption = "Refresh Relative"
                    .Tag = "My_Cell_Control_Tag"
                End With
End With


With MyRefreshSubMenu1

   With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "cbAbsDiscoverRefreshRangeTags"
            .FaceId = 457
            .Caption = "Refresh Selection"
            .Enabled = True
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "cbAbsDiscoverRefreshWSTags"
            .FaceId = 459
            .Caption = "Refresh Worksheet"
            .Enabled = True
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "cbAbsDiscoverRefreshWBTags"
            .FaceId = 1952
            .Caption = "Refresh Workbook"
            .Enabled = True
        End With

End With

With MyRefreshSubMenu2

   With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRelDiscoverRefreshRangeTags"
            .FaceId = 457
            .Caption = "Refresh Selection"
            .Enabled = True
            
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRelDiscoverRefreshWSTags"
            .FaceId = 459
            .Caption = "Refresh Worksheet"
            .Enabled = True
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRelDiscoverRefreshWBTags"
            .FaceId = 1952
            .Caption = "Refresh Workbook"
            .Enabled = True
        End With
End With
    
'''Redundancy menu
'Set MyRedundancySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=4)
        
'With MyRedundancySubMenu
    '.Caption = "Comp Redundancy Menu"
    '.Tag = "My_Cell_Control_Tag"
'End With
    
'With MyRedundancySubMenu 'OnAction Procedures not shown it on a .bas file. One can view it through the xlsm file.

        'With .Controls.Add(Type:=msoControlButton)
            '.OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRunMKCEEDownload"
            '.FaceId = 688
            '.Caption = "Run CEE Download"
            '.Enabled = True
            
        'End With
        'With .Controls.Add(Type:=msoControlButton)
            '.OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRunMKCWEDownload"
            '.FaceId = 688
            '.Caption = "Run CWE Download"
            '.Enabled = True
        'End With
        'With .Controls.Add(Type:=msoControlButton)
            '.OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRunMKItalyDownload"
            '.FaceId = 688
            '.Caption = "Run Italy Download"
            '.Enabled = True
        'End With
        
        
        
        'With .Controls.Add(Type:=msoControlButton)
            '.OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRunMKTempsDownload"
            '.FaceId = 688
            '.Caption = "Run Temperatures Download"
            '.Enabled = True
        'End With
       
       '.Controls(4).BeginGroup = True
       
        'With .Controls.Add(Type:=msoControlButton)
            '.OnAction = "'" & ThisWorkbook.Name & "'!" & "cbRunMKPricesDownload"
            '.FaceId = 688
            '.Caption = "Run Prices Download"
            '.Enabled = True
        'End With

'End With
'''Redundancy menu
    
    'Add seperator to the Cell menu
    'ContextMenu.Controls(5).BeginGroup = True

Set ContextMenu = Nothing: Set MyRefreshSubMenu = Nothing: Set MyRefreshSubMenu1 = Nothing
Set MyRefreshSubMenu2 = Nothing ': Set MyRedundancySubMenu = Nothing

End Sub

Sub DeleteFromCellMenu()
Dim ContextMenu As CommandBar
Dim ctrl As CommandBarControl

'Set ContextMenu to the Cell menu
Set ContextMenu = Application.CommandBars("Cell")

'Delete custom controls with the Tag : My_Cell_Control_Tag
For Each ctrl In ContextMenu.Controls
    If ctrl.Tag = "My_Cell_Control_Tag" Xor ctrl.Caption = "" Then
        ctrl.Delete
    End If
Next ctrl

Set ContextMenu = Nothing
End Sub

