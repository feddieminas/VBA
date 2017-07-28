Attribute VB_Name = "modCmdBar_4_TagActions"

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit
Option Base 1
Option Compare Text

Private Type tCompTAG
    ActualTag As Boolean
    PublicationDate As Double
    Horizontal As Boolean
    ChngNegative As Boolean
    Publisher As String
    MyData As String
    MyType As String
    MyZone As String
    TargetDate As Double
    DeltaDate As Double
End Type

Const cbsMyDirectory As String = "C:\"  '"C:\Users\faidon.dermesonoglou\Desktop\Addinns\GH\"

Public gcbTagsDiscovered As Long
Public gcbTagsUpdated As Long
Public gcbTagsProblems As Long
Public gcbTagsActual As Long
Public gcbTagsOutofDate As Long

Private Sub cbClearStatusBar()
    Application.StatusBar = False
End Sub

Private Sub cbSwitchOff()
With Application
.EnableEvents = False
.Calculation = xlCalculationManual
'.ScreenUpdating = False
'.DisplayStatusBar = False
End With
End Sub
Private Sub cbSwitchOn()
With Application
.EnableEvents = True
.Calculation = xlCalculationAutomatic
'.ScreenUpdating = True
'.DisplayStatusBar = True
End With
End Sub


Private Sub cbPrepFill(cbWB As String, cbWS As String, cbCell As String, cbHorizontal As Boolean)
With Workbooks(CStr(cbWB)).Worksheets(CStr(cbWS)).Range(CStr(cbCell))
    If cbHorizontal Then
        .Resize(1, 24).Interior.ColorIndex = 3
        .Resize(1, 24).Font.ColorIndex = 2
        .Resize(1, 24).Font.Bold = False
    Else
        .Resize(24, 1).Interior.ColorIndex = 3
        .Resize(24, 1).Font.ColorIndex = 2
        .Resize(24, 1).Font.Bold = False
    End If
End With
End Sub

Private Sub cbClearFill(cbWB As String, cbWS As String, cbCell As String, cbHorizontal As Boolean)
With Workbooks(CStr(cbWB)).Worksheets(CStr(cbWS)).Range(CStr(cbCell))
    If cbHorizontal Then
        .Resize(1, 24).Interior.ColorIndex = 0
        .Resize(1, 24).Font.ColorIndex = 1
      '  .Resize(1, 24).Font.Bold = False
    Else
        .Resize(24, 1).Interior.ColorIndex = 0
        .Resize(24, 1).Font.ColorIndex = 1
      '  .Resize(24, 1).Font.Bold = False
    End If
End With
End Sub

Private Sub cbActual(cbWB As String, cbWS As String, cbCell As String, cbHorizontal As Boolean)
With Workbooks(CStr(cbWB)).Worksheets(CStr(cbWS)).Range(CStr(cbCell))
    If cbHorizontal Then
        .Resize(1, 24).Interior.ColorIndex = 19
        .Resize(1, 24).Font.ColorIndex = 1
    Else
        .Resize(24, 1).Interior.ColorIndex = 19
        .Resize(24, 1).Font.ColorIndex = 1
    End If
End With
End Sub
Private Sub cbOutofDate(cbWB As String, cbWS As String, cbCell As String, cbHorizontal As Boolean)
With Workbooks(cbWB).Worksheets(cbWS).Range(cbCell)
    If cbHorizontal Then
        .Resize(1, 24).Interior.ColorIndex = 15
        .Resize(1, 24).Font.ColorIndex = 1
    Else
        .Resize(24, 1).Interior.ColorIndex = 15
        .Resize(24, 1).Font.ColorIndex = 1
    End If
End With
End Sub

Private Function cbfCreateStringAddress(MyRow As Long, MyColumn As Long) As String
Dim cbsAddressResult As String
cbsAddressResult = Cells(MyRow, MyColumn).Resize(24, 1).Address
cbsAddressResult = Replace(cbsAddressResult, "$", "")
cbfCreateStringAddress = cbsAddressResult
End Function

Private Function cbExtractCompTAG(MyString As String) As tCompTAG
Dim MyLength As Long, MyStartPos As Long, MyEndPos As Long
Dim arrCompTAG() As String, MyTag As String
Dim mTAGResult As tCompTAG

MyLength = Len(MyString)
If MyLength > 0 Then

    MyStartPos = InStr(1, MyString, "<<CompTAG")
    MyEndPos = InStr(1, MyString, "CompTAG>>")

        If MyStartPos > 0 And MyEndPos > 0 Then
            MyTag = Mid(MyString, MyStartPos, MyEndPos + 10 - MyStartPos)

            arrCompTAG() = Split(MyTag, "&")
                With mTAGResult
                    .ActualTag = True
                    .PublicationDate = CDbl(arrCompTAG(1))
                    .Horizontal = CBool(arrCompTAG(2))
                    .ChngNegative = CBool(arrCompTAG(3))
                    .Publisher = CStr(arrCompTAG(4))
                    .MyData = CStr(arrCompTAG(5))
                    .MyType = CStr(arrCompTAG(6))
                    .MyZone = CStr(arrCompTAG(7))
                    .TargetDate = CDbl(arrCompTAG(8))
                    .DeltaDate = CDbl(arrCompTAG(9))
                End With

            cbExtractCompTAG = mTAGResult
            Erase arrCompTAG()
        End If
        Exit Function
End If

mTAGResult.ActualTag = False
cbExtractCompTAG = mTAGResult

End Function

Sub cbAbsDiscoverRefreshRangeTags()
Dim cbMyComment As Comment
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String
Dim MyRefreshString As String, MyReplaceString As String, MyFindString As String
Dim MyReplaceStart As Long, MyReplaceEnd As Long


cbPrepRangeTagsforRefresh       'fill all tagged columns with red and white.
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOff
For Each cbMyCell In Selection.Cells
    If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment
             cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
        'for refreshing MK Tags below
        
            With cbMyTAGDetails
                If .ActualTag Then
                    
                    If Not .MyData = "Act" Then             'Actual data does not need to be refreshed
                        If .TargetDate > CDbl(Date) Then    'If Forecast date is past then no longer exists
                            MyDestWB = cbMyCell.Parent.Parent.Name: MyDestWS = cbMyCell.Parent.Name: MyDestCell = Replace(cbMyCell.Address, "$", "")
                            cbTAGRetrieveData cbsMyData:=.MyData, cbsMyType:=.MyType, cbsMyZone:=.MyZone, cbsMyDate1:=.TargetDate, cbsMyDate2:=.DeltaDate, _
                            cbsMyDestWB:=MyDestWB, cbsMyDestWS:=MyDestWS, cbsDestCell:=MyDestCell, cbsHorizontal:=.Horizontal, cbsNegative:=.ChngNegative
                        'update comment box
                        MyRefreshString = cbMyCell.Comment.Text
                        MyReplaceStart = InStr(30, MyRefreshString, "Refresh Date:")
                        MyReplaceEnd = InStrRev(MyRefreshString, "/")
                        MyFindString = Mid(MyRefreshString, MyReplaceStart, MyReplaceEnd - MyReplaceStart + 5)

                        If .DeltaDate > 0 Then
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY") & Chr(10) & "Delta    Date: " & Format(.DeltaDate, "DD/MM/YYYY")
                        Else
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY")
                        End If
                        MyRefreshString = Replace(MyRefreshString, MyFindString, MyReplaceString, 1, 1) 'replace 30,1 with 1,1
                        cbMyCell.Comment.Text MyRefreshString
                        'end of update comment box
                        
                        Else
                        cbOutofDate cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                        gcbTagsOutofDate = gcbTagsOutofDate + 1
                        End If
                    Else
                    cbActual cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsActual = gcbTagsActual + 1
                    End If
                End If
            End With
        End With
    End If
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
Next cbMyCell
Application.StatusBar = "Tag Refresh Completed: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOn
Application.OnTime Now + TimeSerial(0, 0, 30), "cbClearStatusBar"
End Sub

Sub cbAbsDiscoverRefreshWSTags()
Dim cbMyComment As Comment
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String
Dim MyRefreshString As String, MyReplaceString As String, MyFindString As String
Dim MyReplaceStart As Long, MyReplaceEnd As Long

cbPrepWSTagsforRefresh
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOff
For Each cbMyComment In Selection.Parent.Comments
            Set cbMyCell = Selection.Parent.Range(cbMyComment.Parent.Address)
     If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment


            cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
   
            With cbMyTAGDetails
                If .ActualTag Then
                    
                    If Not .MyData = "Act" Then             'Actual data does not need to be refreshed
                        If .TargetDate > CDbl(Date) Then    'If Forecast date is past then no longer exists
                            MyDestWB = cbMyCell.Parent.Parent.Name: MyDestWS = cbMyCell.Parent.Name: MyDestCell = Replace(cbMyCell.Address, "$", "")
                            cbTAGRetrieveData cbsMyData:=.MyData, cbsMyType:=.MyType, cbsMyZone:=.MyZone, cbsMyDate1:=.TargetDate, cbsMyDate2:=.DeltaDate, _
                            cbsMyDestWB:=MyDestWB, cbsMyDestWS:=MyDestWS, cbsDestCell:=MyDestCell, cbsHorizontal:=.Horizontal, cbsNegative:=.ChngNegative
                        'update comment box
                        MyRefreshString = cbMyCell.Comment.Text
                        MyReplaceStart = InStr(30, MyRefreshString, "Refresh Date:")
                        MyReplaceEnd = InStrRev(MyRefreshString, "/")
                        MyFindString = Mid(MyRefreshString, MyReplaceStart, MyReplaceEnd - MyReplaceStart + 5)

                        If .DeltaDate > 0 Then
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY") & Chr(10) & "Delta    Date: " & Format(.DeltaDate, "DD/MM/YYYY")
                        Else
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY")
                        End If
                        MyRefreshString = Replace(MyRefreshString, MyFindString, MyReplaceString, 1, 1)
                        cbMyCell.Comment.Text MyRefreshString
                        'end of update comment box
                        
                        Else
                        cbOutofDate cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                        gcbTagsOutofDate = gcbTagsOutofDate + 1
                        End If
                    Else
                    cbActual cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsActual = gcbTagsActual + 1
                    End If
                End If
            End With
        End With
    End If
Set cbMyCell = Nothing
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
Next cbMyComment
Application.StatusBar = "Tag Refresh Completed: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "    Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOn
Application.OnTime Now + TimeSerial(0, 0, 30), "cbClearStatusBar"
End Sub
Sub cbAbsDiscoverRefreshWBTags()
Dim cbMyComment As Comment
Dim cbMyWS As Worksheet
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String
Dim MyRefreshString As String, MyReplaceString As String
Dim MyReplaceStart As Long, MyReplaceEnd As Long

cbPrepWBTagsforRefresh
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOff
For Each cbMyWS In Selection.Parent.Parent.Worksheets
    For Each cbMyComment In cbMyWS.Comments
           Set cbMyCell = cbMyWS.Range(cbMyComment.Parent.Address)
     If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment
           ' MsgBox cbMyCell.Comment.Text & "is in cell " & cbMyCell.Address
            cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
        
            With cbMyTAGDetails
                If .ActualTag Then
                    
                    If Not .MyData = "Act" Then             'Actual data does not need to be refreshed
                        If .TargetDate > CDbl(Date) Then    'If Forecast date is past then no longer exists
                            MyDestWB = cbMyCell.Parent.Parent.Name: MyDestWS = cbMyCell.Parent.Name: MyDestCell = Replace(cbMyCell.Address, "$", "")
                            cbTAGRetrieveData cbsMyData:=.MyData, cbsMyType:=.MyType, cbsMyZone:=.MyZone, cbsMyDate1:=.TargetDate, cbsMyDate2:=.DeltaDate, _
                            cbsMyDestWB:=MyDestWB, cbsMyDestWS:=MyDestWS, cbsDestCell:=MyDestCell, cbsHorizontal:=.Horizontal, cbsNegative:=.ChngNegative
                         'update comment box
                        MyRefreshString = cbMyCell.Comment.Text
                        MyReplaceStart = InStr(30, MyRefreshString, "Refresh Date:")
                        MyReplaceEnd = InStrRev(MyRefreshString, "/")
                        MyFindString = Mid(MyRefreshString, MyReplaceStart, MyReplaceEnd - MyReplaceStart + 5)

                        If .DeltaDate > 0 Then
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY") & Chr(10) & "Delta    Date: " & Format(.DeltaDate, "DD/MM/YYYY")
                        Else
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY")
                        End If
                        MyRefreshString = Replace(MyRefreshString, MyFindString, MyReplaceString, 1, 1)
                        cbMyCell.Comment.Text MyRefreshString
                        'end of update comment box
                        
                        Else
                        cbOutofDate cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                        gcbTagsOutofDate = gcbTagsOutofDate + 1
                        End If
                    Else
                    cbActual cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsActual = gcbTagsActual + 1
                    End If
                End If
            End With
        End With
    End If
    
    Set cbMyCell = Nothing
    Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "    Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
    Next cbMyComment
Next cbMyWS

Application.StatusBar = "Tag Refresh Completed: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOn
Application.OnTime Now + TimeSerial(0, 0, 30), "cbClearStatusBar"

End Sub

Sub cbRelDiscoverRefreshRangeTags()
Dim cbMyComment As Comment
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String
Dim ShiftDates As Double
Dim MyRefreshString As String, MyReplaceString As String, MyFindString As String
Dim MyReplaceStart As Long, MyReplaceEnd As Long

cbPrepRangeTagsforRefresh       'fill all tagged columns with red and white.
cbSwitchOff
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
For Each cbMyCell In Selection.Cells
    If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment
             cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
        
            With cbMyTAGDetails
                      
                If .ActualTag Then
                ShiftDates = CDbl(Date) - .PublicationDate
                If .DeltaDate > 0 Then                          ' new correction double check - this is cos of zero for non delta retrievals
                    .DeltaDate = .DeltaDate + ShiftDates
                End If
                .TargetDate = .TargetDate + ShiftDates
                                
                    
                    If Not .MyData = "Act" Then             'Actual data does not need to be refreshed
                        If .TargetDate > CDbl(Date) Then    'If Forecast date is past then no longer exists
                            MyDestWB = cbMyCell.Parent.Parent.Name: MyDestWS = cbMyCell.Parent.Name: MyDestCell = Replace(cbMyCell.Address, "$", "")
                            cbTAGRetrieveData cbsMyData:=.MyData, cbsMyType:=.MyType, cbsMyZone:=.MyZone, cbsMyDate1:=.TargetDate, cbsMyDate2:=.DeltaDate, _
                            cbsMyDestWB:=MyDestWB, cbsMyDestWS:=MyDestWS, cbsDestCell:=MyDestCell, cbsHorizontal:=.Horizontal, cbsNegative:=.ChngNegative
                         
                         'update comment box
                        MyRefreshString = cbMyCell.Comment.Text
                        MyReplaceStart = InStr(30, MyRefreshString, "Refresh Date:")
                        MyReplaceEnd = InStrRev(MyRefreshString, "/")
                        MyFindString = Mid(MyRefreshString, MyReplaceStart, MyReplaceEnd - MyReplaceStart + 5)

                        If .DeltaDate > 0 Then
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY") & Chr(10) & "Delta    Date: " & Format(.DeltaDate, "DD/MM/YYYY")
                        Else
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY")
                        End If
                        MyRefreshString = Replace(MyRefreshString, MyFindString, MyReplaceString, 1, 1)
                        cbMyCell.Comment.Text MyRefreshString
                        'end of update comment box
                        
                        
                        Else
                        cbOutofDate cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                        gcbTagsOutofDate = gcbTagsOutofDate + 1
                        End If
                    Else
                    cbActual cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsActual = gcbTagsActual + 1
                    End If
                End If
            End With
        End With
    End If
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
Next cbMyCell
Application.StatusBar = "Tag Refresh Completed: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOn
Application.OnTime Now + TimeSerial(0, 0, 30), "cbClearStatusBar"
End Sub
Sub cbRelDiscoverRefreshWSTags()
Dim cbMyComment As Comment
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String
Dim ShiftDates As Double
Dim MyRefreshString As String, MyReplaceString As String
Dim MyReplaceStart As Long, MyReplaceEnd As Long



cbPrepWSTagsforRefresh
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOff
For Each cbMyComment In Selection.Parent.Comments
            Set cbMyCell = Selection.Parent.Range(cbMyComment.Parent.Address)
     If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment


            cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
   
            With cbMyTAGDetails
                If .ActualTag Then
                ShiftDates = CDbl(Date) - .PublicationDate
                If .DeltaDate > 0 Then
                .DeltaDate = .DeltaDate + ShiftDates
                End If
                .TargetDate = .TargetDate + ShiftDates
                    
                    If Not .MyData = "Act" Then             'Actual data does not need to be refreshed
                        If .TargetDate > CDbl(Date) Then    'If Forecast date is past then no longer exists
                            MyDestWB = cbMyCell.Parent.Parent.Name: MyDestWS = cbMyCell.Parent.Name: MyDestCell = Replace(cbMyCell.Address, "$", "")
                            cbTAGRetrieveData cbsMyData:=.MyData, cbsMyType:=.MyType, cbsMyZone:=.MyZone, cbsMyDate1:=.TargetDate, cbsMyDate2:=.DeltaDate, _
                            cbsMyDestWB:=MyDestWB, cbsMyDestWS:=MyDestWS, cbsDestCell:=MyDestCell, cbsHorizontal:=.Horizontal, cbsNegative:=.ChngNegative
                          
                          'update comment box
                        MyRefreshString = cbMyCell.Comment.Text
                        MyReplaceStart = InStr(30, MyRefreshString, "Refresh Date:")
                        MyReplaceEnd = InStrRev(MyRefreshString, "/")
                        MyFindString = Mid(MyRefreshString, MyReplaceStart, MyReplaceEnd - MyReplaceStart + 5)

                        If .DeltaDate > 0 Then
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY") & Chr(10) & "Delta    Date: " & Format(.DeltaDate, "DD/MM/YYYY")
                        Else
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY")
                        End If
                        MyRefreshString = Replace(MyRefreshString, MyFindString, MyReplaceString, 1, 1) 'replace 30,1 with 1,1
                        cbMyCell.Comment.Text MyRefreshString
                        'end of update comment box
                       
                        
                        
                        Else
                        cbOutofDate cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                        gcbTagsOutofDate = gcbTagsOutofDate + 1
                        End If
                    Else
                    cbActual cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsActual = gcbTagsActual + 1
                    End If
                End If
            End With
        End With
    End If
Set cbMyCell = Nothing
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
Next cbMyComment
Application.StatusBar = "Tag Refresh Completed: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "    Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOn
Application.OnTime Now + TimeSerial(0, 0, 30), "cbClearStatusBar"
End Sub
Sub cbRelDiscoverRefreshWBTags()
Dim cbMyComment As Comment
Dim cbMyWS As Worksheet
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String
Dim ShiftDates As Double
Dim MyRefreshString As String, MyReplaceString As String
Dim MyReplaceStart As Long, MyReplaceEnd As Long

cbPrepWBTagsforRefresh
Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "     Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOff
For Each cbMyWS In Selection.Parent.Parent.Worksheets
    For Each cbMyComment In cbMyWS.Comments
           Set cbMyCell = cbMyWS.Range(cbMyComment.Parent.Address)
     If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment
           ' MsgBox cbMyCell.Comment.Text & "is in cell " & cbMyCell.Address
            cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
        
            With cbMyTAGDetails
                If .ActualTag Then
                 ShiftDates = CDbl(Date) - .PublicationDate
                If .DeltaDate > 0 Then
                .DeltaDate = .DeltaDate + ShiftDates
                End If
                 .TargetDate = .TargetDate + ShiftDates
                    
                    If Not .MyData = "Act" Then             'Actual data does not need to be refreshed
                        If .TargetDate > CDbl(Date) Then    'If Forecast date is past then no longer exists
                            MyDestWB = cbMyCell.Parent.Parent.Name: MyDestWS = cbMyCell.Parent.Name: MyDestCell = Replace(cbMyCell.Address, "$", "")
                            cbTAGRetrieveData cbsMyData:=.MyData, cbsMyType:=.MyType, cbsMyZone:=.MyZone, cbsMyDate1:=.TargetDate, cbsMyDate2:=.DeltaDate, _
                            cbsMyDestWB:=MyDestWB, cbsMyDestWS:=MyDestWS, cbsDestCell:=MyDestCell, cbsHorizontal:=.Horizontal, cbsNegative:=.ChngNegative
                        
                         'update comment box
                        MyRefreshString = cbMyCell.Comment.Text
                        MyReplaceStart = InStr(30, MyRefreshString, "Refresh Date:")
                        MyReplaceEnd = InStrRev(MyRefreshString, "/")
                        MyFindString = Mid(MyRefreshString, MyReplaceStart, MyReplaceEnd - MyReplaceStart + 5)

                        If .DeltaDate > 0 Then
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY") & Chr(10) & "Delta    Date: " & Format(.DeltaDate, "DD/MM/YYYY")
                        Else
                            MyReplaceString = "Refresh Date: " & Format(Date, "DD/MM/YYYY") & Chr(10) & "Target  Date: " & Format(.TargetDate, "DD/MM/YYYY")
                        End If
                        MyRefreshString = Replace(MyRefreshString, MyFindString, MyReplaceString, 1, 1) 'replace 30,1 with 1,1
                        cbMyCell.Comment.Text MyRefreshString
                        'end of update comment box

                        Else
                        cbOutofDate cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                        gcbTagsOutofDate = gcbTagsOutofDate + 1
                        End If
                    Else
                    cbActual cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsActual = gcbTagsActual + 1
                    End If
                End If
            End With
        End With
    End If
    
    Set cbMyCell = Nothing
    Application.StatusBar = "Refreshing Tags: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "    Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
    Next cbMyComment
Next cbMyWS

Application.StatusBar = "Tag Refresh Completed: Tags found: " & gcbTagsDiscovered & "    Tags succcesfully updated: " & gcbTagsUpdated & "     Tags which failed to updated: " & gcbTagsProblems & "     Tags containing Actual Data: " & gcbTagsActual & "     Tags containing forecasts no longer available: " & gcbTagsOutofDate
cbSwitchOn
Application.OnTime Now + TimeSerial(0, 0, 30), "cbClearStatusBar"

End Sub

Private Sub cbTAGRetrieveData(cbsMyData As String, cbsMyType As String, cbsMyZone As String, cbsMyDate1 As Double, cbsMyDate2 As Double, cbsMyDestWB As String, cbsMyDestWS As String, cbsDestCell As String, cbsHorizontal As Boolean, cbsNegative As Boolean) 'This is a test sub we need to adapt to pass thru data thru button subs
Dim cbcMKData As ccbsImportData
Dim cbsMyCurrentCell As String, cbsMyCurrentWS As String, cbsMySubDirectory As String, cbsMyFileNameA As String, cbsMyFileNameB As String
Dim cbsSourceAddressWB1 As String, cbsSourceAddressWB2 As String, cbsMyCurrentWB As String

Dim cbsMyColumnDate1 As Long, cbsMyColumnDate2 As Long
Dim cbsMyRowA As Long, rChangeEnOp As Long, cbsMyRowB As Long
Dim cbsDate As String * 8, cbsDateB As String * 8, cbsDateC As String * 8
Const rITNATI As Integer = 2
Const rITNORD As Integer = 32
Const rITCNORD As Integer = 62
Const rITCSUD As Integer = 92
Const rITSUDS As Integer = 122
Const rITSARD As Integer = 152
Const rITSICI As Integer = 182

Const rCWEDE As Integer = 32
Const rCWEFR As Integer = 182
Const rCWEBE As Integer = 302
Const rCWENL As Integer = 272
Const rCWEAU As Integer = 212
Const rCWECH As Integer = 242

Const rCEEPL As Integer = 32
Const rCEECZ As Integer = 62
Const rCEESK As Integer = 92
Const rCEEHU As Integer = 122


Const rITEnOp As Integer = 210  'difference in rows between operational forecast and ensemble forecast for italy ec00
Const rCWEnOp As Integer = 330 ' this needs to be corrected when known for CWE ie de fr be nl
Const rCEEnOp As Integer = 150
'we need to add select case for understanding the correct file, and a calculation for obtaining the correct column.
'above constants used to determine the cell reference ie Cells(rITCSUD,2).Offset etc

rChangeEnOp = 0
cbsDate = Format(Date, "YYYYMMDD")
cbsDateB = Format(Date - 1, "YYYYMMDD")
cbsDateC = Format(cbsMyDate2 - 1, "YYYYMMDD")
cbsMyCurrentCell = cbsDestCell
cbsMyCurrentWS = cbsMyDestWS
cbsMyCurrentWB = cbsMyDestWB

Select Case cbsMyZone
    Case "CW_DEUT"
        cbsMySubDirectory = "CWE\"
        cbsMyRowA = rCWEDE
        rChangeEnOp = rCWEnOp
    Case Is = "CW_FRAN"
        cbsMySubDirectory = "CWE\"
        cbsMyRowA = rCWEFR
        rChangeEnOp = rCWEnOp
    Case Is = "CW_AUST"
        cbsMySubDirectory = "CWE\"
        cbsMyRowA = rCWEAU
        rChangeEnOp = rCWEnOp
        
    '**********************************
    'Case Is = "IT_NATI"
        'cbsMySubDirectory = "Italy\"
        'cbsMyRowA = rITNATI
        'rChangeEnOp = rITEnOp
    'Case Is = "IT_NORD"
        'cbsMySubDirectory = "Italy\"
        'cbsMyRowA = rITNORD
        'rChangeEnOp = rITEnOp
    'Case Is = "IT_CNOR"
        'cbsMySubDirectory = "Italy\"
        'cbsMyRowA = rITCNORD
        'rChangeEnOp = rITEnOp
    'Case Is = "IT_CSUD"
        'cbsMySubDirectory = "Italy\"
        'cbsMyRowA = rITCSUD
        'rChangeEnOp = rITEnOp
    'Case Is = "IT_SUDS"
        'cbsMySubDirectory = "Italy\"
        'cbsMyRowA = rITSUDS
    'Case Is = "IT_SARD"
        'cbsMySubDirectory = "Italy\"
        'cbsMyRowA = rITSARD
        'rChangeEnOp = rITEnOp
    'Case Is = "IT_SICI"
        'cbsMySubDirectory = "Italy\"
        'cbsMyRowA = rITSICI
        'rChangeEnOp = rITEnOp
    '**********************************
        
    Case Is = "CW_BELG"
        cbsMySubDirectory = "CWE\"
        cbsMyRowA = rCWEBE
         rChangeEnOp = rCWEnOp
    Case Is = "CW_NEDE"
        cbsMySubDirectory = "CWE\"
        cbsMyRowA = rCWENL
        rChangeEnOp = rCWEnOp
        Case Is = "CW_SWIS"
        cbsMySubDirectory = "CWE\"
        cbsMyRowA = rCWECH
        rChangeEnOp = rCWEnOp
        
    '**********************************
    'Case Is = "CE_CZEC"
        'cbsMySubDirectory = "CEE\"
        'cbsMyRowA = rCEECZ
        'rChangeEnOp = rCEEnOp
    'Case Is = "CE_HUNG"
        'cbsMySubDirectory = "CEE\"
        'cbsMyRowA = rCEEHU
        'rChangeEnOp = rCEEnOp
    'Case Is = "CE_POLA"
        'cbsMySubDirectory = "CEE\"
        'cbsMyRowA = rCEEPL
        'rChangeEnOp = rCEEnOp
    'Case Is = "CE_SLKA"
        'cbsMySubDirectory = "CEE\"
        'cbsMyRowA = rCEESK
        'rChangeEnOp = rCEEnOp
    '**********************************

End Select


Select Case cbsMyType

Case Is = "CON"
    
    Select Case cbsMyData
'***********************************************************************************************************************
'beginning of actual

        Case Is = "ACT"
            cbsMyFileNameA = "MK_A_Con_"
        
'end of actual
'***********************************************************************************************************************
'beginning of operational
        
        
        Case Is = "FOR_OP"
            cbsMyFileNameA = "MK_F_Con_"
         
        Case Is = "FAD_OP"
            cbsMyFileNameA = "MK_F_Con_"
            cbsMyFileNameB = "MK_A_Con_"
            cbsMyRowB = cbsMyRowA
            
        Case Is = "FDE_OP"
            cbsMyFileNameA = "MK_F_Con_"
            cbsMyFileNameB = "MK_F_Con_"
            cbsMyRowB = cbsMyRowA
        
        Case Is = "FSH_OP"
            cbsMyFileNameA = "MK_F_Con_"
            cbsMyFileNameB = "MK_F_Con_"
            cbsMyRowB = cbsMyRowA
            
'end of operational
'********************************************************************************************************************
'beginning of ensemble

        Case Is = "FOR_ES"
            cbsMyFileNameA = "MK_F_Con_"
            cbsMyRowA = rChangeEnOp + cbsMyRowA

        Case Is = "FAD_ES"
            cbsMyFileNameA = "MK_F_Con_"
            cbsMyFileNameB = "MK_A_Con_"
            cbsMyRowB = cbsMyRowA
            cbsMyRowA = rChangeEnOp + cbsMyRowA
 
        Case Is = "FDE_ES"
            cbsMyFileNameA = "MK_F_Con_"
            cbsMyFileNameB = "MK_F_Con_"
            cbsMyRowA = rChangeEnOp + cbsMyRowA
            cbsMyRowB = cbsMyRowA
        
        Case Is = "FSH_ES"
            cbsMyFileNameA = "MK_F_Con_"
            cbsMyFileNameB = "MK_F_Con_"
            cbsMyRowA = rChangeEnOp + cbsMyRowA
            cbsMyRowB = cbsMyRowA
    End Select
'end of ensemble
'******************************************************************************************************************



'Case Is = "PVO"

    'Select Case cbsMyData
'******************************************************************************************************************
'beginning of Actual

        'Case Is = "ACT"
            'cbsMyFileNameA = "MK_A_PV_"
'end of actual
'******************************************************************************************************************
'beginning of operational

        'Case Is = "FOR_OP"
            'cbsMyFileNameA = "MK_F_PV_"
        
        
        'Case Is = "FAD_OP"
            'cbsMyFileNameA = "MK_F_PV_"
            'cbsMyFileNameB = "MK_A_PV_"
            'cbsMyRowB = cbsMyRowA

        'Case Is = "FDE_OP"
            'cbsMyFileNameA = "MK_F_PV_"
            'cbsMyFileNameB = "MK_F_PV_"
            'cbsMyRowB = cbsMyRowA
        
        'Case Is = "FSH_OP"
            'cbsMyFileNameA = "MK_F_PV_"
            'cbsMyFileNameB = "MK_F_PV_"
            'cbsMyRowB = cbsMyRowA
        
'end of operational
'******************************************************************************************************************
'beginning of ensemble

        'Case Is = "FOR_ES"
            'cbsMyFileNameA = "MK_F_PV_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
        
        'Case Is = "FAD_ES"
            'cbsMyFileNameA = "MK_F_PV_"
            'cbsMyFileNameB = "MK_A_PV_"
            'cbsMyRowB = cbsMyRowA
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
        
        'Case Is = "FDE_ES"
            'cbsMyFileNameA = "MK_F_PV_"
            'cbsMyFileNameB = "MK_F_PV_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
            'cbsMyRowB = cbsMyRowA
        
        'Case Is = "FSH_ES"
            'cbsMyFileNameA = "MK_F_PV_"
            'cbsMyFileNameB = "MK_F_PV_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
            'cbsMyRowB = cbsMyRowA
'end of ensemble
'******************************************************************************************************************
    'End Select

'Case Is = "WIN"

    'Select Case cbsMyData
'**********************************************************************************************************************
'beginning of actual

        'Case Is = "ACT"
            'cbsMyFileNameA = "MK_A_Win_"
'end of actual
'**********************************************************************************************************************
'beginning of operational

        'Case Is = "FOR_OP"
            'cbsMyFileNameA = "MK_F_Win_"

        'Case Is = "FAD_OP"
            'cbsMyFileNameA = "MK_F_Win_"
            'cbsMyFileNameB = "MK_A_Win_"
            'cbsMyRowB = cbsMyRowA

        'Case Is = "FDE_OP"
            'cbsMyFileNameA = "MK_F_Win_"
            'cbsMyFileNameB = "MK_F_Win_"
            'cbsMyRowB = cbsMyRowA

        'Case Is = "FSH_OP"
            'cbsMyFileNameA = "MK_F_Win_"
            'cbsMyFileNameB = "MK_F_Win_"
            'cbsMyRowB = cbsMyRowA

'end of operational
'***********************************************************************************************************************
'beginning of ensemble
        
        'Case Is = "FOR_ES"
            'cbsMyFileNameA = "MK_F_Win_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA

        'Case Is = "FAD_ES"
            'cbsMyFileNameA = "MK_F_Win_"
            'cbsMyFileNameB = "MK_A_Win_"
            'cbsMyRowB = cbsMyRowA                   'NB cos 2nd wb is actual there is no need for ensemb operational shift
            'cbsMyRowA = rChangeEnOp + cbsMyRowA

        'Case Is = "FDE_ES"
            'cbsMyFileNameA = "MK_F_Win_"
            'cbsMyFileNameB = "MK_F_Win_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
            'cbsMyRowB = cbsMyRowA
        
        'Case Is = "FSH_ES"
            'cbsMyFileNameA = "MK_F_Win_"
            'cbsMyFileNameB = "MK_F_Win_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
            'cbsMyRowB = cbsMyRowA

'end of ensemble
'**********************************************************************************************************************
    
    'End Select
'Case Is = "TEM"

    'Select Case cbsMyData
'**********************************************************************************************************************
'beginning of actual

        'Case Is = "ACT"
            'cbsMyFileNameA = "MK_A_TT_"
'end of actual
'**********************************************************************************************************************
'beginning of operational

        'Case Is = "FOR_OP"
            'cbsMyFileNameA = "MK_F_TT_"


        'Case Is = "FAD_OP"
            'cbsMyFileNameA = "MK_F_TT_"
            'cbsMyFileNameB = "MK_A_TT_"
            'cbsMyRowB = cbsMyRowA


        'Case Is = "FDE_OP"
            'cbsMyFileNameA = "MK_F_TT_"
            'cbsMyFileNameB = "MK_F_TT_"
            'cbsMyRowB = cbsMyRowA

        'Case Is = "FSH_OP"
            'cbsMyFileNameA = "MK_F_TT_"
            'cbsMyFileNameB = "MK_F_TT_"
            'cbsMyRowB = cbsMyRowA

'end of operational
'***********************************************************************************************************************
'beginning of ensemble
        
        'Case Is = "FOR_ES"
            'cbsMyFileNameA = "MK_F_TT_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA

        'Case Is = "FAD_ES"
            'cbsMyFileNameA = "MK_F_TT_"
            'cbsMyFileNameB = "MK_A_TT_"
            'cbsMyRowB = cbsMyRowA                   'NB cos 2nd wb is actual there is no need for ensemb operational shift
            'cbsMyRowA = rChangeEnOp + cbsMyRowA

        'Case Is = "FDE_ES"
            'cbsMyFileNameA = "MK_F_TT_"
            'cbsMyFileNameB = "MK_F_TT_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
            'cbsMyRowB = cbsMyRowA
        
        'Case Is = "FSH_ES"
            'cbsMyFileNameA = "MK_F_TT_"
            'cbsMyFileNameB = "MK_F_TT_"
            'cbsMyRowA = rChangeEnOp + cbsMyRowA
            'cbsMyRowB = cbsMyRowA

'end of ensemble
'**********************************************************************************************************************
    
    'End Select

'Case Is = "PRI"

    'Select Case cbsMyData
'**********************************************************************************************************************
'beginning of actual

        'Case Is = "ACT"
            'cbsMyFileNameA = "MK_A_Price_"
            'cbsDate = Format(Date + 1, "YYYYMMDD")
'End Select

End Select

'calculate column using date

Select Case cbsMyData
    Case Is = "ACT" 'we go for todays files.
     If cbsMyType = "PRI" Then
        cbsMyColumnDate1 = CDbl(Date - cbsMyDate1) + 3      'cos price is labelled with tomorrow file date
        cbsSourceAddressWB1 = cbfCreateStringAddress(cbsMyRowA, cbsMyColumnDate1)
        Else
        cbsMyColumnDate1 = CDbl(Date - cbsMyDate1) + 2
        cbsSourceAddressWB1 = cbfCreateStringAddress(cbsMyRowA, cbsMyColumnDate1)
     End If
    Case "FOR_OP", "FOR_ES"
        cbsMyColumnDate1 = CDbl(cbsMyDate1 - Date) + 2
        cbsSourceAddressWB1 = cbfCreateStringAddress(cbsMyRowA, cbsMyColumnDate1)
    
    Case "FAD_OP", "FAD_ES"
        cbsMyColumnDate1 = CDbl(cbsMyDate1 - Date) + 2
        cbsMyColumnDate2 = CDbl(Date - cbsMyDate2) + 2
        cbsSourceAddressWB1 = cbfCreateStringAddress(cbsMyRowA, cbsMyColumnDate1)
        cbsSourceAddressWB2 = cbfCreateStringAddress(cbsMyRowB, cbsMyColumnDate2)
        
       
    Case "FDE_OP", "FDE_ES"
        cbsMyColumnDate1 = CDbl(cbsMyDate1 - Date) + 2 ' this needs to be corrected just copied and pasted from above
        cbsMyColumnDate2 = CDbl(cbsMyDate2 - (cbsMyDate2 - 1)) + 2 ' this needs to be corrected just copied and pasted from above
        cbsSourceAddressWB1 = cbfCreateStringAddress(cbsMyRowA, cbsMyColumnDate1)
        cbsSourceAddressWB2 = cbfCreateStringAddress(cbsMyRowB, cbsMyColumnDate2)
   
        
    Case "FSH_ES", "FSH_OP"
        cbsMyColumnDate1 = CDbl(cbsMyDate1 - Date) + 2 ' this needs to be corrected just copied and pasted from above
        cbsMyColumnDate2 = CDbl(cbsMyDate2 - (Date - 1)) + 2 ' this needs to be corrected just copied and pasted from above
        cbsSourceAddressWB1 = cbfCreateStringAddress(cbsMyRowA, cbsMyColumnDate1)
        cbsSourceAddressWB2 = cbfCreateStringAddress(cbsMyRowB, cbsMyColumnDate2)

End Select

' We have to separate out again to reflect the need for one workbook retrieval or two workbook retrieval (cos we do a delta calc).

Select Case cbsMyData

    Case "Act", "FOR_OP", "FOR_ES"                  'retrieve data from one book only
        Set cbcMKData = New ccbsImportData
        With cbcMKData
            .sSourceWBNameA = cbsMyDirectory & cbsMySubDirectory & cbsMyFileNameA & cbsDate & ".xls"
            .sSourceWSName = "Data"
            .sSourceWSRangeA = cbsSourceAddressWB1
            .sDestWBName = cbsMyCurrentWB
            .sDestWSName = cbsMyCurrentWS
            .sDestWSRange = cbsMyCurrentCell
            .bboolChngeNeg = cbsNegative
            .bboolHorizontal = cbsHorizontal
            .fImpRangeWB1
        End With
        Set cbcMKData = Nothing

    Case "FSH_OP", "FSH_ES"
         
        Set cbcMKData = New ccbsImportData
        With cbcMKData
            .sSourceWBNameA = cbsMyDirectory & cbsMySubDirectory & cbsMyFileNameA & cbsDate & ".xls"
            .sSourceWBNameB = cbsMyDirectory & cbsMySubDirectory & cbsMyFileNameB & cbsDateB & ".xls"
            .sSourceWSName = "Data"
            .sSourceWSRangeA = cbsSourceAddressWB1
            .sSourceWSRangeB = cbsSourceAddressWB2
            .sDestWBName = cbsMyCurrentWB
            .sDestWSName = cbsMyCurrentWS
            .sDestWSRange = cbsMyCurrentCell
            .bboolChngeNeg = cbsNegative
            .bboolHorizontal = cbsHorizontal
            .fImpRangeWB2
        End With
        Set cbcMKData = Nothing
   
    Case "FDE_OP", "FDE_ES"                                       'retrieve data from two books as we are calculating a delta
        Set cbcMKData = New ccbsImportData
        With cbcMKData
            .sSourceWBNameA = cbsMyDirectory & cbsMySubDirectory & cbsMyFileNameA & cbsDate & ".xls"
            .sSourceWBNameB = cbsMyDirectory & cbsMySubDirectory & cbsMyFileNameB & cbsDateC & ".xls"
            .sSourceWSName = "Data"
            .sSourceWSRangeA = cbsSourceAddressWB1
            .sSourceWSRangeB = cbsSourceAddressWB2
            .sDestWSName = cbsMyCurrentWS
            .sDestWSRange = cbsMyCurrentCell
            .fImpRangeWB2
        End With
        Set cbcMKData = Nothing
    

    Case Else                                       'retrieve data from two books as we are calculating a delta
        Set cbcMKData = New ccbsImportData
        With cbcMKData
            .sSourceWBNameA = cbsMyDirectory & cbsMySubDirectory & cbsMyFileNameA & cbsDate & ".xls"
            .sSourceWBNameB = cbsMyDirectory & cbsMySubDirectory & cbsMyFileNameB & cbsDate & ".xls"
            .sSourceWSName = "Data"
            .sSourceWSRangeA = cbsSourceAddressWB1
            .sSourceWSRangeB = cbsSourceAddressWB2
            .sDestWBName = cbsMyCurrentWB
            .sDestWSName = cbsMyCurrentWS
            .sDestWSRange = cbsMyCurrentCell
            .bboolChngeNeg = cbsNegative
            .bboolHorizontal = cbsHorizontal
            .fImpRangeWB2
        End With
        Set cbcMKData = Nothing

End Select


End Sub

Private Sub cbPrepRangeTagsforRefresh()
Dim cbMyComment As Comment
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String

gcbTagsDiscovered = 0
gcbTagsUpdated = 0
gcbTagsProblems = 0
gcbTagsActual = 0
gcbTagsOutofDate = 0

For Each cbMyCell In Selection.Cells
    If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment
            cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
            With cbMyTAGDetails
                If .ActualTag Then
                     cbPrepFill cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                     gcbTagsDiscovered = gcbTagsDiscovered + 1
                End If
            End With
        End With
    End If
Next cbMyCell
End Sub

Private Sub cbPrepWSTagsforRefresh()
Dim cbMyComment As Comment
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String

gcbTagsDiscovered = 0
gcbTagsUpdated = 0
gcbTagsProblems = 0
gcbTagsActual = 0
gcbTagsOutofDate = 0

For Each cbMyComment In Selection.Parent.Comments
            Set cbMyCell = Selection.Parent.Range(cbMyComment.Parent.Address)
     If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment
           cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
            With cbMyTAGDetails
                If .ActualTag Then
                    cbPrepFill cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsDiscovered = gcbTagsDiscovered + 1
                 End If
            End With
        End With
    End If
Set cbMyCell = Nothing
Next cbMyComment
End Sub

Private Sub cbPrepWBTagsforRefresh()
Dim cbMyComment As Comment
Dim cbMyWS As Worksheet
Dim cbMyCell As Range, cbMyTAGDetails As tCompTAG
Dim MyDestWB As String, MyDestWS As String, MyDestCell As String

gcbTagsDiscovered = 0
gcbTagsUpdated = 0
gcbTagsProblems = 0
gcbTagsActual = 0
gcbTagsOutofDate = 0

For Each cbMyWS In Selection.Parent.Parent.Worksheets
    For Each cbMyComment In cbMyWS.Comments
           Set cbMyCell = cbMyWS.Range(cbMyComment.Parent.Address)
     If Not (cbMyCell.Comment Is Nothing) Then
        With cbMyCell.Comment
           ' MsgBox cbMyCell.Comment.Text & "is in cell " & cbMyCell.Address
            cbMyTAGDetails = cbExtractCompTAG(cbMyCell.Comment.Text)
            With cbMyTAGDetails
                If .ActualTag Then
                    cbPrepFill cbWB:=cbMyCell.Parent.Parent.Name, cbWS:=cbMyCell.Parent.Name, cbCell:=Replace(cbMyCell.Address, "$", ""), cbHorizontal:=.Horizontal
                    gcbTagsDiscovered = gcbTagsDiscovered + 1
                End If
            End With
        End With
    End If
    Set cbMyCell = Nothing
    Next cbMyComment
Next cbMyWS

End Sub

