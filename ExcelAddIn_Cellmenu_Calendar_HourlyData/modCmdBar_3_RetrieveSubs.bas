Attribute VB_Name = "modCmdBar_3_RetrieveSubs"

'*****************
'MK_Data.xlsm

'If one wants to make modifications for the excel add-inn (MK_Data.xlam),
'he/she can make them on this xlsm file and later save a copy of it as an xlam file.
'*****************

Option Explicit
Option Base 1
Option Compare Text

Const cbsMyDirectory As String = "C:\"  '"C:\Users\faidon.dermesonoglou\Desktop\Addinns\GH\"

Private Function cbfCreateStringAddress(MyRow As Long, MyColumn As Long) As String
Dim cbsAddressResult As String
cbsAddressResult = Cells(MyRow, MyColumn).Resize(24, 1).Address
cbsAddressResult = Replace(cbsAddressResult, "$", "")
cbfCreateStringAddress = cbsAddressResult
End Function

Sub cbRetrieveData(cbsMyData As String, cbsMyType As String, cbsMyZone As String, cbsMyDate1 As Double, cbsMyDate2 As Double)
Dim cbcMKData As ccbsImportData
Dim cbsMyCurrentCell As String, cbsMyCurrentWS As String, cbsMySubDirectory As String, cbsMyFileNameA As String, cbsMyFileNameB As String
Dim cbsSourceAddressWB1 As String, cbsSourceAddressWB2 As String

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


Const rITEnOp As Integer = 210  'difference in rows between operational forecast and ensemble forecast for italy ec00 etc
Const rCWEnOp As Integer = 330
Const rCEEnOp As Integer = 150

'above constants used to determine the cell reference ie Cells(rITCSUD,2).Offset etc

rChangeEnOp = 0
cbsDate = Format(Date, "YYYYMMDD")
cbsDateB = Format(Date - 1, "YYYYMMDD")
cbsDateC = Format(cbsMyDate2 - 1, "YYYYMMDD")
cbsMyCurrentCell = Selection.Address
cbsMyCurrentWS = Selection.Worksheet.Name

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
            .sDestWSName = cbsMyCurrentWS
            .sDestWSRange = cbsMyCurrentCell
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
            .sDestWSName = cbsMyCurrentWS
            .sDestWSRange = cbsMyCurrentCell
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
            .sDestWSName = cbsMyCurrentWS
            .sDestWSRange = cbsMyCurrentCell
            .fImpRangeWB2
        End With
        Set cbcMKData = Nothing

End Select


End Sub



