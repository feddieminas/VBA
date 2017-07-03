Attribute VB_Name = "MainPO"
Option Explicit
Option Base 1

Sub POLayout(MyHub As Variant, arrOff1DataTmp() As Double, arrOff2DataTmp() As Double, arrOff3DataTmp() As Double, arrOff4DataTmp() As Double, arrOff5DataTmp() As Double _
, arrOff6DataTmp() As Double, arrOff7DataTmp() As Double, arrOff8DataTmp() As Double, arrOff9DataTmp() As Double _
, arrOff10DataTmp() As Double, arrOff11DataTmp() As Double, arrOff12DataTmp() As Double, arrOff13DataTmp() As Double _
, arrOff14DataTmp() As Double, arrOff15DataTmp() As Double, arrOff16DataTmp() As Double, arrOff17DataTmp() As Double _
, arrOff18DataTmp() As Double, arrOff19DataTmp() As Double, arrOff20DataTmp() As Double, arrOff21DataTmp() As Double _
, arrOff22DataTmp() As Double, arrOff23DataTmp() As Double, arrOff24DataTmp() As Double)
    
    '****Sort out the Distinct Off PricesBO and Accumulate the Quantities Off on that same PriceBO****'
    Dim Filtcounter As Long, i As Long, j As Long
    Dim MyRowData As Long, MyColData As Long
    
    'MyOfferData
Dim arrOff1Data() As Double, arrOff2Data() As Double, arrOff3Data() As Double, arrOff4Data() As Double
Dim arrOff5Data() As Double, arrOff6Data() As Double, arrOff7Data() As Double, arrOff8Data() As Double
Dim arrOff9Data() As Double, arrOff10Data() As Double, arrOff11Data() As Double, arrOff12Data() As Double
Dim arrOff13Data() As Double, arrOff14Data() As Double, arrOff15Data() As Double, arrOff16Data() As Double
Dim arrOff17Data() As Double, arrOff18Data() As Double, arrOff19Data() As Double, arrOff20Data() As Double
Dim arrOff21Data() As Double, arrOff22Data() As Double, arrOff23Data() As Double, arrOff24Data() As Double
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrOff1DataTmp, 1)
    MyColData = UBound(arrOff1DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff1DataTmp(i + 1, 2) <> arrOff1DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff1Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff1Data(i, j) = arrOff1DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff1DataTmp(i, 2) <> arrOff1DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff1Data(Filtcounter, j) = arrOff1DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff1DataTmp(i, 2) = arrOff1DataTmp(i - 1, 2) Then
    arrOff1Data(Filtcounter, 1) = arrOff1Data(Filtcounter, 1) + arrOff1DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff1DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff2DataTmp, 1)
    MyColData = UBound(arrOff2DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff2DataTmp(i + 1, 2) <> arrOff2DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff2Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff2Data(i, j) = arrOff2DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff2DataTmp(i, 2) <> arrOff2DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff2Data(Filtcounter, j) = arrOff2DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff2DataTmp(i, 2) = arrOff2DataTmp(i - 1, 2) Then
    arrOff2Data(Filtcounter, 1) = arrOff2Data(Filtcounter, 1) + arrOff2DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff2DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff3DataTmp, 1)
    MyColData = UBound(arrOff3DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff3DataTmp(i + 1, 2) <> arrOff3DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff3Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff3Data(i, j) = arrOff3DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff3DataTmp(i, 2) <> arrOff3DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff3Data(Filtcounter, j) = arrOff3DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff3DataTmp(i, 2) = arrOff3DataTmp(i - 1, 2) Then
    arrOff3Data(Filtcounter, 1) = arrOff3Data(Filtcounter, 1) + arrOff3DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff3DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff4DataTmp, 1)
    MyColData = UBound(arrOff4DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff4DataTmp(i + 1, 2) <> arrOff4DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff4Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff4Data(i, j) = arrOff4DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff4DataTmp(i, 2) <> arrOff4DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff4Data(Filtcounter, j) = arrOff4DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff4DataTmp(i, 2) = arrOff4DataTmp(i - 1, 2) Then
    arrOff4Data(Filtcounter, 1) = arrOff4Data(Filtcounter, 1) + arrOff4DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff4DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff5DataTmp, 1)
    MyColData = UBound(arrOff5DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff5DataTmp(i + 1, 2) <> arrOff5DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff5Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff5Data(i, j) = arrOff5DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff5DataTmp(i, 2) <> arrOff5DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff5Data(Filtcounter, j) = arrOff5DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff5DataTmp(i, 2) = arrOff5DataTmp(i - 1, 2) Then
    arrOff5Data(Filtcounter, 1) = arrOff5Data(Filtcounter, 1) + arrOff5DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff5DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff6DataTmp, 1)
    MyColData = UBound(arrOff6DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff6DataTmp(i + 1, 2) <> arrOff6DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff6Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff6Data(i, j) = arrOff6DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff6DataTmp(i, 2) <> arrOff6DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff6Data(Filtcounter, j) = arrOff6DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff6DataTmp(i, 2) = arrOff6DataTmp(i - 1, 2) Then
    arrOff6Data(Filtcounter, 1) = arrOff6Data(Filtcounter, 1) + arrOff6DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff6DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff7DataTmp, 1)
    MyColData = UBound(arrOff7DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff7DataTmp(i + 1, 2) <> arrOff7DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff7Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff7Data(i, j) = arrOff7DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff7DataTmp(i, 2) <> arrOff7DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff7Data(Filtcounter, j) = arrOff7DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff7DataTmp(i, 2) = arrOff7DataTmp(i - 1, 2) Then
    arrOff7Data(Filtcounter, 1) = arrOff7Data(Filtcounter, 1) + arrOff7DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff7DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff8DataTmp, 1)
    MyColData = UBound(arrOff8DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff8DataTmp(i + 1, 2) <> arrOff8DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff8Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff8Data(i, j) = arrOff8DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff8DataTmp(i, 2) <> arrOff8DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff8Data(Filtcounter, j) = arrOff8DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff8DataTmp(i, 2) = arrOff8DataTmp(i - 1, 2) Then
    arrOff8Data(Filtcounter, 1) = arrOff8Data(Filtcounter, 1) + arrOff8DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff8DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff9DataTmp, 1)
    MyColData = UBound(arrOff9DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff9DataTmp(i + 1, 2) <> arrOff9DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff9Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff9Data(i, j) = arrOff9DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff9DataTmp(i, 2) <> arrOff9DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff9Data(Filtcounter, j) = arrOff9DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff9DataTmp(i, 2) = arrOff9DataTmp(i - 1, 2) Then
    arrOff9Data(Filtcounter, 1) = arrOff9Data(Filtcounter, 1) + arrOff9DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff9DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff10DataTmp, 1)
    MyColData = UBound(arrOff10DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff10DataTmp(i + 1, 2) <> arrOff10DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff10Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff10Data(i, j) = arrOff10DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff10DataTmp(i, 2) <> arrOff10DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff10Data(Filtcounter, j) = arrOff10DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff10DataTmp(i, 2) = arrOff10DataTmp(i - 1, 2) Then
    arrOff10Data(Filtcounter, 1) = arrOff10Data(Filtcounter, 1) + arrOff10DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff10DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    MyRowData = UBound(arrOff11DataTmp, 1)
    MyColData = UBound(arrOff11DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff11DataTmp(i + 1, 2) <> arrOff11DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff11Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff11Data(i, j) = arrOff11DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff11DataTmp(i, 2) <> arrOff11DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff11Data(Filtcounter, j) = arrOff11DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff11DataTmp(i, 2) = arrOff11DataTmp(i - 1, 2) Then
    arrOff11Data(Filtcounter, 1) = arrOff11Data(Filtcounter, 1) + arrOff11DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff11DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff12DataTmp, 1)
    MyColData = UBound(arrOff12DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff12DataTmp(i + 1, 2) <> arrOff12DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff12Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff12Data(i, j) = arrOff12DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff12DataTmp(i, 2) <> arrOff12DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff12Data(Filtcounter, j) = arrOff12DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff12DataTmp(i, 2) = arrOff12DataTmp(i - 1, 2) Then
    arrOff12Data(Filtcounter, 1) = arrOff12Data(Filtcounter, 1) + arrOff12DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff12DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff13DataTmp, 1)
    MyColData = UBound(arrOff13DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff13DataTmp(i + 1, 2) <> arrOff13DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff13Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff13Data(i, j) = arrOff13DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff13DataTmp(i, 2) <> arrOff13DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff13Data(Filtcounter, j) = arrOff13DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff13DataTmp(i, 2) = arrOff13DataTmp(i - 1, 2) Then
    arrOff13Data(Filtcounter, 1) = arrOff13Data(Filtcounter, 1) + arrOff13DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff13DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff14DataTmp, 1)
    MyColData = UBound(arrOff14DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff14DataTmp(i + 1, 2) <> arrOff14DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff14Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff14Data(i, j) = arrOff14DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff14DataTmp(i, 2) <> arrOff14DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff14Data(Filtcounter, j) = arrOff14DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff14DataTmp(i, 2) = arrOff14DataTmp(i - 1, 2) Then
    arrOff14Data(Filtcounter, 1) = arrOff14Data(Filtcounter, 1) + arrOff14DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff14DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff15DataTmp, 1)
    MyColData = UBound(arrOff15DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff15DataTmp(i + 1, 2) <> arrOff15DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff15Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff15Data(i, j) = arrOff15DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff15DataTmp(i, 2) <> arrOff15DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff15Data(Filtcounter, j) = arrOff15DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff15DataTmp(i, 2) = arrOff15DataTmp(i - 1, 2) Then
    arrOff15Data(Filtcounter, 1) = arrOff15Data(Filtcounter, 1) + arrOff15DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff15DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff16DataTmp, 1)
    MyColData = UBound(arrOff16DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff16DataTmp(i + 1, 2) <> arrOff16DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff16Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff16Data(i, j) = arrOff16DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff16DataTmp(i, 2) <> arrOff16DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff16Data(Filtcounter, j) = arrOff16DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff16DataTmp(i, 2) = arrOff16DataTmp(i - 1, 2) Then
    arrOff16Data(Filtcounter, 1) = arrOff16Data(Filtcounter, 1) + arrOff16DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff16DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff17DataTmp, 1)
    MyColData = UBound(arrOff17DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff17DataTmp(i + 1, 2) <> arrOff17DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff17Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff17Data(i, j) = arrOff17DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff17DataTmp(i, 2) <> arrOff17DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff17Data(Filtcounter, j) = arrOff17DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff17DataTmp(i, 2) = arrOff17DataTmp(i - 1, 2) Then
    arrOff17Data(Filtcounter, 1) = arrOff17Data(Filtcounter, 1) + arrOff17DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff17DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff18DataTmp, 1)
    MyColData = UBound(arrOff18DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff18DataTmp(i + 1, 2) <> arrOff18DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff18Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff18Data(i, j) = arrOff18DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff18DataTmp(i, 2) <> arrOff18DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff18Data(Filtcounter, j) = arrOff18DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff18DataTmp(i, 2) = arrOff18DataTmp(i - 1, 2) Then
    arrOff18Data(Filtcounter, 1) = arrOff18Data(Filtcounter, 1) + arrOff18DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff18DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff19DataTmp, 1)
    MyColData = UBound(arrOff19DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff19DataTmp(i + 1, 2) <> arrOff19DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff19Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff19Data(i, j) = arrOff19DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff19DataTmp(i, 2) <> arrOff19DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff19Data(Filtcounter, j) = arrOff19DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff19DataTmp(i, 2) = arrOff19DataTmp(i - 1, 2) Then
    arrOff19Data(Filtcounter, 1) = arrOff19Data(Filtcounter, 1) + arrOff19DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff19DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff20DataTmp, 1)
    MyColData = UBound(arrOff20DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff20DataTmp(i + 1, 2) <> arrOff20DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff20Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff20Data(i, j) = arrOff20DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff20DataTmp(i, 2) <> arrOff20DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff20Data(Filtcounter, j) = arrOff20DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff20DataTmp(i, 2) = arrOff20DataTmp(i - 1, 2) Then
    arrOff20Data(Filtcounter, 1) = arrOff20Data(Filtcounter, 1) + arrOff20DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff20DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff21DataTmp, 1)
    MyColData = UBound(arrOff21DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff21DataTmp(i + 1, 2) <> arrOff21DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff21Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff21Data(i, j) = arrOff21DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff21DataTmp(i, 2) <> arrOff21DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff21Data(Filtcounter, j) = arrOff21DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff21DataTmp(i, 2) = arrOff21DataTmp(i - 1, 2) Then
    arrOff21Data(Filtcounter, 1) = arrOff21Data(Filtcounter, 1) + arrOff21DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff21DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff22DataTmp, 1)
    MyColData = UBound(arrOff22DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff22DataTmp(i + 1, 2) <> arrOff22DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff22Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff22Data(i, j) = arrOff22DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff22DataTmp(i, 2) <> arrOff22DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff22Data(Filtcounter, j) = arrOff22DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff22DataTmp(i, 2) = arrOff22DataTmp(i - 1, 2) Then
    arrOff22Data(Filtcounter, 1) = arrOff22Data(Filtcounter, 1) + arrOff22DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff22DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff23DataTmp, 1)
    MyColData = UBound(arrOff23DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff23DataTmp(i + 1, 2) <> arrOff23DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff23Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff23Data(i, j) = arrOff23DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff23DataTmp(i, 2) <> arrOff23DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff23Data(Filtcounter, j) = arrOff23DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff23DataTmp(i, 2) = arrOff23DataTmp(i - 1, 2) Then
    arrOff23Data(Filtcounter, 1) = arrOff23Data(Filtcounter, 1) + arrOff23DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff23DataTmp
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrOff24DataTmp, 1)
    MyColData = UBound(arrOff24DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctOffers do we have
    For i = 1 To MyRowData - 1
    If arrOff24DataTmp(i + 1, 2) <> arrOff24DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Off arrays
    ReDim arrOff24Data(1 To Filtcounter, 1 To 2)

    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrOff24Data(i, j) = arrOff24DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrOff24DataTmp(i, 2) <> arrOff24DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrOff24Data(Filtcounter, j) = arrOff24DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrOff24DataTmp(i, 2) = arrOff24DataTmp(i - 1, 2) Then
    arrOff24Data(Filtcounter, 1) = arrOff24Data(Filtcounter, 1) + arrOff24DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrOff24DataTmp
    
    'PO Results
    PORes MyHub:=MyHub, arrOff1Data:=arrOff1Data, arrOff2Data:=arrOff2Data, arrOff3Data:=arrOff3Data _
    , arrOff4Data:=arrOff4Data, arrOff5Data:=arrOff5Data, arrOff6Data:=arrOff6Data _
    , arrOff7Data:=arrOff7Data, arrOff8Data:=arrOff8Data, arrOff9Data:=arrOff9Data _
    , arrOff10Data:=arrOff10Data, arrOff11Data:=arrOff11Data, arrOff12Data:=arrOff12Data _
    , arrOff13Data:=arrOff13Data, arrOff14Data:=arrOff14Data, arrOff15Data:=arrOff15Data _
    , arrOff16Data:=arrOff16Data, arrOff17Data:=arrOff17Data, arrOff18Data:=arrOff18Data _
    , arrOff19Data:=arrOff19Data, arrOff20Data:=arrOff20Data, arrOff21Data:=arrOff21Data _
    , arrOff22Data:=arrOff22Data, arrOff23Data:=arrOff23Data, arrOff24Data:=arrOff24Data
    
End Sub
    
Sub PORes(MyHub As Variant, arrOff1Data() As Double, arrOff2Data() As Double, arrOff3Data() As Double, arrOff4Data() As Double, arrOff5Data() As Double _
, arrOff6Data() As Double, arrOff7Data() As Double, arrOff8Data() As Double, arrOff9Data() As Double _
, arrOff10Data() As Double, arrOff11Data() As Double, arrOff12Data() As Double, arrOff13Data() As Double _
, arrOff14Data() As Double, arrOff15Data() As Double, arrOff16Data() As Double, arrOff17Data() As Double _
, arrOff18Data() As Double, arrOff19Data() As Double, arrOff20Data() As Double, arrOff21Data() As Double _
, arrOff22Data() As Double, arrOff23Data() As Double, arrOff24Data() As Double)
    
  'Retrieve on Excel
   Dim WS As Worksheet
   Dim MyRowsOffData As Long, MyOffset As Long, MyOffsetBid As Long
   
   MyRowsOffData = 0
   'Store MyOffsetBid for the end
   MyOffsetBid = ThisWorkbook.Worksheets("DashBoard").Range("FD13").Value
   MyOffset = ThisWorkbook.Worksheets("DashBoard").Range("FD13").Value
   
   Set WS = ThisWorkbook.Worksheets(MyHub)
   
   
   '****************************MyPublicOffsRetrieve****************************'
   With WS
   
   '****************************Hour 1******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff1Data, 1)
   
   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 1
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff1Data
   'Ofertada (O)/Casada (C)
   Erase arrOff1Data

   '****************************Hour 2******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff2Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 2
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff2Data
   'Ofertada (O)/Casada (C)
   Erase arrOff2Data
   
   '****************************Hour 3******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff3Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 3
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff3Data
   'Ofertada (O)/Casada (C)
   Erase arrOff3Data
   
   '****************************Hour 4******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff4Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 4
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff4Data
   'Ofertada (O)/Casada (C)
   Erase arrOff4Data
   
   '****************************Hour 5******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff5Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 5
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff5Data
   'Ofertada (O)/Casada (C)
   Erase arrOff5Data
   
   '****************************Hour 6******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff6Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 6
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff6Data
   'Ofertada (O)/Casada (C)
   Erase arrOff6Data
     
   '****************************Hour 7******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff7Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 7
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff7Data
   'Ofertada (O)/Casada (C)
   Erase arrOff7Data
   
   '****************************Hour 8******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff8Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 8
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff8Data
   'Ofertada (O)/Casada (C)
   Erase arrOff8Data
   
   '****************************Hour 9******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff9Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 9
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff9Data
   'Ofertada (O)/Casada (C)
   Erase arrOff9Data
   
   '****************************Hour 10******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff10Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 10
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff10Data
   'Ofertada (O)/Casada (C)
   Erase arrOff10Data
   
   '****************************Hour 11******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff11Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 11
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff11Data
   'Ofertada (O)/Casada (C)
   Erase arrOff11Data
   
   '****************************Hour 12******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff12Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 12
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff12Data
   'Ofertada (O)/Casada (C)
   Erase arrOff12Data
   
   '****************************Hour 13******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff13Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 13
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff13Data
   'Ofertada (O)/Casada (C)
   Erase arrOff13Data
   
   '****************************Hour 14******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff14Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 14
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff14Data
   'Ofertada (O)/Casada (C)
   Erase arrOff14Data
   
   '****************************Hour 15******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff15Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 15
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff15Data
   'Ofertada (O)/Casada (C)
   Erase arrOff15Data
   
   '****************************Hour 16******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff16Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 16
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff16Data
   'Ofertada (O)/Casada (C)
   Erase arrOff16Data
   
   '****************************Hour 17******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff17Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 17
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff17Data
   'Ofertada (O)/Casada (C)
   Erase arrOff17Data
   
   '****************************Hour 18******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff18Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 18
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff18Data
   'Ofertada (O)/Casada (C)
   Erase arrOff18Data
   
   '****************************Hour 19******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff19Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 19
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff19Data
   'Ofertada (O)/Casada (C)
   Erase arrOff19Data
   
   '****************************Hour 20******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff20Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 20
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff20Data
   'Ofertada (O)/Casada (C)
   Erase arrOff20Data
   
   '****************************Hour 21******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff21Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 21
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff21Data
   'Ofertada (O)/Casada (C)
   Erase arrOff21Data
   
   '****************************Hour 22******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff22Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 22
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff22Data
   'Ofertada (O)/Casada (C)
   Erase arrOff22Data
   
   '****************************Hour 23******************************************'
   
    MyOffset = MyOffset + MyRowsOffData
    MyRowsOffData = UBound(arrOff23Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 23
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff23Data
   'Ofertada (O)/Casada (C)
   Erase arrOff23Data
   
   '****************************Hour 24******************************************'
   
   MyOffset = MyOffset + MyRowsOffData
   MyRowsOffData = UBound(arrOff24Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsOffData, 1).Value = 24
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsOffData, 2).Value = arrOff24Data
   'Ofertada (O)/Casada (C)
   Erase arrOff24Data
   
   
   'MyTotalRows
    MyOffset = MyOffset + MyRowsOffData
    'Store FilteredRecords
    If MyHub = "ES" Then
    ThisWorkbook.Worksheets("DashBoard").Range("FD15").Value = MyOffset
    ElseIf MyHub = "PT" Then
    ThisWorkbook.Worksheets("DashBoard").Range("FD16").Value = MyOffset
    End If
    
   'Fecha label
   .Range("B4").Offset(MyOffsetBid, 0).Resize(MyOffset - MyOffsetBid, 1).Value = _
    ThisWorkbook.Worksheets("Notepad").Range("B4").Value
   'Pais label
   .Range("C4").Offset(MyOffsetBid, 0).Resize(MyOffset - MyOffsetBid, 1).Value = _
    ThisWorkbook.Worksheets("Notepad").Range("C4").Value
   'Unidad
   'Tipo Oferta
   .Range("E4").Offset(MyOffsetBid, 0).Resize(MyOffset - MyOffsetBid, 1).Value = "V"
   'Energía Compra/Venta and Precio Compra/Venta
   'Ofertada (O)/Casada (C)
   .Range("H4").Offset(MyOffsetBid, 0).Resize(MyOffset - MyOffsetBid, 1).Value = "O"
   End With
   Set WS = Nothing
    
    
    
End Sub










