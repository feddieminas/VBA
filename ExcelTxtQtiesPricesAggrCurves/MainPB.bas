Attribute VB_Name = "MainPB"
Option Explicit
Option Base 1

Sub PBLayout(MyHub As Variant, arrBid1DataTmp() As Double, arrBid2DataTmp() As Double, arrBid3DataTmp() As Double, arrBid4DataTmp() As Double, arrBid5DataTmp() As Double _
, arrBid6DataTmp() As Double, arrBid7DataTmp() As Double, arrBid8DataTmp() As Double, arrBid9DataTmp() As Double _
, arrBid10DataTmp() As Double, arrBid11DataTmp() As Double, arrBid12DataTmp() As Double, arrBid13DataTmp() As Double _
, arrBid14DataTmp() As Double, arrBid15DataTmp() As Double, arrBid16DataTmp() As Double, arrBid17DataTmp() As Double _
, arrBid18DataTmp() As Double, arrBid19DataTmp() As Double, arrBid20DataTmp() As Double, arrBid21DataTmp() As Double _
, arrBid22DataTmp() As Double, arrBid23DataTmp() As Double, arrBid24DataTmp() As Double)
    
'****Sort out the Distinct Bid PricesBO and Accumulate the Quantities Bid on that same PriceBO****'
Dim Filtcounter As Long, i As Long, j As Long
Dim MyRowData As Long, MyColData As Long
    
'MyBidData
Dim arrBid1Data() As Double, arrBid2Data() As Double, arrBid3Data() As Double, arrBid4Data() As Double
Dim arrBid5Data() As Double, arrBid6Data() As Double, arrBid7Data() As Double, arrBid8Data() As Double
Dim arrBid9Data() As Double, arrBid10Data() As Double, arrBid11Data() As Double, arrBid12Data() As Double
Dim arrBid13Data() As Double, arrBid14Data() As Double, arrBid15Data() As Double, arrBid16Data() As Double
Dim arrBid17Data() As Double, arrBid18Data() As Double, arrBid19Data() As Double, arrBid20Data() As Double
Dim arrBid21Data() As Double, arrBid22Data() As Double, arrBid23Data() As Double, arrBid24Data() As Double
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid1DataTmp, 1)
    MyColData = UBound(arrBid1DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid1DataTmp(i + 1, 2) <> arrBid1DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid1Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid1Data(i, j) = arrBid1DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid1DataTmp(i, 2) <> arrBid1DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid1Data(Filtcounter, j) = arrBid1DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid1DataTmp(i, 2) = arrBid1DataTmp(i - 1, 2) Then
    arrBid1Data(Filtcounter, 1) = arrBid1Data(Filtcounter, 1) + arrBid1DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid1DataTmp
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
    MyRowData = UBound(arrBid2DataTmp, 1)
    MyColData = UBound(arrBid2DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid2DataTmp(i + 1, 2) <> arrBid2DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid2Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid2Data(i, j) = arrBid2DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid2DataTmp(i, 2) <> arrBid2DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid2Data(Filtcounter, j) = arrBid2DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid2DataTmp(i, 2) = arrBid2DataTmp(i - 1, 2) Then
    arrBid2Data(Filtcounter, 1) = arrBid2Data(Filtcounter, 1) + arrBid2DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid2DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid3DataTmp, 1)
    MyColData = UBound(arrBid3DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid3DataTmp(i + 1, 2) <> arrBid3DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid3Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid3Data(i, j) = arrBid3DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid3DataTmp(i, 2) <> arrBid3DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid3Data(Filtcounter, j) = arrBid3DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid3DataTmp(i, 2) = arrBid3DataTmp(i - 1, 2) Then
    arrBid3Data(Filtcounter, 1) = arrBid3Data(Filtcounter, 1) + arrBid3DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid3DataTmp
      
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid4DataTmp, 1)
    MyColData = UBound(arrBid4DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid4DataTmp(i + 1, 2) <> arrBid4DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid4Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid4Data(i, j) = arrBid4DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid4DataTmp(i, 2) <> arrBid4DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid4Data(Filtcounter, j) = arrBid4DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid4DataTmp(i, 2) = arrBid4DataTmp(i - 1, 2) Then
    arrBid4Data(Filtcounter, 1) = arrBid4Data(Filtcounter, 1) + arrBid4DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid4DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid5DataTmp, 1)
    MyColData = UBound(arrBid5DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid5DataTmp(i + 1, 2) <> arrBid5DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid5Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid5Data(i, j) = arrBid5DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid5DataTmp(i, 2) <> arrBid5DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid5Data(Filtcounter, j) = arrBid5DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid5DataTmp(i, 2) = arrBid5DataTmp(i - 1, 2) Then
    arrBid5Data(Filtcounter, 1) = arrBid5Data(Filtcounter, 1) + arrBid5DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid5DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid6DataTmp, 1)
    MyColData = UBound(arrBid6DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid6DataTmp(i + 1, 2) <> arrBid6DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid6Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid6Data(i, j) = arrBid6DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid6DataTmp(i, 2) <> arrBid6DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid6Data(Filtcounter, j) = arrBid6DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid6DataTmp(i, 2) = arrBid6DataTmp(i - 1, 2) Then
    arrBid6Data(Filtcounter, 1) = arrBid6Data(Filtcounter, 1) + arrBid6DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid6DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid7DataTmp, 1)
    MyColData = UBound(arrBid7DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid7DataTmp(i + 1, 2) <> arrBid7DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid7Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid7Data(i, j) = arrBid7DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid7DataTmp(i, 2) <> arrBid7DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid7Data(Filtcounter, j) = arrBid7DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid7DataTmp(i, 2) = arrBid7DataTmp(i - 1, 2) Then
    arrBid7Data(Filtcounter, 1) = arrBid7Data(Filtcounter, 1) + arrBid7DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid7DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid8DataTmp, 1)
    MyColData = UBound(arrBid8DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid8DataTmp(i + 1, 2) <> arrBid8DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid8Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid8Data(i, j) = arrBid8DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid8DataTmp(i, 2) <> arrBid8DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid8Data(Filtcounter, j) = arrBid8DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid8DataTmp(i, 2) = arrBid8DataTmp(i - 1, 2) Then
    arrBid8Data(Filtcounter, 1) = arrBid8Data(Filtcounter, 1) + arrBid8DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid8DataTmp
      
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid9DataTmp, 1)
    MyColData = UBound(arrBid9DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid9DataTmp(i + 1, 2) <> arrBid9DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid9Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid9Data(i, j) = arrBid9DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid9DataTmp(i, 2) <> arrBid9DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid9Data(Filtcounter, j) = arrBid9DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid9DataTmp(i, 2) = arrBid9DataTmp(i - 1, 2) Then
    arrBid9Data(Filtcounter, 1) = arrBid9Data(Filtcounter, 1) + arrBid9DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid9DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid10DataTmp, 1)
    MyColData = UBound(arrBid10DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid10DataTmp(i + 1, 2) <> arrBid10DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid10Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid10Data(i, j) = arrBid10DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid10DataTmp(i, 2) <> arrBid10DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid10Data(Filtcounter, j) = arrBid10DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid10DataTmp(i, 2) = arrBid10DataTmp(i - 1, 2) Then
    arrBid10Data(Filtcounter, 1) = arrBid10Data(Filtcounter, 1) + arrBid10DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid10DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid11DataTmp, 1)
    MyColData = UBound(arrBid11DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid11DataTmp(i + 1, 2) <> arrBid11DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid11Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid11Data(i, j) = arrBid11DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid11DataTmp(i, 2) <> arrBid11DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid11Data(Filtcounter, j) = arrBid11DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid11DataTmp(i, 2) = arrBid11DataTmp(i - 1, 2) Then
    arrBid11Data(Filtcounter, 1) = arrBid11Data(Filtcounter, 1) + arrBid11DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid11DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid12DataTmp, 1)
    MyColData = UBound(arrBid12DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid12DataTmp(i + 1, 2) <> arrBid12DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid12Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid12Data(i, j) = arrBid12DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid12DataTmp(i, 2) <> arrBid12DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid12Data(Filtcounter, j) = arrBid12DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid12DataTmp(i, 2) = arrBid12DataTmp(i - 1, 2) Then
    arrBid12Data(Filtcounter, 1) = arrBid12Data(Filtcounter, 1) + arrBid12DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid12DataTmp
      
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid13DataTmp, 1)
    MyColData = UBound(arrBid13DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid13DataTmp(i + 1, 2) <> arrBid13DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid13Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid13Data(i, j) = arrBid13DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid13DataTmp(i, 2) <> arrBid13DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid13Data(Filtcounter, j) = arrBid13DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid13DataTmp(i, 2) = arrBid13DataTmp(i - 1, 2) Then
    arrBid13Data(Filtcounter, 1) = arrBid13Data(Filtcounter, 1) + arrBid13DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid13DataTmp
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid14DataTmp, 1)
    MyColData = UBound(arrBid14DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid14DataTmp(i + 1, 2) <> arrBid14DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid14Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid14Data(i, j) = arrBid14DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid14DataTmp(i, 2) <> arrBid14DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid14Data(Filtcounter, j) = arrBid14DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid14DataTmp(i, 2) = arrBid14DataTmp(i - 1, 2) Then
    arrBid14Data(Filtcounter, 1) = arrBid14Data(Filtcounter, 1) + arrBid14DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid14DataTmp
  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid15DataTmp, 1)
    MyColData = UBound(arrBid15DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid15DataTmp(i + 1, 2) <> arrBid15DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid15Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid15Data(i, j) = arrBid15DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid15DataTmp(i, 2) <> arrBid15DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid15Data(Filtcounter, j) = arrBid15DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid15DataTmp(i, 2) = arrBid15DataTmp(i - 1, 2) Then
    arrBid15Data(Filtcounter, 1) = arrBid15Data(Filtcounter, 1) + arrBid15DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid15DataTmp
      
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid16DataTmp, 1)
    MyColData = UBound(arrBid16DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid16DataTmp(i + 1, 2) <> arrBid16DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid16Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid16Data(i, j) = arrBid16DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid16DataTmp(i, 2) <> arrBid16DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid16Data(Filtcounter, j) = arrBid16DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid16DataTmp(i, 2) = arrBid16DataTmp(i - 1, 2) Then
    arrBid16Data(Filtcounter, 1) = arrBid16Data(Filtcounter, 1) + arrBid16DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid16DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid17DataTmp, 1)
    MyColData = UBound(arrBid17DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid17DataTmp(i + 1, 2) <> arrBid17DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid17Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid17Data(i, j) = arrBid17DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid17DataTmp(i, 2) <> arrBid17DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid17Data(Filtcounter, j) = arrBid17DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid17DataTmp(i, 2) = arrBid17DataTmp(i - 1, 2) Then
    arrBid17Data(Filtcounter, 1) = arrBid17Data(Filtcounter, 1) + arrBid17DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid17DataTmp
     
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid18DataTmp, 1)
    MyColData = UBound(arrBid18DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid18DataTmp(i + 1, 2) <> arrBid18DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid18Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid18Data(i, j) = arrBid18DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid18DataTmp(i, 2) <> arrBid18DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid18Data(Filtcounter, j) = arrBid18DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid18DataTmp(i, 2) = arrBid18DataTmp(i - 1, 2) Then
    arrBid18Data(Filtcounter, 1) = arrBid18Data(Filtcounter, 1) + arrBid18DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid18DataTmp
       
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid19DataTmp, 1)
    MyColData = UBound(arrBid19DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid19DataTmp(i + 1, 2) <> arrBid19DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid19Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid19Data(i, j) = arrBid19DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid19DataTmp(i, 2) <> arrBid19DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid19Data(Filtcounter, j) = arrBid19DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid19DataTmp(i, 2) = arrBid19DataTmp(i - 1, 2) Then
    arrBid19Data(Filtcounter, 1) = arrBid19Data(Filtcounter, 1) + arrBid19DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid19DataTmp
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
    MyRowData = UBound(arrBid20DataTmp, 1)
    MyColData = UBound(arrBid20DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid20DataTmp(i + 1, 2) <> arrBid20DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid20Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid20Data(i, j) = arrBid20DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid20DataTmp(i, 2) <> arrBid20DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid20Data(Filtcounter, j) = arrBid20DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid20DataTmp(i, 2) = arrBid20DataTmp(i - 1, 2) Then
    arrBid20Data(Filtcounter, 1) = arrBid20Data(Filtcounter, 1) + arrBid20DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid20DataTmp
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyRowData = UBound(arrBid21DataTmp, 1)
    MyColData = UBound(arrBid21DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid21DataTmp(i + 1, 2) <> arrBid21DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid21Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid21Data(i, j) = arrBid21DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid21DataTmp(i, 2) <> arrBid21DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid21Data(Filtcounter, j) = arrBid21DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid21DataTmp(i, 2) = arrBid21DataTmp(i - 1, 2) Then
    arrBid21Data(Filtcounter, 1) = arrBid21Data(Filtcounter, 1) + arrBid21DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid21DataTmp
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid22DataTmp, 1)
    MyColData = UBound(arrBid22DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid22DataTmp(i + 1, 2) <> arrBid22DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid22Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid22Data(i, j) = arrBid22DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid22DataTmp(i, 2) <> arrBid22DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid22Data(Filtcounter, j) = arrBid22DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid22DataTmp(i, 2) = arrBid22DataTmp(i - 1, 2) Then
    arrBid22Data(Filtcounter, 1) = arrBid22Data(Filtcounter, 1) + arrBid22DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid22DataTmp
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid23DataTmp, 1)
    MyColData = UBound(arrBid23DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid23DataTmp(i + 1, 2) <> arrBid23DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid23Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData
    
    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid23Data(i, j) = arrBid23DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid23DataTmp(i, 2) <> arrBid23DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid23Data(Filtcounter, j) = arrBid23DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid23DataTmp(i, 2) = arrBid23DataTmp(i - 1, 2) Then
    arrBid23Data(Filtcounter, 1) = arrBid23Data(Filtcounter, 1) + arrBid23DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid23DataTmp
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    MyRowData = UBound(arrBid24DataTmp, 1)
    MyColData = UBound(arrBid24DataTmp, 2)
    
    Filtcounter = 1
    'Count How Many DistinctBids do we have
    For i = 1 To MyRowData - 1
    If arrBid24DataTmp(i + 1, 2) <> arrBid24DataTmp(i, 2) Then Filtcounter = Filtcounter + 1
    Next i
    
    'Redefine your Bid arrays
    ReDim arrBid24Data(1 To Filtcounter, 1 To 2)

    'Insert values on the BidArray
    Filtcounter = 1
    For i = 1 To MyRowData

    'If is the Very First Row of the Array just copy the whole row to the new created array
    If i = 1 Then
    For j = 1 To MyColData
    arrBid24Data(i, j) = arrBid24DataTmp(i, j)
    Next j
    
    'If from the second row there are no duplicate prices then copy the whole row to an added row of
    'the new created array
    ElseIf arrBid24DataTmp(i, 2) <> arrBid24DataTmp(i - 1, 2) Then
    Filtcounter = Filtcounter + 1
    For j = 1 To MyColData
    arrBid24Data(Filtcounter, j) = arrBid24DataTmp(i, j)
    Next j
    
    'If there are duplicate price values then accumulate on the same price the relative quantities
    ElseIf arrBid24DataTmp(i, 2) = arrBid24DataTmp(i - 1, 2) Then
    arrBid24Data(Filtcounter, 1) = arrBid24Data(Filtcounter, 1) + arrBid24DataTmp(i, 1)
    End If
    
    Next i
    
    Erase arrBid24DataTmp
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'PB Results
    PBRes MyHub:=MyHub, arrBid1Data:=arrBid1Data, arrBid2Data:=arrBid2Data, arrBid3Data:=arrBid3Data _
    , arrBid4Data:=arrBid4Data, arrBid5Data:=arrBid5Data, arrBid6Data:=arrBid6Data _
    , arrBid7Data:=arrBid7Data, arrBid8Data:=arrBid8Data, arrBid9Data:=arrBid9Data _
    , arrBid10Data:=arrBid10Data, arrBid11Data:=arrBid11Data, arrBid12Data:=arrBid12Data _
    , arrBid13Data:=arrBid13Data, arrBid14Data:=arrBid14Data, arrBid15Data:=arrBid15Data _
    , arrBid16Data:=arrBid16Data, arrBid17Data:=arrBid17Data, arrBid18Data:=arrBid18Data _
    , arrBid19Data:=arrBid19Data, arrBid20Data:=arrBid20Data, arrBid21Data:=arrBid21Data _
    , arrBid22Data:=arrBid22Data, arrBid23Data:=arrBid23Data, arrBid24Data:=arrBid24Data
    
End Sub

Sub PBRes(MyHub As Variant, arrBid1Data() As Double, arrBid2Data() As Double, arrBid3Data() As Double, arrBid4Data() As Double, arrBid5Data() As Double _
, arrBid6Data() As Double, arrBid7Data() As Double, arrBid8Data() As Double, arrBid9Data() As Double _
, arrBid10Data() As Double, arrBid11Data() As Double, arrBid12Data() As Double, arrBid13Data() As Double _
, arrBid14Data() As Double, arrBid15Data() As Double, arrBid16Data() As Double, arrBid17Data() As Double _
, arrBid18Data() As Double, arrBid19Data() As Double, arrBid20Data() As Double, arrBid21Data() As Double _
, arrBid22Data() As Double, arrBid23Data() As Double, arrBid24Data() As Double)
    
   'Retrieve on Excel
   Dim WS As Worksheet
   Dim MyRowsBidData As Long, MyOffset As Long
   MyRowsBidData = 0
   MyOffset = 0
   
   Set WS = ThisWorkbook.Worksheets(MyHub)
   
   
   '****************************MyPublicBidsRetrieve****************************'
   With WS
   
   '****************************Hour 1******************************************'
   
   MyRowsBidData = UBound(arrBid1Data, 1)
   
   'Hora
   .Range("A4").Resize(MyRowsBidData, 1).Value = 1
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Resize(MyRowsBidData, 2).Value = arrBid1Data
   'Ofertada (O)/Casada (C)
   Erase arrBid1Data

   '****************************Hour 2******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid2Data, 1)
   
   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 2
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid2Data
   'Ofertada (O)/Casada (C)
   Erase arrBid2Data
   
   '****************************Hour 3******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid3Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 3
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid3Data
   'Ofertada (O)/Casada (C)
   Erase arrBid3Data
   
   '****************************Hour 4******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid4Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 4
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid4Data
   'Ofertada (O)/Casada (C)
   Erase arrBid4Data
   
   '****************************Hour 5******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid5Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 5
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid5Data
   'Ofertada (O)/Casada (C)
   Erase arrBid5Data
   
   '****************************Hour 6******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid6Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 6
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid6Data
   'Ofertada (O)/Casada (C)
   Erase arrBid6Data
     
   '****************************Hour 7******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid7Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 7
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid7Data
   'Ofertada (O)/Casada (C)
   Erase arrBid7Data
   
   '****************************Hour 8******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid8Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 8
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid8Data
   'Ofertada (O)/Casada (C)
   Erase arrBid8Data
   
   '****************************Hour 9******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid9Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 9
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid9Data
   'Ofertada (O)/Casada (C)
   Erase arrBid9Data
   
   '****************************Hour 10******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid10Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 10
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid10Data
   'Ofertada (O)/Casada (C)
   Erase arrBid10Data
   
   '****************************Hour 11******************************************'
    
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid11Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 11
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid11Data
   'Ofertada (O)/Casada (C)
   Erase arrBid11Data
   
   '****************************Hour 12******************************************'
    
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid12Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 12
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid12Data
   'Ofertada (O)/Casada (C)
   Erase arrBid12Data
   
   '****************************Hour 13******************************************'
    
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid13Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 13
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid13Data
   'Ofertada (O)/Casada (C)
   Erase arrBid13Data
   
   '****************************Hour 14******************************************'
   
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid14Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 14
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid14Data
   'Ofertada (O)/Casada (C)
   Erase arrBid14Data
   
   '****************************Hour 15******************************************'
    
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid15Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 15
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid15Data
   'Ofertada (O)/Casada (C)
   Erase arrBid15Data
   
   '****************************Hour 16******************************************'
    
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid16Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 16
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid16Data
   'Ofertada (O)/Casada (C)
   Erase arrBid16Data
   
   '****************************Hour 17******************************************'
    
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid17Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 17
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid17Data
   'Ofertada (O)/Casada (C)
   Erase arrBid17Data
   
   '****************************Hour 18******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid18Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 18
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid18Data
   'Ofertada (O)/Casada (C)
   Erase arrBid18Data
   
   '****************************Hour 19******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid19Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 19
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid19Data
   'Ofertada (O)/Casada (C)
   Erase arrBid19Data
   
   '****************************Hour 20******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid20Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 20
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid20Data
   'Ofertada (O)/Casada (C)
   Erase arrBid20Data
   
   '****************************Hour 21******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid21Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 21
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid21Data
   'Ofertada (O)/Casada (C)
   Erase arrBid21Data
   
   '****************************Hour 22******************************************'
   
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid22Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 22
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid22Data
   'Ofertada (O)/Casada (C)
   Erase arrBid22Data
   
   '****************************Hour 23******************************************'
    
    MyOffset = MyOffset + MyRowsBidData
    MyRowsBidData = UBound(arrBid23Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 23
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid23Data
   'Ofertada (O)/Casada (C)
   Erase arrBid23Data
   
   '****************************Hour 24******************************************'
   
   MyOffset = MyOffset + MyRowsBidData
   MyRowsBidData = UBound(arrBid24Data, 1)

   'Hora
   .Range("A4").Offset(MyOffset, 0).Resize(MyRowsBidData, 1).Value = 24
   'Fecha
   'Pais
   'Unidad
   'Tipo Oferta
   'Energía Compra/Venta and Precio Compra/Venta
   .Range("F4").Offset(MyOffset, 0).Resize(MyRowsBidData, 2).Value = arrBid24Data
   'Ofertada (O)/Casada (C)
   Erase arrBid24Data
   
   'MyTotalBidRows
    MyOffset = MyOffset + MyRowsBidData
    'ThisWorkbook.Worksheets("DashBoard").Range("FD14").Value = MyOffset
    
   'Fecha label
   .Range("B4").Resize(MyOffset, 1).Value = ThisWorkbook.Worksheets("Notepad").Range("B4").Value
   'Pais label
   .Range("C4").Resize(MyOffset, 1).Value = ThisWorkbook.Worksheets("Notepad").Range("C4").Value
   'Unidad
   'Tipo Oferta
   .Range("E4").Resize(MyOffset, 1).Value = "C"
   'Energía Compra/Venta and Precio Compra/Venta
   'Ofertada (O)/Casada (C)
   .Range("H4").Resize(MyOffset, 1).Value = "O"
   End With
     
   'MyTotal BidArray Offset
   ThisWorkbook.Worksheets("DashBoard").Range("FD13").Value = MyOffset
   Set WS = Nothing
    
 End Sub

