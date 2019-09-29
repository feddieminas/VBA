Attribute VB_Name = "modFunctions"
Option Explicit

' Worked in Windows 10 and Excel 2016

Function MyDate(MyChoice As String)
MyDate = DateSerial(CInt(Right(CStr(MyChoice), 4)), CInt(Mid(CStr(MyChoice), 4, 2)), CInt(Left(CStr(MyChoice), 2)))
End Function

Function FileFolderExists(strFullPath As String) As Boolean
    On Error GoTo EarlyExit
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
EarlyExit:
    On Error GoTo 0
End Function

Function MyItalianMonth(MyMonth As Long) As String

Select Case MyMonth

Case Is = 1
MyItalianMonth = "gennaio"
Case Is = 2
MyItalianMonth = "febbraio"
Case Is = 3
MyItalianMonth = "marzo"
Case Is = 4
MyItalianMonth = "aprile"
Case Is = 5
MyItalianMonth = "maggio"
Case Is = 6
MyItalianMonth = "giugno"
Case Is = 7
MyItalianMonth = "luglio"
Case Is = 8
MyItalianMonth = "agosto"
Case Is = 9
MyItalianMonth = "settembre"
Case Is = 10
MyItalianMonth = "ottobre"
Case Is = 11
MyItalianMonth = "novembre"
Case Is = 12
MyItalianMonth = "dicembre"

End Select

End Function

Function MyWeekDay(MyDate As Double) As String

Dim FullWD As Long
FullWD = Weekday(MyDate, vbMonday)
Select Case FullWD
Case Is = 1
MyWeekDay = "M"
Case Is = 2
MyWeekDay = "T"
Case Is = 3
MyWeekDay = "W"
Case Is = 4
MyWeekDay = "T"
Case Is = 5
MyWeekDay = "F"
Case Is = 6
MyWeekDay = "S"
Case Is = 7
MyWeekDay = "S"
End Select

End Function

Function MyDay(MyDate As Double) As String
MyDay = Format(MyDate, "DD")
End Function

Function CountFilesInFolder(strDir As String, Optional strType As String) As Integer ' By Ryan Wells (wellsr.com)
    Dim file As Variant, i As Integer
    i = 0
    If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
    file = Dir(strDir & strType)
    While (file <> "")
        i = i + 1
        file = Dir
    Wend
    CountFilesInFolder = i
End Function






