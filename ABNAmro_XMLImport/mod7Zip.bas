Attribute VB_Name = "mod7Zip"
#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#End If


Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103


Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub



'With this example you browse to the zip or 7z file you want to unzip
'The zip file will be unzipped in a new folder in: Application.DefaultFilePath
'Normal if you have not change it this will be your Documents folder
'The name of the folder that the code create in this folder is the Date/Time
'You can change this folder to this if you want to use a fixed folder:
'NameUnZipFolder = "C:\Users\Ron\TestFolder\"
'Read the comments in the code about the commands/Switches in the ShellStr
'There is no need to change the code before you test it

Sub A_UnZip_Zip_File_Browse()
    Dim PathZipProgram As String, NameUnzipFolder As String
    Dim FileNameZip As Variant, ShellStr As String

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        MsgBox "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    'Create path and name of the normal folder to unzip the files in
    'In this example we use: Application.DefaultFilePath
    'Normal if you have not change it this will be your Documents folder
    'The name of the folder that the code create in this folder is the Date/Time
    NameUnzipFolder = Application.DefaultFilePath & "\" & Format(Now, "yyyy-mm-dd h-mm-ss")
    'You can also use a fixed path like
    'NameUnZipFolder = "C:\Users\Ron\TestFolder"

    'Select the zip file (.zip or .7z files)
    FileNameZip = Application.GetOpenFilename(filefilter:="Zip Files, *.zip, 7z Files, *.7z", _
                                              MultiSelect:=False, Title:="Select the file that you want to unzip")

    'Unzip the files/folders from the zip file in the NameUnZipFolder folder
    If FileNameZip = False Then
        'do nothing
    Else
        'There are a few commands/Switches that you can change in the ShellStr
        'We use x command now to keep the folder stucture, replace it with e if you want only the files
        '-aoa Overwrite All existing files without prompt.
        '-aos Skip extracting of existing files.
        '-aou aUto rename extracting file (for example, name.txt will be renamed to name_1.txt).
        '-aot auto rename existing file (for example, name.txt will be renamed to name_1.txt).
        'Use -r if you also want to unzip the subfolders from the zip file
        'You can add -ppassword if you want to unzip a zip file with password (only 7zip files)
        'Change "*.*" to for example "*.txt" if you only want to unzip the txt files
        'Use "*.xl*" for all Excel files: xls, xlsx, xlsm, xlsb
        ShellStr = PathZipProgram & "7z.exe x -aoa -r" _
                 & " " & Chr(34) & FileNameZip & Chr(34) _
                 & " -o" & Chr(34) & NameUnzipFolder & Chr(34) & " " & "*.*"

        ShellAndWait ShellStr, vbHide
        MsgBox "Look in " & NameUnzipFolder & " for extracted files"

    End If
End Sub




'With this example you unzip a fixed zip file: FileNameZip = "C:\Users\Ron\Test.zip"
'Note this file must exist, this is the only thing that you must change before you test it
'The zip file will be unzipped in a new folder in: Application.DefaultFilePath
'Normal if you have not change it this will be your Documents folder
'The name of the folder that the code create in this folder is the Date/Time
'You can change this folder to this if you want to use a fixed folder:
'NameUnZipFolder = "C:\Users\Ron\TestFolder\"
'Read the comments in the code about the commands/Switches in the ShellStr

Sub MyUnzip7(FileNameZip As Variant, NameUnzipFolder As String)
    Dim PathZipProgram As String ', NameUnzipFolder As String
    Dim ShellStr As String  ',FileNameZip As Variant

    'Path of the Zip program
    PathZipProgram = "C:\program files\7-Zip\"
    If Right(PathZipProgram, 1) <> "\" Then
        PathZipProgram = PathZipProgram & "\"
    End If

    'Check if this is the path where 7z is installed.
    If Dir(PathZipProgram & "7z.exe") = "" Then
        MsgBox "Please find your copy of 7z.exe and try again"
        Exit Sub
    End If

    'Create path and name of the normal folder to unzip the files in
    'In this example we use: Application.DefaultFilePath
    'Normal if you have not change it this will be your Documents folder
    'The name of the folder that the code create in this folder is the Date/Time
    '******** NameUnzipFolder = Application.DefaultFilePath & "\" & Format(Now, "yyyy-mm-dd h-mm-ss")********original code
    'You can also use a fixed path like
    'NameUnZipFolder = "C:\Users\Ron\TestFolder\"

    'Name of the zip file that you want to unzip (.zip or .7z files)
    '******** FileNameZip = "data.zip on Users folder" ********original code

    'There are a few commands/Switches that you can change in the ShellStr
    'We use x command now to keep the folder stucture, replace it with e if you want only the files
    '-aoa Overwrite All existing files without prompt.
    '-aos Skip extracting of existing files.
    '-aou aUto rename extracting file (for example, name.txt will be renamed to name_1.txt).
    '-aot auto rename existing file (for example, name.txt will be renamed to name_1.txt).
    'Use -r if you also want to unzip the subfolders from the zip file
    'You can add -ppassword if you want to unzip a zip file with password (only .7z files)
    'Change "*.*" to for example "*.txt" if you only want to unzip the txt files
    'Use "*.xl*" for all Excel files: xls, xlsx, xlsm, xlsb
    ShellStr = PathZipProgram & "7z.exe x -aoa -r" _
             & " " & Chr(34) & FileNameZip & Chr(34) _
             & " -o" & Chr(34) & NameUnzipFolder & Chr(34) & " " & "*.*"

    ShellAndWait ShellStr, vbHide
  '  MsgBox "Look in " & NameUnzipFolder & " for extracted files"

End Sub


