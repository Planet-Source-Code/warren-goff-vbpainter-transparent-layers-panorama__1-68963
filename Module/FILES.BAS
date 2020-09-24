Attribute VB_Name = "Module2"
'Windows API functions
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function DiskSpaceFree Lib "SETUPKIT.DLL" () As Long
Function CheckFileThere(PathName As String, filename As String) As Integer
    
    ' Sub/Function Name       : CheckFileThere
    ' Purpose                 : Check if a file is on disk
    ' Parameters              : PathName (Path were file is located)
    '                           FileName (Filename of file)
    ' Returns                 : True = File is there, False = File is not there
    ' Created by              : Paul Treffers
    ' Date Created            : 14/09/94

    Dim CheckThere As Integer
    Dim SearchFile As String
    
    
    PathName = UCase$(Trim$(PathName))
    filename = UCase$(Trim$(filename))
    If Right$(PathName, 1) <> "\MicrophoneData\" Then
        PathName = PathName & "\MicrophoneData\"
    End If
    SearchFile = PathName & filename
    On Error Resume Next
    If UCase$(Trim$(Dir(SearchFile, 0))) = filename Then
        If Err = 76 Then
            CheckThere = False
        Else
            CheckThere = True
        End If
    Else
        CheckThere = False
    End If
    CheckFileThere = CheckThere
End Function

Function GetDirName(ScanString$) As String
    
    ' Sub/Function Name       : GetDirName
    ' Purpose                 : Gets a full directory from a string (with filename)
    ' Parameters              : String to scan for directory name
    ' Return                  : Directory Path
    ' Created by              : Paul Treffers
    ' Date Created            : 19/11/94
    
    Dim ExitWhile As Integer
    Dim Pos%, PosSave%
    ExitWhile = True
    Pos% = 1
    Do While ExitWhile = True
        Pos% = InStr(Pos%, ScanString$, "\MicrophoneData\")
        If Pos% = 0 Then
            Exit Do
        Else
            Pos% = Pos% + 1
            PosSave% = Pos% - 1
        End If
    Loop
    GetDirName = Left$(ScanString$, PosSave%)
End Function

Function GetFilename(Path As String, NoExt As Boolean) As String
    ' Sub/Function Name       : GetFileName
    ' Purpose                 : Gets a filename from string that contains directory also
    ' Parameters              : String to scan for filename
    ' Return                  : Directory Path
    ' Created by              : Paul Treffers
    ' Date Created            : 19/11/94
    
    ExitWhile = True
    Pos1% = 1
    Do While ExitWhile = True
        Pos1% = InStr(Pos1%, ScanString$, "\")
        If Pos1% = 0 Then
            Exit Do
        Else
            Pos1% = Pos1% + 1
            PosSave% = Pos1% - 1
        End If
    Loop
    GetFilename = Trim$(Mid$(ScanString$, PosSave% + 1, Len(ScanString$)))

End Function

Function GetFreeDiskSpace(Drive As String) As Long
    
    ' Sub/Function Name       : GetFreeDiskSpace
    ' Purpose                 : Determine Free Space on Disk
    ' Parameters              : Drive to Check
    ' Returns                 : Free Disk Space
    ' Created by              : Paul Treffers
    ' Date Created            : 14/09/94
    
    On Error GoTo FErrorHandler
    Dim TempDrive As String
    Dim XValue As Long
    Dim DirTest As String
    TempDrive = Left$(CurDir$, 2)
    ChDrive Drive
    DirTest = Dir$(Drive & "\*.*")
    XValue = DiskSpaceFree&()
    ChDrive TempDrive
    GetFreeDiskSpace = XValue
    Exit Function

FErrorHandler:
    
    GetFreeDiskSpace = -Err
    On Error GoTo 0
    ChDrive TempDrive
    Err = 0
    Exit Function

End Function

Function GetWindowsDir() As String
    ' Sub/Function Name       : GetWindowsDir
    ' Purpose                 : Returns Windows Directory
    ' Parameters              : None
    ' Returns                 : Windows Directore
    ' Created by              : Paul Treffers
    ' Date Created            : 14/09/94
    
    Dim WinDir As String
    Dim File As String
    Dim Res As Integer
    WinDir = Space$(20)
    Res = GetWindowsDirectory(WinDir, 20)
    File = Left$(WinDir, InStr(1, WinDir, Chr$(0)) - 1)
    GetWindowsDir = Trim$(File) & "\MicrophoneData\"
End Function

