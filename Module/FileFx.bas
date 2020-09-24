Attribute VB_Name = "FileFx"
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function SHFileOperation Lib _
    "shell32.dll" Alias "SHFileOperationA" _
    (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40

Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Global BaseHeight As Long
Global BaseWidth As Long




Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
    End Type
    
Public Function filename(fName As String) As String
Dim STime As String
Dim ST As Integer
T1 = Year(Now)
T2 = Month(Now)
If Len(T2) = 1 Then T2 = "0" & T2
T3 = Day(Now)
If Len(T3) = 1 Then T3 = "0" & T3
Tday = T1 & T2 & T3
STime = Time
If right(STime, 2) = "PM" And Val(left(STime, InStr(1, STime, ":"))) <> 12 Then
    ST = Str(Val(left(STime, InStr(1, STime, ":"))) + 12)
    STime = Replace(STime, "PM", "")
    STime = ST & Mid(STime, (InStr(1, STime, ":") + 1), (Len(STime) - InStr(1, STime, ":")))
Else
    STime = Replace(STime, "AM", "")
End If
If InStr(1, STime, ":") = 2 Then STime = "0" & STime
yy = Trim(Replace(STime, ":", ""))
filename = Tday & yy
'MsgBox FileName
End Function


Public Sub ShellDeleteOne(sFile As String, ActionFlag As Long)

Dim SHFileOp As SHFILEOPSTRUCT
Dim R As Long

sFile = sFile & Chr$(0)
'MsgBox sFile
With SHFileOp
  .wFunc = FO_DELETE
  .pFrom = sFile
  .fFlags = ActionFlag
End With

R = SHFileOperation(SHFileOp)

End Sub

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds


    Do Until GetTickCount > EndTime


        DoEvents
        Loop
    End Function

Public Function AddBackSlash(ByVal sPath As String) As String
    'Returns sPath with a trailing backslash
    '     if sPath does not
    'already have a trailing backslash. Othe
    '     rwise, returns sPath.
    sPath = Trim$(sPath)


    If Len(sPath) > 0 Then
        sPath = sPath & IIf(right$(sPath, 1) <> "\", "\", "")
    End If
    AddBackSlash = sPath
    
End Function


Public Function GetLongFilename(ByVal sShortFilename As String) As String
    'Returns the Long Filename associated wi
    '     th sShortFilename
    Dim lRet As Long
    Dim sLongFilename As String
    'First attempt using 1024 character buff
    '     er.
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    
    'If buffer is too small lRet contains bu
    '     ffer size needed.


    If lRet > Len(sLongFilename) Then
        'Increase buffer size...
        sLongFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    End If
    
    'lRet contains the number of characters
    '     returned.


    If lRet > 0 Then
        GetLongFilename = left$(sLongFilename, lRet)
    End If
    
End Function


Public Function GetShortFilename(ByVal sLongFilename As String) As String
    'Returns the Short Filename associated w
    '     ith sLongFilename
    Dim lRet As Long
    Dim sShortFilename As String
    'First attempt using 1024 character buff
    '     er.
    sShortFilename = String$(1024, " ")
    lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))
    
    'If buffer is too small lRet contains bu
    '     ffer size needed.


    If lRet > Len(sShortFilename) Then
        'Increase buffer size...
        sShortFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetShortPathName(sLongFilename, sShortFilename, Len(sShortFilename))
    End If
    
    'lRet contains the number of characters
    '     returned.


    If lRet > 0 Then
        GetShortFilename = left$(sShortFilename, lRet)
    End If
    
End Function


Public Function RemoveBackSlash(ByVal sPath As String) As String
    'Returns sPath without a trailing backsl
    '     ash if sPath
    'has one. Otherwise, returns sPath.
    
    sPath = Trim$(sPath)


    If Len(sPath) > 0 Then
        sPath = left$(sPath, Len(sPath) - IIf(right$(sPath, 1) = "\", 1, 0))
    End If
    RemoveBackSlash = sPath
    
End Function


Public Function AppPath() As String
    'Returns App.Path with backslash "\"
    Dim sPath As String
    sPath = App.path
    AppPath = sPath & IIf(right$(sPath, 1) <> "\", "\", "")
    
End Function


Public Function Exists(ByVal sFileName As String) As Boolean
    'Returns True if File Exists.
    'Else returns False.


    If Len(Trim$(sFileName)) > 0 Then
        On Error Resume Next
        sFileName = Dir$(sFileName)
        Exists = ((Err.Number = 0) And (Len(sFileName) > 0))
    Else
        Exists = False
    End If
    
End Function


Public Function GetFilePath(ByVal sFileName As String, Optional ByVal bAddBackslash As Boolean) As String
    'Returns Path Without FileTitle
    Dim lPos As Long
    lPos = InStrRev(sFileName, "\")


    If lPos > 0 Then
        GetFilePath = left$(sFileName, lPos - 1) _
        & IIf(bAddBackslash, "\", "")
    Else
        GetFilePath = ""
    End If
    
End Function


Public Function GetFileTitle(ByVal sFileName As String) As String
    'Returns FileTitle Without Path
    Dim lPos As Long
    lPos = InStrRev(sFileName, "\")


    If lPos > 0 Then


        If lPos < Len(sFileName) Then
            GetFileTitle = Mid$(sFileName, lPos + 1)
        Else
            GetFileTitle = ""
        End If
    Else
        GetFileTitle = sFileName
    End If
    
End Function



