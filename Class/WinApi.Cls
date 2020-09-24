VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim Akhil As Long

'VolumeLabel
Private Declare Function GetVolumeInformation Lib "kernel32" _
Alias "GetVolumeInformationA" _
(ByVal lpRootPathName As String, _
ByVal lpVolumeNameBuffer As String, _
ByVal nVolumeNameSize As Long, _
lpVolumeSerialNumber As Long, _
lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, _
ByVal lpFileSystemNameBuffer As String, _
ByVal nFileSystemNameSize As Long) As Long

'DiskFreeSpace
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long

'FindFiles
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFilename As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const MAX_PATH = 260, MAXDWORD = &HFFFF, INVALID_HANDLE_VALUE = -1, FILE_ATTRIBUTE_ARCHIVE = &H20, FILE_ATTRIBUTE_DIRECTORY = &H10, FILE_ATTRIBUTE_HIDDEN = &H2, FILE_ATTRIBUTE_NORMAL = &H80, FILE_ATTRIBUTE_READONLY = &H1, FILE_ATTRIBUTE_SYSTEM = &H4, FILE_ATTRIBUTE_TEMPORARY = &H100
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long:    ftCreationTime As FILETIME:    ftLastAccessTime As FILETIME:    ftLastWriteTime As FILETIME:    nFileSizeHigh As Long:    nFileSizeLow As Long:    dwReserved0 As Long:    dwReserved1 As Long:    cFileName As String * MAX_PATH:    cAlternate As String * 14
End Type
Private Declare Function SendMessageList Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Const LB_ADDSTRING = &H180

'ComputerName
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const REG_NONE = 0&: Private Const REG_SZ = 1&: Private Const REG_EXPAND_SZ = 2&
Private Const REG_BINARY = 3&: Private Const REG_DWORD = 4&: Private Const REG_DWORD_LITTLE_ENDIAN = 4&: Private Const REG_DWORD_BIG_ENDIAN = 5&: Private Const REG_LINK = 6&: Private Const REG_MULTI_SZ = 7&: Private Const REG_RESOURCE_LIST = 8&: Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9&: Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20&
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE = KEY_READ

Private Declare Function RegEnumKeyEx& Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey&, ByVal dwIndex&, ByVal lpname$, lpcbName&, ByVal lpReserved&, ByVal lpClass$, lpcbClass&, lpftLastWriteTime As FILETIME)
Private Declare Function RegEnumValue& Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey&, ByVal dwIndex&, ByVal lpname$, lpcbName&, ByVal lpReserved&, lpdwType&, lpValue As Any, lpcbValue&)
Private Declare Function RegQueryInfoKey& Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey&, ByVal lpClass$, lpcbClass&, ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, lpcbMaxClassLen&, lpcValues&, lpcbMaxValueNameLen&, lpcbMaxValueLen&, lpcbSecurityDescriptor&, lpftLastWriteTime As FILETIME)
Private Const KEY_ALL_ACCESS = &H3: Private Const REG_OPTION_NON_VOLATILE = 0

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExBinary Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Byte, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

'SystemParametersInfo
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPIF_UPDATEINIFILE = &H1, SPIF_SENDWININICHANGE = &H2
Private Type RECT
        Left As Long: Top As Long: Right As Long: Bottom As Long
End Type
Dim AkhilSt As RECT

'ScreenSaver
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const WM_SYSCOMMAND = &H112&, SC_SCREENSAVE = &HF140&

'GetMousePos
Private Type POINTAPI
        x As Long: y As Long
End Type: Private AkhilApi As POINTAPI
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Transparent form
Private Const GWL_EXSTYLE = (-20), WS_EX_TRANSPARENT = &H20&, SWP_FRAMECHANGED = &H20, SWP_NOMOVE = &H2, SWP_NOSIZE = &H1, SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE, HWND_NOTOPMOST = -2, HWND_TOPMOST = -1
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'ShellNotifyIcon
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
 Private Type NOTIFYICONDATA
    cbSize As Long:    hWnd As Long:    uID As Long:    uFlags As Long:    uCallbackMessage As Long:    hIcon As Long:    szTip As String * 64
End Type
Const NIM_ADD = &H0, NIM_MODIFY = &H1, NIM_DELETE = &H2, NIF_MESSAGE = &H1, NIF_ICON = &H2, NIF_TIP = &H4
Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Const WM_MOUSEMOVE = &H200, WM_LBUTTONDBLCLK = &H203, WM_LBUTTONDOWN = &H201, WM_RBUTTONDOWN = &H204

'MenuBitmaps
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Const MF_BITMAP = &H4&

'MoveFormWithoutBorders
Private Declare Function SendMessageMFWB Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1, HTCAPTION = 2

'GetSystemFolders
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Keybd_event
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_LWIN = &H5B, KEYEVENTF_KEYUP = &H2, VK_APPS = &H5D

'PlayAVI
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'SetDeskWallPaper
Private Declare Function SystemParameters Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

'ExitWindows
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const EWX_FORCE = 4, EWX_LOGOFF = 0, EWX_REBOOT = 2, EWX_SHUTDOWN = 1

'ShowHideTaskBar
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Const SWP_HIDEWINDOW = &H80, SWP_SHOWWINDOW = &H40

'Retrieve Icon
Private Const SHGFI_DISPLAYNAME = &H200, SHGFI_EXETYPE = &H2000, SHGFI_SYSICONINDEX = &H4000, SHGFI_LARGEICON = &H0, SHGFI_SMALLICON = &H1, SHGFI_SHELLICONSIZE = &H4, SHGFI_TYPENAME = &H400, ILD_TRANSPARENT = &H1, BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Type SHFILEINFO
   hIcon As Long: iIcon As Long: dwAttributes As Long: szDisplayName As String * MAX_PATH: szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal I As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Private shinfo As SHFILEINFO, sshinfo As SHFILEINFO

'DocumentList
Private Declare Sub SHAddToRecentDocs Lib "shell32.dll" (ByVal uFlags As Long, ByVal pv As String)

'Hotkey
Private Declare Function SendMessageHotKey Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETHOTKEY = &H32, WM_SHOWWINDOW = &H18

'ShowCursor
Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)

'Detect Sound Card
Private Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long

'ShellExecute
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL As Long = 1, SW_SHOWMAXIMIZED As Long = 3, SW_SHOWDEFAULT As Long = 10

'sndPlaySound
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Const SND_ALIAS = &H10000, SND_ASYNC = &H1, SND_NOWAIT = &H2000, SND_LOOP = &H8

'SetCursorPos
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

'Change ToolBar Style
Private Const WM_USER = &H400, TB_SETSTYLE = WM_USER + 56, TB_GETSTYLE = WM_USER + 57, TBSTYLE_FLAT = &H800
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'Default Property Values:
Const m_def_x = 0, m_def_y = 0, m_def_WorkAreaLeft = 0, m_def_WorkAreaTop = 0, m_def_WorkAreaRight = 0, m_def_WorkAreaBottom = 0

'Property Variables:
Dim m_WorkAreaLeft, m_WorkAreaTop, m_WorkAreaRight, m_WorkAreaBottom, m_x, m_y As Integer

Public Enum ShowCursors
chtHide = False
chtShow = True
End Enum
Public Enum Hotkeys
keyShiftA = &H141
keyShiftB = &H142
keyShiftC = &H143
keyShiftD = &H144
keyShiftE = &H145
keyShiftF = &H146
keyShiftG = &H147
keyShiftH = &H148
keyShiftI = &H149
keyShiftJ = &H14A
keyShiftK = &H14B
keyShiftL = &H14C
keyShiftM = &H14D
keyShiftN = &H14E
keyShiftO = &H14F
keyShiftP = &H150
keyShiftQ = &H151
keyShiftR = &H152
keyShiftS = &H153
keyShiftT = &H154
keyShiftU = &H155
keyShiftV = &H156
keyShiftW = &H157
keyShiftX = &H158
keyShiftY = &H159
keyShiftZ = &H15A

keyControlA = &H241
keyControlB = &H242
keyControlC = &H243
keyControlD = &H244
keyControlE = &H245
keyControlF = &H246
keyControlG = &H247
keyControlH = &H248
keyControlI = &H249
keyControlJ = &H24A
keyControlK = &H24B
keyControlL = &H24C
keyControlM = &H24D
keyControlN = &H24E
keyControlO = &H24F
keyControlP = &H250
keyControlQ = &H251
keyControlR = &H252
keyControlS = &H253
keyControlT = &H254
keyControlU = &H255
keyControlV = &H256
keyControlW = &H257
keyControlX = &H258
keyControlY = &H259
keyControlZ = &H25A

keyAltA = &H441
keyAltB = &H442
keyAltC = &H443
keyAltD = &H444
keyAltE = &H445
keyAltF = &H446
keyAltG = &H447
keyAltH = &H448
keyAltI = &H449
keyAltJ = &H44A
keyAltK = &H44B
keyAltL = &H44C
keyAltM = &H44D
keyAltN = &H44E
keyAltO = &H44F
keyAltP = &H450
keyAltQ = &H451
keyAltR = &H452
keyAltS = &H453
keyAltT = &H454
keyAltU = &H455
keyAltV = &H456
keyAltW = &H457
keyAltX = &H458
keyAltY = &H459
keyAltZ = &H45A
End Enum
Public Enum PlaySoundSettings
    sndWavFile = 0
    sndWindowsStart = 1
    sndWindowsExit = 2
    sndRestoreUp = 3
    sndRestoreDown = 4
    sndApplicationError = 5
    sndQuestion = 6
    sndAsterisk = 7
    sndExclamation = 8
    sndSystemHand = 9
End Enum
Public Enum TrayIconReturnValues
    xyzMouseMove = &H200
    xyzLeftButtonDown = &H201
    xyzLeftButtonUp = &H202
    xyzLeftButtonDoubleClick = &H203
    xyzRightButtonDown = &H204
    xyzRightButtonUp = &H205
    xyzRightButtonDoubleClick = &H206
    xyzMiddleButtonDown = &H207
    xyzMiddleButtonUp = &H208
    xyzMiddleButtonDoubleClick = &H209
End Enum
Public Enum SystemDirs
    dirWindows = 0
    dirSystem = 1
    dirTemp = 2
End Enum
Public Enum StartMenuItems
    strtExplorer
    strtFind
    strtMinimize
    strtRun
    strtStartMenu
    strtHelp
End Enum
Public Enum IconRetrieve
ricnLarge = 32
ricnSmall = 16
End Enum
Public Enum Reg
    CLASSES_ROOT = &H80000000
    CURRENT_USER = &H80000001
    LOCAL_MACHINE = &H80000002
    USERS = &H80000003
    PERFORMANCE_DATA = &H80000004
    CURRENT_CONFIG = &H80000005
    DYN_DATA = &H80000006
End Enum
Public Enum RegDataType
    regString = 1
    regBinary = 3
    regDword = 4
End Enum
Public Enum ShellIcon
    shelliconAdd = 0
    shelliconModify = 1
    shelliconDelete = 2
End Enum
Public Enum ComputerNames
    cmpGetComputerName = 0
    cmpSetComputerName = 1
    cmpGetUserName = 2
End Enum
Enum WindowsShutDown
    winShutDown = 0
    winReboot = 1
    winForce = 2
    winLogOff = 3
End Enum
Public Enum AutoRunDrives
    runFloppy = 251
    runHard = 247
    runFloppyHard = 243
    runDataCD = 223
    runAudioCD = 255
    runAudioDataCD = 223
    runAllDrives = 131
    runRAMDrive = 191
    runNetworkDrive = 239
    runNetworkRAMDrives = 175
End Enum
Public Sub PlaySound(func As PlaySoundSettings, Optional WavFileName As String)
Dim temp As String: Select Case func
Case 0:        Akhil = sndPlaySound(WavFileName, SND_ASYNC + SND_NOWAIT)
Case 1:        temp = "SystemStart":        GoTo PlayAlias
Case 2:        temp = "SystemExit":        GoTo PlayAlias
Case 3:        temp = "RestoreUp":        GoTo PlayAlias
Case 4:        temp = "RestoreDown":        GoTo PlayAlias
Case 5:        temp = "AppGPFault":        GoTo PlayAlias
Case 6:        temp = "SystemQuestion":        GoTo PlayAlias
Case 7:        temp = "SystemAsterisk":        GoTo PlayAlias
Case 8:        temp = "SystemExclamation":        GoTo PlayAlias
Case 9:        temp = "SystemHand":        GoTo PlayAlias: End Select: Exit Sub
PlayAlias: Akhil = sndPlaySound(temp, SND_ALIAS + SND_ASYNC + SND_NOWAIT)
End Sub
Public Sub SetMyPosTopMost(H As Long, Optional TakeOffMyPosTopMost As Boolean)
If TakeOffMyPosTopMost = True Then
SetWindowPos H, -2, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
Else
SetWindowPos H, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End If
End Sub
Public Sub SetMyPosOnDesktop(H As Long)
SetWindowPos H, 1, 0&, 0&, 0&, 0&, &H1 Or &H2 Or &H10
End Sub
Public Sub HideWindow(H As Long, Optional ShowMe As Boolean)
If ShowMe = False Then
SetWindowPos H, 1, 0&, 0&, 0&, 0&, SWP_HIDEWINDOW
Else
SetWindowPos H, 1, 0&, 0&, 0&, 0&, SWP_SHOWWINDOW
End If
End Sub
Public Sub SetSysWorkArea(l As Integer, t As Integer, r As Integer, B As Integer)
AkhilSt.Left = l
AkhilSt.Top = t
AkhilSt.Right = r
AkhilSt.Bottom = B
Akhil = SystemParametersInfo(47, 0, AkhilSt, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
End Sub
Public Sub ChangeToolBarStyle(ToolBarh As Long)
'Refresh after using the function to get the results. eg: Toolbar1.Refresh
 Dim Style As Long, hToolbar As Long, r As Long
 hToolbar = FindWindowEx(ToolBarh, 0&, "ToolbarWindow32", vbNullString)
 Style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
 Style = Style Or TBSTYLE_FLAT
 Call SendMessageLong(hToolbar, TB_SETSTYLE, 0, Style)
 End Sub
Public Sub MakeTransparentForm(H As Long)
SetWindowLong H, GWL_EXSTYLE, WS_EX_TRANSPARENT
SetWindowPos H, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME
End Sub
Public Function GetSystemFolders(func As SystemDirs)
Dim r, nSize As Long, tmp As String
 tmp = Space$(256):    nSize = Len(tmp)
Select Case func
   Case 0
      r = GetWindowsDirectory(tmp, nSize):     GetSystemFolders = TrimNull(tmp)
    Case 1
      r = GetSystemDirectory(tmp, nSize):      GetSystemFolders = TrimNull(tmp)
    Case 2
       r = GetTempPath(nSize, tmp):       GetSystemFolders = TrimNull(tmp)
    End Select
End Function
Public Sub TrayIcon(func As ShellIcon, H As Long, sName As String, P As StdPicture)
Dim Tic As NOTIFYICONDATA
Tic.cbSize = Len(Tic): Tic.hWnd = H: Tic.uID = 1&: Tic.uFlags = NIF_DOALL
Tic.uCallbackMessage = WM_MOUSEMOVE: Tic.hIcon = P: Tic.szTip = sName & Chr$(0)
Select Case func
Case 0
erg = Shell_NotifyIcon(NIM_ADD, Tic)
Case 1
erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
Case 2
erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Select
End Sub
Public Function TryIcnRtnMsg(x As Single) As TrayIconReturnValues
x = x \ Screen.TwipsPerPixelX
Select Case x
    Case xyzMouseMove: TryIcnRtnMsg = &H200
    Case xyzLeftButtonDown: TryIcnRtnMsg = &H201
    Case xyzLeftButtonUp: TryIcnRtnMsg = &H202
    Case xyzLeftButtonDoubleClick: TryIcnRtnMsg = &H203
    Case xyzRightButtonDown: TryIcnRtnMsg = &H204
    Case xyzRightButtonUp: TryIcnRtnMsg = &H205
    Case xyzRightButtonDoubleClick: TryIcnRtnMsg = &H206
    Case xyzMiddleButtonDown: TryIcnRtnMsg = &H207
    Case xyzMiddleButtonUp: TryIcnRtnMsg = &H208
    Case xyzMiddleButtonDoubleClick: TryIcnRtnMsg = &H209
End Select
End Function
Public Sub MenuBitmaps(H As Long, subMenu As Byte, MenuID As Byte, chkbmp As StdPicture)
hMenu& = GetMenu(H)
hSubMenu& = GetSubMenu(hMenu&, subMenu)
hID& = GetMenuItemID(hSubMenu&, MenuID)
SetMenuItemBitmaps hMenu&, hID&, MF_BITMAP, chkbmp, chkbmp
End Sub
Public Sub MoveFormWithoutBorders(H As Long)
Dim lngReturnValue As Long
Call ReleaseCapture
lngReturnValue = SendMessageMFWB(H, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
Public Sub Launch(func As StartMenuItems)
Dim VK_ACTION As Long
Select Case func
      Case strtExplorer: VK_ACTION = &H45
      Case strtFind: VK_ACTION = &H46
      Case strtMinimize: VK_ACTION = &H4D
      Case strtRun: VK_ACTION = &H52
      Case strtStartMenu: VK_ACTION = &H5B
      Case strtHelp: VK_ACTION = &H70
End Select
   Call keybd_event(VK_LWIN, 0, 0, 0):   Call keybd_event(VK_ACTION, 0, 0, 0)
   Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub
'Public Sub DocumentList(func As Byte, Optional sName As String)
'Select Case func
'Case 0
'If sName <> "" Then Call SHAddToRecentDocs(2, sName)
'Case 1
'Call SHAddToRecentDocs(2, vbNullString)
'End Select
'End Sub
Public Sub SetDeskWallPaper(func As Byte, Optional sBitmapFile As String)
Dim lRetVal As Long
Select Case func
Case 1
lRetVal = SystemParameters(ByVal 20, 0&, ByVal sBitmapFile, _
SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
Case 0
lRetVal = SystemParameters(20, 0&, "(None)", _
SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Select
End Sub
Public Sub ExitWindows(func As WindowsShutDown)
Dim lRetVal As Long
Select Case func
Case 0
lRetVal = ExitWindowsEx(EWX_SHUTDOWN, 0)
Case 1
lRetVal = ExitWindowsEx(EWX_REBOOT, 0)
Case 2
lRetVal = ExitWindowsEx(EWX_FORCE, 0)
Case 3
lRetVal = ExitWindowsEx(EWX_LOGOFF, 0)
End Select
End Sub
Public Sub RetrieveIcon(fName As String, DC As Long, icnSize As IconRetrieve)
Dim hImgSmall, hImgLarge As Long  'the handle to the system image list
Select Case icnSize
Case ricnSmall
hImgSmall = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
Call ImageList_Draw(hImgSmall, shinfo.iIcon, DC, 0, 0, ILD_TRANSPARENT)
Case ricnLarge
hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
Call ImageList_Draw(hImgLarge, shinfo.iIcon, DC, 0, 0, ILD_TRANSPARENT)
End Select
End Sub
Public Function RetrieveFileTypeName(fName As String) As String
ImgSmall& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
RetrieveFileTypeName = Left$(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr$(0)) - 1)
End Function
Public Sub TaskBar(func As Byte)
Dim Thwnd As Long
Select Case func
Case 0
Thwnd = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
Case 1
 Thwnd = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Select
End Sub
Public Sub ShowMouse(cur As ShowCursors)
Select Case cur
    Case chtHide
        Call ShowCursor(False)
     Case chtShow
        Call ShowCursor(True)
End Select
End Sub
Public Sub StartScreenSaver(Handle As Long)
Dim result As Long
result = SendMessage(Handle, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
Public Sub SetHotKey(H As Long, key As Hotkeys)
erg& = SendMessageHotKey(H, WM_SETHOTKEY, key, 0)
If erg& <> 1 Then
MsgBox "You need another hotkey", vbOKOnly, "Error"
End If
erg& = DefWindowProc(H, WM_SHOWWINDOW, 0, 0)
End Sub
Public Function SoundCardDetect() As Boolean
Dim I As Integer
I = waveOutGetNumDevs()
If I > 0 Then
    SoundCardDetect = True
Else
    SoundCardDetect = False
End If
End Function
Public Sub Delay(HowLong As Date)
TempTime = DateAdd("s", HowLong, Now)
While TempTime > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub
Public Sub DocumentList(func As String, Optional sName As String)
func = UCase$(func)
Select Case func
    Case "CLEAR"
     Call SHAddToRecentDocs(2, vbNullString)
    Case "ADD"
    Call SHAddToRecentDocs(2, sName)
End Select
End Sub
Public Sub PlayAvi(sFile As String)
Dim returnstring As String
returnstring = Space(127)
erg = mciSendString("open " & Chr$(34) & sFile & Chr$(34) & " type avivideo alias video", returnstring, 127, 0)
erg = mciSendString("set video time format ms", returnstring, 127, 0)
erg = mciSendString("play video from 0", returnstring, 127, 0)
End Sub
Public Sub ShowRecycleBin(H As Long)
   Dim success As Long
   success = ShellExecute(H, "Open", "explorer.exe", "/root,::{645FF040-5081-101B-9F08-00AA002F954E}", 0&, SW_SHOWNORMAL)
End Sub
Public Sub OpenWebSite(sFile As String, Handle As Long)
   Dim success As Long
   success = ShellExecute(Handle, "Open", sFile, 0&, 0&, 3)
   If success < 32 Then
     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
End Sub
Public Sub ExecuteFile(sFile As String, Handle As Long, Optional sParams As String)
   Dim success As Long
   success = ShellExecute(Handle, "Open", sFile, sParams, 0&, 3)
   If success < 32 Then
     Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
End Sub

Public Sub SendEmail(address As String, Handle As Long)
   address = "mailto:" & address
   Dim success As Long
   success = ShellExecute(Handle, "Open", address, 0&, 0&, 3)
 End Sub
Public Function PtoT(no As Variant, Optional YValue As Boolean)
' This function converts Pixels to Twips
If Not YValue Then
PtoT = no * Screen.TwipsPerPixelX
Else
PtoT = no * Screen.TwipsPerPixelY
End If
End Function
Public Function TtoP(no As Variant, Optional YValue As Boolean)
' This function does the reverse of the above function
If Not YValue Then
TtoP = no \ Screen.TwipsPerPixelX
Else
TtoP = no \ Screen.TwipsPerPixelY
End If
End Function
Public Sub SetMousePos(xx As Long, yy As Long)
Call SetCursorPos(xx, yy)
End Sub
Public Property Get WorkAreaLeft() As Integer
    Class_Initialize
    WorkAreaLeft = m_WorkAreaLeft
End Property
Public Property Get WorkAreaTop() As Integer
    Class_Initialize
    WorkAreaTop = m_WorkAreaTop
End Property
Public Property Get WorkAreaRight() As Integer
    Class_Initialize
    WorkAreaRight = m_WorkAreaRight
End Property
Public Property Get MouseX() As Integer
    Class_Initialize
    MouseX = m_x
End Property
Public Property Get MouseY() As Integer
    Class_Initialize
    MouseY = m_y
End Property
Public Property Get WorkAreaBottom() As Integer
        Class_Initialize
        WorkAreaBottom = m_WorkAreaBottom
End Property

Private Sub Class_Initialize()
    m_WorkAreaLeft = m_def_WorkAreaLeft
    m_WorkAreaTop = m_def_WorkAreaTop
    m_WorkAreaRight = m_def_WorkAreaRight
    m_WorkAreaBottom = m_def_WorkAreaBottom
    m_x = m_def_x
    m_y = m_def_y
    Akhil = SystemParametersInfo(48, 0, AkhilSt, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
    m_WorkAreaLeft = AkhilSt.Left * Screen.TwipsPerPixelX
    m_WorkAreaTop = AkhilSt.Top * Screen.TwipsPerPixelY
    m_WorkAreaRight = AkhilSt.Right * Screen.TwipsPerPixelX
    m_WorkAreaBottom = AkhilSt.Bottom * Screen.TwipsPerPixelY
    GetCursorPos AkhilApi
    m_x = AkhilApi.x * Screen.TwipsPerPixelX
    m_y = AkhilApi.y * Screen.TwipsPerPixelY
End Sub
Public Sub ListFunc()
'Temporary Function
'Lists all the functions of this WinApi Class
Dim Msg As String
Msg = Msg & vbNewLine & "Delay "
Msg = Msg & vbNewLine & "DiskSpace "
Msg = Msg & vbNewLine & "DocumentList "
Msg = Msg & vbNewLine & "ExecuteFile "
Msg = Msg & vbNewLine & "ExecuteFile "
Msg = Msg & vbNewLine & "ExitWindows "
Msg = Msg & vbNewLine & "FindFilesAPI "
Msg = Msg & vbNewLine & "Launch "
Msg = Msg & vbNewLine & "MakeTransparentForm "
Msg = Msg & vbNewLine & "MenuBitmaps "
Msg = Msg & vbNewLine & "MoveFormWithoutBorders "
Msg = Msg & vbNewLine & "OpenWebSite "
Msg = Msg & vbNewLine & "PlayAudioCD "
Msg = Msg & vbNewLine & "PlayAvi "
Msg = Msg & vbNewLine & "PlaySound "
Msg = Msg & vbNewLine & "RegistryValueS "
Msg = Msg & vbNewLine & "RegistryValueD "
Msg = Msg & vbNewLine & "RegistryRun "
Msg = Msg & vbNewLine & "RegistryDeleteValue "
Msg = Msg & vbNewLine & "RegistryNewKey "
Msg = Msg & vbNewLine & "RegistryDeleteKey "
Msg = Msg & vbNewLine & "RegistryCreateRightClickAccess "
Msg = Msg & vbNewLine & "RegistryCreateAssociation "
Msg = Msg & vbNewLine & "RetrieveIcon "
Msg = Msg & vbNewLine & "SendEmail "
Msg = Msg & vbNewLine & "SetDeskWallPaper "
Msg = Msg & vbNewLine & "SetHotKey "
Msg = Msg & vbNewLine & "SetMousePos "
Msg = Msg & vbNewLine & "SetMyPosOnDesktop "
Msg = Msg & vbNewLine & "SetMyPosTopMost "
Msg = Msg & vbNewLine & "SetSysWorkArea "
Msg = Msg & vbNewLine & "ShowMouse "
Msg = Msg & vbNewLine & "ShowRecycleBin "
Msg = Msg & vbNewLine & "StartScreenSaver "
Msg = Msg & vbNewLine & "TaskBar "
Msg = Msg & vbNewLine & "TrayIcon "
Msg = Msg & vbNewLine & "**********Functions**********"
Msg = Msg & vbNewLine & "GetSystemFolders()"
Msg = Msg & vbNewLine & "RetrieveFileTypeName()"
Msg = Msg & vbNewLine & "SoundCardDetect()"
Msg = Msg & vbNewLine & "TtoP()"
Msg = Msg & vbNewLine & "PtoT()"
Msg = Msg & vbNewLine & "**********Properties**********"
Msg = Msg & vbNewLine & "WorkAreaLeft, WorkAreaRight, WorkAreaTop, WorkAreaBottom"
Msg = Msg & vbNewLine & "MouseX, MouseY "
MsgBox Msg
End Sub
Public Sub RegistryValueS(sKey As String, sValueName As String, sValue As String, home As Reg)
   SetKeyValue sKey, sValueName, sValue, REG_SZ, home
End Sub
Public Sub RegistryValueD(sKey As String, sValueName As String, sValue As String, home As Reg)
   SetKeyValue sKey, sValueName, sValue, REG_DWORD, home
End Sub
Public Sub RegistryValueB(sKey As String, sValueName As String, sValue As Byte, home As Reg)
   SetKeyValue sKey, sValueName, sValue, REG_BINARY, home
End Sub
Public Function RegistryGetValue(sKey As String, sValueName As String, sValue As Variant, nValuetype As RegDataType, home As Reg)
Call GetKeyValue(sKey, sValueName, sValue, nValuetype, home)
If nValuetype = regString Then sValue = StripNulls(CStr(sValue))
End Function
Public Sub RegistryRun(sValueName As String, sValue As String)
      SetKeyValue "software\microsoft\windows\currentversion\run", sValueName, sValue, REG_SZ, HKEY_LOCAL_MACHINE
End Sub
Public Sub RegistryDeleteValue(sKey As String, sValueName As String, home As Reg)
      DeleteValue home, sKey, sValueName
End Sub
Public Sub RegistryNewKey(sKeyName As String, home As Reg)
    CreateNewKey sKeyName, home
End Sub
Public Sub RegistryDeletekey(sSubKey As String, sKey As String, home As Reg)
      DeleteKey home, sKey, sSubKey
End Sub
Public Sub RegistryCreateRightClickAccess(sName As String, spath As String)
   CreateNewKey "*\shell\" & sName & "\Command", HKEY_CLASSES_ROOT
   If Right(App.path, 1) = "\" Then
   spath = App.path & App.EXEName & ".exe" & " " & Command & " " & "%1"
   Else
   spath = App.path & "\" & App.EXEName & ".exe" & " " & Command & " " & "%1"
   End If
   SetKeyValue "*\shell\" & sName & "\Command", "", spath, REG_SZ, HKEY_CLASSES_ROOT
End Sub
Public Sub RegistryCreateAssociation(sFileExtension_3Chars As String, _
 sProgramType As String, spath As String, Optional sProgramName As String)
  CreateNewKey "." & sFileExtension_3Chars, HKEY_CLASSES_ROOT
  Call SetKeyValue("." & sFileExtension_3Chars, "", sProgramType, REG_SZ, HKEY_CLASSES_ROOT)
  CreateNewKey sProgramType & "\shell\open\command", HKEY_CLASSES_ROOT
  SetKeyValue sProgramType, "", sProgramName, REG_SZ, HKEY_CLASSES_ROOT
  spath = spath & " %1"
  SetKeyValue sProgramType & "\shell\open\command", "", spath, REG_SZ, HKEY_CLASSES_ROOT
End Sub
Public Sub RegistryEnumerateKeys(sValue As String, handleListbox As Long, home As Reg)
Call RegEnumKeys(sValue, handleListbox, home)
End Sub
Public Sub RegistryEnumerateValues(sValue As String, handleListboxEntries As Long, handleListboxValues As Long, home As Reg)
Call RegEnumValues(sValue, handleListboxEntries, handleListboxValues, home)
End Sub
Public Sub DisableCtrlAltDelete(bDisable As Boolean)
    Dim x As Long
    x = SystemParametersInfo(97, bDisable, CStr(1), 0)
End Sub
Public Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer, Optional handleListbox As Long, Optional SendBackInFilename_PathFormat As Boolean, Optional DontLookInSubFolders As Boolean)
Dim Filename As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim I As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim temp As String
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
If Right(path, 1) <> "\" Then path = path & "\"
If Not DontLookInSubFolders Then
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
    Do While Cont
    DirName = StripNulls(WFD.cFileName)
    If (DirName <> ".") And (DirName <> "..") Then
        If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
            dirNames(nDir) = DirName
            DirCount = DirCount + 1
            nDir = nDir + 1
            ReDim Preserve dirNames(nDir)
        End If
    End If
    Cont = FindNextFile(hSearch, WFD)
    DoEvents
    Loop
    Cont = FindClose(hSearch)
End If
End If
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
    While Cont
    Filename = StripNulls(WFD.cFileName)
    DoEvents
    If (Filename <> ".") And (Filename <> "..") Then
        FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
        MAXDWORD) + WFD.nFileSizeLow
        FileCount = FileCount + 1
        temp = path & Filename
 If SendBackInFilename_PathFormat Then temp = Filename & vbTab & path
 SendMessageList handleListbox, LB_ADDSTRING, -1, ByVal temp
    End If
    Cont = FindNextFile(hSearch, WFD) ' Get next file
    Wend
    Cont = FindClose(hSearch)
End If
If Not DontLookInSubFolders Then
If nDir > 0 Then
    For I = 0 To nDir - 1
    DoEvents
    Debug.Print path & dirNames(I) & "\"
    If Not SendBackInFilename_PathFormat Then
    FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(I) & "\", SearchStr, FileCount, DirCount, handleListbox)
    Else
    FindFilesAPI = FindFilesAPI + FindFilesAPI(path & dirNames(I) & "\", SearchStr, FileCount, DirCount, handleListbox, True)
    End If
    Next I
End If
End If
End Function
Public Function DiskSpace(spath As String, dFree As Double, dUsed As Double, dTotal As Double)
 Dim sDrive As String
 Dim lReturn As Long
 Dim l1 As Long   'l1 = Sectors Per Cluster
 Dim l2 As Long   'l2 = Bytes Per Sector
 Dim l3 As Long   'l3 = Number Of Free Clusters
 Dim l4 As Long 'l4 = Total Number Of Clusters
sDrive = Left$(spath, 1) & ":\" 'Get drive letter from path
lReturn = GetDiskFreeSpace(sDrive, l1, l2, l3, l4)
'DiskSpace = l1 * l2 * l3
dFree = l1 * l2 * l3
dTotal = l1 * l2 * l4
dUsed = dTotal - dFree
End Function
Public Sub PlayAudioCd(TrackNum As Integer, Optional StopPlaying As Boolean)
On Local Error Resume Next
Dim I As Long, RS As String, cb As Long, t#
RS = Space$(128)
If StopPlaying Then
I = mciSendString("stop cdaudio", RS, 128, cb)
I = mciSendString("close cdaudio", RS, 128, cb)
Else
I = mciSendString("stop cdaudio", RS, 128, cb)
I = mciSendString("close cdaudio", RS, 128, cb)
RS = Space$(128)
I = mciSendString("status cdaudio position track " & TrackNum, RS, 128, cb)
I = mciSendString("open cdaudio", RS, 128, cb)
I = mciSendString("set cdaudio time format milliseconds", RS, 128, cb)
I = mciSendString("play cdaudio", RS, 128, cb)
End If
End Sub
Public Sub AutoRun(drv As AutoRunDrives)
RegistryValueB "software\microsoft\windows\currentversion\policies\explorer", "NoDriveTypeAutoRun", CByte(drv), CURRENT_USER
End Sub
Public Function ComputerName(func As ComputerNames, Optional CmpName As String) As String
Dim temp As String
temp = Space$(255)
Select Case func
    Case 0
        GetComputerName temp, 255
        ComputerName = StripNulls(temp)
    Case 1
        If CmpName <> "" Then SetComputerName CmpName
    Case 2
        GetUserName temp, 255
        ComputerName = StripNulls(temp)
End Select
End Function
Public Function DiskVolumeInfo(spath As String, sDriveName As String, lLength As Long, lSerialNo As Long, lMaximumChar As Long, lFileSystemFlags As Long, sNameOfFileSystem As String, lLeng As Long)
t = GetVolumeInformation(spath, sDriveName, lLength, lSerialNo, lMaximumChar, lFileSystemFlags, sNameOfFileSystem, lLeng)
End Function

Public Function RegEnumKeys&(rgeSubKey As String, handleListbox As Long, rgeMainKey As Long)
    Dim sRoot$, sRoot2$, temp As String
         
    ' --------------------------------------------------------
    ' This function will load all subkeys into the TreeView
    ' --------------------------------------------------------
    Dim lRtn&       ' Returned by registry functions, should be 0&
    Dim hKey&       ' Return handle to opened key
    Dim strucLastWriteTime    As FILETIME
    Dim sSubKeyName$
    Dim sClassString$
    Dim lLenSubKey&
    Dim lLenClass&
    Dim lKeyIndx&
    Dim lRet&
    Dim hKey2&
    Dim sSubKey2$
    Dim sNewKey$
    
    '---------------------------------------------
    'values for QueryInfoKey:
    '---------------------------------------------
    Dim sClassName$
    Dim lClassLen&
    Dim lSubKeys&
    Dim lMaxSubKey&
    Dim sMaxSubKey$
    Dim lMaxClass&
    Dim sMaxClass$
    Dim lValues&
    Dim lMaxValueName&
    Dim lMaxValueData&
    Dim lSecurityDesc&
    
    ' -----------------------------------------------------
    ' Open key
    ' -----------------------------------------------------
    lRtn = RegOpenKeyEx(rgeMainKey, rgeSubKey, 0&, KEY_READ, hKey)
        
    ' -----------------------------------------------------
    ' A call to RegQueryInfoKey will tell us the maximum
    '   keyname length
    ' -----------------------------------------------------
    sClassName = Space$(255)
    lClassLen = CLng(Len(sClassName))
    lRet = RegQueryInfoKey(hKey, sClassName, lClassLen, 0&, lSubKeys, lMaxSubKey, lMaxClass, lValues, lMaxValueName, lMaxValueData, lSecurityDesc, strucLastWriteTime)
    sMaxSubKey = Space$(lMaxSubKey + 1)
    sMaxClass = Space$(lMaxClass + 1)

    ' -----------------------------------------------------
    ' Enumerate the keys
    ' -----------------------------------------------------
    lKeyIndx = 0&
    Do While lRtn = ERROR_SUCCESS
        
        ' -----------------------------------------------------
        ' If the enumeration fails due to a buffer over-run,
        '   we will loop back to this point with larger buffers.
        ' -----------------------------------------------------
ReTryKeyEnumeration:
            
        ' --------------------------------------------------
        ' Set variables
        ' --------------------------------------------------
        sSubKeyName = sMaxSubKey
        lLenSubKey = lMaxSubKey
        sClassString = sMaxClass
        lLenClass = lMaxClass
    
        
        ' --------------------------------------------------
        ' Call the enumeration function
        ' --------------------------------------------------
        lRtn = RegEnumKeyEx(hKey, lKeyIndx, sSubKeyName, lLenSubKey, 0&, sClassString, lLenClass, strucLastWriteTime)
        If InStr(sSubKeyName, Chr$(0)) > 1 Then
            sSubKeyName = Left$(sSubKeyName, InStr(sSubKeyName, Chr$(0)) - 1)
        End If
        
        
        ' --------------------------------------------------
        ' Check for success
        ' --------------------------------------------------
        If lRtn = ERROR_SUCCESS Then
            lNewKey = lNewKey + 1
            sNewKey = "A" & Format$(lNewKey, "000000")
                    temp = sSubKeyName
                    SendMessageList handleListbox, LB_ADDSTRING, -1, ByVal temp
'            If bFullEnumeration = True Then
'                sSubKey2 = sSubKeyName
'                If rgeSubKey <> "" Then
'                    sSubKey2 = Trim(rgeSubKey) & "\" & sSubKeyName
'                 End If
'
'                ' -----------------------------------------------
'                ' Use RegQueryInfoKey to find out if this key has
'                '   subkeys
'                ' -----------------------------------------------
'                lRet = RegOpenKeyEx(rgeMainKey, sSubKey2, 0&, KEY_READ, hKey2)
'                '------------------------------------------------------
'                'We are fully enumerating a key, so set images and
'                'Recurse a single SubKey to set + indicator if there are
'                'subkeys below this one
'                '------------------------------------------------------
'                lRet = RegQueryInfoKey(hKey2, vbNullString, 0&, 0&, lSubKeys, 0&, 0&, 0&, 0&, 0&, 0&, strucLastWriteTime)
'                    ' --------------------------------------------------
'                    ' Check for success.  If lSubKeys is greater than zero
'                    ' there are subkeys for this key, and we will set a fake
'                    ' node under this one to make a + symbol.
'                    ' --------------------------------------------------
'                    If lRet = ERROR_SUCCESS And lSubKeys > 0 Then
'                        sRoot2 = nodX.key
'                        lNewKey = lNewKey + 1
'                        sNewKey = "A" & Format$(lNewKey, "000000")
'                        Set nodX = Angst.TreeView1.Nodes.Add(sRoot2, tvwChild, sNewKey, "PlaceHolder", 1)
'                    End If
'                    lRet = RegCloseKey(hKey2)
'                End If
'            Else
'                Exit Do
'            End If
            lKeyIndx = lKeyIndx + 1
        ElseIf lRtn = ERROR_MORE_DATA Then
            ' -----------------------------------------------
            ' This error means that, despite querying the key
            '   we have not set one of the buffers large
            '   enough.Increment the buffer sizes and try
            '   again
            ' -----------------------------------------------
            lMaxSubKey = lMaxSubKey + 5
            lMaxClass = lMaxClass + 5
            sMaxSubKey = Space$(lMaxSubKey + 1)
            sMaxClass = Space$(lMaxClass + 1)
            GoTo ReTryKeyEnumeration
        ElseIf lRtn = ERROR_NO_MORE_ITEMS Then
            ' -----------------------------------------------
            ' Not an error, just end of list -- exit the
            '   loop
            ' -----------------------------------------------
            lRtn = ERROR_SUCCESS
            Exit Do
        ElseIf lRtn <> ERROR_SUCCESS Then
            ' --------------------------------------------------
            ' Key still open, so display the error and fall
            '   thru to the close function below
            ' --------------------------------------------------
            
            Exit Do
        End If
    Loop

    
    ' -----------------------------------------------------
    ' Set the return to the last error
    ' -----------------------------------------------------
    RegEnumKeys = lRtn

    ' -----------------------------------------------------
    ' Always close opened keys!
    ' -----------------------------------------------------
    lRtn = RegCloseKey(hKey)
    
End Function
Public Sub RegEnumValues(rgeSubKey As String, handleListboxEntries As Long, handleListboxValues As Long, rgeMainKey As Long)
 Dim sTemps As String, rgeEntry As String, rgeValue As String
    Dim lRtn&        ' Returned by registry functions, should be 0&
    Dim hKey&       ' Return handle to opened key
    Dim lLenValueName&
    Dim lLenValue&
    Dim lKeyIndx&
    Dim sBinaryString$
    Dim iTempInt%
    
    Dim sClassName$
    Dim lClassLen&
    Dim lSubKeys&
    Dim lMaxSubKey&
    Dim lMaxClass&
    Dim lValues&
    Dim lMaxValueName&
    Dim lMaxValueData&
    Dim lSecurityDesc&
    Dim strucLastWriteTime As FILETIME
    Dim pbytValueName() As Byte
    Dim pbytValue() As Byte
    Dim DataType&
    Dim ValueName$
    Dim ValueVal$
    
    Dim iListWidth%  'Used to set listbox scrollbar
    
    lRtn = RegOpenKeyEx(rgeMainKey, rgeSubKey, 0&, KEY_READ, hKey)
    If lRtn <> ERROR_SUCCESS Then
        Exit Sub
    End If
   
    sClassName = Space$(255)
    lClassLen = CLng(Len(sClassName))
    lRtn = RegQueryInfoKey(hKey, sClassName, lClassLen, 0&, lSubKeys, lMaxSubKey, lMaxClass, lValues, lMaxValueName, lMaxValueData, lSecurityDesc, strucLastWriteTime)
   
    lKeyIndx = 0&
    Do While lRtn = ERROR_SUCCESS And lKeyIndx < lValues
        sBinaryString = ""

ReTryValueEnumeration:
        
        ValueName = Space$(lMaxValueName + 1)
        lLenValueName = CLng(Len(ValueName)) '+ 1
        ValueVal = Space$(lMaxValueData + 1)
        lLenValue = CLng(Len(ValueVal))       '+ 1
        
        lRtn = RegEnumValue(hKey, lKeyIndx, ValueName, _
        lLenValueName, 0&, DataType, ByVal ValueVal, _
        lLenValue)
        
        If lRtn = ERROR_SUCCESS Then
            
            rgeEntry = ValueName
            rgeDataType = DataType

            rgeValue = ValueVal
            
            'rgeValue = pbytValue()
            '-----------------------------------------------
            'Add an item to the list box
            '-----------------------------------------------
                
            ' -----------------------------------------------
            ' Start building the entry to put in the list box:
            ' -----------------------------------------------
            rgeEntry = Mid$(rgeEntry, 1, lLenValueName)
        
            ' -----------------------------------------------
            ' Default values don't have a name.
            ' -----------------------------------------------
            If lLenValueName = 0 Then
                rgeEntry = "(Default)"
            End If
            
            rgeValue = Mid$(rgeValue, 1, lLenValue)
            Select Case rgeDataType
                Case REG_MULTI_SZ
                    ' --------------------------------------
                    ' REG_MULTI_SZ strings are a series of
                    '   zero terminated strings. If we don't
                    '   strip out the zeros, only the first
                    '   one will display.
                    ' We will replace them with spaces.
                    ' --------------------------------------
                     Do While InStr(rgeValue, Chr$(0))
                        rgeValue = Left$(rgeValue, InStr(rgeValue, Chr$(0)) - 1) & " " & Right$(rgeValue, Len(rgeValue) - InStr(rgeValue, Chr$(0)))
                    Loop
                Case REG_SZ
                    ' --------------------------------------
                    ' REG_SZ values are zero-terminated
                    '   strings, and are the most common
                    '   values.
                    ' --------------------------------------
                        '---------------------------------------
                    'Put quotes around the string
                    '---------------------------------------
                    rgeValue = """" & Left$(rgeValue, lLenValue - 1) & """"
                
                Case REG_EXPAND_SZ
                    '---------------------------------------
                    'Environmental variables that are binary
                    'but evaluate as strings.  Not edited by
                    'this program.
                    '---------------------------------------
                    Case REG_FULL_RESOURCE_DESCRIPTOR
                    ' --------------------------------------
                    ' Resource Descriptors require a special
                    '   editor to properly be displayed or
                    '   edited.
                    ' --------------------------------------
    
                    rgeValue = "REG_FULL_RESOURCE_DESCRIPTOR"

                Case REG_DWORD
                    ' --------------------------------------
                    ' REG_DWORD values are 32-bit unsigned
                    '   integers
                    ' Tortuous manipulation to make values
                    '   above 7FFFFFFF appear as positive
                    '   values.
                    ' VB Longs would display them as
                    '   negative numbers.
                    ' --------------------------------------
    
                    fTempDbl = Asc(Mid$(rgeValue, 1, 1)) + &H100& * Asc(Mid$(rgeValue, 2, 1)) + &H10000 * Asc(Mid$(rgeValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(rgeValue, 4, 1)))
                    If fTempDbl > &H7FFFFFFF Then
                        rgeValue = "&H" & Hex$(fTempDbl - 4294967296#)
                    Else
                        rgeValue = "&H" & Hex$(fTempDbl)
                    End If
                    rgeValue = rgeValue & " (" & Format$(fTempDbl) & ")"

                Case REG_BINARY
              
                    ' --------------------------------------
                    ' Binary values may be of any length,
                    '   and may represent text or other data.
                    ' They require a special editor to
                    '   modify them.
                    ' --------------------------------------
    
                    For iTempInt = 1 To Len(rgeValue)
                        sBinaryString = sBinaryString & Format$(Hex(Asc(Mid$(rgeValue, iTempInt, 1))), "00") & " "
                    Next iTempInt
                    rgeValue = sBinaryString
            End Select

            If Len(rgeValue) = 0 Then
                rgeValue = "(value not set)"
            End If

         
            ' -----------------------------------------------
            ' Enter the value into the list box
            ' -----------------------------------------------
            temp = rgeEntry
            SendMessageList handleListboxEntries, LB_ADDSTRING, -1, ByVal temp
            temp = rgeValue
            SendMessageList handleListboxValues, LB_ADDSTRING, -1, ByVal temp
                        
            ' -----------------------------------------------
            ' Increment the key and do it again.
            ' -----------------------------------------------
            lKeyIndx = lKeyIndx + 1

        ElseIf lRtn = ERROR_MORE_DATA Then
            ' -----------------------------------------------
            ' This error means that, despite querying the key,
            '   we have not set one of the buffers large
            '   enough.  Increment the buffer sizes and try
            '   again
            ' -----------------------------------------------
            lMaxValueData = lMaxValueData + 5
            lMaxValueName = lMaxValueName + 5
            GoTo ReTryValueEnumeration

        ElseIf lRtn = ERROR_NO_MORE_ITEMS Then
            ' -----------------------------------------------
            ' Not an error, just end of list -- exit the
            '   loop
            ' -----------------------------------------------
            lRtn = ERROR_SUCCESS
            Exit Do

        Else
            ' --------------------------------------------------
            ' Key still open, so display the error and fall
            '   thru to the close function below
            ' --------------------------------------------------
            
            Exit Do
        End If
    Loop
    
    ' -----------------------------------------------------
    ' Always close opened keys!
    ' -----------------------------------------------------
    lRtn = RegCloseKey(hKey)

End Sub

Private Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
   Dim nValue As Long:   Dim sValue As String
   Select Case lType
      Case REG_BINARY
            nValue = vValue
         SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, nValue, 4)
      Case REG_SZ
         sValue = vValue & Chr$(0)
         SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
      Case REG_DWORD
         nValue = vValue
         SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, nValue, 4)
   End Select
 End Function
Private Function GetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, nValue As Variant) As Long
    Dim sValue As String:    sValue = Space$(255)
    Select Case lType
      Case REG_BINARY
          GetValueEx = RegQueryValueExLong(hKey, sValueName, 0&, lType, nValue, 4)
      Case REG_SZ
          GetValueEx = RegQueryValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
                                nValue = StripNulls(sValue)
      Case REG_DWORD
          GetValueEx = RegQueryValueExLong(hKey, sValueName, 0&, lType, nValue, 4)
 End Select
 End Function
Private Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
    OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function
Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
     Dim hKey As Long
      Call RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, RX)
      Call RegCloseKey(hKey)
End Sub
Private Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long, t As Variant)
  Dim hKey As Long
  Call RegOpenKeyEx(t, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  Call SetValueEx(hKey, sValueName, lValueType, vValueSetting)
  Call RegCloseKey(hKey)
End Sub
Private Function GetKeyValue(sKeyName As String, sValueName As String, _
                       vValueSetting As Variant, lValueType As Long, t As Variant) As Long
  Dim hKey As Long
  Call RegOpenKeyEx(t, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  GetKeyValue = GetValueEx(hKey, sValueName, lValueType, vValueSetting)
  Call RegCloseKey(hKey)
End Function
Private Sub DeleteKey(t As Variant, sKey As String, SubKey As String)
Dim hKey As Long
  Call RegOpenKeyEx(t, sKey, 0, KEY_ALL_ACCESS, hKey)
  Call RegDeleteKey(hKey, SubKey)
  Call RegCloseKey(hKey)
End Sub
Private Sub DeleteValue(t As Variant, sKey As String, sValueName As String)
Dim hKey As Long
  Call RegOpenKeyEx(t, sKey, 0, KEY_ALL_ACCESS, hKey)
  Call RegDeleteValue(hKey, sValueName)
  Call RegCloseKey(hKey)
End Sub
Private Function TrimNull(Item As String)
    Dim pos As Integer:    pos = InStr(Item, Chr$(0))
    If pos Then
          TrimNull = Left$(Item, pos - 1)
    Else: TrimNull = Item
    End If
End Function

