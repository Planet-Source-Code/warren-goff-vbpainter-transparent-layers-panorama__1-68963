Attribute VB_Name = "Global"
Option Explicit
Global Loopit As Boolean
Global Moose As Boolean
Global Scanned As Boolean
Global Scannn As Integer
Global x$
Global Effect As Boolean
Global Thumb As Boolean
Global Iconic As Boolean
Global DefaultThumb As String
Global BMPtoICO As String
Global Captured As Boolean
Global PictureExt(15) As String
Global VideoExt(10) As String
Global ShowOn As Boolean
Global j1 As Long
Global Crop As Boolean
Global Stringy As String
Global LoadRegion As Boolean
Global Painting As Boolean
Global Reverse As Boolean
Global Forward As Boolean
Global SpeedPer As Long
Global AsBut As Boolean
Global Patherino As String
Global Exterino As String
Global Mht As Boolean
Global MooseSnobbler As Boolean
Global PickFlag, ReplaceFlag As Boolean
Global PickColor, ReplaceColor As String
Global GlobalFIlter As String
Global frmCaptureFlag As Boolean
Global UpFlag As Boolean
Global Flicker As Boolean
Global SeeFlag As String
Global XX1 As Single
Global YY1 As Single
Global XX2 As Single
Global YY2 As Single
Global TheMovie As String
Global Filenamer As String
Global MoosePosition As Long
Global Schnorbel As Boolean
Global SlideShowFlag As Boolean
Global ImagesPath As String
Global ScannedO As Boolean
'Global IR As RECT
Global RegionsF As Long
Global RegionsFlag As Boolean
Global wID As String, Hei As String
Global PictTop As Long, PictLeft As Long
Global Outahere As Boolean, PasteClip As Boolean
Global Element As Integer
Global Elementary As Boolean
Global intsave As Integer

Public Enum enmStatusBar
  conStPaintArea = 0
  conStColorBox = 1
  conStForeColorBox = 2
  conStBackColorBox = 3
  conStFiltering = 4
  conStRetrieveingColor = 5
End Enum
Global sng As Single

Public Enum enmTool
  'the values below must match optTools index
  conTSelect = 0
  conTPick = 1
  conTEraser = 2
  conTFill = 3
  conTPencil = 4
  conTLine = 5
  conTRect = 6
  conTEllipse = 7
  conTText = 8
  conTArrow = 9
  conTAirBrush = 10
  conTRoundRect = 11
  conTPolygon = 12
  conTCurve = 13
  conTFilter = 14
  conTZoom = 15
  conTBrush = 16
  conTHand = 17
End Enum

Public Enum enmFillStyle
  conTsBorderOnly = 0
  conTsBorderFill = 1
  conTsFillOnly = 2
End Enum

Public Enum enmBrushShape
  'the values below must match imgBrush index
  conFilledRect = 0
  conFilledCircle = 1
  conRect = 2
  conCircle = 3
  conCross = 4
  conDiagonalCross = 5
  conUpwardDiagonal = 6
  conDownwardDiagonal = 7
  conHorizontal = 8
  conVertical = 9
End Enum



Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Const BN_CLICKED = 0
Private Const WM_COMMAND = &H111
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetDlgCtrlID Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Public Const SHERB_NOPROGRESSUI = &H2

Public mlngCurWindow As Long
Public mstrPhrase As String
'**************************************
'Windows API/Global Declarations for :In
'     stall a font (under 10 lines of code)
'**************************************
Public Const WM_FONTCHANGE = &H1D
Public Const HWND_BROADCAST = &HFFFF&


Declare Function AddFontResource& Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFilename As String)


Declare Function SendMessageBynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long)
'declare for moving the form
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Sub Delay(HowLong As Date)
Dim TempTIme As Long
TempTIme = DateAdd("s", HowLong, Now)
While TempTIme > Now
DoEvents 'Allows windows to handle other stuff
Wend
End Sub
Public Sub UnloadAll(Optional activefrm As Form)
Dim frm As Form
Dim ctl As Control
On Error Resume Next
        For Each frm In Forms
        ' Don't unload active form or we will reload it when we return to it
        ' Allow active form to unload itself

        If frm.Name <> activefrm.Name Then

            ' This is extra protection and may not be needed

            For Each ctl In frm.Controls
                Set ctl = Nothing
            Next

            Unload frm
            Set frm = Nothing

        End If

    Next

End Sub


Public Function GetFileExtension(filename As String)
    On Error Resume Next
    Dim TempStr As String
    TempStr = Right(filename, 2)


    If Left(TempStr, 1) = "." Then
        GetFileExtension = Right(filename, 1)
        Exit Function
    Else
        TempStr = Right(filename, 3)


        If Left(TempStr, 1) = "." Then
            GetFileExtension = Right(filename, 2)
            Exit Function
        Else
            TempStr = Right(filename, 4)


            If Left(TempStr, 1) = "." Then
                GetFileExtension = Right(filename, 3)
                Exit Function
            Else
                TempStr = Right(filename, 5)


                If Left(TempStr, 1) = "." Then
                    GetFileExtension = Right(filename, 4)
                    Exit Function
                Else
                    GetFileExtension = "Unknown"
                End If
            End If
        End If
    End If
    
End Function

'Function called each time EnumChildWindows finds a child
Public Function EnumAChild(ByVal hwnd As Long, ByVal lParam As Long) As Boolean

Dim lngButtonID As Long
Dim lngLength As Long
Dim strTitle As String
Dim lngResult As Long
    
    EnumAChild = True 'Let EnumChildWindows find another child
    lngLength = GetWindowTextLength(hwnd) 'get length of the caption
    If lngLength > 0 Then
        strTitle = String$(100, Chr(0))
        'get the caption
        lngResult = GetWindowText(hwnd, strTitle, lngLength + 1) 'length + \0
        strTitle = Left$(strTitle, lngLength)
        'if the caption of the control is the same as the word you said
        If UCase(strTitle) = UCase(mstrPhrase) Or (UCase(strTitle) = "OK" And UCase(mstrPhrase) = "OKAY") Then
            lngButtonID = GetDlgCtrlID(hwnd)
            lngResult = PostMessage(mlngCurWindow, WM_COMMAND, lngButtonID, BN_CLICKED * &H10000 + hwnd)
            If lngResult <> 0 Then
                'You found what you wanted, stop the EnumChildWindows function
                EnumAChild = False
            End If
        End If
    End If

End Function
