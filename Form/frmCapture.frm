VERSION 5.00
Begin VB.Form frmCapture 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   DrawMode        =   6  'Mask Pen Not
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmCapture.frx":0000
   ScaleHeight     =   218
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   105
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   108
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnCapturing As Boolean
Dim X1 As Single
Dim Y1 As Single

Private Sub Form_Load()
On Error GoTo Errorr
  'Capture desktop and make it this forms background picture
  Dim DeskhWnd As Long, DeskDC As Long
  Me.WindowState = vbMaximized
  DeskhWnd& = GetDesktopWindow()
  DeskDC& = GetDC(DeskhWnd&)
  BitBlt Me.hDC, 0&, 0&, Screen.Width, Screen.Height, DeskDC&, 0&, 0&, SRCCOPY
  Me.Picture = Me.Image
Exit Sub
Errorr:
MsgBox "Unable to Perform capture at this time!  Memory Resources are Short!"
Unload Me
Set frmCapture = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'User pressed escape so unload
  If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not blnCapturing Then  'Start capture
    MousePointer = 99 'Change our mousepointer to custom
    X1 = X: Y1 = Y  'Set our starting x & y
    blnCapturing = True 'Turn capturing bit on
  ElseIf blnCapturing = True Then 'Done capturing
    If Button = vbRightButton Then  'User clicked right mouse so cancel but stay capturing
      blnCapturing = False  'Turn capturing bit off
      MousePointer = vbNormal 'Set our mousepointer back to normal
      Cls 'Clear anything we drew to the form
    ElseIf Button = vbLeftButton Then 'User clicked left mouse button so capture
      CaptureIt X1, X, Y1, Y  'Do the capture
      Unload Me 'Unload form
    End If
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If blnCapturing Then  'If we are capturing then draw box and dimensions
    Cls 'Clear the form
    Line (X1, Y1)-(X, Y), , B 'Draw our box where the mouse selection is
    
    'Get left, right, top and bottom regarldess of where they started and ended
    Dim Left As Long, Top As Long, Right As Long, Bottom As Long
    Dim lWidth As Long, lHeight As Long
    Left = IIf(X1 > X, X, X1)
    Right = IIf(X1 < X, X, X1)
    Top = IIf(Y1 > Y, Y, Y1)
    Bottom = IIf(Y1 < Y, Y, Y1)
    lWidth = (Right - Left)
    lHeight = (Bottom - Top)

    Dim strOut As String
    strOut = lWidth & "x" & lHeight 'Setup our dimensions string

    'If the text will fit in our selection then draw to screen
    If lWidth > TextWidth(strOut) And lHeight > TextHeight(strOut) Then

      Dim tX As Single, tY As Single
      Dim cx As Single, cy As Single
      cx = Right - (lWidth / 2) 'Get our center of our rectangle's x position
      cy = Bottom - (lHeight / 2) 'Get our center of our rectangle's y position
      tX = cx - TextWidth(strOut) / 2 'Get our offset from x center with text width
      tY = cy - TextHeight(strOut) / 2  'Get our offset from y center with text height

      If Me.Point(cx, cy) < 62255 / 2 Then  'Center of selection color was darker
        ForeColor = vbWhite 'Set font color to white
      Else  'Color was lighter
        ForeColor = vbBlack 'Set font color to black
      End If

      TextOut Me.hDC, tX, tY, strOut, Len(strOut) 'Draw our dimensions text on the form

    End If
  End If
End Sub

Private Sub CaptureIt(xStart As Single, xEnd As Single, yStart As Single, yEnd As Single)
  Dim Left As Long, Top As Long, Right As Long, Bottom As Long
  Dim lWidth As Long, lHeight As Long, MMM As String
On Error Resume Next
  blnCapturing = False

  'Get left, right, top and bottom regarldess of where they started and ended
  Left = IIf(xStart > xEnd, xEnd, xStart)
  Right = IIf(xStart < xEnd, xEnd, xStart)
  Top = IIf(yStart > yEnd, yEnd, yStart)
  Bottom = IIf(yStart < yEnd, yEnd, yStart)
  lWidth = (Right - Left)
  lHeight = (Bottom - Top)
  'DoEvents
 ' Open App.path & "\SSTT" For Output As #11
 '   Print #11, "Left= " & Left
 '   Print #11, "Right= " & Right
 '   Print #11, "Top= " & Top
 '   Print #11, "Bottom= " & Bottom
 '   Print #11, "lWidth= " & lWidth
 '   Print #11, "lHeight= " & lHeight
 'Close #11
  If lWidth <= 0 Or lHeight <= 0 Then GoTo PROC_TOOSMALL  'Nothing to capture
  
  With picTemp
    .Cls  'Clear our picture box that holds the image till copied to clipboar
    .Width = lWidth 'Set it's hight and width
    .Height = lHeight
  End With
  Me.Cls  'Clear screen so we don't get the box and dimensions
  BitBlt picTemp.hDC, 0, 0, lWidth, lHeight, Me.hDC, Left, Top, SRCCOPY 'Copy screen to picture box

  SavePicture picTemp.Image, App.path & "\" & "Captured.bmp"
  Captured = True

    Unload Me
PROC_EXIT:
  Exit Sub
  
PROC_TOOSMALL:
  GoTo PROC_EXIT
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPaint.Transister
Unload Me
Set frmCapture = Nothing

End Sub

