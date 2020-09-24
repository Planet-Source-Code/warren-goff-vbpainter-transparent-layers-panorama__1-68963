VERSION 5.00
Begin VB.Form frmCaptur 
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
   MouseIcon       =   "frmCaptur.frx":0000
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
Attribute VB_Name = "frmCaptur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnCapturing As Boolean
Dim X1 As Single
Dim Y1 As Single

Private Sub Form_Activate()
  Dim Left As Long, Top As Long, Right As Long, Bottom As Long
  Dim lWidth As Long, lHeight As Long, MMM As String
lWidth = (frmPaint.picPaint.Width - 50) / Screen.TwipsPerPixelX '444
Left = (frmPaint.Left + frmPaint.picPaint.Left + 100) / Screen.TwipsPerPixelX '266
Right = Left + lWidth ' 710
Top = (frmPaint.Top + frmPaint.picPaint.Top + 485) / Screen.TwipsPerPixelY '123
lHeight = (frmPaint.picPaint.Height - 50) / Screen.TwipsPerPixelY ' 346
Bottom = (Top + lHeight)    ' / Screen.TwipsPerPixelY '469
  'If lWidth <= 0 Or lHeight <= 0 Then GoTo PROC_TOOSMALL  'Nothing to capture
  With picTemp
    .Cls  'Clear our picture box that holds the image till copied to clipboar
    .Width = lWidth 'Set it's hight and width
    .Height = lHeight
  End With
  Me.Cls  'Clear screen so we don't get the box and dimensions
  BitBlt picTemp.hDC, 0, 0, lWidth, lHeight, Me.hDC, Left, Top, SRCCOPY 'Copy screen to picture box
SavePicture picTemp.Image, App.path & "\" & "Captured.bmp"
Captured = True
frmPaint.picPaint.Picture = picTemp.Image

            frmPaint.TransPicBox1.Visible = False
            frmPaint.TransPicBox2.Visible = False
            frmPaint.TransPicBox3.Visible = False
            frmPaint.TransPicBox4.Visible = False
            frmPaint.TransPicBox5.Visible = False
            frmPaint.TransPicBox6.Visible = False
            frmPaint.TransPicBox7.Visible = False
            frmPaint.TransPicBox8.Visible = False
            frmPaint.TransPicBox9.Visible = False
            frmPaint.TransPicBox10.Visible = False

Unload Me

End Sub

Private Sub Form_Load()
  'Capture desktop and make it this forms background picture
  Dim DeskhWnd As Long, DeskDC As Long
  Me.WindowState = vbMaximized
  DeskhWnd& = GetDesktopWindow()
  DeskDC& = GetDC(DeskhWnd&)
  BitBlt Me.hDC, 0&, 0&, Screen.Width, Screen.Height, DeskDC&, 0&, 0&, SRCCOPY
  Me.Picture = Me.Image
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmPaint.Transister
Set frmCaptur = Nothing
End Sub

