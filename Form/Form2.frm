VERSION 5.00
Begin VB.Form Formaa 
   BackColor       =   &H80000008&
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7635
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5820
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5820
      Left            =   -45
      ScaleHeight     =   5760
      ScaleWidth      =   7620
      TabIndex        =   0
      Top             =   -30
      Width           =   7680
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   240
         Left            =   5475
         TabIndex        =   3
         Top             =   4305
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.FileListBox File1 
         Height          =   480
         Left            =   5715
         Pattern         =   "*.bmp"
         TabIndex        =   2
         Top             =   4755
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   2910
         Top             =   4980
      End
      Begin VB.Image tmpimg 
         Height          =   660
         Left            =   30
         Top             =   30
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   -1000
      TabIndex        =   1
      Text            =   "Available Frames to edit"
      Top             =   -1000
      Width           =   7560
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run"
   End
   Begin VB.Menu mnuPause 
      Caption         =   "Pause"
   End
   Begin VB.Menu mnuDF 
      Caption         =   "Delete Frame"
   End
   Begin VB.Menu mnySA 
      Caption         =   "Save AVI"
   End
End
Attribute VB_Name = "Formaa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pix As Long

Private Sub Combo1_Change()
Command9_Click

End Sub

Private Sub Combo1_Click()
Command9_Click

End Sub

Private Sub Command9_Click()
'Dim tmpimg As VB.Image
'Set tmpimg = Controls.Add("VB.Image", "ctlName", Controller)
On Error Resume Next
tmpimg.Refresh
Picture1.Picture = LoadPicture("")
tmpimg.Picture = LoadPicture(App.path & "\Images\" & Combo1.List(Combo1.ListIndex))   'App.Path & "\Images\180705172743.BMP") 'change to your picture path

Dim xImg, yImg As Single
Dim xPic, yPic As Single
xImg = tmpimg.Width
yImg = tmpimg.Height
xPic = Picture1.Width
yPic = Picture1.Height

Dim xRatio, yRatio As Single
xRatio = xImg / xPic
yRatio = yImg / yPic

If xRatio >= yRatio Then
Picture1.PaintPicture tmpimg.Picture, 0, 0, (tmpimg.Width / xRatio), (tmpimg.Height / xRatio)
Else
Picture1.PaintPicture tmpimg.Picture, 0, 0, (tmpimg.Width / yRatio), (tmpimg.Height / yRatio)
End If
Picture1.Width = tmpimg.Width / yRatio
'Picture1.Height = tmpimg.Height / xRatio
Picture1.left = (Me.Width - Picture1.Width) / 2
Picture1.Refresh

End Sub

Private Sub Form_Activate()
Flicker = True
PauseFlag = False
Picture1.Refresh
'Combo1.SetFocus
End Sub

Private Sub Form_Load()
Picture1.top = 0
Dim i As Long
'SetTopMostWindow Me.hWnd, True
Me.Height = 4620
Me.Width = 4605
Me.top = 0
Me.left = 0
File1.path = App.path & "\Images"
File1.Refresh
Combo1.Clear
For i = 0 To File1.ListCount - 1
    Combo1.AddItem File1.List(i)
Next
Pix = -1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Combo1.SetFocus
End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.Width = Me.Width
Picture1.Height = Me.Height
Command9_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Set Formaa = Nothing
End Sub

Private Sub mnuDF_Click()
On Error Resume Next
mnuPause_Click
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to continue?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Delete Frame"   ' Define title.
Help = "DEMO.HLP"   ' Define Help file.
Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then   ' User chose Yes.
    Combo1.RemoveItem Combo1.ListIndex
    Kill App.path & "\Images\" & Combo1.List(Combo1.ListIndex)
End If


End Sub

Private Sub mnuPause_Click()
Timer1.Enabled = False
End Sub

Private Sub mnuRun_Click()
Timer1.Enabled = True
End Sub

Private Sub mnySA_Click()
    Shell App.path & "\AVICreator6.exe", vbNormalFocus
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Combo1.SetFocus

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Pix = Combo1.ListCount Then Pix = -1
Pix = Pix + 1
Combo1.ListIndex = Pix
Combo1.Text = Combo1.List(Pix)

End Sub
