VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VideoLibrary 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9735
   ClientLeft      =   -990
   ClientTop       =   -990
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000007&
   Icon            =   "VideoLibrary.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "VideoLibrary.frx":08CA
   ScaleHeight     =   649
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   Begin VB.CommandButton Command21 
      Height          =   225
      Left            =   7665
      TabIndex        =   86
      Top             =   8055
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   7515
      Top             =   5910
   End
   Begin VB.CommandButton AnyShape26 
      BackColor       =   &H000C4ACC&
      Height          =   495
      Left            =   7425
      Picture         =   "VideoLibrary.frx":16726
      Style           =   1  'Graphical
      TabIndex        =   85
      ToolTipText     =   "Open Multiple Files"
      Top             =   7440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Haburabadooda27 
      BackColor       =   &H00244ABC&
      Height          =   555
      Left            =   7440
      Picture         =   "VideoLibrary.frx":16C71
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Loads default directory c:\1down\"
      Top             =   6450
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H000C4ACC&
      Height          =   495
      Left            =   7440
      Picture         =   "VideoLibrary.frx":172B4
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "Browse for a Directory of Files to load"
      Top             =   7005
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H80000008&
      Caption         =   "Select"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   8925
      Value           =   -1  'True
      Width           =   750
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H80000008&
      Caption         =   "Loop"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   9150
      Width           =   750
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H80000008&
      Caption         =   "Random"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   9360
      Width           =   750
   End
   Begin VB.CommandButton Haburabadooda12 
      BackColor       =   &H0074BADC&
      Height          =   195
      Left            =   6960
      Picture         =   "VideoLibrary.frx":17B7E
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "Minimize"
      Top             =   45
      Width           =   345
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0074BADC&
      Height          =   345
      Left            =   7335
      Picture         =   "VideoLibrary.frx":17F7D
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Toggle Keep on Top"
      Top             =   945
      Width           =   495
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H000000FF&
      Caption         =   "mp3"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   6765
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0000FFFF&
      Caption         =   "Ren"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Ok for this Slider"
      Top             =   6540
      Width           =   480
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00333333&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   1695
      TabIndex        =   75
      Top             =   6585
      Width           =   4680
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0000FF00&
      Caption         =   "View Clip"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   9420
      Width           =   1155
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1035
      Picture         =   "VideoLibrary.frx":18847
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9030
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FFFF&
      Caption         =   "MPG 2 AVI"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   405
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9030
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   210
      TabIndex        =   64
      Top             =   8415
      Visible         =   0   'False
      Width           =   4185
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Height          =   495
         Left            =   195
         Picture         =   "VideoLibrary.frx":18C89
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   75
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0000FFFF&
         Caption         =   "  AVI    2 BMP"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   1425
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0000FFFF&
         Caption         =   "AVI  2 MPG"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   765
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   90
         Width           =   555
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H000000FF&
         Caption         =   " AVI IN"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   2955
         TabIndex        =   69
         Top             =   120
         Width           =   570
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   3540
         TabIndex        =   68
         Top             =   120
         Width           =   570
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFF00&
         Caption         =   "Process"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2955
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   720
         Width           =   1170
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FF8080&
         Caption         =   "Begin"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2955
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   420
         Width           =   570
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FF8080&
         Caption         =   "End"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   420
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   1470
         TabIndex        =   74
         Top             =   135
         Width           =   1365
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Wordpro"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   14
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Send to WordPro"
      Top             =   9210
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Science"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   13
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Send to Science"
      Top             =   9210
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Telecom"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   12
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Send to Telecom"
      Top             =   9210
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Music"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   11
      Left            =   4485
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Send to Music"
      Top             =   9210
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Multi"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   10
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Send to Multimedia"
      Top             =   9210
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Personal"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   9
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Send to Personal"
      Top             =   8490
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Util"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Send to Utilities"
      Top             =   8880
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Gif"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   4485
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Send to Gif"
      Top             =   8880
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Legal"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   1
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Send to legal"
      Top             =   8880
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Receipt"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Send to Receipts"
      Top             =   8880
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "1down"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   3
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Send to 1down"
      Top             =   8880
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Mics"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   4
      Left            =   4485
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Send to microphones"
      Top             =   8490
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Vb"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Send to Vb"
      Top             =   8490
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Guitar"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   6
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Send to Guitar"
      Top             =   8490
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00242614&
      Caption         =   "Medical"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Send to Medical"
      Top             =   8490
      Width           =   600
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   7095
      TabIndex        =   42
      Top             =   0
      Width           =   840
      Begin VB.Image Image1 
         Height          =   480
         Left            =   225
         Picture         =   "VideoLibrary.frx":19ACB
         Top             =   90
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   300
         Picture         =   "VideoLibrary.frx":1A395
         Top             =   120
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin AngstArt.ocxFormShape ocxFormShape1 
      Left            =   7650
      Top             =   8160
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H000080FF&
      Caption         =   "Full Screen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4950
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7470
      Width           =   1305
   End
   Begin VB.OptionButton Options 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Index           =   4
      Left            =   7530
      Picture         =   "VideoLibrary.frx":1AC5F
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9975
      Width           =   510
   End
   Begin VB.ComboBox Moviee 
      BackColor       =   &H00000000&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      ItemData        =   "VideoLibrary.frx":1B92D
      Left            =   330
      List            =   "VideoLibrary.frx":1B92F
      TabIndex        =   35
      Text            =   "Video Cache"
      Top             =   8115
      Width           =   7110
   End
   Begin VB.CommandButton Command6 
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
      Height          =   495
      Left            =   2865
      Picture         =   "VideoLibrary.frx":1B931
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   7470
      Width           =   495
   End
   Begin VB.Frame frainfo 
      BackColor       =   &H80000012&
      Caption         =   "Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   975
      Left            =   11235
      TabIndex        =   25
      Top             =   3645
      Visible         =   0   'False
      Width           =   4980
      Begin VB.TextBox txtlength 
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtplayback 
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtremain 
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtelapse 
         BackColor       =   &H00808080&
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbllength 
         BackStyle       =   0  'Transparent
         Caption         =   "Length:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblelapse 
         Caption         =   "Elapsed Time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblplayback 
         BackStyle       =   0  'Transparent
         Caption         =   "Playback Speed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   2235
         TabIndex        =   31
         Top             =   255
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblremain 
         Caption         =   "Time Remaining:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   30
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog cmdopen 
      Left            =   7695
      Top             =   9870
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7575
      Top             =   8985
   End
   Begin VB.OptionButton Options 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   3420
      Picture         =   "VideoLibrary.frx":1BD73
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Play"
      Top             =   7470
      Width           =   375
   End
   Begin VB.OptionButton Options 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   495
      Index           =   2
      Left            =   3780
      Picture         =   "VideoLibrary.frx":1BFC9
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Pause"
      Top             =   7470
      Width           =   375
   End
   Begin VB.OptionButton Options 
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      Height          =   495
      Index           =   3
      Left            =   4140
      Picture         =   "VideoLibrary.frx":1C21F
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Stop"
      Top             =   7470
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton Options 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Index           =   5
      Left            =   7710
      Picture         =   "VideoLibrary.frx":1C475
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9900
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Frame frabal 
      BackColor       =   &H00336699&
      BorderStyle     =   0  'None
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Left            =   -30
      TabIndex        =   11
      Top             =   1380
      Width           =   2550
      Begin MSComctlLib.Slider sldbalance 
         Height          =   195
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   344
         _Version        =   393216
         LargeChange     =   1500
         SmallChange     =   500
         Min             =   -5000
         Max             =   5000
         TickFrequency   =   1000
      End
      Begin VB.Label lblleft 
         BackStyle       =   0  'Transparent
         Caption         =   "Left"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   225
         Width           =   735
      End
      Begin VB.Label lblright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   1485
         TabIndex        =   13
         Top             =   225
         Width           =   825
      End
   End
   Begin VB.Frame fravol 
      BackColor       =   &H00336699&
      BorderStyle     =   0  'None
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Left            =   4755
      TabIndex        =   15
      Top             =   1380
      Width           =   3015
      Begin MSComctlLib.Slider sldvolume 
         Height          =   195
         Left            =   30
         TabIndex        =   16
         Top             =   30
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   344
         _Version        =   393216
         LargeChange     =   25
         Min             =   -5000
         Max             =   0
         TickFrequency   =   500
      End
      Begin VB.Label lblmin 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   225
         Width           =   735
      End
      Begin VB.Label lblmax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   255
         Left            =   2355
         TabIndex        =   17
         Top             =   225
         Width           =   615
      End
   End
   Begin VB.Frame fraplayrate 
      BackColor       =   &H00336699&
      BorderStyle     =   0  'None
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   570
      Left            =   2475
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   1380
      Width           =   2430
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3435
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   900
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Caption         =   "50 %"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   930
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF00&
         Caption         =   "200%"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "250%"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   900
         Visible         =   0   'False
         Width           =   465
      End
      Begin MSComctlLib.Slider sldplayrate 
         Height          =   210
         Left            =   0
         TabIndex        =   8
         Top             =   30
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   370
         _Version        =   393216
         LargeChange     =   2
         Min             =   75
         Max             =   150
         SelStart        =   100
         TickFrequency   =   5
         Value           =   100
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0FFC0&
         BorderWidth     =   3
         X1              =   810
         X2              =   810
         Y1              =   240
         Y2              =   350
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   570
         TabIndex        =   87
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "75%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   30
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "150%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D8E9EC&
      BorderStyle     =   0  'None
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   375
      TabIndex        =   1
      Top             =   7065
      Width           =   7005
      Begin MSComctlLib.Slider Slider1 
         Height          =   225
         Left            =   45
         TabIndex        =   2
         Top             =   15
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   226
         SelStart        =   1
         TickStyle       =   3
         TickFrequency   =   10
         Value           =   1
      End
      Begin VB.Label Lab2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   5520
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Lab1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   45
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   1140
         X2              =   1140
         Y1              =   915
         Y2              =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   1260
         X2              =   1260
         Y1              =   900
         Y2              =   1305
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   7125
      TabIndex        =   40
      Top             =   9540
      Visible         =   0   'False
      Width           =   1605
   End
   Begin AngstArt.FileOPS FileOPS1 
      Height          =   465
      Left            =   6285
      TabIndex        =   43
      Top             =   765
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   820
   End
   Begin VB.PictureBox picVideoWindow 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00666666&
      BorderStyle     =   0  'None
      DrawMode        =   1  'Blackness
      DrawStyle       =   5  'Transparent
      ForeColor       =   &H80000004&
      Height          =   4185
      Left            =   1065
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   4185
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   2160
      Width           =   5715
   End
   Begin VB.CommandButton Command19 
      Height          =   495
      Left            =   6345
      Picture         =   "VideoLibrary.frx":1D143
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   7470
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Frame Number = "
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   705
      TabIndex        =   62
      Top             =   7770
      Width           =   1605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Frames = "
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   690
      TabIndex        =   61
      Top             =   7515
      Width           =   1605
   End
   Begin VB.Label MovieName 
      BackColor       =   &H00CC0033&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   270
      Left            =   2370
      TabIndex        =   24
      Top             =   6645
      Width           =   2970
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Height          =   960
      Left            =   75
      TabIndex        =   59
      Top             =   8415
      Width           =   7695
   End
   Begin VB.Menu mnuLoop 
      Caption         =   "Menu Loop"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "VideoLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim Point1, Point2 As Long

Dim rtn As Long

'Tells us wheather the movie/music is looping
Private bloop As Boolean
'Same except tells us if fullscreen
Private bfullscreen As Boolean
'How fast the media is playing
Private playrate As Double
Private mcontrol As New FilgraphManager
Private audio As IBasicAudio
Private video As IVideoWindow
Private mposition As IMediaPosition
Dim Loopit As Boolean
Dim MoveFlag As Boolean
Dim Pix As Long
Dim PositionSec As Single
Dim FrameNumber As Long
Dim FrameSec As Single
Dim DurationSec As Single
Dim i As Long

Dim InitFlag As Boolean

Private Sub Combo1_Change()
Command12_Click

End Sub

Private Sub Combo1_Click()
Load Formaa
Formaa.Show
Command12_Click

End Sub

Private Sub AnyShape26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
                ''SetTopMostWindow Me.hWnd, False
                'Me.Hide
                On Error Resume Next
                Dim M1, M2, M3 As String
                Filenamer = ""
                'Option8.Value = False
                'Set our controls to nothing. If we didnt, we would here any music or sound
                'from the previous file.
                Set mcontrol = Nothing
                Set audio = Nothing
                Set mposition = Nothing
                'CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
                  "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
                M1 = "Video Files (*.mpg;*.mpeg;*.m2v;*.avi;*.asf;*.mov|*.mpg;*.mpeg;*.m2v;*.avi;*.asf;*.mov"
'                M1 = "Media Files (*.bmp;*.jpg;*.mpg;*.mpeg;*.mid;*.avi;*.asf*.mov;*.wav;*.mp2;*.mp3)|*.bmp;*.jpg;*.mpg;*.mpeg;*.mid;*.avi;*.asf;*.mov;*.wav;*.mp2;*.mp3"
                M1 = M1 & "|Mpg Files" & "  (*.mpg)|*.mpg"
                M1 = M1 & "|Mpeg Files" & "  (*.mpeg)|*.mpeg"
                M1 = M1 & "|Asf Files" & "  (*.asf)|*.asf"
                M1 = M1 & "|Avi Files" & "  (*.avi)|*.avi"
                M1 = M1 & "|Mov Files" & "  (*.mov)|*.mov"
                M1 = M1 & "|Mp3 Files" & "  (*.mp3)|*.mp3"
                M1 = M1 & "|Mid Files" & "  (*.mid)|*.mid"
                M1 = M1 & "|Wav Files" & "  (*.wav)|*.wav"
                M1 = M1 & "|Mp2 Files" & "  (*.mp2)|*.mp2"
                M1 = M1 & "|Jpg Files" & "  (*.jpg)|*.jpg"
                M1 = M1 & "|Bmp Files" & "  (*.bmp)|*.bmp"
                M1 = M1 & "|Wav Files" & "  (*.wav)|*.wav"
                M1 = M1 & "|Mp3 Files" & "  (*.mp3)|*.bmp"
                M1 = M1 & "|WMA Files" & "  (*.wma)|*.bmp"
                
                'Sets so only *.mpg, *.avi etc can be opened and viewed
                cmdopen.Filter = M1   '"Media Files (*.bmp;*.jpg;*.mpg;*.avi;*.mov;*.wav;*.mp2;*.mp3)|*.bmp;*.jpg;*.mpg;*.avi;*.mov;*.wav;*.mp2;*.mp3"
                'Shows the open dialog box.
                cmdopen.ShowOpen
                'store the filename into a variable
                Filenamer = cmdopen.filename
                
                TheMovie = Filenamer
                If LCase(Right(TheMovie, 3)) = "avi" Then
                    Command7.Visible = True
                    Frame2.Visible = True
                Else
                    Command7.Visible = False
                    Frame2.Visible = False
                End If
                If LCase(Right(TheMovie, 3)) = "mpg" Or LCase(Right(TheMovie, 4)) = "mpeg" Then
                    Command8.Visible = True
                Else
                    Command8.Visible = False
                End If

                MovieName.Caption = GetFileTitle(cmdopen.filename)
                'Renders the file so that it can be played. Otherwise you will not
                'here anything.
                
                'Refrences our audio the the main object
                Set audio = mcontrol
                'Sets the volume and balance from the slider.
                audio.Volume = 0 'sldvolume.Value
                sldvolume.Value = 0
                audio.Balance = sldbalance.Value
                'Refrences our video to the main object
                mcontrol.RenderFile Filenamer
                Set video = mcontrol
                video.WindowStyle = CLng(&H6000000)
                video.AutoShow = True
                video.Top = 0  'picVideoWindow.Top - 50
                'This makes it so that no caption appears in the picture box
                video.Left = 0  'picVideoWindow.Left
                'sets the height and width to the same as the picture box.
                'Otherwise not all of the movie would be seen.
                video.Height = picVideoWindow.Height
                video.Width = picVideoWindow.Width
                'This sets where the video will be run from.
                'You can run it from anything such as a textbox, frame or even the slider control!
                video.owner = picVideoWindow.hWnd
                'Me.Top = Me.Top + 1
                'refrences our positioning to the main object
                Set mposition = mcontrol
                
                mposition.Rate = 1
                
                'txtlength = Round(mposition.Duration, 2)
                
                
                Slider1.Max = mposition.Duration
                ''SetTopMostWindow Me.hWnd, True
                Me.Show
                Me.Refresh
                mcontrol.Run

End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then
    SetTopMostWindow Me.hWnd, True
Else
    SetTopMostWindow Me.hWnd, False
End If

End Sub

Private Sub Command10_Click()
On Error Resume Next
Kill App.path & "\Images\*.*"
Kill App.path & "\MyAvi.avi"
'MsgBox TheMovie
FileCopy TheMovie, App.path & "\MyAvi.avi"
Shell App.path & "\AVI2BMP.exe", vbNormalFocus
ImagesPath = App.path & "\Images"
'Unload Me
End Sub

Private Sub Command11_Click()
    Dim Mpgest, Mpger, TheMovie As String
    SetTopMostWindow Me.hWnd, False
    Dim RetVal
    TheMovie = Moviee.Text
    Mpger = Replace(TheMovie, ".", "") & ".mpg"
    'MsgBox Mpger
    'Mpger = App.path & "\MyMPEG1.mpg"
    Mpgest = App.path & "\avi2mpg1.exe -f1 " & TheMovie ' & " " & Mpger
    RetVal = Shell(Mpgest, 1)

End Sub

Private Sub Command12_Click()
    Load Formaa
    Formaa.Show
End Sub

Private Sub Command13_Click()
    Dim Res As Long         'result code
    Dim ofd As cFileDlg     'OpenFileDialog class
    Dim szFile As String    'filename
    Dim pAVIFile As Long    'pointer to AVI file interface (PAVIFILE handle)
    Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
    Dim numFrames As Long   'number of frames in video stream
    Dim firstFrame As Long  'position of the first video frame
    Dim fileInfo As AVI_FILE_INFO       'file info struct
    Dim streamInfo As AVI_STREAM_INFO   'stream info struct
    Dim dib As cDib
    Dim pGetFrameObj As Long    'pointer to GetFrame interface
    Dim pDIB As Long            'pointer to packed DIB in memory
    Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
    Dim i As Long
    Dim szzFile As String, sFile As String
    Dim message, Title, Default, MyValue As String
    Dim ii As Long, FN As Long
On Error Resume Next
Options_Click (2)
    Screen.MousePointer = 11
    File1.path = App.path & "\Images"
    Kill App.path & "\Images\*.*"
    File1.Refresh
    szFile = TheMovie
    
    'Open the AVI File and get a file interface pointer (PAVIFILE)
    Res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If Res <> AVIERR_OK Then GoTo ErrorOut
 
    'Get the first available video stream (PAVISTREAM)
    Res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut 'this function returns -1 on error
    
    'get the length of video stream in frames
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error
    PositionSec = Val(Trim(mposition.CurrentPosition))
    DurationSec = Val(mposition.Duration)
    FrameSec = numFrames / DurationSec
    FrameNumber = FrameSec * PositionSec
   
'    MsgBox "PAVISTREAM handle is " & pAVIStream & vbCrLf & _
'            "Video stream length - " & numFrames & vbCrLf & _
'            "Stream starts on frame #" & firstFrame & vbCrLf & _
'            "File and Stream info will be written to Immediate Window (from IDE - Ctrl+G to view)", vbInformation, App.title
'
    'get file info struct (UDT)
    Res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print file info to Debug Window
'    Call DebugPrintAVIFileInfo(fileInfo)
    
    'get stream info struct (UDT)
    Res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print stream info to Debug Window
'    Call DebugPrintAVIStreamInfo(streamInfo)
    
    'set bih attributes which we want GetFrame functions to return
    With bih
        .biBitCount = 24
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = BI_RGB
        .biHeight = streamInfo.rcFrame.Bottom - streamInfo.rcFrame.Top
        .biPlanes = 1
        .biSize = 40
        .biWidth = streamInfo.rcFrame.Right - streamInfo.rcFrame.Left
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight 'calculate total size of RGBQUAD scanlines (DWORD aligned)
    End With
    
    'init AVISTreamGetFrame* functions and create GETFRAME object
    'pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, ByVal AVIGETFRAMEF_BESTDISPLAYFMT) 'tell AVIStream API what format we expect and input stream
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih) 'force function to return 24bit DIBS
    If pGetFrameObj = 0 Then
        MsgBox "No suitable decompressor found for this video stream!", vbInformation, App.Title
        GoTo ErrorOut
    End If
    
    If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Then
        GoTo ErrorOut
        Exit Sub
        message = "Total Frames= " & numFrames   ' Set prompt.
        Title = "Choose total frames to process!!"   ' Set title.
        Default = Str(numFrames) ' Set default.
        
        MyValue = InputBox(message, Title, Default, 100, 100)
        ii = Val(MyValue)
    Else
        firstFrame = Val(Text1.Text)
        ii = Val(Text2.Text)
    End If
    'create a DIB class to load the frames into
    Set dib = New cDib
    Load Formaa
    FN = Val(Text1.Text)
    For i = firstFrame To ii + firstFrame  '(numFrames - 1) + firstFrame
    DoEvents
        If i / 2 = Int(i / 2) Then
            Label4.Caption = "---"
        Else
            Label4.Caption = " | "
        End If
        
        pDIB = AVIStreamGetFrame(pGetFrameObj, i)  'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) Then
        
here:
        Select Case Len(Trim(szzFile))
            Case 1
                szzFile = "0000" & szzFile
            Case 2
                szzFile = "000" & szzFile
            Case 3
                szzFile = "00" & szzFile
            Case 4
                szzFile = "0" & szzFile
        End Select
        szzFile = Replace(szzFile, " ", "")
        sFile = App.path & "\Images\" & szzFile & ".bmp"
        If Dir(sFile) <> "" Then szzFile = Str(Val(szzFile) + 1): GoTo here
        'MsgBox sFile
        Formaa.Combo1.AddItem szzFile & ".bmp"
            Call dib.WriteToFile(sFile)       '(App.Path & "\" & i & ".bmp")
        Else
            
        End If
    DoEvents
    Next
    
    Set dib = Nothing
    Formaa.Show
    Label4.Caption = ""
    Screen.MousePointer = Default

ErrorOut:
    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources and interface
    End If
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream) '//closes video stream
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    If (Res <> AVIERR_OK) Then 'if there was an error then show feedback to user
        MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.Title
    End If
    Screen.MousePointer = Default

End Sub

Private Sub Command14_Click()
Text1.Text = FrameNumber
End Sub

Private Sub Command15_Click()
    Dim Res As Long         'result code
    Dim ofd As cFileDlg     'OpenFileDialog class
    Dim szFile As String    'filename
    Dim pAVIFile As Long    'pointer to AVI file interface (PAVIFILE handle)
    Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
    Dim numFrames As Long   'number of frames in video stream
    Dim firstFrame As Long  'position of the first video frame
    Dim fileInfo As AVI_FILE_INFO       'file info struct
    Dim streamInfo As AVI_STREAM_INFO   'stream info struct
    Dim dib As cDib
    Dim pGetFrameObj As Long    'pointer to GetFrame interface
    Dim pDIB As Long            'pointer to packed DIB in memory
    Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
    Dim i As Long
    Dim szzFile As String, sFile As String
    Dim message, Title, Default, MyValue As String
    Dim ii As Long
    Load Formaa
    File1.path = App.path & "\Images"
    File1.Refresh
    Formaa.Combo1.Clear
    
    'Get the name of an AVI file to work with
    Set ofd = New cFileDlg
    With ofd
        .OwnerHwnd = Me.hWnd
        .Filter = "AVI Files|*.avi"
        .DlgTitle = "Open AVI File"
    End With
    Res = ofd.VBGetOpenFileNamePreview(szFile)
    If Res = False Then GoTo ErrorOut
    
    'Open the AVI File and get a file interface pointer (PAVIFILE)
    Res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If Res <> AVIERR_OK Then GoTo ErrorOut
 
    'Get the first available video stream (PAVISTREAM)
    Res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut 'this function returns -1 on error
    
    'get the length of video stream in frames
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error
    
'    MsgBox "PAVISTREAM handle is " & pAVIStream & vbCrLf & _
'            "Video stream length - " & numFrames & vbCrLf & _
'            "Stream starts on frame #" & firstFrame & vbCrLf & _
'            "File and Stream info will be written to Immediate Window (from IDE - Ctrl+G to view)", vbInformation, App.title
'
    'get file info struct (UDT)
    Res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print file info to Debug Window
'    Call DebugPrintAVIFileInfo(fileInfo)
    
    'get stream info struct (UDT)
    Res = AVIStreamInfo(pAVIStream, streamInfo, Len(streamInfo))
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
'    'print stream info to Debug Window
'    Call DebugPrintAVIStreamInfo(streamInfo)
    
    'set bih attributes which we want GetFrame functions to return
    With bih
        .biBitCount = 24
        .biClrImportant = 0
        .biClrUsed = 0
        .biCompression = BI_RGB
        .biHeight = streamInfo.rcFrame.Bottom - streamInfo.rcFrame.Top
        .biPlanes = 1
        .biSize = 40
        .biWidth = streamInfo.rcFrame.Right - streamInfo.rcFrame.Left
        .biXPelsPerMeter = 0
        .biYPelsPerMeter = 0
        .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight 'calculate total size of RGBQUAD scanlines (DWORD aligned)
    End With
    
    'init AVISTreamGetFrame* functions and create GETFRAME object
    'pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, ByVal AVIGETFRAMEF_BESTDISPLAYFMT) 'tell AVIStream API what format we expect and input stream
    pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih) 'force function to return 24bit DIBS
    If pGetFrameObj = 0 Then
        MsgBox "No suitable decompressor found for this video stream!", vbInformation, App.Title
        GoTo ErrorOut
    End If
    
    If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Then
        message = "Total Frames= " & numFrames   ' Set prompt.
        Title = "Choose total frames to process!!"   ' Set title.
        Default = Str(numFrames) ' Set default.
        
        MyValue = InputBox(message, Title, Default, 100, 100)
        ii = Val(MyValue)
    Else
        firstFrame = Val(Text1.Text)
        ii = Val(Text2.Text)
    End If
    'create a DIB class to load the frames into
    Set dib = New cDib
    For i = firstFrame To ii + firstFrame  '(numFrames - 1) + firstFrame
        pDIB = AVIStreamGetFrame(pGetFrameObj, i)  'returns "packed DIB"
        If dib.CreateFromPackedDIBPointer(pDIB) Then
        

here:
        Select Case Len(Trim(szzFile))
            Case 1
                szzFile = "0000" & szzFile
            Case 2
                szzFile = "000" & szzFile
            Case 3
                szzFile = "00" & szzFile
            Case 4
                szzFile = "0" & szzFile
        End Select
        szzFile = Replace(szzFile, " ", "")
        sFile = App.path & "\Images\" & szzFile & ".bmp"
        If Dir(sFile) <> "" Then szzFile = Str(Val(szzFile) + 1): GoTo here
        'MsgBox sFile
        Formaa.Combo1.AddItem szzFile & ".bmp"
            Call dib.WriteToFile(sFile)       '(App.Path & "\" & i & ".bmp")
        Else
            
        End If
    Next
    
    Set dib = Nothing
    Formaa.Show


ErrorOut:
    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources and interface
    End If
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream) '//closes video stream
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    If (Res <> AVIERR_OK) Then 'if there was an error then show feedback to user
        MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.Title
    End If


End Sub

Private Sub Command16_Click()
Text2.Text = FrameNumber

End Sub

Private Sub Command17_Click()
On Error Resume Next
If Moviee.Text = "" Then Exit Sub
Dim ActionFlag As Long
Dim i As Integer, k As Long, TempFile As String
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Me.Visible = False
Msg = "Are You sure you want to Rename: " & Filenamer & " ?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Rename Media"   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then   ' User chose Yes.
        k = Moviee.ListIndex
        'Kill Moviee.Text
        TempFile = Patherino & Text3.Text & Exterino
        Name TheMovie As TempFile
        Moviee.RemoveItem k
        Moviee.ListIndex = k
        Moviee.Text = Moviee.List(k)
        Set mcontrol = Nothing
        Set audio = Nothing
        Set mposition = Nothing
        Set audio = mcontrol
        mcontrol.RenderFile Moviee.Text
        Set video = mcontrol
        video.WindowStyle = CLng(&H6000000)
        video.AutoShow = True
        video.Top = 0  'picVideoWindow.Top - 50
        video.Left = 0  'picVideoWindow.Left
        video.Height = picVideoWindow.Height
        video.Width = picVideoWindow.Width
        video.owner = picVideoWindow.hWnd
        Set mposition = mcontrol
        mposition.Rate = 1
        txtlength = Round(mposition.Duration, 2)
        Slider1.Max = txtlength     'mposition.Duration
        Me.Show
        Me.Refresh
        mcontrol.Run
        sldbalance.Refresh
        sldvolume.Refresh
        sldplayrate.Refresh
        Slider1.Value = 1
        Slider1.Refresh
        sldvolume.Value = -1000
        audio.Volume = -1000  'sldvolume.Value
End If
Me.Visible = True
    
End Sub

Private Sub Command18_Click()
Moviee.Text = "Loading " & Moviee.ListCount & " Files into Mp3-Renamer.  Please Wait!"
Shell App.path & "\Mp3renamer.exe", vbNormalFocus
End Sub

Private Sub Command19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
                ''SetTopMostWindow Me.hWnd, False
                'Me.Hide
                On Error Resume Next
                Dim M1, M2, M3 As String
                Filenamer = ""
                'Option8.Value = False
                'Set our controls to nothing. If we didnt, we would here any music or sound
                'from the previous file.
                Set mcontrol = Nothing
                Set audio = Nothing
                Set mposition = Nothing
                'CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
                  "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"
                M1 = "Video Files (*.mpg;*.mpeg;*.m2v;*.avi;*.asf;*.mov|*.mpg;*.mpeg;*.m2v;*.avi;*.asf;*.mov"
'                M1 = "Media Files (*.bmp;*.jpg;*.mpg;*.mpeg;*.mid;*.avi;*.asf*.mov;*.wav;*.mp2;*.mp3)|*.bmp;*.jpg;*.mpg;*.mpeg;*.mid;*.avi;*.asf;*.mov;*.wav;*.mp2;*.mp3"
                M1 = M1 & "|Mpg Files" & "  (*.mpg)|*.mpg"
                M1 = M1 & "|Mpeg Files" & "  (*.mpeg)|*.mpeg"
                M1 = M1 & "|Asf Files" & "  (*.asf)|*.asf"
                M1 = M1 & "|Avi Files" & "  (*.avi)|*.avi"
                M1 = M1 & "|Mov Files" & "  (*.mov)|*.mov"
                M1 = M1 & "|Mp3 Files" & "  (*.mp3)|*.mp3"
                M1 = M1 & "|Mid Files" & "  (*.mid)|*.mid"
                M1 = M1 & "|Wav Files" & "  (*.wav)|*.wav"
                M1 = M1 & "|Mp2 Files" & "  (*.mp2)|*.mp2"
                M1 = M1 & "|Jpg Files" & "  (*.jpg)|*.jpg"
                M1 = M1 & "|Bmp Files" & "  (*.bmp)|*.bmp"
                M1 = M1 & "|Wav Files" & "  (*.wav)|*.wav"
                M1 = M1 & "|Mp3 Files" & "  (*.mp3)|*.bmp"
                M1 = M1 & "|WMA Files" & "  (*.wma)|*.bmp"
                
                'Sets so only *.mpg, *.avi etc can be opened and viewed
                cmdopen.Filter = M1   '"Media Files (*.bmp;*.jpg;*.mpg;*.avi;*.mov;*.wav;*.mp2;*.mp3)|*.bmp;*.jpg;*.mpg;*.avi;*.mov;*.wav;*.mp2;*.mp3"
                'Shows the open dialog box.
                cmdopen.ShowOpen
                'store the filename into a variable
                Filenamer = cmdopen.filename
                
                TheMovie = Filenamer
                If LCase(Right(TheMovie, 3)) = "avi" Then
                    Command7.Visible = True
                    Frame2.Visible = True
                Else
                    Command7.Visible = False
                    Frame2.Visible = False
                End If
                If LCase(Right(TheMovie, 3)) = "mpg" Or LCase(Right(TheMovie, 4)) = "mpeg" Then
                    Command8.Visible = True
                Else
                    Command8.Visible = False
                End If

                MovieName.Caption = GetFileTitle(cmdopen.filename)
                'Renders the file so that it can be played. Otherwise you will not
                'here anything.
                
                'Refrences our audio the the main object
                Set audio = mcontrol
                'Sets the volume and balance from the slider.
                audio.Volume = 0 'sldvolume.Value
                sldvolume.Value = 0
                audio.Balance = sldbalance.Value
                'Refrences our video to the main object
                mcontrol.RenderFile Filenamer
                Set video = mcontrol
                video.WindowStyle = CLng(&H6000000)
                video.AutoShow = True
                video.Top = 0  'picVideoWindow.Top - 50
                'This makes it so that no caption appears in the picture box
                video.Left = 0  'picVideoWindow.Left
                'sets the height and width to the same as the picture box.
                'Otherwise not all of the movie would be seen.
                video.Height = picVideoWindow.Height
                video.Width = picVideoWindow.Width
                'This sets where the video will be run from.
                'You can run it from anything such as a textbox, frame or even the slider control!
                video.owner = picVideoWindow.hWnd
                'Me.Top = Me.Top + 1
                'refrences our positioning to the main object
                Set mposition = mcontrol
                
                mposition.Rate = 1
                
                'txtlength = Round(mposition.Duration, 2)
                
                
                Slider1.Max = mposition.Duration
                ''SetTopMostWindow Me.hWnd, True
                Me.Show
                Me.Refresh
                mcontrol.Run

End Sub

Private Sub Command21_Click()
    J = Me.Top
        For i = -10000 To J * 2
            Me.Top = Int(i / 2)
        Next
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Moviee.Text = "" Then Exit Sub
Dim ActionFlag As Long
Dim i As Integer, k As Long
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Me.Visible = False
Msg = "Are You sure you want to DELETE: " & Filenamer & " ?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Delete Media"   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then   ' User chose Yes.
        k = Moviee.ListIndex
        Kill Moviee.Text
        Moviee.RemoveItem k
        Moviee.ListIndex = k
        Moviee.Text = Moviee.List(k)
        Set mcontrol = Nothing
        Set audio = Nothing
        Set mposition = Nothing
        Set audio = mcontrol
        mcontrol.RenderFile Moviee.Text
        Set video = mcontrol
        video.WindowStyle = CLng(&H6000000)
        video.AutoShow = True
        video.Top = 0  'picVideoWindow.Top - 50
        video.Left = 0  'picVideoWindow.Left
        video.Height = picVideoWindow.Height
        video.Width = picVideoWindow.Width
        video.owner = picVideoWindow.hWnd
        Set mposition = mcontrol
        mposition.Rate = 1
        txtlength = Round(mposition.Duration, 2)
        Slider1.Max = txtlength     'mposition.Duration
        Me.Show
        Me.Refresh
        mcontrol.Run
        sldbalance.Refresh
        sldvolume.Refresh
        sldplayrate.Refresh
        Slider1.Value = 1
        Slider1.Refresh
        sldvolume.Value = -1000
        audio.Volume = -1000  'sldvolume.Value
End If
Me.Visible = True

End Sub
Private Sub Refill()
Dim i As Integer
Dim Filepath As String
On Error Resume Next
Moviee.Clear
If Right(File1.path, 1) = "\" Then
    Filepath = File1.path
Else
    Filepath = File1.path & "\"
End If
For i = 0 To File1.ListCount - 1
    Moviee.AddItem Filepath & File1.List(i)
Next
Moviee.Text = Moviee.List(0)
Moviee.ListIndex = 0
Filenamer = Moviee.ListIndex
End Sub


Private Sub Command6_Click()
On Error Resume Next
    Set audio = mcontrol
    sldvolume.Value = -5000
    audio.Volume = -5000  'sldvolume.Value
    sldvolume.Refresh
End Sub

Private Sub Command7_Click()
        On Error Resume Next
        MovieName.Caption = "Saving Frame to .bmp"
        sldbalance.Refresh
        sldvolume.Refresh
        sldplayrate.Refresh
        Slider1.Refresh
        Open App.path & "\Avitemp" For Output As #1
            Print #1, TheMovie
            Print #1, mposition.CurrentPosition
            Print #1, mposition.Duration
        Close #1
        Dim RetVal
        RetVal = Shell(App.path & "\FrameCap.exe", 1)
        Options(4).Visible = False
        'Line Input #1, Fille  = Movie
        'Line Input #1, xx = mposition.currentposition
        ' PositionSec = Val(Trim(xx))
        ' Line Input #1, xx
        'DurationSec = Val(xx)
        'FrameSec = numFrames / DurationSec
        'FrameNumber = FrameSec * PositionSec

End Sub


Private Sub Command8_Click()
    Dim Mpgest, Mpger As String
    SetTopMostWindow Me.hWnd, False
    Dim RetVal
    Mpger = Replace(LCase(TheMovie), ".mpg", "")
    Mpger = Replace(LCase(TheMovie), ".mpeg", "")
    Mpgest = App.path & "\m2apx3g.exe -b " & TheMovie & " -f -o7 " & Mpger & ".avi"
    RetVal = Shell(Mpgest, 1)
    Options(4).Visible = False
End Sub



Private Sub Command9_Click()
If Trim(TheMovie) <> "" Then
    'On Error Resume Next
    Schnorbel = True
    Me.Visible = False
    MoosePosition = mposition.CurrentPosition
    mcontrol.Stop
    Set mcontrol = Nothing
    Set audio = Nothing
    Set video = Nothing
    Set mposition = Nothing
    'hide the taskbar
    rtn = FindWindow("Shell_traywnd", "") 'get the Window
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW) 'hide the Tasbar
    FullScreen.Slider1.Value = Slider1.Value
    Load FullScreen
    FullScreen.Show
End If
End Sub

Private Sub Yes_Click()
Dim Moove As Boolean
Dim J, k As Long
Dim ToFileName As String
On Error Resume Next
'Me.Visible = False
k = Moviee.ListIndex
Moove = False
ToFileName = Trim(GetFileTitle(Moviee.Text))
Do While Moove = False
    If Dir(Patherino & ToFileName) = "" Then
        If Dir(Moviee.Text) <> "" Then
        'MsgBox Patherino & ToFileName
            FileOPS1.MoveFile Moviee.Text, Patherino & ToFileName
        End If
        Moove = True
    Else
        ToFileName = ToFileName & Str(J)
        J = J + 1
    End If
Loop
If k = -1 Then k = 0
Kill Moviee.Text
Moviee.RemoveItem Moviee.ListIndex
Moviee.ListIndex = k
Moviee.Text = Moviee.List(k)
Set mcontrol = Nothing
Set audio = Nothing
Set mposition = Nothing
Set audio = mcontrol
mcontrol.RenderFile Moviee.Text
Set video = mcontrol
video.WindowStyle = CLng(&H6000000)
video.AutoShow = True
video.Top = 0  'picVideoWindow.Top - 50
video.Left = 0  'picVideoWindow.Left
video.Height = picVideoWindow.Height
video.Width = picVideoWindow.Width
video.owner = picVideoWindow.hWnd
Set mposition = mcontrol
mposition.Rate = 1
txtlength = Round(mposition.Duration, 2)
Slider1.Max = txtlength     'mposition.Duration
mcontrol.Run
sldbalance.Refresh
sldvolume.Refresh
sldplayrate.Refresh
Slider1.Value = 1
Slider1.Refresh
sldvolume.Value = -1000
audio.Volume = -1000  'sldvolume.Value
'Me.Show
Me.Refresh
End Sub

Private Sub Form_Activate()
On Error Resume Next
Angst.Visible = False
Moviee_Click
End Sub
Sub OldInit()
'If Trim(Moviee.List(0)) <> "" Then
'If Schnorbel = False Then
    TheMovie = Moviee.Text
    'Moviee.ListIndex = 0
'Else
    'Moviee.Text = TheMovie
    'Moviee.ListIndex = 0
'End If
Text3.Text = TheMovie
    If LCase(Right(TheMovie, 3)) = "avi" Then
        Command7.Visible = True
        Frame2.Visible = True
        Command10.Visible = True
        Command11.Visible = True
    Else
        Command7.Visible = False
        Frame2.Visible = False
        Command10.Visible = False
        Command11.Visible = False
    End If
    If LCase(Right(TheMovie, 3)) = "mp3" Then
        Command18.Visible = True
        Open App.path & "\Mp3z" For Output As #1
        For i = 0 To Moviee.ListCount - 1
            Print #1, Moviee.List(i)
        Next
        Close #1
    Else
        Command18.Visible = False
    End If
    If LCase(Right(TheMovie, 3)) = "mpg" Or LCase(Right(TheMovie, 4)) = "mpeg" Then
        Command8.Visible = True
    Else
        Command8.Visible = False
    End If
    If Trim(TheMovie) <> "" Then
        Set mcontrol = Nothing
        Set audio = Nothing
        Set mposition = Nothing
        Filenamer = TheMovie
        Set audio = mcontrol
        mcontrol.RenderFile Filenamer
        Set video = mcontrol
        video.WindowStyle = CLng(&H6000000)
        video.AutoShow = True
        video.Top = 0  'picVideoWindow.Top - 50
        video.Left = 0  'picVideoWindow.Left
        video.Height = picVideoWindow.Height
        video.Width = picVideoWindow.Width
        video.owner = picVideoWindow.hWnd
        Set mposition = mcontrol
        mposition.Rate = 1
        txtlength = Round(mposition.Duration, 2)
        Slider1.Max = txtlength     'mposition.Duration
        Me.Show
        Me.Refresh
        audio.Volume = -1000  'sldvolume.Value
        mposition.CurrentPosition = MoosePosition
        mcontrol.Run
        sldbalance.Refresh
        sldvolume.Refresh
        sldplayrate.Refresh
        Slider1.Refresh
        sldvolume.Value = -1000
    End If
'End If

End Sub
Private Sub Form_Initialize()
On Error Resume Next
Call AVIFileInit   '// opens AVIFile library
File1.path = Angst.File1.path
File1.Pattern = "*.mpg;*.mpeg;*.m2v;*.avi;*.asf;*.mov;*.wmv;*.mp3;*.wma;*.wav;*m1v"
Screen.MousePointer = Default
Moviee.ListIndex = 0
Moviee.Text = Moviee.List(0)
Moviee_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
ImagesPath = ""
Filenamer = ""
Point1 = 0: Point2 = 0
Loopit = False
For i = 0 To 7
    Options(i).Enabled = True
Next
Command5.Enabled = True
MovieName.Caption = ""

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MovieName.Caption = X & "   " & Y
'Dim i As Long
'If X >= 61 And X <= 464 And Y >= 165 And Y <= 450 And Tegosw1.Value = False Then
    'picVideoWindow.Visible = True
    'For i = 0 To 7
        'Options(i).Enabled = True
    'Next
    'Command5.Enabled = True
    'Tegosw1.Value = True
    'MovieName.Caption = ""
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mcontrol = Nothing
    Set audio = Nothing
    Set video = Nothing
    Set mposition = Nothing
    Unload Me
    Set VideoLibrary = Nothing
    FullScreen.Command1_Click
    Schnorbel = False
    Call AVIFileExit   '// releases AVIFile library
    'If Trim(Angst.CList4.List(0)) <> "" Or ImagesPath <> "" Then
    Angst.Show
    'If ImagesPath <> "" Then
        'Call Angst.Images1
    'End If
    'Else
        'Unload Angst
        'UnloadAll
        'Set Angst = Nothing
    'End If
End Sub

Private Sub frabal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub fraplayrate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub fravol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Haburabadooda12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = 1
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = False
Image2.Visible = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub mnuLoop_Click()
If Filenamer = "" Then Exit Sub
If Options(7).Value = True Then
    Options(7).Value = False
    Line1.Visible = False
    Line2.Visible = False
Else
    Options(7).Value = True
End If
Point1 = 0: Point2 = 0

End Sub

Private Sub Moviee_Change()
If LCase(Right(TheMovie, 3)) = "avi" Then
    Command7.Visible = True
    Frame2.Visible = True
    Command10.Visible = True
    Command11.Visible = True
Else
    Command7.Visible = False
    Frame2.Visible = False
    Command10.Visible = False
    Command11.Visible = False
End If

End Sub

Public Sub Moviee_Click()
On Error Resume Next
TheMovie = Moviee.List(Moviee.ListIndex)
Text3.Text = TheMovie
If LCase(Right(TheMovie, 3)) = "mp3" Then
    Command18.Visible = True
    Open App.path & "\Mp3z" For Output As #1
    For i = 0 To Moviee.ListCount - 1
        Print #1, Moviee.List(i)
    Next
    Close #1
Else
    Command18.Visible = False
End If

If LCase(Right(TheMovie, 3)) = "avi" Then
    Command7.Visible = True
    Frame2.Visible = True
    Command10.Visible = True
    Command11.Visible = True
Else
    Command7.Visible = False
    Frame2.Visible = False
    Command10.Visible = False
    Command11.Visible = False
End If
If LCase(Right(TheMovie, 3)) = "mpg" Or LCase(Right(TheMovie, 4)) = "mpeg" Then
    Command8.Visible = True
Else
    Command8.Visible = False
End If
        
If Trim(TheMovie) <> "" Then
Set mcontrol = Nothing
Set audio = Nothing
Set mposition = Nothing
        Filenamer = TheMovie
        Set audio = mcontrol
        audio.Volume = -1000 'sldvolume.Value
        sldvolume.Value = -1000
        audio.Balance = sldbalance.Value
        mcontrol.RenderFile Filenamer
        Set video = mcontrol
        video.WindowStyle = CLng(&H6000000)
        video.AutoShow = True
        video.Top = 0  'picVideoWindow.Top - 50
        video.Left = 0  'picVideoWindow.Left
        video.Height = picVideoWindow.Height
        video.Width = picVideoWindow.Width
        video.owner = picVideoWindow.hWnd
        Set mposition = mcontrol
        mposition.Rate = 1
        txtlength = Round(mposition.Duration, 2)
        Slider1.Max = txtlength     'mposition.Duration
        Me.Show
        Me.Refresh
        mcontrol.Run
        sldbalance.Refresh
        sldvolume.Refresh
        sldplayrate.Refresh
        Slider1.Refresh
        picVideoWindow.Refresh
End If
End Sub

Private Sub MovieName_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Ontop_Click()

End Sub

Private Sub OnBottom_Click()

End Sub

Private Sub Opt_Click(Index As Integer)

    If Lab1.Visible = True Then
        Lab1.Visible = False
        Lab2.Visible = False
        Loopit = False
        Line1.Visible = False
        Line2.Visible = False
        Point1 = 0: Point2 = 0
        Exit Sub
     End If

End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Select Case Index

Case 0
    Patherino = "p:\1down\1\Gif\"
Case 1
    Patherino = "p:\1down\1\Legal Security Government\"
Case 2
    Patherino = "p:\1down\1\Economic Financial\"
Case 3
    Patherino = "p:\1down\"
Case 4
    Patherino = "p:\1down\1\Microphones\"
Case 5
    Patherino = "p:\1down\1\Programming\"
Case 6
    Patherino = "p:\1down\1\Guitar\"
Case 7
    Patherino = "p:\1down\1\Medical\"
Case 8
    Patherino = "p:\1down\1\Utilities Cracks\"
Case 9
    Patherino = "p:\1down\1\Personals\"
Case 10
    Patherino = "p:\1down\1\Multimedia & Graphic\"
Case 11
    Patherino = "p:\1down\1\Telecom Internet FTP Networking\"
Case 12
    Patherino = "p:\1down\1\Science Education\"
Case 13
    Patherino = "p:\1down\1\Wordpro Spreadsheet Datapro\"
Case 14
    Patherino = "p:\1down\1\Multimedia & Graphic\"
End Select
Yes_Click
Option1(Index).Value = False
End Sub

Private Sub Options_Click(Index As Integer)
Dim Framespos As Single
Select Case Index   ' Evaluate Number.
Case 0 ' Open
Case 1 'Play
        On Error Resume Next
        If Filenamer = "" Then Exit Sub
        mcontrol.Run
Case 2 'Pause
        If Filenamer = "" Then Exit Sub
        On Error Resume Next
        mcontrol.Pause
Case 3 'Stop
        'Filenamer = ""
        On Error Resume Next
        mcontrol.RenderFile ""
        Set video = mcontrol
        video.WindowStyle = CLng(&H6000000)
        video.AutoShow = True
        video.owner = picVideoWindow.hWnd
        Set mposition = mcontrol
        mposition.Rate = 1
        Slider1.Max = mposition.Duration
        Me.Show
        Me.Refresh
        mcontrol.Run
        mcontrol.Stop
        mcontrol.RenderFile ""
        mposition.CurrentPosition = 0
        Line1.Visible = False
        Line2.Visible = False
        Timer1.Enabled = False
Case 4 'Snapshot
        On Error Resume Next
        MovieName.Caption = "This may take a few minutes..."
        sldbalance.Refresh
        sldvolume.Refresh
        sldplayrate.Refresh
        Slider1.Refresh
        Open App.path & "\Avitemp" For Output As #1
            Print #1, TheMovie
            Print #1, mposition.CurrentPosition
            Print #1, mposition.Duration
        Close #1
        Dim RetVal
        RetVal = Shell(App.path & "\AVI2BMP.exe", 1)
        Options(4).Visible = False
Case 5  'Looper
        On Error Resume Next
        If Filenamer = "" Then Exit Sub
        If mnuLoop.Checked = True Then
            mnuLoop.Checked = False
            Line1.Visible = False
            Line2.Visible = False
        Else
            mnuLoop.Checked = True
        End If
        Point1 = 0: Point2 = 0
Case 6  'Exit
        Unload Me
Case 7
    If Filenamer = "" Then Exit Sub
    If Loopit = True Then
        Loopit = False
        Line1.Visible = False
        Line2.Visible = False
    Else
       Loopit = True
    End If
    Point1 = 0: Point2 = 0

Case 8
        On Error Resume Next
        mcontrol.Stop
        mcontrol.RenderFile App.path & "\Database.BMP"
                Set video = mcontrol
                video.WindowStyle = CLng(&H6000000)
                video.AutoShow = True
                video.Top = 0  'picVideoWindow.Top - 50
                'This makes it so that no caption appears in the picture box
                video.Left = 0  'picVideoWindow.Left
                'sets the height and width to the same as the picture box.
                'Otherwise not all of the movie would be seen.
                video.Height = picVideoWindow.Height
                video.Width = picVideoWindow.Width
                'This sets where the video will be run from.
                'You can run it from anything such as a textbox, frame or even the slider control!
                video.owner = picVideoWindow.hWnd
                'Me.Top = Me.Top + 1
                'refrences our positioning to the main object
                Set mposition = mcontrol
                
                mposition.Rate = 1
                
                'txtlength = Round(mposition.Duration, 2)
                
                
                Slider1.Max = mposition.Duration
                ''SetTopMostWindow Me.hWnd, True
                Me.Show
                Me.Refresh
                mcontrol.Run
        
End Select


End Sub

Private Sub Picture2_Click()
Unload Me
Set VideoLibrary = Nothing
End Sub

Private Sub picVideoWindow_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XY, txtlength As Long
On Error Resume Next
        Moviee.Clear
        For XY = 1 To Data.Files.Count
            Moviee.AddItem Data.Files(XY), XY - 1
        Next XY
        'LoadFileAndUpdateDisplay Moviee.List(0)
        TheMovie = Moviee.List(0)
        If LCase(Right(TheMovie, 3)) = "avi" Then
            Command7.Visible = True
            Frame2.Visible = True
        Else
            Command7.Visible = False
            Frame2.Visible = False
        End If
        If LCase(Right(TheMovie, 3)) = "mp3" Then
            Command18.Visible = True
            Open App.path & "\Mp3z" For Output As #1
                Print #1, TheMovie
            Close #1
        Else
            Command18.Visible = False
        End If
        If LCase(Right(TheMovie, 3)) = "mpg" Or LCase(Right(TheMovie, 4)) = "mpeg" Or LCase(Right(TheMovie, 4)) = "mp3" Or LCase(Right(TheMovie, 4)) = "wav" Then
            Command8.Visible = True
        Else
            Command8.Visible = False
        End If
If Trim(TheMovie) <> "" Then
Set mcontrol = Nothing
Set audio = Nothing
Set mposition = Nothing
        Filenamer = TheMovie
        Set audio = mcontrol
        mcontrol.RenderFile Filenamer
        Set video = mcontrol
        video.WindowStyle = CLng(&H6000000)
        video.AutoShow = True
        video.Top = 0  'picVideoWindow.Top - 50
        video.Left = 0  'picVideoWindow.Left
        video.Height = picVideoWindow.Height
        video.Width = picVideoWindow.Width
        video.owner = picVideoWindow.hWnd
        Set mposition = mcontrol
        mposition.Rate = 1
        txtlength = Round(mposition.Duration, 2)
        Slider1.Max = txtlength     'mposition.Duration
        Me.Show
        Me.Refresh
        mcontrol.Run
        sldbalance.Refresh
        sldvolume.Refresh
        sldplayrate.Refresh
        Slider1.Refresh
End If
Moviee.ListIndex = 0
Moviee_Click
End Sub

Private Sub sldbalance_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
audio.Balance = sldbalance.Value

End Sub

Private Sub sldplayrate_Change()
On Error Resume Next
sldplayrate_MouseUp 0, 0, 0, 0

End Sub

Private Sub sldplayrate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'Changes the play rate.
'Note: only values from .1 - 2.26 can be used guess it
'cant handle any faster playback
playrate = sldplayrate.Value & "%"
txtplayback = sldplayrate.Value & "%"
mposition.Rate = sldplayrate.Value / 100

End Sub

Private Sub sldvolume_Change()
On Error Resume Next
sldvolume.Refresh
End Sub

Private Sub sldvolume_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
audio.Volume = sldvolume.Value

End Sub

Private Sub Slider1_Change()
'On Error Resume Next
If Slider1.Value = Slider1.Max And Option4.Value = True Then
    Slider1.Value = 1
    If Moviee.ListIndex = Moviee.ListCount - 1 Then
        Moviee.ListIndex = 0
        Moviee_Click
    Else
        Moviee.ListIndex = Moviee.ListIndex + 1
        Moviee_Click
    End If
End If
If Slider1.Value = Slider1.Max And Option3.Value = True Then
    Dim MyValue
    Randomize   ' Initialize random-number generator.
    MyValue = Int((Moviee.ListCount - 1 * Rnd))    ' Generate random value between 1 and 6.
    Slider1.Value = 1
    Moviee.ListIndex = MyValue
    Moviee_Click
End If

    Dim Res As Long         'result code
    Dim ofd As cFileDlg     'OpenFileDialog class
    Dim szFile As String    'filename
    Dim xx As String
    Dim pAVIFile As Long    'pointer to AVI file interface (PAVIFILE handle)
    Dim pAVIStream As Long  'pointer to AVI stream interface (PAVISTREAM handle)
    Dim numFrames As Long   'number of frames in video stream
    Dim firstFrame As Long  'position of the first video frame
    Dim fileInfo As AVI_FILE_INFO       'file info struct
    Dim streamInfo As AVI_STREAM_INFO   'stream info struct
    Dim dib As cDib
    Dim pGetFrameObj As Long    'pointer to GetFrame interface
    Dim pDIB As Long            'pointer to packed DIB in memory
    Dim bih As BITMAPINFOHEADER 'infoheader to pass to GetFrame functions
    Dim i As Long
    
    If Slider1.Value >= Point2 And Loopit = True And Point2 > Point1 Then
        mposition.CurrentPosition = Point1
    End If
    szFile = TheMovie
    Res = AVIFileOpen(pAVIFile, szFile, OF_SHARE_DENY_WRITE, 0&)
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
    'Get the first available video stream (PAVISTREAM)
    Res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    If Res <> AVIERR_OK Then GoTo ErrorOut
    
    'get the starting position of the stream (some streams may not start simultaneously)
    firstFrame = AVIStreamStart(pAVIStream)
    If firstFrame = -1 Then GoTo ErrorOut 'this function returns -1 on error
    'MsgBox "moose"
    'get the length of video stream in frames
    numFrames = AVIStreamLength(pAVIStream)
    If numFrames = -1 Then GoTo ErrorOut ' this function returns -1 on error
    PositionSec = Val(Trim(mposition.CurrentPosition))
    DurationSec = Val(mposition.Duration)
    FrameSec = numFrames / DurationSec
    FrameNumber = FrameSec * PositionSec
    Label2.Caption = "Total Frames= " & numFrames
    Label3.Caption = "Frame Number " & FrameNumber
    
Exit Sub
ErrorOut:
    If pGetFrameObj <> 0 Then
        Call AVIStreamGetFrameClose(pGetFrameObj) '//deallocates the GetFrame resources and interface
    End If
    If pAVIStream <> 0 Then
        Call AVIStreamRelease(pAVIStream) '//closes video stream
    End If
    If pAVIFile <> 0 Then
        Call AVIFileRelease(pAVIFile) '// closes the file
    End If
    
    If (Res <> AVIERR_OK) Then 'if there was an error then show feedback to user
        'MsgBox "There was an error working with the file:" & vbCrLf & szFile, vbInformation, App.title
    End If

End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Timer1.Enabled = False

End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Loopit = False Then
    Point1 = 0: Point2 = 0
    Line1.Visible = False
    Line2.Visible = False
End If
If Loopit = True And Point1 = 0 Then
    Point1 = Slider1.Value
    Lab1.Caption = Point1
    Lab1.Visible = True
    Line1.X1 = X
    Line1.X2 = X
    Line1.Visible = False 'True
    Timer1.Enabled = True
    Exit Sub
End If
If Loopit = True And Point1 <> 0 And Point2 = 0 Then
    Point2 = Slider1.Value
    Line2.X1 = X
    Line2.X2 = X
    Line2.Visible = False   'True
    Lab2.Caption = Point2
    Lab2.Visible = True

End If

If Loopit = True And Point1 <> 0 And Point2 <> 0 Then
    mposition.CurrentPosition = Point1
Else
    mposition.CurrentPosition = Slider1.Value
End If
'MsgBox Point1
'MsgBox Point2
Timer1.Enabled = True

End Sub


Private Sub Text1_Change()
If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
    Command13.Enabled = True
Else
    Command13.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Trim(Text1.Text) <> "" And Trim(Text2.Text) <> "" Then
    Command13.Enabled = True
Else
    Command13.Enabled = False
End If

End Sub

Private Sub Text3_Change()
On Error Resume Next
Dim TempFile As String, i As Integer
Patherino = Replace(GetFilePath(Text3.Text) & "\", "\\", "\")
Open App.path & "\LastPath" For Output As #1
    Print #1, Patherino
Close #1
Exterino = Right(Text3.Text, 4)
i = InStrRev(GetFileTitle(Text3.Text), ".")
TempFile = GetFileTitle(Text3.Text)
Text3.Text = Left(TempFile, i - 1)

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Slider1.Value = mposition.CurrentPosition

End Sub

Private Sub Timer2_Timer()
Dim i As Long, J As Long
Timer2.Enabled = False
Me.Hide
SetTopMostWindow Me.hWnd, True
If InitFlag = False Then
        Me.Left = (Screen.Width - Me.Width) / 2
        For i = -1000 To (Screen.Height - Me.Height) / 2
            Me.Top = i
        Next
Else
    J = Me.Top
        'Me.left = (Screen.Width - Me.Width) / 2
        For i = -1000 To J
            Me.Top = i
        Next
End If
Me.Visible = True
InitFlag = True
If Check1.Value = 1 Then
    SetTopMostWindow Me.hWnd, True
Else
    SetTopMostWindow Me.hWnd, False
End If
End Sub
