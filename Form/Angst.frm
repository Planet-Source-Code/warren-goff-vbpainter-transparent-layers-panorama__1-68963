VERSION 5.00
Object = "{C1A6E3E0-74BB-11D6-97C3-0000B4BDB148}#4.5#0"; "THBImg45.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDB6A620-59F5-4221-94F4-C61F12CF4572}#1.0#0"; "DirProcess.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RichTx32.ocx"
Object = "{CE116DEF-695F-4EE1-851A-E677CC4BD3DE}#1.0#0"; "Mousewheel.ocx"
Begin VB.Form Angst 
   BorderStyle     =   0  'None
   ClientHeight    =   8580
   ClientLeft      =   7140
   ClientTop       =   4110
   ClientWidth     =   9435
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Angst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Angst.frx":08CA
   Picture         =   "Angst.frx":1194
   ScaleHeight     =   572
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   629
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      Caption         =   "   Custom Resize"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2385
      Left            =   5505
      TabIndex        =   58
      Top             =   5955
      Visible         =   0   'False
      Width           =   2085
      Begin VB.CommandButton Command21 
         BackColor       =   &H003476AC&
         Caption         =   "Apply"
         Height          =   315
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Apply"
         Top             =   1830
         Width           =   555
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   315
         Left            =   1650
         Max             =   -1
         Min             =   -32767
         TabIndex        =   86
         Top             =   690
         Value           =   -1
         Width           =   255
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   315
         Left            =   1650
         Max             =   -1
         Min             =   -32767
         TabIndex        =   85
         Top             =   300
         Value           =   -1
         Width           =   255
      End
      Begin VB.TextBox Spinner2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005CA6BC&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   780
         TabIndex        =   84
         Text            =   "1"
         Top             =   690
         Width           =   855
      End
      Begin VB.TextBox Spinner1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005CA6BC&
         Height          =   315
         Left            =   780
         TabIndex        =   83
         Text            =   "1"
         Top             =   300
         Width           =   855
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H0073AEB6&
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
         Height          =   690
         Left            =   390
         TabIndex        =   82
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0073AEB6&
         Caption         =   "Keep Aspect Ratio"
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
         Height          =   315
         Left            =   60
         TabIndex        =   80
         Top             =   2130
         Width           =   1980
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   3180
         TabIndex        =   71
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox Text10 
         Height          =   315
         Left            =   4680
         TabIndex        =   70
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox Text9 
         Height          =   315
         Left            =   4440
         TabIndex        =   69
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Height          =   315
         Left            =   4200
         TabIndex        =   68
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox Text7 
         Height          =   315
         Left            =   4680
         TabIndex        =   67
         Top             =   2100
         Width           =   255
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   4440
         TabIndex        =   66
         Top             =   2100
         Width           =   255
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   4200
         TabIndex        =   65
         Top             =   2100
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Height          =   315
         Left            =   4680
         TabIndex        =   64
         Top             =   1770
         Width           =   255
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   4440
         TabIndex        =   63
         Top             =   1770
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   4200
         TabIndex        =   62
         Top             =   1770
         Width           =   255
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H003476AC&
         Caption         =   "Ok"
         Height          =   285
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "OK Apply"
         Top             =   1830
         Width           =   615
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H000000FF&
         Caption         =   "x"
         Height          =   210
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Load Graphic"
         Top             =   30
         Width           =   225
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FF80FF&
         Caption         =   "Rnd"
         Height          =   270
         Left            =   2370
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Load Graphic"
         Top             =   1470
         Width           =   405
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox14 
         Height          =   300
         Left            =   2310
         TabIndex        =   72
         Top             =   1170
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox15 
         Height          =   300
         Left            =   2340
         TabIndex        =   73
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox16 
         Height          =   300
         Left            =   2340
         TabIndex        =   74
         Top             =   930
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox17 
         Height          =   300
         Left            =   2340
         TabIndex        =   75
         Top             =   1230
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox18 
         Height          =   300
         Left            =   2940
         TabIndex        =   76
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox19 
         Height          =   300
         Left            =   2940
         TabIndex        =   77
         Top             =   930
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox20 
         Height          =   300
         Left            =   2940
         TabIndex        =   78
         Top             =   1230
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin VB.Image Image1 
         Height          =   3705
         Left            =   30
         Picture         =   "Angst.frx":17AD56
         Top             =   210
         Width           =   3720
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height: "
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
         Height          =   255
         Index           =   1
         Left            =   60
         TabIndex        =   81
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width: "
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
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   79
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0073AEB6&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   5610
      TabIndex        =   25
      Top             =   6645
      Visible         =   0   'False
      Width           =   1995
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FF80FF&
         Caption         =   "Rnd"
         Height          =   270
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Load Graphic"
         Top             =   300
         Width           =   435
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H000000FF&
         Caption         =   "x"
         Height          =   210
         Left            =   1770
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Load Graphic"
         Top             =   15
         Width           =   225
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080FF80&
         Caption         =   "Ok"
         Height          =   270
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Load Graphic"
         Top             =   300
         Width           =   345
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox10 
         Height          =   300
         Left            =   90
         TabIndex        =   45
         Top             =   300
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin VB.TextBox tb21 
         Height          =   315
         Left            =   4200
         TabIndex        =   35
         Top             =   1770
         Width           =   255
      End
      Begin VB.TextBox tb22 
         Height          =   315
         Left            =   4440
         TabIndex        =   34
         Top             =   1770
         Width           =   255
      End
      Begin VB.TextBox tb23 
         Height          =   315
         Left            =   4680
         TabIndex        =   33
         Top             =   1770
         Width           =   255
      End
      Begin VB.TextBox tb31 
         Height          =   315
         Left            =   4200
         TabIndex        =   32
         Top             =   2100
         Width           =   255
      End
      Begin VB.TextBox tb32 
         Height          =   315
         Left            =   4440
         TabIndex        =   31
         Top             =   2100
         Width           =   255
      End
      Begin VB.TextBox tb33 
         Height          =   315
         Left            =   4680
         TabIndex        =   30
         Top             =   2100
         Width           =   255
      End
      Begin VB.TextBox tb11 
         Height          =   315
         Left            =   4200
         TabIndex        =   29
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox tb12 
         Height          =   315
         Left            =   4440
         TabIndex        =   28
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox tb13 
         Height          =   315
         Left            =   4680
         TabIndex        =   27
         Top             =   1440
         Width           =   255
      End
      Begin VB.TextBox tbDivisor 
         Height          =   285
         Left            =   3180
         TabIndex        =   26
         Top             =   1440
         Width           =   495
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox1 
         Height          =   300
         Left            =   90
         TabIndex        =   36
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox2 
         Height          =   300
         Left            =   90
         TabIndex        =   37
         Top             =   930
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox3 
         Height          =   300
         Left            =   90
         TabIndex        =   38
         Top             =   1230
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox4 
         Height          =   300
         Left            =   690
         TabIndex        =   39
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox5 
         Height          =   300
         Left            =   690
         TabIndex        =   40
         Top             =   930
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox6 
         Height          =   300
         Left            =   690
         TabIndex        =   41
         Top             =   1230
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox7 
         Height          =   300
         Left            =   1290
         TabIndex        =   42
         Top             =   630
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox8 
         Height          =   300
         Left            =   1290
         TabIndex        =   43
         Top             =   930
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin AngstArt.ASpinnerBox ASpinnerBox9 
         Height          =   300
         Left            =   1290
         TabIndex        =   44
         Top             =   1230
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         ForeColor       =   0
         BorderColor     =   12083200
         Value           =   0
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Devisor"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   46
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0080FF80&
      Caption         =   "Color to replace"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8055
      Picture         =   "Angst.frx":185A46
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   7290
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0073AEB6&
      Caption         =   "BMP 640x480"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   157
      Top             =   7935
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0073AEB6&
      Caption         =   "JPG 640x480"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3450
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   7665
      Width           =   1200
   End
   Begin WheelCtl.MouseWheel MouseWheel1 
      Left            =   3300
      Top             =   8070
      _ExtentX        =   529
      _ExtentY        =   503
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   375
      Left            =   11565
      TabIndex        =   151
      Top             =   45
      Width           =   1005
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H0073AEB6&
      Caption         =   "Move"
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
      Height          =   555
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   150
      Top             =   6825
      Width           =   480
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H0073AEB6&
      Caption         =   "Clip"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   149
      ToolTipText     =   "Rename Between these 2 Frames"
      Top             =   7110
      Width           =   525
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H00000000&
      Height          =   525
      Left            =   2505
      Picture         =   "Angst.frx":189DEF
      Style           =   1  'Graphical
      TabIndex        =   148
      ToolTipText     =   "Overlay"
      Top             =   7890
      Width           =   585
   End
   Begin VB.CommandButton Command39 
      BackColor       =   &H0073AEB6&
      Caption         =   "Ren"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   147
      ToolTipText     =   "Rename Between these 2 Frames"
      Top             =   6825
      Width           =   525
   End
   Begin VB.CommandButton Command37 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8625
      Picture         =   "Angst.frx":18A6B9
      Style           =   1  'Graphical
      TabIndex        =   145
      ToolTipText     =   "Resize and Run Slide Show"
      Top             =   5970
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1260
      TabIndex        =   144
      Top             =   7380
      Width           =   3360
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7875
      Picture         =   "Angst.frx":18AD42
      Style           =   1  'Graphical
      TabIndex        =   143
      ToolTipText     =   "Slide Show Plus"
      Top             =   5970
      Width           =   675
   End
   Begin VB.ListBox List8 
      Height          =   1110
      Left            =   11655
      Sorted          =   -1  'True
      TabIndex        =   142
      Top             =   3795
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0073AEB6&
      Height          =   495
      Left            =   105
      Picture         =   "Angst.frx":18B60C
      Style           =   1  'Graphical
      TabIndex        =   141
      ToolTipText     =   "Browse for a Directory of Files to load"
      Top             =   2265
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0073AEB6&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   8625
      Picture         =   "Angst.frx":18BED6
      Style           =   1  'Graphical
      TabIndex        =   139
      ToolTipText     =   "Resize Wizard"
      Top             =   5370
      Width           =   675
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0073AEB6&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   7860
      Picture         =   "Angst.frx":18C479
      Style           =   1  'Graphical
      TabIndex        =   138
      ToolTipText     =   "Copy Machine"
      Top             =   5370
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   540
      Width           =   495
   End
   Begin THBImageLibCtl.THBImage THBImage1 
      Height          =   5505
      Left            =   900
      TabIndex        =   0
      Top             =   1005
      Width           =   6810
      _cx             =   12012
      _cy             =   9710
      KeepAspect      =   -1  'True
      PositionPic     =   4
      StretchPic      =   3
      BackColor       =   0
      BackColorOn     =   -1  'True
      BorderWidth     =   0
      BorderColor     =   0
      BorderStyle     =   0
      MirrorVertical  =   0   'False
      MirrorHorizontal=   0   'False
      PictureIsTransparent=   -1  'True
      Scrolling       =   -1  'True
      DynamicZoom     =   -1  'True
      Gradient        =   0
      GradientSimpColFrom=   1073741824
      GradientSimpColTo=   1073742079
      GradientSimpHorizontal=   0   'False
      GradientUseSystemFunctions=   0   'False
      Halftone16Bit   =   0
      PopupMenuNewStyle=   -1  'True
      PopupMenu       =   -1  'True
      MoveWithPreview =   -1  'True
      HyperlinkURL    =   ""
      PreviewScrollWindow=   -1  'True
      PreviewScrollWidth=   200
      PreviewScrollHeight=   200
      Clipboard       =   -1  'True
      ImgContainer    =   0
      MousePointer    =   32512
      EnableKeyboardHandling=   -1  'True
      BackgroundPicPosition=   7
      BackgroundPicStretch=   3
      BackgroundPicKeepAspect=   0   'False
      MagnificationWindowWidth=   160
      MagnificationWindowHeight=   100
      SysKey          =   $"Angst.frx":18CD43
      NumGradientCols =   2
      NumGradientRows =   2
      GradientPoint000000PercX=   0
      GradientPoint000000PercY=   0
      GradientPoint000000ThbCol=   1073741824
      GradientPoint001000PercX=   100
      GradientPoint001000PercY=   0
      GradientPoint001000ThbCol=   1073742079
      GradientPoint000001PercX=   0
      GradientPoint000001PercY=   100
      GradientPoint000001ThbCol=   1073807104
      GradientPoint001001PercX=   100
      GradientPoint001001PercY=   100
      GradientPoint001001ThbCol=   1090453504
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H0073AEB6&
      Height          =   450
      Left            =   105
      Picture         =   "Angst.frx":18CEE1
      Style           =   1  'Graphical
      TabIndex        =   134
      ToolTipText     =   "Load all graphic video files from Selected Directory"
      Top             =   4815
      Width           =   495
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00547AAC&
      Height          =   450
      Left            =   8355
      Picture         =   "Angst.frx":18D7AB
      Style           =   1  'Graphical
      TabIndex        =   133
      ToolTipText     =   "Show Text Page"
      Top             =   4785
      Width           =   495
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00343634&
      Height          =   450
      Left            =   7845
      Picture         =   "Angst.frx":18DCD3
      Style           =   1  'Graphical
      TabIndex        =   132
      ToolTipText     =   "Show Graphic Page"
      Top             =   4785
      Width           =   495
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00242A24&
      Height          =   450
      Left            =   8850
      Picture         =   "Angst.frx":18E1B1
      Style           =   1  'Graphical
      TabIndex        =   131
      ToolTipText     =   "Show Htm/Pdf Page"
      Top             =   4785
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      Picture         =   "Angst.frx":18E6F4
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   6525
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSComctlLib.Slider Slider4 
      Height          =   5505
      Left            =   735
      TabIndex        =   129
      ToolTipText     =   "Up is Zoom Out Down is Zoom In"
      Top             =   990
      Visible         =   0   'False
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   9710
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   1
      Min             =   -100
      Max             =   100
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   10110
      ScaleHeight     =   1215
      ScaleWidth      =   1545
      TabIndex        =   127
      Top             =   7125
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0074BADC&
      Height          =   345
      Left            =   8340
      Picture         =   "Angst.frx":18E976
      Style           =   1  'Graphical
      TabIndex        =   125
      ToolTipText     =   "Toggle Keep on Top"
      Top             =   540
      Width           =   495
   End
   Begin VB.CommandButton AnyShape32 
      BackColor       =   &H000C4ACC&
      Height          =   615
      Left            =   0
      Picture         =   "Angst.frx":18F240
      Style           =   1  'Graphical
      TabIndex        =   124
      ToolTipText     =   "MSPaint Like interface"
      Top             =   1020
      Width           =   765
   End
   Begin VB.CommandButton Haburabadooda30 
      BackColor       =   &H0073AEB6&
      Height          =   345
      Left            =   945
      Picture         =   "Angst.frx":19011E
      Style           =   1  'Graphical
      TabIndex        =   122
      ToolTipText     =   "Removes Region from Display"
      Top             =   7890
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton Haburabadooda27 
      BackColor       =   &H0073AEB6&
      Height          =   555
      Left            =   105
      Picture         =   "Angst.frx":190519
      Style           =   1  'Graphical
      TabIndex        =   121
      ToolTipText     =   "Loads default directory c:\1down\"
      Top             =   1695
      Width           =   495
   End
   Begin VB.CommandButton Haburabadooda16 
      BackColor       =   &H0073AEB6&
      Height          =   645
      Left            =   8055
      Picture         =   "Angst.frx":190B5C
      Style           =   1  'Graphical
      TabIndex        =   120
      ToolTipText     =   "Video Library and Functions /Extraction/ Conversion"
      Top             =   7785
      Width           =   945
   End
   Begin VB.CommandButton AnyShape26 
      BackColor       =   &H001C4694&
      Height          =   495
      Left            =   105
      Picture         =   "Angst.frx":19182A
      Style           =   1  'Graphical
      TabIndex        =   119
      ToolTipText     =   "Open Multiple Files"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Haburabadooda26 
      BackColor       =   &H000C4ACC&
      Height          =   255
      Left            =   105
      Picture         =   "Angst.frx":191D75
      Style           =   1  'Graphical
      TabIndex        =   118
      ToolTipText     =   "Save Active file and overwrite the original"
      Top             =   6015
      Width           =   525
   End
   Begin VB.CommandButton Haburabadooda19 
      BackColor       =   &H0073AEB6&
      Height          =   405
      Left            =   105
      Picture         =   "Angst.frx":19215E
      Style           =   1  'Graphical
      TabIndex        =   117
      ToolTipText     =   "Save As"
      Top             =   6285
      Width           =   525
   End
   Begin VB.CommandButton Haburabadooda29 
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      Height          =   525
      Left            =   1230
      MaskColor       =   &H00000000&
      Picture         =   "Angst.frx":192533
      Style           =   1  'Graphical
      TabIndex        =   116
      ToolTipText     =   "Select a Rectangular Region of the Image to Modify"
      Top             =   7890
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton Haburabadooda25 
      BackColor       =   &H0073AEB6&
      Height          =   555
      Left            =   8700
      Picture         =   "Angst.frx":192DFD
      Style           =   1  'Graphical
      TabIndex        =   115
      ToolTipText     =   "Step Capture of any Selected Area (Images dir)"
      Top             =   3300
      Width           =   525
   End
   Begin VB.CommandButton Haburabadooda2 
      BackColor       =   &H0074BADC&
      Height          =   315
      Left            =   7860
      Picture         =   "Angst.frx":1936C7
      Style           =   1  'Graphical
      TabIndex        =   114
      ToolTipText     =   "Toggle View any GIF as Animated GIF"
      Top             =   2910
      Width           =   1365
   End
   Begin VB.CommandButton Haburabadooda4 
      BackColor       =   &H0073AEB6&
      Height          =   510
      Left            =   4215
      Picture         =   "Angst.frx":193DE2
      Style           =   1  'Graphical
      TabIndex        =   113
      ToolTipText     =   "Paste from Clipboard"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Haburabadooda15 
      BackColor       =   &H0074BADC&
      Height          =   315
      Left            =   7920
      Picture         =   "Angst.frx":194271
      Style           =   1  'Graphical
      TabIndex        =   112
      ToolTipText     =   "Thumbnails of all prior Images"
      Top             =   1390
      Width           =   1245
   End
   Begin VB.CommandButton Haburabadooda23 
      BackColor       =   &H0073AEB6&
      Height          =   510
      Left            =   5895
      Picture         =   "Angst.frx":19490D
      Style           =   1  'Graphical
      TabIndex        =   111
      ToolTipText     =   "Makes any Window Translucent"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Haburabadooda3 
      BackColor       =   &H0073AEB6&
      Height          =   510
      Left            =   5295
      Picture         =   "Angst.frx":195093
      Style           =   1  'Graphical
      TabIndex        =   110
      ToolTipText     =   "Print Active Display or Selected Region"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Haburabadooda12 
      BackColor       =   &H0074BADC&
      Height          =   195
      Left            =   8340
      Picture         =   "Angst.frx":19551B
      Style           =   1  'Graphical
      TabIndex        =   109
      ToolTipText     =   "Minimize"
      Top             =   60
      Width           =   345
   End
   Begin VB.CommandButton Haburabadooda22 
      BackColor       =   &H0074BADC&
      Height          =   345
      Left            =   8310
      Picture         =   "Angst.frx":19591A
      Style           =   1  'Graphical
      TabIndex        =   108
      ToolTipText     =   "Open Active Document in Word Processor"
      Top             =   990
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Haburabadooda5 
      BackColor       =   &H0073AEB6&
      Height          =   510
      Left            =   4725
      Picture         =   "Angst.frx":1961E4
      Style           =   1  'Graphical
      TabIndex        =   107
      ToolTipText     =   "Copy to Clipboard"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Haburabadooda28 
      BackColor       =   &H0073AEB6&
      Height          =   555
      Left            =   7935
      Picture         =   "Angst.frx":196696
      Style           =   1  'Graphical
      TabIndex        =   106
      ToolTipText     =   "Automated Capture of Selected area on Desktop (Images1 dir) etc."
      Top             =   3300
      Width           =   525
   End
   Begin VB.CommandButton Haburabadooda24 
      BackColor       =   &H0073AEB6&
      Height          =   555
      Left            =   8700
      Picture         =   "Angst.frx":196CE2
      Style           =   1  'Graphical
      TabIndex        =   104
      ToolTipText     =   "Convert Most AVI's into Animated testgif.gif in App Directory"
      Top             =   3915
      Width           =   585
   End
   Begin VB.CommandButton Haburabadooda9 
      BackColor       =   &H0073AEB6&
      Height          =   555
      Left            =   7935
      Picture         =   "Angst.frx":19732F
      Style           =   1  'Graphical
      TabIndex        =   103
      ToolTipText     =   "Convert BMP's in Image Directory to AVI in Application Directory"
      Top             =   3915
      Width           =   525
   End
   Begin VB.CommandButton Haburabadooda18 
      BackColor       =   &H0073AEB6&
      Height          =   585
      Left            =   1365
      Picture         =   "Angst.frx":197857
      Style           =   1  'Graphical
      TabIndex        =   102
      ToolTipText     =   "Permanently Delete All Pictures Prior to the Present"
      Top             =   0
      Width           =   465
   End
   Begin VB.CommandButton Haburabadooda13 
      BackColor       =   &H0073AEB6&
      Height          =   465
      Left            =   0
      Picture         =   "Angst.frx":197DF4
      Style           =   1  'Graphical
      TabIndex        =   101
      ToolTipText     =   "Cover Display"
      Top             =   -45
      Width           =   525
   End
   Begin VB.CommandButton Haburabadooda1 
      BackColor       =   &H00242A24&
      Height          =   555
      Left            =   105
      Picture         =   "Angst.frx":1983B8
      Style           =   1  'Graphical
      TabIndex        =   100
      ToolTipText     =   "Load thumbnails"
      Top             =   5400
      Width           =   525
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H0073AEB6&
      Height          =   450
      Left            =   105
      Picture         =   "Angst.frx":1989FA
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   "Load all htm files from Selected Directory"
      Top             =   4290
      Width           =   495
   End
   Begin VB.ListBox List7 
      Height          =   690
      Left            =   11325
      Sorted          =   -1  'True
      TabIndex        =   98
      Top             =   3510
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox List6 
      Height          =   690
      Left            =   10995
      Sorted          =   -1  'True
      TabIndex        =   91
      Top             =   3330
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox CList4 
      BackColor       =   &H00C56A31&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   390
      Left            =   900
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   570
      Width           =   6795
   End
   Begin VB.ListBox List5 
      Height          =   690
      Left            =   10635
      TabIndex        =   88
      Top             =   3090
      Width           =   1275
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   10275
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   2865
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   10005
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command25 
      BackColor       =   &H0073AEB6&
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   96
      ToolTipText     =   "Delete Between these 2 Frames"
      Top             =   6825
      Width           =   525
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   93
      Text            =   "to HERE"
      Top             =   7080
      Width           =   1455
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1605
      Locked          =   -1  'True
      TabIndex        =   92
      Text            =   "<<<Here>>>"
      Top             =   6810
      Width           =   1455
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00244244&
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   5
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Save Formatted File Name ICON"
      Top             =   6765
      Width           =   555
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00244244&
      Caption         =   "Pdf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   4
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Save Formatted File Name PDF"
      Top             =   7380
      Width           =   555
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00244244&
      Caption         =   "Bmp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   3
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7080
      Width           =   555
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00244244&
      Caption         =   "Jpg"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   2
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Save Formatted File Name JPG"
      Top             =   7380
      Width           =   555
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00244244&
      Caption         =   "Tif"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   1
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Save Tif and OCR"
      Top             =   7080
      Width           =   555
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00244244&
      Caption         =   "Gif"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Index           =   0
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Save Formatted File Name GIF"
      Top             =   6780
      Width           =   555
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0073AEB6&
      Caption         =   "ok"
      Height          =   210
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Ok for this Slider"
      Top             =   6870
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0073AEB6&
      Caption         =   "ok"
      Height          =   210
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Ok for this Slider"
      Top             =   7470
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0073AEB6&
      Caption         =   "ok"
      Height          =   210
      Left            =   9060
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Ok for this Slider"
      Top             =   7180
      Visible         =   0   'False
      Width           =   315
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   210
      Left            =   7860
      TabIndex        =   18
      Top             =   7180
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   370
      _Version        =   393216
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   210
      Left            =   7860
      TabIndex        =   16
      Top             =   7470
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   370
      _Version        =   393216
      TickStyle       =   3
   End
   Begin THBImageLibCtl.THBImageFrame THBImageFrame1 
      Height          =   855
      Left            =   8040
      TabIndex        =   13
      Top             =   1770
      Width           =   975
      _cx             =   1246904
      _cy             =   1246692
      BackColor       =   16777152
      BackColorOn     =   0   'False
      BorderWidth     =   0
      BorderColor     =   255
      BorderStyle     =   4
      Begin THBImageLibCtl.THBImage THBImage2 
         Height          =   855
         Left            =   30
         TabIndex        =   14
         ToolTipText     =   "Undo Last Item"
         Top             =   0
         Width           =   900
         _cx             =   1587
         _cy             =   1508
         KeepAspect      =   -1  'True
         PositionPic     =   4
         StretchPic      =   3
         BackColor       =   12582912
         BackColorOn     =   -1  'True
         BorderWidth     =   0
         BorderColor     =   0
         BorderStyle     =   3
         MirrorVertical  =   0   'False
         MirrorHorizontal=   0   'False
         PictureIsTransparent=   -1  'True
         Scrolling       =   -1  'True
         DynamicZoom     =   -1  'True
         Gradient        =   0
         GradientSimpColFrom=   1073741824
         GradientSimpColTo=   1073742079
         GradientSimpHorizontal=   0   'False
         GradientUseSystemFunctions=   0   'False
         Halftone16Bit   =   0
         PopupMenuNewStyle=   -1  'True
         PopupMenu       =   -1  'True
         MoveWithPreview =   -1  'True
         HyperlinkURL    =   ""
         PreviewScrollWindow=   -1  'True
         PreviewScrollWidth=   200
         PreviewScrollHeight=   200
         Clipboard       =   -1  'True
         ImgContainer    =   0
         MousePointer    =   32512
         EnableKeyboardHandling=   -1  'True
         BackgroundPicPosition=   7
         BackgroundPicStretch=   3
         BackgroundPicKeepAspect=   0   'False
         MagnificationWindowWidth=   160
         MagnificationWindowHeight=   100
         SysKey          =   $"Angst.frx":1992C4
         NumGradientCols =   2
         NumGradientRows =   2
         GradientPoint000000PercX=   0
         GradientPoint000000PercY=   0
         GradientPoint000000ThbCol=   1073741824
         GradientPoint001000PercX=   100
         GradientPoint001000PercY=   0
         GradientPoint001000ThbCol=   1073742079
         GradientPoint000001PercX=   0
         GradientPoint000001PercY=   100
         GradientPoint000001ThbCol=   1073807104
         GradientPoint001001PercX=   100
         GradientPoint001001PercY=   100
         GradientPoint001001ThbCol=   1090453504
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   9810
      TabIndex        =   4
      Top             =   2505
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H8000000D&
      Caption         =   "Date"
      ForeColor       =   &H000000FF&
      Height          =   250
      Left            =   -45
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Sort by Date"
      Top             =   795
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H8000000D&
      Caption         =   "Name"
      ForeColor       =   &H00FFFF00&
      Height          =   250
      Left            =   -45
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Sort by Name"
      Top             =   555
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00343634&
      Height          =   450
      Left            =   105
      Picture         =   "Angst.frx":199462
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Load all graphic files from Selected Directory"
      Top             =   3375
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00547AAC&
      Height          =   450
      Left            =   105
      Picture         =   "Angst.frx":199D2C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Load all text files from Selected Directory"
      Top             =   3840
      Width           =   495
   End
   Begin AngstArt.FileOPS FileOPS1 
      Height          =   465
      Left            =   9570
      TabIndex        =   5
      Top             =   105
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   820
   End
   Begin AngstArt.cmdopen CmDlg 
      Left            =   11475
      Top             =   285
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin DirectoryProcesses.DirProcess DirProcess 
      Height          =   495
      Left            =   -540
      TabIndex        =   1
      Top             =   9645
      Visible         =   0   'False
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   873
   End
   Begin VB.FileListBox File1 
      Height          =   1140
      Left            =   9510
      Pattern         =   "*.psd;*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga"
      ReadOnly        =   0   'False
      System          =   -1  'True
      TabIndex        =   3
      Top             =   1140
      Visible         =   0   'False
      Width           =   1755
   End
   Begin THBImageLibCtl.THBImage THBImage3 
      Height          =   1455
      Left            =   9840
      TabIndex        =   17
      Top             =   5745
      Visible         =   0   'False
      Width           =   2280
      _cx             =   4022
      _cy             =   2566
      KeepAspect      =   -1  'True
      PositionPic     =   4
      StretchPic      =   3
      BackColor       =   12582912
      BackColorOn     =   -1  'True
      BorderWidth     =   0
      BorderColor     =   0
      BorderStyle     =   3
      MirrorVertical  =   0   'False
      MirrorHorizontal=   0   'False
      PictureIsTransparent=   -1  'True
      Scrolling       =   -1  'True
      DynamicZoom     =   -1  'True
      Gradient        =   0
      GradientSimpColFrom=   1073741824
      GradientSimpColTo=   1073742079
      GradientSimpHorizontal=   0   'False
      GradientUseSystemFunctions=   0   'False
      Halftone16Bit   =   0
      PopupMenuNewStyle=   -1  'True
      PopupMenu       =   -1  'True
      MoveWithPreview =   -1  'True
      HyperlinkURL    =   ""
      PreviewScrollWindow=   -1  'True
      PreviewScrollWidth=   200
      PreviewScrollHeight=   200
      Clipboard       =   -1  'True
      ImgContainer    =   0
      MousePointer    =   32512
      EnableKeyboardHandling=   -1  'True
      BackgroundPicPosition=   7
      BackgroundPicStretch=   3
      BackgroundPicKeepAspect=   0   'False
      MagnificationWindowWidth=   160
      MagnificationWindowHeight=   100
      SysKey          =   $"Angst.frx":19A5F6
      NumGradientCols =   2
      NumGradientRows =   2
      GradientPoint000000PercX=   0
      GradientPoint000000PercY=   0
      GradientPoint000000ThbCol=   1073741824
      GradientPoint001000PercX=   100
      GradientPoint001000PercY=   0
      GradientPoint001000ThbCol=   1073742079
      GradientPoint000001PercX=   0
      GradientPoint000001PercY=   100
      GradientPoint000001ThbCol=   1073807104
      GradientPoint001001PercX=   100
      GradientPoint001001PercY=   100
      GradientPoint001001ThbCol=   1090453504
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   210
      Left            =   7860
      TabIndex        =   23
      Top             =   6870
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   370
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10080
      Top             =   6360
   End
   Begin VB.CommandButton Command24 
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      Caption         =   ">>"
      Height          =   285
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   95
      ToolTipText     =   "Select Last Frame to Delete TO"
      Top             =   7080
      Width           =   345
   End
   Begin VB.CommandButton Command23 
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      Caption         =   "<<"
      Height          =   270
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "Select First Frame to DELETE"
      Top             =   6810
      Width           =   345
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   8190
      Picture         =   "Angst.frx":19A794
      ScaleHeight     =   615
      ScaleWidth      =   645
      TabIndex        =   90
      Top             =   2250
      Visible         =   0   'False
      Width           =   675
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10770
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   17
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Angst.frx":314356
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Angst.frx":314739
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Angst.frx":314AFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Angst.frx":314EF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Angst.frx":3152B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Angst.frx":3153C5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5415
      Left            =   915
      TabIndex        =   97
      Top             =   1095
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   9551
      _Version        =   393217
      BackColor       =   16576
      ScrollBars      =   2
      TextRTF         =   $"Angst.frx":3154CA
   End
   Begin VB.CommandButton Haburabadooda14 
      Appearance      =   0  'Flat
      BackColor       =   &H0073AEB6&
      Height          =   585
      Left            =   810
      Picture         =   "Angst.frx":315572
      Style           =   1  'Graphical
      TabIndex        =   105
      ToolTipText     =   "Delete All Pictures To the Recycle Bin"
      Top             =   0
      Width           =   525
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   1365
      Left            =   3540
      TabIndex        =   89
      Top             =   3180
      Visible         =   0   'False
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   2408
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   915
      TabIndex        =   8
      Top             =   1095
      Visible         =   0   'False
      Width           =   6600
      ExtentX         =   11642
      ExtentY         =   9551
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   -1  'True
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.FileListBox File2 
      Height          =   300
      Left            =   4710
      Pattern         =   "*.bmp"
      TabIndex        =   137
      Top             =   8070
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   7935
      ScaleHeight     =   435
      ScaleWidth      =   900
      TabIndex        =   140
      Top             =   2325
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00400040&
      Height          =   525
      Left            =   1875
      Picture         =   "Angst.frx":315AE2
      Style           =   1  'Graphical
      TabIndex        =   126
      ToolTipText     =   "Select a Rectangular Region of the Image to Modify"
      Top             =   7890
      Width           =   585
   End
   Begin VB.CommandButton Command33 
      BackColor       =   &H00000000&
      Height          =   525
      Left            =   1890
      Picture         =   "Angst.frx":3163AC
      Style           =   1  'Graphical
      TabIndex        =   128
      ToolTipText     =   "Select a Rectangular Region of the Image to Modify"
      Top             =   7890
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton Haburabadooda10 
      BackColor       =   &H00364048&
      Height          =   375
      Left            =   8970
      Picture         =   "Angst.frx":316C76
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "Select Any Area and Capture"
      Top             =   540
      Width           =   345
   End
   Begin AngstArt.ocxFormShape ocxFormShape1 
      Left            =   8895
      Top             =   585
      _ExtentX        =   794
      _ExtentY        =   873
      Shape           =   4
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H80000007&
      Height          =   1680
      Left            =   4650
      ScaleHeight     =   1620
      ScaleWidth      =   3090
      TabIndex        =   158
      Top             =   6555
      Width           =   3150
      Begin VB.OptionButton Option1 
         BackColor       =   &H00242614&
         Caption         =   "Literat"
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
         Index           =   19
         Left            =   2445
         Style           =   1  'Graphical
         TabIndex        =   178
         ToolTipText     =   "Send to microphones"
         Top             =   1170
         Width           =   600
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00242614&
         Caption         =   "Travel"
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
         Index           =   18
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   177
         ToolTipText     =   "Send to microphones"
         Top             =   1170
         Width           =   600
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00242614&
         Caption         =   "MP3"
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
         Index           =   17
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   176
         ToolTipText     =   "Send to microphones"
         Top             =   1170
         Width           =   600
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00242614&
         Caption         =   "Lang"
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
         Index           =   16
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   175
         ToolTipText     =   "Send to microphones"
         Top             =   1170
         Width           =   600
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00242614&
         Caption         =   "MyPics"
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
         Index           =   15
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   174
         ToolTipText     =   "Send to microphones"
         Top             =   1170
         Width           =   600
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
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   173
         ToolTipText     =   "Send to WordPro"
         Top             =   795
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
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   172
         ToolTipText     =   "Send to Science"
         Top             =   795
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
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   171
         ToolTipText     =   "Send to Telecom"
         Top             =   795
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   170
         ToolTipText     =   "Send to Music"
         Top             =   795
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
         Left            =   2445
         Style           =   1  'Graphical
         TabIndex        =   169
         ToolTipText     =   "Send to Multimedia"
         Top             =   795
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
         Left            =   2445
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Send to Personal"
         Top             =   75
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
         Left            =   2445
         Style           =   1  'Graphical
         TabIndex        =   167
         ToolTipText     =   "Send to Utilities"
         Top             =   465
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   166
         ToolTipText     =   "Send to Gif"
         Top             =   465
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
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "Send to legal"
         Top             =   465
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
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   164
         ToolTipText     =   "Send to Receipts"
         Top             =   465
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
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Send to 1down"
         Top             =   465
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
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   162
         ToolTipText     =   "Send to microphones"
         Top             =   75
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
         Left            =   630
         Style           =   1  'Graphical
         TabIndex        =   161
         ToolTipText     =   "Send to Vb"
         Top             =   75
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
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   160
         ToolTipText     =   "Send to Guitar"
         Top             =   75
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
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Send to Medical"
         Top             =   75
         Width           =   600
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Move"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   5955
      TabIndex        =   155
      Top             =   8235
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   195
      TabIndex        =   154
      Top             =   7755
      Width           =   765
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   300
      Left            =   11520
      TabIndex        =   153
      Top             =   855
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   285
      Left            =   11580
      TabIndex        =   152
      Top             =   705
      Width           =   825
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   1065
      TabIndex        =   146
      Top             =   6525
      Width           =   45
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   270
      Left            =   2055
      TabIndex        =   135
      Top             =   8655
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   10095
      Top             =   8565
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Filters"
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
      Height          =   435
      Index           =   7
      Left            =   7740
      TabIndex        =   55
      Top             =   6630
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label FxLabel2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6480
      TabIndex        =   24
      Top             =   6900
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label FxLabel1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6480
      TabIndex        =   19
      Top             =   7180
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label FxLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6480
      TabIndex        =   15
      Top             =   7470
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   7545
      Picture         =   "Angst.frx":317085
      Top             =   -30
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   7635
      Picture         =   "Angst.frx":31794F
      Top             =   -30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   8895
      Picture         =   "Angst.frx":318219
      Top             =   30
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   8940
      Picture         =   "Angst.frx":318AE3
      Top             =   15
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuRegion 
      Caption         =   "Region"
      Visible         =   0   'False
      Begin VB.Menu mnuInvertr 
         Caption         =   "Invert"
      End
      Begin VB.Menu mnuEdger 
         Caption         =   "Find Edge"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
      End
      Begin VB.Menu mnuMzoom 
         Caption         =   "Manual Zoom"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "Resize"
         Begin VB.Menu mnu16 
            Caption         =   "16x16"
         End
         Begin VB.Menu mnu32 
            Caption         =   "32x32"
         End
         Begin VB.Menu mnu65 
            Caption         =   "65x65"
         End
         Begin VB.Menu mnu320 
            Caption         =   "320x200"
         End
         Begin VB.Menu mnu640 
            Caption         =   "640x480"
         End
         Begin VB.Menu mnu800 
            Caption         =   "800x600"
         End
         Begin VB.Menu mnu1024 
            Caption         =   "1024x768"
         End
         Begin VB.Menu mnuStretch 
            Caption         =   "Stretch"
         End
         Begin VB.Menu mnuCustom 
            Caption         =   "Custom"
         End
      End
      Begin VB.Menu mnuResizette 
         Caption         =   "Resizette"
         Visible         =   0   'False
         Begin VB.Menu mn640 
            Caption         =   "640x480"
         End
         Begin VB.Menu mn800 
            Caption         =   "800x600"
         End
         Begin VB.Menu mn1024 
            Caption         =   "1024x768"
         End
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "Crop"
      End
      Begin VB.Menu mnuCropCapture 
         Caption         =   "Crop-Capture"
      End
      Begin VB.Menu mnuRotate 
         Caption         =   "Rotate"
         Begin VB.Menu mnu180 
            Caption         =   "180 CW"
         End
         Begin VB.Menu mnu90CW 
            Caption         =   "90 CW"
         End
         Begin VB.Menu mnu90CCW 
            Caption         =   "90 CCW"
         End
         Begin VB.Menu mnuHoriz 
            Caption         =   "Mirror Horizontal"
         End
         Begin VB.Menu mnuVertical 
            Caption         =   "Mirror Vertical"
         End
         Begin VB.Menu mnurCustom 
            Caption         =   "Custom"
         End
      End
      Begin VB.Menu mnuGetBMP 
         Caption         =   "Get BMP from Resources"
      End
      Begin VB.Menu mnuEffects 
         Caption         =   "Effx"
         Begin VB.Menu mnuAtal 
            Caption         =   "Atalasoft"
         End
         Begin VB.Menu mnuBright 
            Caption         =   "Brightness/Contrast"
         End
         Begin VB.Menu mnuColors 
            Caption         =   "Colors"
            Begin VB.Menu mnuHSV 
               Caption         =   "HSV"
            End
            Begin VB.Menu mnuGray 
               Caption         =   "Gray Scale"
            End
            Begin VB.Menu mnuBlackWhite 
               Caption         =   "Black & White"
            End
            Begin VB.Menu mnuInvert 
               Caption         =   "Invert"
            End
            Begin VB.Menu mnuPickColor 
               Caption         =   "Pick Color to Replace"
            End
            Begin VB.Menu mnuRGB 
               Caption         =   "RGB"
            End
         End
         Begin VB.Menu mnuFilterz 
            Caption         =   "Filters"
            Begin VB.Menu mnuSharpen 
               Caption         =   "Sharpen"
               Begin VB.Menu mnuSharpenScr1 
                  Caption         =   "Sharpen 1"
               End
               Begin VB.Menu mnuSharpenScr2 
                  Caption         =   "Mean Removal"
               End
               Begin VB.Menu mnuSharpenMatrix1 
                  Caption         =   "Sharpen Matrix (-1 5 -1)"
                  Visible         =   0   'False
               End
               Begin VB.Menu mnuSharpenScr3 
                  Caption         =   "Sharpen Matrix (-2 11 -2)"
               End
            End
            Begin VB.Menu mnuBlur 
               Caption         =   "Blur"
               Begin VB.Menu mnuBlurru 
                  Caption         =   "Blur"
               End
               Begin VB.Menu mnuMediumBlur 
                  Caption         =   "Median Blur"
               End
               Begin VB.Menu mnuGauss 
                  Caption         =   "Gaussian Blur"
               End
            End
            Begin VB.Menu mnuEmboss 
               Caption         =   "Emboss"
               Begin VB.Menu mnuEmboss1 
                  Caption         =   "Emboss1"
               End
               Begin VB.Menu mnuEmboss2 
                  Caption         =   "Emboss2"
               End
               Begin VB.Menu mnuEmboss5 
                  Caption         =   "South"
               End
               Begin VB.Menu mnuEmboss3 
                  Caption         =   "East Emboss"
               End
               Begin VB.Menu mnuEmboss4 
                  Caption         =   "Southeast"
               End
            End
            Begin VB.Menu mnuEdge 
               Caption         =   "Edge Detection"
               Begin VB.Menu mnuDetect1 
                  Caption         =   "Detect1"
               End
               Begin VB.Menu mnuDetect2 
                  Caption         =   "Detect2"
               End
               Begin VB.Menu mnuDetect3 
                  Caption         =   "Detect3"
               End
               Begin VB.Menu mnuDetect4 
                  Caption         =   "Detect4"
                  Visible         =   0   'False
               End
               Begin VB.Menu mnuDetect5 
                  Caption         =   "Detect5"
                  Visible         =   0   'False
               End
               Begin VB.Menu mnuDetect6 
                  Caption         =   "Detect4"
               End
            End
            Begin VB.Menu mnuRandomFilter 
               Caption         =   "Random Filter"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuUserDefined 
               Caption         =   "User Defined"
            End
         End
         Begin VB.Menu mnuantialias 
            Caption         =   "Antialias"
         End
         Begin VB.Menu mnuDropshadow 
            Caption         =   "Drop Shadow"
         End
      End
      Begin VB.Menu mnuScan 
         Caption         =   "Scan"
      End
      Begin VB.Menu mnuRecycle 
         Caption         =   "Recycle"
         Begin VB.Menu mnuExplore 
            Caption         =   "Explore Recycle Bin"
         End
         Begin VB.Menu mnuDeleteRB 
            Caption         =   "Delete to Recycle Bin"
         End
         Begin VB.Menu mnuEmptyRB 
            Caption         =   "Empty Recycle Bin"
         End
      End
      Begin VB.Menu mnuAssociate 
         Caption         =   "Associate Graphic Files to this program"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Slide Show"
         Visible         =   0   'False
         Begin VB.Menu mnuReverse 
            Caption         =   "Reverse"
         End
         Begin VB.Menu mnuReset 
            Caption         =   "Reset"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuChangeUp 
            Caption         =   "Interval"
            Visible         =   0   'False
         End
      End
   End
End
Attribute VB_Name = "Angst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim clsWinApi As CWinAPI 'Declare a class variable
Dim ie As THBImageEdit
Dim ie1 As THBImageEdit
Dim ie2 As THBImageEdit

Private WithEvents ieevents As THBImageEdit
Attribute ieevents.VB_VarHelpID = -1
Dim rtn As Long
Dim OcrTiff As String
Dim nMarkerPictureCounter As Long   'Just to use different marker pictures
Dim gnMode As Long
'Dim PDFTEST As New PDF
Const MODE_PULLQUOTE As Long = 1
Const MODE_ZOOMTORECT As Long = 2
Dim ieLimit As New THBImageEdit
Dim lngLeft As Long
Dim lngTop As Long
Dim lngRight As Long
Dim lngBottom As Long
Dim pBytes As Long
Dim lngSizeBytes As Long
Dim rg As THBRegion
Dim FormatNow As String
Dim MoveFlag As Boolean
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Const CB_FINDSTRINGEXACT = &H158

'Private Declare Function TWAIN_AcquireToFilename Lib "TWAIN32d.DLL" (ByVal hwndApp As Long, ByVal bmpFileName As String) As Integer
'Private Declare Function TWAIN_IsAvailable Lib "TWAIN32d.DLL" () As Long
'Private Declare Function TWAIN_SelectImageSource Lib "TWAIN32d.DLL" (ByVal hwndApp As Long) As Long
Dim OK, OK1, OK2 As Boolean
Dim Fx As Integer
Dim dPercent As Double
Dim dPercent1 As Double
Dim dPercent2 As Double
Dim nHue As Long
Dim nSat As Long
Dim nValue As Long
Dim nDeltaX As Long
Dim nDeltaY As Long
Dim errorstring As String
Dim sourcebmp, targetgif As String
Dim lngNewWidth As Long
Dim lngNewHeight, RetVal As Long
Dim Ration As Single
Dim arMatrix() As Long
Dim nDivisor As Long
Dim varMatrix As Variant
Dim AspectWidth, AspectHeight As Long
Dim AspectH, AspectW As Long
Dim k As Long
Dim ReturnValuePath As String
Dim ReturnValue As Long
Dim TextExt(4) As String
Dim TextExt1(2) As String
Dim szzFile As String
Dim Step As Boolean
Dim Crap As Boolean
Dim Limit As Boolean
Dim RTF As Boolean
Dim MoveMe As Boolean
Dim PicThere As Boolean
Dim TxtThere As Boolean
Dim Txt1There As Boolean
Dim VideoThere As Boolean
Dim ScanFlag As Boolean
Private Resizing As Boolean
Dim SlideSHowFlagget As Boolean
Dim Resizette As Integer
Dim Overlaid As Boolean
Dim ReNameMoveFile As String

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = False
Image6.Visible = True
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub MouseWheel1_WheelScroll(ByVal ConnectedTo As Long, ByVal Direction As Long, ByVal Shift As WheelCtl.KeyDown)
On Error Resume Next
    Select Case Shift 'shift could also be used to control horizontal or vertical scrolling
      Case WheelCtl.KeyCntl
        Direction = Direction * 2
      Case WheelCtl.KeyShift
        Direction = Direction * 4
      Case WheelCtl.KeyBoth
        Direction = Direction * 8
    End Select
    Select Case ConnectedTo
      Case THBImage1.hwnd
        CList4.ListIndex = CList4.ListIndex - Direction
        CList4.Text = CList4.List(CList4.ListIndex)
      Case Else  'NOT CONNECTEDTO...
      
    End Select

End Sub

Private Sub AnyShape1_MouseEnter()
'Tips.Visible = True
'Tips.Caption = "Box Capture"
End Sub

Private Sub AnyShape1_MouseExit()
'Tips.Visible = False
End Sub

Private Sub AnyShape1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If RegionsFlag = False Then Undoit
CList4.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
SendKeys "{ESC}", True
''SetTopMostWindow Me.hWnd, False
'Ontop.Visible = False
'OnBottom.Visible = True
Load frmCapture
frmCapture.Show
End Sub


Public Sub Anyshape10_Click()
If Mht = True Or RTF = True Then Exit Sub
Dim i As Long
Angst.Visible = False
Load Form2
Form2.Show
Form2.CList4.Clear
SlideShowFlag = True
For i = 0 To Angst.CList4.ListCount - 1
        Form2.CList4.AddItem Angst.CList4.List(i)
Next
End Sub



Private Sub AnyShape11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DirectSR1.Deactivate
Unload Me
Set Angst = Nothing
'End

End Sub



Private Sub AnyShape12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.WindowState = 1

End Sub


Private Sub AnyShape13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Hide
Load VideoLibrary
If List1.ListCount <> 0 Then
    VideoLibrary.Moviee.Clear
    For i = 0 To List1.ListCount - 1
            VideoLibrary.Moviee.AddItem List1.List(i)
            VideoLibrary.Moviee.Text = VideoLibrary.Moviee.List(0)
    Next
End If
'VideoLibrary.Show

End Sub

Private Sub AnyShape14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2

End Sub


Private Sub AnyShape15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ActionFlag As Long
Dim i As Integer
On Error Resume Next
If CList4.Text <> "Drag & Drop on me now, if you dare" And CList4.Text <> "" Then
    If Dir(CList4.Text) <> "" Then
        Screen.MousePointer = 11
        ActionFlag = FOF_ALLOWUNDO
        ShellDeleteOne CList4.Text, ActionFlag
        For i = 0 To CList4.ListCount - 1
            If Dir(CList4.List(i)) = "" Then CList4.RemoveItem i
        Next
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(CList4.ListIndex)
        LoadFileAndUpdateDisplay CList4.List(CList4.ListIndex + 1)
        'clist4.setfocus
        Screen.MousePointer = 0
    End If
End If
End Sub

Private Sub AnyShape16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Moove As Boolean
Dim J As Long
Dim ToFileName As String
Moove = False
If CList4.Text <> "Drag & Drop on me now, if you dare" Then
    If Dir(CList4.Text) <> "" Then
        ToFileName = GetFileTitle(CList4.Text)
        Do While Moove = False
        'MsgBox ToFileName
        If Dir(App.path & "\SaveBin\" & ToFileName) = "" Then
            FileOPS1.MoveFile CList4.Text, App.path & "\SaveBin\" & ToFileName
            LoadFileAndUpdateDisplay CList4.Text
            Moove = True
        Else
            ToFileName = ToFileName & Str(J)
            J = J + 1
        End If
        Loop
    End If
End If
'clist4.setfocus

End Sub


Private Sub AnyShape17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to Delete All Pictures in the List?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Clear Box and Optional Delete"   ' Define title.
'Help = "DEMO.HLP"   ' Define Help file.
Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If RegionsFlag = False Then Undoit
If Response = vbYes Then   ' User chose Yes.
    Dim ActionFlag As Long
    Dim i As Integer
    On Error Resume Next
    If CList4.Text <> "Drag & Drop on me now, if you dare" And CList4.Text <> "" Then
        If Dir(CList4.Text) <> "" Then
            Screen.MousePointer = 11
            ''SetTopMostWindow Me.hWnd, False
            ActionFlag = FOF_ALLOWUNDO
            For i = 0 To CList4.ListCount - 1
                If Exists(CList4.List(i)) = True Then
                    ShellDeleteOne CList4.List(i), ActionFlag
                End If
            Next
            File1.Refresh
            '''SetTopMostWindow Me.hWnd, True
            Screen.MousePointer = 0
        End If
    End If
End If
CList4.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
List1.Clear
End Sub



Private Sub AnyShape18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If RegionsFlag = False Then Undoit
    hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
    With Angst
        .CmDlg.InitialDir = App.path & "\SaveBin"
        '.CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = False   'True 'Allow multi select
        .CmDlg.DialogTitle = "Select file (s) to open" 'Set dialog title
        '.CmDlg.Filter = "All Files (*.*)|*.*"

'FileDialog.sFilter = "All Graphic Files (*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga)|*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga|Word Files (*.doc)|*.doc|Html Files (*.htm;*.html)|*.htm;*.html|Batch Files (*.bat)|*.bat|INI Files (*.ini)|*.ini|All Files (*.*)|*.*|"
        .CmDlg.Filter = "All Graphic Files (*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga;*.psd)" & Chr$(0) & "*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga;*.psd" & Chr$(0) _
        & "Gif Files (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "Jpeg Files (*.jpeg;*.jpg)" & Chr$(0) & "*.jpeg;*.jpg" & Chr$(0) & "Icon/Cursor (*.ico;*.cur)" & Chr$(0) & "*.ico;*.cur" & Chr$(0) & "BMP Files (*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0) & "Meta Files (*.wmf;*.emf)" & Chr$(0) & "*.wmf;*.emf" & Chr$(0) & "PCX Files (*.pcx)" & Chr$(0) & "*.pcx" & Chr$(0) & "TIF Files (*.tif;*.tiff)" & Chr$(0) & "*.tif;*.tiff" _
        & Chr$(0) & "PNG Files (*.png)" & Chr$(0) & "*.png" & Chr$(0) & "TGA Files (*.tga)" & Chr$(0) & "*.tga" & Chr$(0) & "Photoshop Files (*.psd)" & Chr$(0) & "*.psd" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"


        '"Psd Files (*.psd)|*.psd; *.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga;*.txt
        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen 'Show open dialog
        If hHook Then UnhookWindowsHookEx hHook
        ThePicture = CmDlg.cFileName(1)
        CList4.Clear
        CList4.AddItem CmDlg.cFileName(1)
        CList4.Text = CList4.List(0)
        LoadFileAndUpdateDisplay CmDlg.cFileName(1)
    End With
End Sub

Private Sub AnyShape19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
THBImage1.ZoomFit
THBImage1.Redraw
If RegionsFlag = False Then Undoit
End Sub


Private Sub AnyShape2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If RegionsFlag = False Then Undoit
''SetTopMostWindow Me.hWnd, False
'Ontop.Visible = False
'OnBottom.Visible = True
CList4.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
Load Rdouble
Rdouble.Show
Load frmCapture
frmCapture.Show
End Sub

Private Sub AnyShape20_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
SendKeys "%{DOWN}", True

End Sub

Private Sub AnyShape21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
SendKeys "%{UP}", True
End Sub

Private Sub AnyShape23_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Private Sub AnyShape24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub AnyShape25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Private Sub AnyShape22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub

Private Sub AnyShape26_MouseEnter()
'AnyShape26.ToolTipText = "Browse For Directory. Presently: " & File1.path

End Sub

Public Sub Aa_Click()
    On Error Resume Next
    Set ie.Picture = Picture2.Picture
    Set THBImage1.Picture = ie.THBStdPicture
    Unload frmPaint
End Sub

Private Sub AnyShape26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
Dim FileExtension As String
Dim Filenamme As String, ThbIT As String, i As Long
On Error Resume Next
    With Angst
        .CmDlg.InitialDir = File1.path
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = True   'True 'Allow multi select
        .CmDlg.DialogTitle = "Open All" 'Set dialog title
        '.CmDlg.DefaultFilename = Format(Now, "ddmmyyhhmmss")
       '.CmDlg.Filter = "All Files" & Chr$(0) & "*.gif;*.jpg;*.ico;*.bmp;*.wmf;*.pcx;*.tif;*.png;*.tga;*.psd;*.m1v;*.mpg;*.mpeg;*.m2v;*.avi;*.asf;*.mov;*.wmv;*.mp3;*.wma;*.wav;*.mht;*.html;*.htm;*.pdf;*.txt;*.rtf" & Chr$(0) &
        .CmDlg.Filter = "All Graphic/Text Files" & Chr$(0) & "*.gif;*.jpg;*.ico;*.bmp;*.wmf;*.pcx;*.tif;*.png;*.tga;*.psd;*.mht;*.html;*.htm;*.pdf;*.txt;*.rtf" & Chr$(0) & _
        "Gif Files (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "Jpeg Files (*.jpg)" & Chr$(0) & "*.jpg" & Chr$(0) & "Icon (*.ico)" & Chr$(0) & "*.ico" & Chr$(0) & "BMP Files (*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0) & "Meta Files (*.wmf)" & Chr$(0) & "*.wmf" & Chr$(0) & "PCX Files (*.pcx)" & Chr$(0) & "*.pcx" & Chr$(0) & "TIF Files (*.tif)" & Chr$(0) & "*.tif" _
        & Chr$(0) & "PNG Files (*.png)" & Chr$(0) & "*.png" & Chr$(0) & "TGA Files (*.tga)" & Chr$(0) & "*.tga" & Chr$(0) & "Photoshop Files (*.psd)" & Chr$(0) & "*.psd" & Chr$(0) & "m1v Files (*.m1v)" & Chr$(0) & "*.m1v" _
        & Chr$(0) & "mht Files (*.mht)" & Chr$(0) & "*.mht" & Chr$(0) & "html Files (*.html)" & Chr$(0) & "*.html" _
        & Chr$(0) & "htm Files (*.htm)" & Chr$(0) & "*.htm" & Chr$(0) & "pdf Files (*.pdf)" & Chr$(0) & "*.pdf" & Chr$(0) & "txt Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "rtf Files (*.rtf)" & Chr$(0) & "*.rtf"
        '& Chr$(0) & "MPG Files (*.mpg)" & Chr$(0) & "*.mpg" & Chr$(0) & "mpeg Files (*.mpeg)" & Chr$(0) & "*.mpeg" & Chr$(0) & "m2v Files (*.m2v)" & Chr$(0) & "*.m2v" & Chr$(0) & "avi Files (*.avi)" & Chr$(0) & "*.avi" _
        '& Chr$(0) & "asf Files (*.asf)" & Chr$(0) & "*.asf" & Chr$(0) & "mov Files (*.mov)" & Chr$(0) & "*.mov" & Chr$(0) & "wmv Files (*.wmv)" & Chr$(0) & "*.wmv" & Chr$(0) & "mp3 Files (*.mp3)" & Chr$(0) & "*.mp3" _
        '& Chr$(0) & "wma Files (*.wma)" & Chr$(0) & "*.wma" & Chr$(0) & "wav Files (*.wav)" & Chr$(0) & "*.wav" & Chr$(0) & "mht Files (*.mht)" & Chr$(0) & "*.mht" & Chr$(0) & "html Files (*.html)" & Chr$(0) & "*.html" _
        '& Chr$(0) & "htm Files (*.htm)" & Chr$(0) & "*.htm" & Chr$(0) & "pdf Files (*.pdf)" & Chr$(0) & "*.pdf" & Chr$(0) & "txt Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "rtf Files (*.rtf)" & Chr$(0) & "*.rtf"
        
        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen
    End With
    'List8.Clear
    If Angst.CmDlg.cFileName(1) = "" Then Exit Sub
    File1.path = GetFilePath(Angst.CmDlg.cFileName(1))
    Open App.path & "\LastPath" For Output As #1
        Print #1, File1.path
    Close #1
    Patherino = File1.path
    Dim TempFile, Lisstcount
    Lisstcount = Angst.CList4.ListCount - 1
    
    For i = 1 To Angst.CmDlg.cFileName.Count
        If Trim(Angst.CmDlg.cFileName(i)) <> "" Then
            'List8.AddItem Angst.CmDlg.cFileName(i)
            CList4.AddItem Angst.CmDlg.cFileName(i)
        End If
    Next i
    Angst.CList4.ListIndex = Lisstcount + 1
    Angst.CList4_Click

'LoadClist4a
End Sub

Private Sub AnyShape27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu mnuRecycle
End Sub


Private Sub StartDir()
Dim Filepath As String
On Error Resume Next
Open App.path & "\LastPath" For Input As #1
    Line Input #1, Filepath
Close #1
If Trim(Filepath) <> "" Then
    File1.path = Filepath
Else
    File1.path = App.path
End If
File1.Refresh
LoadClist4
'clist4.setfocus
Label4.Caption = CList4.Text
If CList4.ListCount <> 0 Then CList4.ListIndex = 0
End Sub

Public Sub Images1()
Dim Filepath As String
On Error Resume Next
Filepath = App.path & "\Images"
File1.path = Filepath
File1.Refresh
LoadClist4
'clist4.setfocus
Label4.Caption = CList4.Text
If CList4.ListCount <> 0 Then CList4.ListIndex = 0

End Sub


Private Sub AnyShape29_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
File1.path = App.path & "\SaveBin"
LoadClist4
'clist4.setfocus

End Sub


Private Sub AnyShape3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If RegionsFlag = False Then Undoit
''SetTopMostWindow Me.hWnd, False
'Ontop.Visible = False
'OnBottom.Visible = True
CList4.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
Load Clipper
Load Rdouble1
Rdouble1.Show
End Sub


Private Sub AnyShape32_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
On Error Resume Next
'If CList4.Text = "Drag & Drop on me now, if you dare" Then Exit Sub
Painting = True
Undoit
Angst.Hide
Load frmPaint
frmPaint.Show
''ie.BMPUseRLE = True
'ie.SavePictureToFile App.path & "\Paint.bmp", thbifBMP
'Dim RetVal
'RetVal = Shell(App.path & "\Vbpaint.exe " & App.path & "\Paint.bmp", 1)
'RetVal = Shell(App.path & "\Vbpaint.exe " & CList4.Text, 1)

End Sub
Private Sub AnyShape5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'clist4.setfocus
End Sub

Private Sub AnyShape5_MouseEnter()
    'clist4.setfocus
    If Trim(CList4.Text) = "" Then CList4.Text = "Drag & Drop on me now, if you dare": WebBrowser1.Navigate2 ("about:blank")
    'Tips.Visible = True
    'Tips.Caption = "Activates Wheel Mouse Scroll"
End Sub

Private Sub AnyShape6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
butClipboardCopyTo_Click
End Sub


Private Sub AnyShape7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
butClipboardPasteFrom_Click
If RegionsFlag = False Then Undoit
End Sub


Private Sub AnyShape8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngPrintWidth  As Long
    Dim lngPrintHeight As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    CmDlg.ShowPrinter
    Printer.Print Space(1)
    'Calculate destination rect
    lngPrintWidth = 1.65 * Printer.Width
    lngPrintHeight = 1.65 * Printer.Height
    lngLeft = 300   'lngPrintWidth / 4
    lngTop = 500   'lngPrintHeight / 4
    lngRight = lngLeft + lngPrintWidth / 2
    lngBottom = lngTop + lngPrintHeight / 2
    ie.PrintPicAligned Printer.hDC, lngLeft, lngTop, lngRight, lngBottom, thbguTwips, True, thbPosCC, thbStretchBoth
    Printer.EndDoc

End Sub

Private Sub AnyShape9_MouseEnter()
'Tips.Visible = True
'Tips.Caption = "Reset Substance"
    'ShowOn = False
     'Timer2.Enabled = False
    'Slide_Show.Value = False
End Sub


Private Sub AnyShape9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RegionsFlag = False Then Undoit
LoadFileAndUpdateDisplay App.path & "\" & "Munc.BMP"
ThePicture = App.path & "\" & "Munc.BMP"
Set THBImage1.Picture = ie.THBStdPicture
'CList4.Clear
'CList4.Text = "Drag & Drop on me now, if you dare"
'clist4.setfocus
End Sub

Private Sub butClipboardCopyTo_Click()
On Error Resume Next
    ie.ClipboardCopyTo
    UpdatePicInfo
End Sub
Public Sub butClipboardPasteFrom_Click()
    On Error Resume Next
    ie.ClipboardPasteFrom
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End Sub

Private Sub Clearit_Click()
On Error Resume Next
LoadFileAndUpdateDisplay App.path & "\" & "help.bmp"
ThePicture = App.path & "\" & "help.bmp"
If RegionsFlag = False Then Undoit
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Check1.Value = 1 Then
    SetTopMostWindow Me.hwnd, True
Else
    SetTopMostWindow Me.hwnd, False
End If
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Open App.path & "\Aspect" For Output As #1
    Print #1, Check3.Value
Close #1
 On Error Resume Next
If Check3.Value = 1 Then
    If Trim(Spinner1.Text) <> "" Then
        VScroll2.Value = -1 * Int(Spinner1.Text * AspectH / AspectW)
    End If
    If Trim(Spinner2.Text) <> "" Then
        VScroll1.Value = -1 * Int(Spinner1.Text * AspectW / AspectH)
    End If
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
    Picture4.BackColor = &HFF&
    Text1.BackColor = &HFF&
Else
    Picture4.BackColor = &H0&
    Text1.BackColor = &H73AEB6
End If
End Sub

Private Sub CList4_Change()
If CList4.Text = "" Then
    Option2.Visible = False
    Option3.Visible = False
    'CList4.Text = "Drag & Drop on me now, if you dare"
Else
        Option2.Visible = True
        Option3.Visible = True
End If

End Sub
Private Sub Propertiz()
DoEvents
    Dim varImageProperties As Variant
    Dim i As Long: Dim SSTT As String
    varImageProperties = ie.ImageProperties
    If Not IsEmpty(varImageProperties) Then
        For i = 0 To UBound(varImageProperties) Step 2
            If InStr(varImageProperties(i), "WIDTH") <> 0 Or InStr(varImageProperties(i), "HEIGHT") <> 0 Then
                Select Case varImageProperties(i)
                Case "WIDTH"
                     wID = varImageProperties(i + 1)
                     
                Case "HEIGHT"
                    Hei = varImageProperties(i + 1)
                End Select
            End If
        Next i
    End If
Text1.Text = CList4.Text
Label14 = "Number " & CList4.ListIndex & " of " & CList4.ListCount & "   Dimensions " & wID & " / " & Hei
End Sub

Public Sub CList4_Click()
On Error Resume Next
Dim TempFile As String, i As Integer
Patherino = Replace(GetFilePath(CList4.Text) & "\", "\\", "\")
Open App.path & "\LastPath" For Output As #1
    Print #1, Patherino
Close #1
Exterino = Right(CList4.Text, 4)
i = InStrRev(GetFileTitle(CList4.Text), ".")
TempFile = GetFileTitle(CList4.Text)
Text1.Text = Left(TempFile, i - 1)

Label5.Caption = CList4.Text
Dim Extntion As String, EExt As Integer
Extntion = Trim(LCase(Right(CList4.Text, Len(CList4.Text) - InStrRev(CList4.Text, "."))))
'MsgBox Extntion
For EExt = 0 To 13
    If Extntion = PictureExt(EExt) Then
        Mht = False
        RTF = False
        WebBrowser1.Visible = False
        RichTextBox1.Visible = False
        THBImage1.Visible = True
        Exit For
    End If
Next
For EExt = 0 To 3
    If Extntion = TextExt(EExt) Then
        THBImage1.Visible = False
        Mht = True
        RTF = False
        WebBrowser1.Visible = True
        RichTextBox1.Visible = False
        Exit For
    End If
Next
For EExt = 0 To 1
    If Extntion = TextExt1(EExt) Then
        THBImage1.Visible = False
        Mht = False
        RTF = True
        WebBrowser1.Visible = False
        RichTextBox1.Visible = True
        Exit For
    End If
Next

    If CList4.List(CList4.ListIndex) <> "Drag & Drop on me now, if you dare" And CList4.List(CList4.ListIndex) <> "" Then
        Haburabadooda22.Visible = False
        If Mht = False And RTF = False Then
          THBImage1.Visible = True
          WebBrowser1.Visible = False
          RichTextBox1.Visible = False
          LoadFileAndUpdateDisplay CList4.List(CList4.ListIndex)
          ThePicture = CList4.List(CList4.ListIndex)
          TheNumber = CList4.ListIndex
          Propertiz
          If Haburabadooda2.Visible = True Then
              WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & Trim(ThePicture) & "'></img></p></body></html>"
              WebBrowser2.Top = THBImage1.Top
              WebBrowser2.Left = THBImage1.Left
              WebBrowser2.Height = THBImage1.Height
              WebBrowser2.Width = THBImage1.Width
          End If
          If LCase(Right(CList4.List(CList4.ListIndex), 3)) = "gif" Then
              Haburabadooda2.Visible = True
              Mht = False
              RTF = False
              RichTextBox1.Visible = False
          Else
              THBImage1.Visible = True
              WebBrowser1.Visible = False
              Mht = False
              RTF = False
              RichTextBox1.Visible = False
              Haburabadooda2.Visible = False
          End If
     Else
        Haburabadooda2.Visible = False
        If Mht = True And RTF = False Then
           THBImage1.Visible = False
           WebBrowser1.Visible = True
           RichTextBox1.Visible = False
           WebBrowser1.Navigate CList4.List(CList4.ListIndex)
        Else
            If Mht = False And RTF = True Then
                Haburabadooda22.Visible = True
                THBImage1.Visible = False
                WebBrowser1.Visible = False
                RichTextBox1.Visible = True
                RichTextBox1.LoadFile CList4.List(CList4.ListIndex)
                If LCase(Right(CList4.List(CList4.ListIndex), 3)) = "txt" Then
                    RichTextBox1.Font.Size = 10
                End If
            End If
        End If
     End If
    End If

End Sub

Private Sub CList4_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyDelete Then
    Label1.Caption = CList4.ListIndex
End Sub

Private Sub CList4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then
    Dim Filepath As String
    Dim ii As Long, NewMic As String, i As Long
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    On Error Resume Next
    NewMic = Label5.Caption
    Msg = "Do you want to Permanently DELETE This File?"   ' Define message.
    Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
    Title = "Permanently Delete This File"   ' Define title.
    Response = MsgBox(Msg, Style, Title)
    Screen.MousePointer = 11
If Response = vbYes Then
        If Mht = False And RTF = False Then
            If RegionsFlag = False Then Undoit
        End If
        Kill NewMic
        CList4.RemoveItem Val(Label1.Caption)
        'StartDir1
        CList4.ListIndex = Val(Label1.Caption)
        CList4.Text = CList4.List(CList4.ListIndex)
        CList4_Click
        ThePicture = CList4.Text
        If Mht = False And RTF = False Then
            RichTextBox1.Visible = False
            WebBrowser1.Visible = False
            THBImage1.Visible = True
            LoadFileAndUpdateDisplay CList4.List(0)
        End If
        If Mht = True And RTF = False Then
            RichTextBox1.Visible = False
            WebBrowser1.Visible = True
            THBImage1.Visible = False
            WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & Trim(ThePicture) & "'></img></p></body></html>"
            WebBrowser2.Top = THBImage1.Top
            WebBrowser2.Left = THBImage1.Left
            WebBrowser2.Height = THBImage1.Height
            WebBrowser2.Width = THBImage1.Width
        End If
        If Mht = False And RTF = True Then
            RichTextBox1.Visible = True
            WebBrowser1.Visible = False
            THBImage1.Visible = False
            RichTextBox1.LoadFile CList4.List(0)
        End If
        'clist4.setfocus
    End If
End If
If KeyCode = 18 Then OK1 = True
If KeyCode = 115 And OK1 = True Then Unload Me
Screen.MousePointer = 0
End Sub

Private Sub CList4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim XY As Long
Dim Please As String
On Error Resume Next
    CList4.Clear
    Option2.Value = True        'Sort by Name
    List1.Clear 'video
    List2.Clear 'pics
    List3.Clear 'pics
    List6.Clear 'mht, htm, pdf
    List7.Clear 'txt, rtf
    PicThere = False
    TxtThere = False
    Txt1There = False
    VideoThere = False
    
    File1.path = GetFilePath(Data.Files(1))
    File1.Refresh
    'set file path to whereever drag from
    'Open App.path & "\LastPath" For Output As #1
        'Print #1, File1.path
    'Close #1
    
    'loading lists and clist4
    For XY = 1 To Data.Files.Count
        Please = Data.Files(XY)
        If Trim(Please) = "" Then GoTo EmptyOne     'Don't add empty strings
        For i = 0 To 13 'pics
            If LCase(GetFileExtension(Please)) = PictureExt(i) Then
                    PicThere = True
                    CList4.AddItem Please
                    List3.AddItem Please    'sort name
                    List2.AddItem Format(FileDateTime(Please), "YYYYMMDDHHMMSS") & "*" & Please 'sort date
                Exit For
            End If
        Next
        For i = 0 To 3  'htm pdf
            If LCase(GetFileExtension(Please)) = TextExt(i) Then
                    TxtThere = True
                    List6.AddItem Please
                    Exit For
            End If
        Next
        For i = 0 To 1  'txt rtf
            If LCase(GetFileExtension(Please)) = TextExt1(i) Then
                    Txt1There = True
                    List7.AddItem Please
                    Exit For
            End If
        Next
        For i = 0 To 10   'videos
            If LCase(GetFileExtension(Please)) = VideoExt(i) Then
                VideoThere = True
                    List1.AddItem Please
                    Exit For
            End If
        Next
EmptyOne:
    Next XY
    
    If PicThere = True Then
        Undoit
        WebBrowser1.Visible = False
        RichTextBox1.Visible = False
        THBImage1.Visible = True
        Mht = False
        RTF = False
        File1.Pattern = "*.psd;*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga"
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        Label4.Caption = CList4.Text
        LoadFileAndUpdateDisplay CList4.List(0)
        Exit Sub
    End If
    If TxtThere = True Then
        WebBrowser1.Visible = True      'use browser
        RichTextBox1.Visible = False
        THBImage1.Visible = False
        File1.Pattern = "*.mht;*.htm;*.html;*.pdf"
        File1.Refresh
        Mht = True
        RTF = False
        For i = 0 To List6.ListCount - 1
                CList4.AddItem List6.List(i)
        Next
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & _
            " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & _
            Trim(ThePicture) & "'></img></p></body></html>"
            
        WebBrowser2.Top = THBImage1.Top
        WebBrowser2.Left = THBImage1.Left
        WebBrowser2.Height = THBImage1.Height
        WebBrowser2.Width = THBImage1.Width
        Exit Sub
    End If
    If Txt1There = True Then
        Haburabadooda22.Visible = True
        THBImage1.Visible = False
        WebBrowser1.Visible = False
        Mht = False
        RTF = True
        For i = 0 To List7.ListCount - 1
                CList4.AddItem List7.List(i)
        Next
        File1.Pattern = "*.txt;*.rtf"
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        RichTextBox1.Visible = True
        RichTextBox1.LoadFile CList4.List(0)
        Exit Sub
    End If
    If VideoThere = True Then
        Me.Hide
        Load VideoLibrary
        'VideoLibrary.Moviee.Clear
        For i = 0 To List1.ListCount - 1
                VideoLibrary.Moviee.AddItem List1.List(i)
        Next
        VideoLibrary.Moviee.Text = VideoLibrary.Moviee.List(0)
        'VideoLibrary.Show
        Exit Sub
     End If
     CList4.ListIndex = 0
      
End Sub
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, "", strURL, "", _
    "c:\", 10)
End Function

Private Sub Command13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu mnuResize
End Sub

Private Sub Command15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngNewWidth As Long, i As Long, i1 As Long
Dim lngNewHeight As Long
Dim Ration As Single
Dim varImageProperties As Variant
Dim SSTT As String
On Error Resume Next
If Mht = False And RTF = False Then
  THBImage1.Visible = True
  WebBrowser1.Visible = False
  RichTextBox1.Visible = False
    Kill App.path & "\Resized\*.*"
    For i = 0 To CList4.ListCount - 1
        DoEvents
        LoadFileAndUpdateDisplay CList4.List(i)
        varImageProperties = ie.ImageProperties
        If Not IsEmpty(varImageProperties) Then
            For i1 = 0 To UBound(varImageProperties) Step 2
                If InStr(varImageProperties(i1), "WIDTH") <> 0 Or InStr(varImageProperties(i1), "HEIGHT") <> 0 Then
                    Select Case varImageProperties(i1)
                    Case "WIDTH"
                         wID = varImageProperties(i1 + 1)
                         
                    Case "HEIGHT"
                        Hei = varImageProperties(i1 + 1)
                    End Select
                End If
            Next i1
        End If
        Ration = wID / Hei
        Ration = Int(640 / Ration)
        'MsgBox wID
        'MsgBox Hei
        'MsgBox wID / Hei
        'MsgBox Ration
        lngNewWidth = CLng(640)
        lngNewHeight = CLng(Ration)
        ie.Resize lngNewWidth, lngNewHeight, 1
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
        ThePicture = CList4.List(i)
        TheNumber = i
        ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.BMPUseRLE = True
        ie.SavePictureToFile App.path & "\Resized\" & i & ".bmp", thbifBMP
    Next
End If

End Sub

Private Sub Command16_Click()
Frame1.Visible = False
Picture4.Visible = True
Fx = 0
End Sub

Private Sub Command17_Click()
mnuRandomFilter_Click
End Sub
'Update Pictureinfo
Private Sub UpdatePicInfo()
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim strSizeInch As String
    Dim strSizeMM As String
       
    'Convert image width from Pixel to Twips
    lngWidth = ie.Width
    lngHeight = ie.Height
    ie.CnvPixelToTwipsImg lngWidth, lngHeight
    strSizeInch = Format(CDbl(lngWidth) / 1440, "0.00") & "x" & _
                Format(CDbl(lngHeight) / 1440, "0.00") & "inch"

    'Convert image width from Pixel to HiMetric
    lngWidth = ie.Width
    lngHeight = ie.Height
    ie.CnvPixelToHiMetricImg lngWidth, lngHeight
    strSizeMM = CDbl(lngWidth) / 100 & "x" & CDbl(lngHeight) / 100 & "mm"

    
End Sub
Private Sub UpdatePicInfo1()
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim strSizeInch As String
    Dim strSizeMM As String
       
    'Convert image width from Pixel to Twips
    lngWidth = ie.Width
    lngHeight = ie.Height
    ieLimit.CnvPixelToTwipsImg lngWidth, lngHeight
    strSizeInch = Format(CDbl(lngWidth) / 1440, "0.00") & "x" & _
                Format(CDbl(lngHeight) / 1440, "0.00") & "inch"

    'Convert image width from Pixel to HiMetric
    lngWidth = ie.Width
    lngHeight = ie.Height
    ieLimit.CnvPixelToHiMetricImg lngWidth, lngHeight
    strSizeMM = CDbl(lngWidth) / 100 & "x" & CDbl(lngHeight) / 100 & "mm"

    
End Sub

Public Sub LoadFileAndUpdateDisplay(strFile As String)
    LoadFile strFile, ie
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End Sub



    Private Sub Invert()
        Dim ieLimit As New THBImageEdit
        Dim lngLeft As Long
        Dim lngTop As Long
        Dim lngRight As Long
        Dim lngBottom As Long
        Dim pBytes As Long
        Dim lngSizeBytes As Long
        Dim rg As THBRegion
        
        On Error GoTo ErrHandler
        
        'We need one Region
        If THBImage1.RegionCount <> 1 Then
            MsgBox "No Fence defined!"
            Exit Sub
        End If
        
        'We need rectangle region
        Set rg = THBImage1.RegionGetByIndex(0)
        If rg.NumPoints <> 5 Then
            MsgBox "Invalid Fence!"
            Exit Sub
        End If
        
        'Define the part we are interested in
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        'Create a new image from the cropped part
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        
        'Do the operation
        'ieLimit.Invert
        'ieLimit.AutoLevels 0.5, 0.5
        'ieLimit.ConvertToBlackWhite
        'ieLimit.ClipboardCopyTo
        'ieLimit.ClipboardPasteFrom
        'ieLimit.ConvertToBPP thbbppBW, thbDitherFS, True 'NOTTT
        'ieLimit.Despeckle 100, 100
        'ieLimit.Grayscale
        'ieLimit.Rotate 45
        'ieLimit.OverlayWithTransparency ieLimit, lngLeft, lngTop, 100, False
        ieLimit.Clear
        'Copy the result back to the original image
                ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        Exit Sub
        
        
ErrHandler:
        MsgBox Err.Description
    End Sub


Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
CList4.Clear
End Sub

Private Sub Command10_Click()
MouseWheel1.WheelDisconnect
'On Error Resume Next
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to print to the default printer?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Choose NO to Select your printer!"   ' Define title.
Help = "DEMO.HLP"   ' Define Help file.
Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then   ' User chose Yes.
     ' Perform some action.
Else   ' User chose No.
     ' Perform some action.
     CmDlg.ShowPrinter
     If CmDlg.CancelError = False Then Exit Sub
End If

Load frmTest
'frmTest.Show
frmTest.Hide
frmTest.mnuScanSelectSource_Click
frmTest.mnuScanAcquire_Click
Set ie.Picture = frmTest.ImgXCtrl1.Picture
Set THBImage1.Picture = ie.THBStdPicture
THBImage1.ZoomFit
End Sub


Private Sub Command11_Click()
    
    On Error GoTo ErrHandler
    nDivisor = CLng(ASpinnerBox10.Value)
    
    'Emboss Filter
    ' -2 -1  0
    ' -1  1  1
    '  0  1  2
    ReDim arMatrix(0 To 8)
    arMatrix(0) = CLng(ASpinnerBox1.Value): arMatrix(1) = CLng(ASpinnerBox2.Value): arMatrix(2) = CLng(ASpinnerBox3.Value)
    arMatrix(3) = CLng(ASpinnerBox4.Value): arMatrix(4) = CLng(ASpinnerBox5.Value): arMatrix(5) = CLng(ASpinnerBox6.Value)
    arMatrix(6) = CLng(ASpinnerBox7.Value): arMatrix(7) = CLng(ASpinnerBox8.Value): arMatrix(8) = CLng(ASpinnerBox9.Value)
    
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    ie2.FilterUserDefined arMatrix, nDivisor, 3, 3, 100 ', False, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ReDim arMatrix(0 To 8) As Long
        arMatrix(0) = -2: arMatrix(1) = -1: arMatrix(2) = 0
        arMatrix(3) = -1: arMatrix(4) = 1: arMatrix(5) = 1
        arMatrix(6) = 0: arMatrix(7) = 1: arMatrix(8) = 2
        ieLimit.FilterUserDefined arMatrix, nDivisor, 3, 3, 100 ', False, True
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
End If
ErrHandler:
    'MsgBox Err.Description
If Fx <> 1000 Then Frame1.Visible = False: Picture4.Visible = True
End Sub

Public Sub Command12_Click()
MouseWheel1.WheelDisconnect
Me.Enabled = False
If Command12.Caption = "Color to replace" Then
    Command12.Caption = "Replace with color"
    Load frmColor
    frmColor.Show
    Exit Sub
End If
If Command12.Caption = "Replace with color" Then
    ReplaceFlag = True
    PickFlag = False
    Me.Enabled = False
    Command12.Caption = "Color to replace"
    'Slider1.Visible = True
    'Slider1.Min = 0
    'Slider1.Max = 1000
    'Slider1.Value = 1000
    'FxLabel.Caption = "Tolerance"
    'FxLabel.Visible = True
    'Command4.Visible = True
    Fx = 5
    Load frmColor
    frmColor.Show
    Command12.Caption = "Perform Task"
    Exit Sub
End If
If Command12.Caption = "Perform Task" Then
    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Do you want to Swap these colors?"   ' Define message.
    Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
    Title = "Color Replacement"   ' Define title.
    On Error Resume Next
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then
        Command12.Visible = False
        Command12.Caption = "Color to replace"
        Command4_Click
    Else
        Command12.Visible = False
        Command12.Caption = "Color to replace"
    End If
    Me.Enabled = True
End If
End Sub

Private Sub Command13_Click()
    'DMSlider3.Value = 2000
End Sub

Private Sub Command14_Click()
MouseWheel1.WheelDisconnect
On Error Resume Next
'clist4.setfocus
If RegionsFlag = False Then Undoit
CList4.Clear
List1.Clear
List2.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
ReturnValuePath = BrowseForFolder(Me.hwnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN)
File1.path = ReturnValuePath
Open App.path & "\LastPath" For Output As #1
    Print #1, ReturnValuePath
Close #1
LoadClist4
CList4.ListIndex = 0
End Sub


Private Sub Command19_Click()
Frame3.Visible = False
Picture4.Visible = True

End Sub

Private Sub Command2_Click()
Dim lngNewWidth As Long, i As Long, i1 As Long
Dim lngNewHeight As Long
Dim Ration As Single
Dim varImageProperties As Variant
Dim SSTT As String
On Error Resume Next
If Mht = False And RTF = False Then
  THBImage1.Visible = True
  WebBrowser1.Visible = False
  RichTextBox1.Visible = False
    Kill App.path & "\Resized\*.*"
    For i = 0 To CList4.ListCount - 1
        DoEvents
        LoadFileAndUpdateDisplay CList4.List(i)
        varImageProperties = ie.ImageProperties
        If Not IsEmpty(varImageProperties) Then
            For i1 = 0 To UBound(varImageProperties) Step 2
                If InStr(varImageProperties(i1), "WIDTH") <> 0 Or InStr(varImageProperties(i1), "HEIGHT") <> 0 Then
                    Select Case varImageProperties(i1)
                    Case "WIDTH"
                         wID = varImageProperties(i1 + 1)
                         
                    Case "HEIGHT"
                        Hei = varImageProperties(i1 + 1)
                    End Select
                End If
            Next i1
        End If
        Ration = wID / Hei
        Ration = Int(640 / Ration)
        'MsgBox wID
        'MsgBox Hei
        'MsgBox wID / Hei
        'MsgBox Ration
        lngNewWidth = CLng(640)
        lngNewHeight = CLng(Ration)
        ie.Resize lngNewWidth, lngNewHeight, 1
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
        ThePicture = CList4.List(i)
        TheNumber = i
        ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.JPGProgressive = False
        ie.JPGQuality = 60  'Compression quality from 1-100
        ie.JPGGrayscale = False 'Export to 8bit grayscale JPEG
        ie.SavePictureToFile App.path & "\Resized\" & i & ".jpg", thbifJPG
    Next
End If
End Sub

Private Sub Command23_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub


Private Sub Command20_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    On Error Resume Next
    lngNewWidth = CLng(Val(Spinner1))
    lngNewHeight = CLng(Val(Spinner2))
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
    Frame3.Visible = False
    Picture4.Visible = True
End Sub

Private Sub Command21_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    On Error Resume Next
    lngNewWidth = CLng(Val(Spinner1.Text))
    lngNewHeight = CLng(Val(Spinner2.Text))
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
    mnuCustom_Click
End Sub



Private Sub Command22_Click()
'clist4.setfocus
    THBImage1.Visible = False
    Mht = True
    RTF = False
    WebBrowser1.Visible = True
    RichTextBox1.Visible = False
    File1.Pattern = "*.mht;*.htm;*.html;*.pdf"
    StartDir
End Sub

Private Sub Command23_Click()
    Text12.Text = CList4.ListIndex
End Sub

Private Sub Command24_Click()
If Val(Text12.Text) > CList4.ListIndex Then
    MsgBox "You must select a higher frame to delete!"
    Exit Sub
End If
Text13.Text = CList4.ListIndex

End Sub

Private Sub Command25_Click()
If Text12.Text = "<<<Here>>>" Or Text13.Text = "to Here" Or Val(Text12.Text) - Val(Text13.Text) <= 0 Then
    CList4_KeyUp vbKeyDelete, 0
Exit Sub
End If

Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to Permanently DELETE These Files?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Permanently Delete Files Up to present"   ' Define title.
On Error Resume Next
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
    List2.Clear
    List3.Clear
    For i = 0 To Val(Text12.Text) - 1
            List3.AddItem CList4.List(i)
            List2.AddItem Format(FileDateTime(CList4.List(i)), "YYYYMMDDHHMMSS") & "*" & CList4.List(i)
    Next
    For i = Val(Text13.Text) + 1 To CList4.ListCount - 1
            List3.AddItem CList4.List(i)
            List2.AddItem Format(FileDateTime(CList4.List(i)), "YYYYMMDDHHMMSS") & "*" & CList4.List(i)
    Next
    For i = Val(Text12.Text) To Val(Text13.Text)
        Kill CList4.List(i)
    Next
    CList4.Clear
    For i = 0 To List3.ListCount - 1
            CList4.AddItem List3.List(i)
    Next
    CList4.Refresh
    CList4.ListIndex = Val(Text12.Text)
    CList4.Text = CList4.List(Val(Text12.Text))
    LoadFileAndUpdateDisplay CList4.Text
    ThePicture = CList4.Text
    Text12.Text = "<<<Here>>>"
    Text13.Text = "to Here"
End If

End Sub

Private Sub Command26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub


Private Sub Command27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command28_Click()
    If List6.ListCount = 0 Then Exit Sub
    THBImage1.Visible = False
    Mht = True
    RTF = False
    WebBrowser1.Visible = True
    RichTextBox1.Visible = False
    File1.Pattern = "*.mht;*.htm;*.html;*.pdf"
    File1.Refresh
    CList4.Clear
    For i = 0 To List6.ListCount - 1
        If Trim(List6.List(i)) <> "" Then
            CList4.AddItem List6.List(i)
        End If
    Next
    CList4.Text = CList4.List(0)
    CList4_Click
    
End Sub

Private Sub Command28_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub


Private Sub Command29_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command3_Click()
Slider2.Visible = False
FxLabel1.Visible = False
Command3.Visible = False
    Label8(7).Visible = False
End Sub

Private Sub Command30_Click()
MouseWheel1.WheelDisconnect
Dim lngNewWidth As Long, i As Long
Dim lngNewHeight As Long
Dim Ration As Single
On Error Resume Next
If Mht = False And RTF = False Then
  THBImage1.Visible = True
  WebBrowser1.Visible = False
  RichTextBox1.Visible = False
    Kill App.path & "\Slideshow\*.*"
    Open App.path & "\SlideShow\SlideShow.shw" For Output As #1
    For i = 0 To CList4.ListCount - 1
        DoEvents
        LoadFileAndUpdateDisplay CList4.List(i)
        
        If SlideSHowFlagget = True And Resizette <> 0 Then
            Ration = ie.Height / ie.Width
            Ration = Int(Ration * Resizette)
            lngNewWidth = CLng(Resizette)
            lngNewHeight = CLng(Ration)
            ie.Resize lngNewWidth, lngNewHeight, 1
            UpdatePicInfo
            Set THBImage1.Picture = ie.THBStdPicture
        End If
        
        ThePicture = CList4.List(i)
        TheNumber = i
        Print #1, i & ".jpg"
        ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.JPGProgressive = False
        ie.JPGQuality = 60  'Compression quality from 1-100
        ie.JPGGrayscale = False 'Export to 8bit grayscale JPEG
        ie.SavePictureToFile App.path & "\SlideShow\" & i & ".jpg", thbifJPG
    Next
    Close #1
    SlideSHowFlagget = False
    Me.WindowState = 1
    Shell App.path & "\Snotnose.exe " & App.path & "\SlideShow\SlideShow.shw", vbNormalFocus
End If
End Sub

Private Sub Command32_Click()
If List3.ListCount = 0 Then Exit Sub
Dim i As Long
Mht = False
RTF = False
WebBrowser1.Visible = False
RichTextBox1.Visible = False
THBImage1.Visible = True
File1.Pattern = "*.psd;*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga"
CList4.Clear
For i = 0 To List3.ListCount - 1
    If Trim(List3.List(i)) <> "" Then
        CList4.AddItem List3.List(i)
    End If
Next
CList4.Text = CList4.List(0)
CList4_Click
End Sub

Private Sub Command32_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command33_Click()
    Dim lngPixel As Long
    Set ieLimit.Picture = ie.Picture
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    ieLimit.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    lngPixel = ie.GetPixelRGB(1, 1)
    'MsgBox lngPixel
    'Clipboard.SetText lngPixel
    'MsgBox lngPixel
    mnuUndo_Click

    ie.OverlayWithTransparency ieLimit, 0, 0, lngPixel  'RGB(0, 0, 0)

Set THBImage1.Picture = ie.THBStdPicture

    Command31.Visible = True
    Command33.Visible = False

End Sub

Private Sub Command34_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command35_Click()
MouseWheel1.WheelDisconnect
    Screen.MousePointer = 11
    File1.Pattern = "*.mpg;*.mpeg;*.m2v;*.avi;*.asf;*.mov;*.wmv;*mp3"
    StartDir
End Sub

Private Sub Command31_Click()
On Error GoTo ErrHandler
MouseWheel1.WheelDisconnect
Image2.Picture = ie.Picture
Undoit
Load Regions
Regions.Show
Regions.PIC(0).Height = Image2.Height
Regions.PIC(0).Width = Image2.Height
Regions.PIC(1).Height = Image2.Height
Regions.PIC(1).Width = Image2.Height
Regions.PIC(0).Picture = ie.Picture
Regions.PIC(1).Picture = ie.Picture
RegionsFlag = True
    'ie.PositionPic = thbPosCC
    'ie.StretchPic = thbStretchBoth
    'ie.KeepAspect = True
    'ie.DrawPic Regions.PIC(0).hdc, thbguPixel, 0, 0, CLng(THBImage1.Width * 1.36), CLng(THBImage1.Height * 2.1)
    'ie.DrawPic Regions.PIC(1).hdc, thbguPixel, 0, 0, CLng(THBImage1.Width * 1.36), CLng(THBImage1.Height * 2.1)
    'ie.DrawPic Regions.PIC(0).hdc, thbguPixel, 0, 0, CLng(Image2.Width), CLng(Image2.Height)
    'ie.DrawPic Regions.PIC(1).hdc, thbguPixel, 0, 0, CLng(Image2.Width), CLng(Image2.Height)
    Regions.PIC(0).Refresh
    Regions.PIC(1).Refresh
    Regions.NewPicture
    Angst.Enabled = False
    Command31.Visible = False
    Command33.Visible = True
    Exit Sub

ErrHandler:
    MsgBox Err.Description
End Sub


Private Sub Command34_Click()
If List7.ListCount = 0 Then Exit Sub
THBImage1.Visible = False
Mht = False
RTF = True
WebBrowser1.Visible = False
RichTextBox1.Visible = True
File1.Pattern = "*.txt;*.rtf"
File1.Refresh
CList4.Clear
For i = 0 To List7.ListCount - 1
    If Trim(List7.List(i)) <> "" Then
        CList4.AddItem List7.List(i)
    End If
Next
CList4.Text = CList4.List(0)
CList4_Click

End Sub

Private Sub Command35_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'clist4.setfocus

End Sub

Private Sub Command36_Click()

End Sub

Private Sub Command37_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
    SlideSHowFlagget = True
    Resizette = 0
    PopupMenu mnuResizette
End Sub

Private Sub Command38_Click()

End Sub

Private Sub Command39_Click()
On Error Resume Next
If Trim(CList4.Text) = "" And Trim(Clipboard.GetText) <> "" Then Exit Sub
Dim i As Integer, k As Long, TempFile As String
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
'Text1.Text = Clipboard.GetText
CleanTxt1
ReNameMoveFile = ""
TempFile = Patherino & Text1.Text & Exterino
Msg = "Are You sure you want to Rename: " & CList4.Text & vbCrLf & " to: " & TempFile & " ?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Rename Media"   ' Define title.
Ctxt = 1000   ' Define topic
'Response = MsgBox(Msg, Style, Title, Help, Ctxt)
'If Response = vbYes Then   ' User chose Yes.
        k = CList4.ListIndex
        Name CList4.Text As TempFile
        CList4.RemoveItem k
        CList4.ListIndex = k
        CList4.Text = CList4.List(k)
        'CList4.Text = TempFile
        'Msg = "Do You want to Move: " & TempFile & " ?"   ' Define message.
        'Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
        'Title = "Move"   ' Define title.
        'Ctxt = 1000   ' Define topic
        'Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Check4.Value = 1 Then
            ReNameMoveFile = TempFile
        Else
            ReNameMoveFile = ""
        End If
'End If


End Sub

Private Sub Command4_Click()
Slider1.Visible = False
FxLabel.Visible = False
Command4.Visible = False
    Label8(7).Visible = False
If Fx = 5 Then
    Undoit
    Dim lngTolerance As Long
    Dim colOld As Long
    Dim colNew As Long
    colOld = CDec(PickColor)
    colNew = CDec(ReplaceColor)
    lngTolerance = CLng(1000)
    If Limit = False Then
        ie.ReplaceColor colOld, colNew, lngTolerance
        If Not ie.ThreadRunning Then Set THBImage1.Picture = ie.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.ReplaceColor colOld, colNew, lngTolerance
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If

End If
End Sub
Private Sub Command4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command40_Click()
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
  ' Define title.
Help = "DEMO.HLP"   ' Define Help file.
Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.


If Overlaid = False Then
    Msg = "Do you want to Overlay the present Image, " & vbCrLf & _
    CList4.Text & " on another?"   ' Define message.
    Title = "Overlay"
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    If Response = vbYes Then   ' User chose Yes.
        ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.BMPUseRLE = True
        ie.SavePictureToFile App.path & "\Overlayer.bmp", thbifBMP
        Undoit
        MsgBox "Select the Receiving Image and CLICK Again"
        ie.Grayscale8Bit
        ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.BMPUseRLE = True
        ie.SavePictureToFile App.path & "\Masker.bmp", thbifBMP
        
        
    Else   ' User chose No.
       Exit Sub
    End If
    Overlaid = True
    'Command40.Picture = LoadPicture(App.path & "\Overlay1.ico")
Else
    'Command40.Picture = LoadPicture(App.path & "\Overlay.ico")
    Overlaid = False
        ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.BMPUseRLE = True
        ie.SavePictureToFile App.path & "\Overlayed.bmp", thbifBMP
        Load frmOverlay
        frmOverlay.Show
        Me.Enabled = False
        'Set frmOverlay.THBImageResult.Picture = ie.THBStdPicture
End If
End Sub
Private Sub butConvertToBW_Click()
    On Error GoTo ErrHandler
    
    ie.ConvertToBlackWhite
    
    'Assign the loaded image to the Image Viewer
    'control THBImage1. This does NOT duplicate
    'the image data. We just assign a reference to
    'the THBImageEdit object.
    'If Not ie.ThreadRunning Then Set THBImage1.Picture = ie.THBStdPicture
    
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description
End Sub

Private Sub Command41_Click()
On Error Resume Next
If Trim(CList4.Text) = "" And Trim(Clipboard.GetText) <> "" Then Exit Sub
Dim i As Integer, k As Long, TempFile As String
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Text1.Text = Clipboard.GetText
CleanTxt1
TempFile = Patherino & Text1.Text & Exterino
ReNameMoveFile = ""
Msg = "Are You sure you want to Rename: " & CList4.Text & vbCrLf & " to: " & TempFile & " ?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Rename Media"   ' Define title.
Ctxt = 1000   ' Define topic
'Response = MsgBox(Msg, Style, Title, Help, Ctxt)
'If Response = vbYes Then   ' User chose Yes.
        k = CList4.ListIndex
        Name CList4.Text As TempFile
        CList4.RemoveItem k
        CList4.ListIndex = k
        CList4.Text = CList4.List(k)
        'Msg = "Do You want to Move: " & TempFile & " ?"   ' Define message.
        'Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
        'Title = "Move"   ' Define title.
        'Ctxt = 1000   ' Define topic
        'Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        If Check4.Value = 1 Then
            ReNameMoveFile = TempFile
        Else
            ReNameMoveFile = ""
        End If

'End If

End Sub

Private Sub Command5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command6_Click()
Dim lngNewWidth As Long
Dim lngNewHeight As Long
Dim Ration As Single
Dim RetVal
On Error Resume Next
''SetTopMostWindow Me.hWnd, False
'Ontop.Visible = False
'OnBottom.Visible = True
Iconic = True
ie.BMPUseRLE = True
ie.SavePictureToFile App.path & "\" & "Iconest.bmp", thbifBMP
ie.ConvertToBPP thbbpp8Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
Ration = ie.Height / ie.Width
Ration = Int(Ration * 25)
lngNewWidth = CLng(32)
lngNewHeight = CLng(32)   'Ration)
ie.Resize lngNewWidth, lngNewHeight, 1
ie.BMPUseRLE = True
ie.SavePictureToFile App.path & "\" & "Icon.bmp", thbifBMP
'RetVal = Shell(App.path & "\Icon.exe", 1)
'LoadFileAndUpdateDisplay App.path & "\Iconest.bmp"
Iconic = False
'clist4.setfocus
End Sub

Private Sub Command7_Click()
'clist4.setfocus
THBImage1.Visible = False
Mht = False
RTF = True
WebBrowser1.Visible = False
RichTextBox1.Visible = True
File1.Pattern = "*.txt;*.rtf"
StartDir
End Sub

Private Sub Command8_Click()
'clist4.setfocus
Mht = False
RTF = False
WebBrowser1.Visible = False
RichTextBox1.Visible = False
THBImage1.Visible = True
File1.Pattern = "*.psd;*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga"
StartDir
End Sub

Private Sub Command8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Command9_Click()
    Slider3.Visible = False
    FxLabel2.Visible = False
    Command9.Visible = False
    Label8(7).Visible = False
End Sub


Private Sub Command5_Click()
Command5.Visible = False
Slider4.Visible = False
End Sub


Private Sub DMSlider1_Change(NewValue As Double)
End Sub




Private Sub Crapper_Click()
mnuCrop_Click
End Sub

Private Sub DMSlider3_Change(NewValue As Double)
'If DMSlider3.Value = 0 Then DMSlider3.Value = 1
'If DMSlider3.Value < 1000 Then
'    'mnuReverse.Checked = True
'    Timer2.Interval = DMSlider3.Value * 2
'''Tips.Visible = True
'Else
'    'mnuReverse.Checked = False
'    Timer2.Interval = DMSlider3.Value * 2
'End If
''Tips.Caption = CLng(DMSlider3.Value)

'If DMSlider3.Value < 1000 Then
    'Timer2.Interval = DMSlider3.Value * 2
'Else
    'Timer2.Interval = DMSlider3.Value * 2
'End If

End Sub

Private Sub DMSlider4_Change(NewValue As Double)
    On Error Resume Next
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 65)
    
    lngNewWidth = NewValue  'CLng(DMSlider5.Value)
    lngNewHeight = NewValue * Ration   'CLng(DMSlider5.Value * Ration)  'CLng(DMSlider4.Value)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture

End Sub
Private Sub LoadClist4()
On Error Resume Next
Dim ii As Long, i As Long
Dim Filepath As String
CList4.Clear
List3.Clear
List2.Clear
List6.Clear
List7.Clear
List1.Clear
For ii = 0 To File1.ListCount - 1
    If Right(File1.path, 1) = "\" Then
        Filepath = File1.path & File1.List(ii)
    Else
        Filepath = File1.path & "\" & File1.List(ii)
    End If
    Filepath = Trim(Filepath)
    If Trim(File1.List(ii)) <> "" Then
        For i = 0 To 13
            Option2.Visible = False
            Option3.Visible = False
            If LCase(GetFileExtension(Filepath)) = PictureExt(i) Then
                    CList4.AddItem Filepath
                    List2.AddItem Format(FileDateTime(Filepath), "YYYYMMDDHHMMSS") & "*" & Filepath
                    List3.AddItem Filepath
                    Option2.Visible = True
                    Option3.Visible = True
                    Exit For
            End If
        Next
        For i = 0 To 10
            If LCase(GetFileExtension(Filepath)) = VideoExt(i) Then
                    List1.AddItem Filepath
                    Exit For
            End If
        Next
        For i = 0 To 3
            If LCase(GetFileExtension(Filepath)) = TextExt(i) Then
                    List6.AddItem Filepath
                    Exit For
            End If
        Next
        For i = 0 To 1
            If LCase(GetFileExtension(Filepath)) = TextExt1(i) Then
                    List7.AddItem Filepath
                    Exit For
            End If
        Next
    End If
Next
'MsgBox List1.ListCount
If Mht = True And RTF = False And List6.ListCount > 0 Then
    WebBrowser1.Visible = True
    RichTextBox1.Visible = False
    For i = 0 To List6.ListCount - 1
            CList4.AddItem List6.List(i)
    Next
    File1.Refresh
    CList4.Refresh
    CList4.Text = CList4.List(0)
    ThePicture = CList4.List(0)
    WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & Trim(ThePicture) & "'></img></p></body></html>"
    WebBrowser2.Top = THBImage1.Top
    WebBrowser2.Left = THBImage1.Left
    WebBrowser2.Height = THBImage1.Height
    WebBrowser2.Width = THBImage1.Width
End If
If Mht = False And RTF = True And List7.ListCount > 0 Then
    WebBrowser1.Visible = False
    THBImage1.Visible = False
    For i = 0 To List7.ListCount - 1
            CList4.AddItem List7.List(i)
    Next
    File1.Refresh
    CList4.Refresh
    CList4.Text = CList4.List(0)
    ThePicture = CList4.List(0)
    RichTextBox1.Visible = True
    RichTextBox1.LoadFile CList4.List(0)
End If
If Mht = False And RTF = False And List2.ListCount > 0 Then
        WebBrowser1.Visible = False
        RichTextBox1.Visible = False
        THBImage1.Visible = True
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        LoadFileAndUpdateDisplay CList4.List(0)
End If
THBImage1.Redraw
MooseSnobbler = False
If CList4.ListCount = 0 And List1.ListCount > 0 And Mht = False And RTF = False Then
    Me.Hide
    MooseSnobbler = True
    Load VideoLibrary
    VideoLibrary.Moviee.Clear
    For i = 0 To List1.ListCount - 1
            VideoLibrary.Moviee.AddItem List1.List(i)
            VideoLibrary.Moviee.Text = VideoLibrary.Moviee.List(0)
    Next
    VideoLibrary.Show
    VideoLibrary.Moviee.ListIndex = 0
    VideoLibrary.Moviee_Click
    Exit Sub
End If
If Trim(CList4.Text) = "" Then
    CList4.Text = "Drag & Drop on me now, if you dare"
    WebBrowser1.Navigate2 ("about:blank")
    Option2.Visible = False
    Option3.Visible = False
Else
    Option2.Visible = True
    Option3.Visible = True
End If
Screen.MousePointer = Default

End Sub
Private Sub LoadClist4a()
On Error Resume Next
Dim ii As Long, i As Long
Dim Filepath As String
CList4.Clear
List3.Clear
List2.Clear
List6.Clear
List7.Clear
List1.Clear
For ii = 0 To List8.ListCount - 1
    Filepath = Trim(List8.List(ii))
    'MsgBox Filepath
    If Trim(List8.List(ii)) <> "" Then
        For i = 0 To 13
            Option2.Visible = False
            Option3.Visible = False
            If LCase(GetFileExtension(Filepath)) = PictureExt(i) Then
                    CList4.AddItem Filepath
                    List2.AddItem Format(FileDateTime(Filepath), "YYYYMMDDHHMMSS") & "*" & Filepath
                    List3.AddItem Filepath
                    Option2.Visible = True
                    Option3.Visible = True
                    Exit For
            End If
        Next
        For i = 0 To 10
            If LCase(GetFileExtension(Filepath)) = VideoExt(i) Then
                    List1.AddItem Filepath
                    Exit For
            End If
        Next
        For i = 0 To 3
            If LCase(GetFileExtension(Filepath)) = TextExt(i) Then
                    List6.AddItem Filepath
                    Exit For
            End If
        Next
        For i = 0 To 1
            If LCase(GetFileExtension(Filepath)) = TextExt1(i) Then
                    List7.AddItem Filepath
                    Exit For
            End If
        Next
    End If
Next
'MsgBox List1.ListCount
If Mht = True And RTF = False And List6.ListCount > 0 Then
    WebBrowser1.Visible = True
    RichTextBox1.Visible = False
    For i = 0 To List6.ListCount - 1
            CList4.AddItem List6.List(i)
    Next
    List8.Refresh
    CList4.Refresh
    CList4.Text = CList4.List(0)
    ThePicture = CList4.List(0)
    WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & Trim(ThePicture) & "'></img></p></body></html>"
    WebBrowser2.Top = THBImage1.Top
    WebBrowser2.Left = THBImage1.Left
    WebBrowser2.Height = THBImage1.Height
    WebBrowser2.Width = THBImage1.Width
End If
If Mht = False And RTF = True And List7.ListCount > 0 Then
    WebBrowser1.Visible = False
    THBImage1.Visible = False
    For i = 0 To List7.ListCount - 1
            CList4.AddItem List7.List(i)
    Next
    List8.Refresh
    CList4.Refresh
    CList4.Text = CList4.List(0)
    ThePicture = CList4.List(0)
    RichTextBox1.Visible = True
    RichTextBox1.LoadFile CList4.List(0)
End If
If Mht = False And RTF = False And List2.ListCount > 0 Then
        WebBrowser1.Visible = False
        RichTextBox1.Visible = False
        THBImage1.Visible = True
        List8.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        LoadFileAndUpdateDisplay CList4.List(0)
End If
THBImage1.Redraw
MooseSnobbler = False
If CList4.ListCount = 0 And List1.ListCount > 0 And Mht = False And RTF = False Then
    Me.Hide
    MooseSnobbler = True
    Load VideoLibrary
    VideoLibrary.Moviee.Clear
    For i = 0 To List1.ListCount - 1
            VideoLibrary.Moviee.AddItem List1.List(i)
            VideoLibrary.Moviee.Text = VideoLibrary.Moviee.List(0)
    Next
    VideoLibrary.Show
    Exit Sub
End If
If Trim(CList4.Text) = "" Then
    CList4.Text = "Drag & Drop on me now, if you dare"
    WebBrowser1.Navigate2 ("about:blank")
    Option2.Visible = False
    Option3.Visible = False
Else
    Option2.Visible = True
    Option3.Visible = True
End If
End Sub

Private Sub DMSlider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then
    'DMSlider3.Value = 1000
'End If
'If Button = 1 Then DMSlider3.Visible = True

End Sub

Private Sub DMSlider5_Change(NewValue As Double)
    On Error Resume Next
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 65)
    
    lngNewWidth = NewValue / Ration
    lngNewHeight = NewValue 'CLng(DMSlider4.Value)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture

End Sub

Public Sub AssignImage(ie As THBImageEdit)
    Set THBImage1.Picture = ie.THBStdPicture
    THBImage1.ZoomFit
    UpdatePicInfo

End Sub

Private Sub FlatScrollBar1_Change()

End Sub

Private Sub Command9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Form_DblClick()
'clist4.setfocus

End Sub

Private Sub Form_GotFocus()
'clist4.setfocus
End Sub

Private Sub Form_Initialize()
Dim Filepath As String
Dim i As Integer
Dim t As Single
Dim rtn As Long
Dim Holder As String
Dim Holder1 As String
Dim Please As String
Dim aPlease As String
Dim Pleased As Boolean
On Error Resume Next
RichTextBox1.ZOrder 0
RTF = False
t = Timer
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
If Dir(App.path & "\Undo\*.BMP") <> "" Then
    Kill App.path & "\Undo\*.BMP"
End If
If Exists(App.path & "\LastPath") = False Then
    Filepath = App.path
    Open App.path & "\LastPath" For Output As #1
        Print #1, Filepath
    Close #1
Else
    Open App.path & "\LastPath" For Input As #1
        Line Input #1, Filepath
    Close #1
    File1.path = Filepath
End If
If Command$ <> "" Then
    CList4.Clear
    Option2.Value = True        'Sort by Name
    List1.Clear 'video
    List2.Clear 'pics
    List3.Clear 'pics
    List6.Clear 'mht, htm, pdf
    List7.Clear 'txt, rtf
    PicThere = False
    TxtThere = False
    Txt1There = False
    VideoThere = False
        aPlease = Replace(Command$, Chr(34), "")
        Holder = Left(Trim(aPlease), 1)
        Holder1 = " " & Holder & ":"
        Holder = vbCr & vbLf & Holder & ":"
        Please = Replace(aPlease, Holder1, Holder)
        Open App.path & "\Stuff" For Output As #1
            Print #1, Please
        Close #1
        Open App.path & "\Stuff" For Input As #1
spot:
                Line Input #1, Please
                If Trim(Please) = "" Then GoTo spot
                File1.path = GetFilePath(Please)
        Close #1
        Open App.path & "\LastPath" For Output As #1
            Print #1, File1.path
        Close #1
        
    Open App.path & "\Stuff" For Input As #1
    Do While Not EOF(1)
        Line Input #1, Please
        If Trim(Please) = "" Then GoTo EmptyOne     'Don't add empty strings
        For i = 0 To 13 'pics
            If LCase(GetFileExtension(Please)) = PictureExt(i) Then
                    PicThere = True
                    CList4.AddItem Please
                    List3.AddItem Please    'sort name
                    List2.AddItem Format(FileDateTime(Please), "YYYYMMDDHHMMSS") & "*" & Please 'sort date
                Exit For
            End If
        Next
        For i = 0 To 3  'htm pdf
            If LCase(GetFileExtension(Please)) = TextExt(i) Then
                    TxtThere = True
                    List6.AddItem Please
                    Exit For
            End If
        Next
        For i = 0 To 1  'txt rtf
            If LCase(GetFileExtension(Please)) = TextExt1(i) Then
                    Txt1There = True
                    List7.AddItem Please
                    Exit For
            End If
        Next
        For i = 0 To 10   'videos
            If LCase(GetFileExtension(Please)) = VideoExt(i) Then
                VideoThere = True
                    List1.AddItem Please
                    Exit For
            End If
        Next
EmptyOne:
    Loop
    Close #1
        
        
    If PicThere = True Then
        Undoit
        WebBrowser1.Visible = False
        RichTextBox1.Visible = False
        THBImage1.Visible = True
        Mht = False
        RTF = False
        File1.Pattern = "*.psd;*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga"
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        Label4.Caption = CList4.Text
        LoadFileAndUpdateDisplay CList4.List(0)
        CList4.ListIndex = 0
        Exit Sub
    End If
    If TxtThere = True Then
        WebBrowser1.Visible = True      'use browser
        RichTextBox1.Visible = False
        THBImage1.Visible = False
        File1.Pattern = "*.mht;*.htm;*.html;*.pdf"
        File1.Refresh
        Mht = True
        RTF = False
        For i = 0 To List6.ListCount - 1
           CList4.AddItem List6.List(i)
        Next
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & _
            " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & _
            Trim(ThePicture) & "'></img></p></body></html>"
            
        WebBrowser2.Top = THBImage1.Top
        WebBrowser2.Left = THBImage1.Left
        WebBrowser2.Height = THBImage1.Height
        WebBrowser2.Width = THBImage1.Width
        Exit Sub
    End If
    If Txt1There = True Then
        Haburabadooda22.Visible = True
        THBImage1.Visible = False
        WebBrowser1.Visible = False
        Mht = False
        RTF = True
        For i = 0 To List7.ListCount - 1
                CList4.AddItem List7.List(i)
        Next
        File1.Pattern = "*.txt;*.rtf"
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        RichTextBox1.Visible = True
        RichTextBox1.LoadFile CList4.List(0)
        Exit Sub
    End If
    If VideoThere = True Then
        Load VideoLibrary
        VideoLibrary.Moviee.Clear
        For i = 0 To List1.ListCount - 1
                VideoLibrary.Moviee.AddItem List1.List(i)
        Next
        VideoLibrary.Show
        Exit Sub
     End If
Else
    If Clipboard.GetData <> 0 Then
        On Error Resume Next
        butClipboardPasteFrom_Click
    Else
        Set ie.Picture = Picture1.Picture
        Set THBImage1.Picture = ie.THBStdPicture
    End If
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
THBImage1.Refresh
End Sub

Private Sub FxLabel_Change()
Label8(7).Visible = True

End Sub

Private Sub FxLabel1_Change()
Label8(7).Visible = True

End Sub

Private Sub FxLabel2_Change()
Label8(7).Visible = True

End Sub

Private Sub Haburabadooda19_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

Dim FileExtension As String
Dim Filenamme, ThbIT As String
On Error GoTo SaveAsError
    With Angst
        .CmDlg.InitialDir = File1.path
        .CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = False   'True 'Allow multi select
        .CmDlg.DialogTitle = "Save As" 'Set dialog title
        .CmDlg.DefaultFilename = Format(Now, "ddmmyyhhmmss")
        .CmDlg.Filter = "Gif Files (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "Jpeg Files (*.jpg)" & Chr$(0) & "*.jpg" & Chr$(0) & "Icon (*.ico)" & Chr$(0) & "*.ico" & Chr$(0) & "BMP Files (*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0) & "Meta Files (*.wmf)" & Chr$(0) & "*.wmf" & Chr$(0) & "PCX Files (*.pcx)" & Chr$(0) & "*.pcx" & Chr$(0) & "TIF Files (*.tif)" & Chr$(0) & "*.tif" _
        & Chr$(0) & "PNG Files (*.png)" & Chr$(0) & "*.png" & Chr$(0) & "TGA Files (*.tga)" & Chr$(0) & "*.tga" & Chr$(0) & "Photoshop Files (*.psd)" & Chr$(0) & "*.psd"
        
        '.CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowSave
    End With
    If Angst.CmDlg.cFileName(1) = "" Then Exit Sub
    Filenamme = Angst.CmDlg.cFileName(1)
'MsgBox GlobalFIlter
Select Case GlobalFIlter
Case 1
    FileExtension = ".gif"
Case 2
    FileExtension = ".jpg"
Case 3
    FileExtension = ".ico"
Case 4
    FileExtension = ".bmp"
Case 5
    FileExtension = ".wmf"
Case 6
    FileExtension = ".pcx"
Case 7
    FileExtension = ".tif"
Case 8
    FileExtension = ".png"
Case 9
    FileExtension = ".tga"
Case 10
    FileExtension = ".psd"
Case 11
    FileExtension = ".pdf"
End Select
If LCase(Right(Filenamme, 4)) <> FileExtension Then
    Filenamme = Filenamme & FileExtension
End If

Select Case FileExtension
Case ".gif"
        ie.ConvertToBPP thbbpp8Bit, thbDitherFS, True
        Dim arComments() As String
        Dim varComments As Variant
        ReDim arComments(0 To 3)
        arComments(0) = "Angst Draw"
        arComments(1) = "Warren S. Goff, D.O."
        arComments(2) = ""
        arComments(3) = ""
        varComments = arComments
        'Export LZW
        ie.GIFSettings thbGIFComp_LZW, varComments, -1, 100, 0, 0, 0, 1
        ie.SavePictureToFile Filenamme, thbifGIF
Case ".jpg"
        If ie.BitsPerPixel = thbbppBW Then
            ie.Grayscale8Bit
            ie.JPGGrayscale = True 'Export to 8bit grayscale JPEG
        Else
            ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
            ie.JPGGrayscale = False 'Do not export to 8bit grayscale JPEG
        End If
        ie.JPGProgressive = False
        ie.JPGQuality = 60 'Compression quality from 1-100
        ie.SavePictureToFile Filenamme, thbifJPG
Case ".ico"
    Option4_Click (5)
Case ".bmp"
        ie.BMPUseRLE = True
        ie.SavePictureToFile Filenamme, thbifBMP
Case ".wmf"
    ie.SavePictureToFile Filenamme, thbifWBMP
Case ".pcx"
    ie.SavePictureToFile Filenamme, thbifPCX
Case ".tif"
        If ie.BitsPerPixel = thbbppBW Then
            'Export TIFF with CCITT Group4 Compression
            ie.TIFFCompression = thbTIFFCompCCITTFAX4
            ie.TIFFFaxMode = thbTIFFFaxModeClassF
            ie.TIFFGroup4Options = thbTIFFGroup4None
        ElseIf ie.BitsPerPixel = thbbpp8Bit Then
            'Export TIFF with PACKBITS Compression
            ie.TIFFCompression = thbTIFFCompPACKBITS
        Else
            'Export TIFF with Lossy JPEG Compression
            'Not compatible with all viewers
            'ie.TIFFCompression = thbTIFFCompJPEG

            'Export TIFF with Lossy JPEG Compression
            'Not compatible with all viewers but works
            'for windows imaging viewer
            'ie.TIFFCompression = thbTIFFCompOldJPEG

            'Export TIFF with Lossless PACKBITS RLE Compression
            'Compatible with all viewers but compression ratio
            'is not as good as jpeg compression
            ie.TIFFCompression = thbTIFFCompPACKBITS
            ie.JPGQuality = 75
        End If
        ie.SavePictureToFile Filenamme, thbifTIF

Case ".png"
        If ie.BitsPerPixel = thbbppBW Then
            ie.PNGColType = thbPNGColTypePalette
        ElseIf ie.BitsPerPixel = thbbpp8Bit Then
            ie.PNGColType = thbPNGColTypePalette
        Else
            ie.PNGColType = thbPNGColTypeRGB
        End If
        ie.PNGInterlace = thbPNGInterlaceNone
        ie.SavePictureToFile Filenamme, thbifPNG
Case ".tga"
    ie.SavePictureToFile Filenamme, thbifTGA
Case ".psd"
    ie.SavePictureToFile Filenamme, thbifPSD
Case ".pdf"
        ie.PDFSettings "AngstArt", Format(Now), _
            "MooseNoseInc", "AngstImage", _
            "", "", "PDF, Picture"
        If ie.BitsPerPixel = thbbppBW Then
            'Export PDF with CCITT Group4 Compression
            ie.TIFFCompression = thbTIFFCompCCITTFAX4
            ie.TIFFFaxMode = thbTIFFFaxModeClassF
            ie.TIFFGroup4Options = thbTIFFGroup4None
        Else
            ie.TIFFCompression = thbTIFFCompPACKBITS
        End If
        ie.SavePictureToFile Filenamme, thbifPDF
End Select

SaveAsError:
'clist4.setfocus

End Sub

Private Sub Haburabadooda21_Click()

End Sub

Private Sub Haburabadooda22_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
RetVal = Shell(App.path & "\WordWebEdit.exe " & CList4.Text, 1)

End Sub

Private Sub Haburabadooda23_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
Dim RetVal
RetVal = Shell(App.path & "\WindowPicker.exe", 1)   ' Run Calculator
End Sub

Private Sub Haburabadooda24_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
    Shell App.path & "\AVItoGIF.exe", vbNormalFocus
End Sub

Private Sub Haburabadooda25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
Shell App.path & "\CaptureAVI.exe", vbNormalFocus

End Sub

Private Sub Haburabadooda26_Click()
On Error Resume Next
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to Overwrite" & CList4.Text & "?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Save file and overwrite the original!!"   ' Define title.
Help = "DEMO.HLP"   ' Define Help file.
Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then   ' User chose Yes.
     ' Perform some action.
Else   ' User chose No.
    Exit Sub
End If
'clist4.setfocus
Dim FileExx As String
FileExx = LCase(Right(CList4.Text, 3))
Select Case FileExx
Case "gif"  '
        ie.ConvertToBPP thbbpp8Bit, thbDitherFS, True
        Dim arComments() As String
        Dim varComments As Variant
        ReDim arComments(0 To 3)
        arComments(0) = "Angst Draw"
        arComments(1) = "Warren S. Goff, D.O."
        arComments(2) = ""
        arComments(3) = ""
        varComments = arComments
        'Export LZW
        targetgif = CList4.Text
        ie.GIFSettings thbGIFComp_LZW, varComments, -1, 100, 0, 0, 0, 1
        ie.SavePictureToFile targetgif, thbifGIF
    'clist4.setfocus
Case "bmp"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.BMPUseRLE = True
        ie.SavePictureToFile CList4.Text, thbifBMP
    'clist4.setfocus
Case "tif"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.TIFFCompression = thbTIFFCompPACKBITS
        ie.SavePictureToFile CList4.Text, thbifTIF
    'clist4.setfocus
Case "jpg"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    ie.JPGProgressive = False
    ie.JPGQuality = 60  'Compression quality from 1-100
    ie.JPGGrayscale = False 'Export to 8bit grayscale JPEG
    ie.SavePictureToFile CList4.Text, thbifJPG
    'clist4.setfocus
Case "psd"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    ie.SavePictureToFile CList4.Text, thbifPSD
    'clist4.setfocus
Case "pdf"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.PDFSettings "AngstArt", Format(Now), _
            "MooseNoseInc", "AngstImage", _
            "", "", "PDF, Picture"
        If ie.BitsPerPixel = thbbppBW Then
            'Export PDF with CCITT Group4 Compression
            ie.TIFFCompression = thbTIFFCompCCITTFAX4
            ie.TIFFFaxMode = thbTIFFFaxModeClassF
            ie.TIFFGroup4Options = thbTIFFGroup4None
        Else
            ie.TIFFCompression = thbTIFFCompPACKBITS
        End If
        ie.SavePictureToFile CList4.Text, thbifPDF
End Select

End Sub

Private Sub Haburabadooda26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Haburabadooda27_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'clist4.setfocus
File1.path = "c:\1down" 'App.path
File1.Refresh
THBImage1.Visible = True
RichTextBox1.Visible = False
WebBrowser1.Visible = False
Mht = False
RTF = False
Open App.path & "\LastPath" For Output As #1
    Print #1, "c:\1down"        'App.path
Close #1
LoadClist4
CList4.ListIndex = 0
End Sub

Private Sub Haburabadooda28_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
Shell App.path & "\CaptureAVI1.exe", vbNormalFocus
End Sub

Private Sub Haburabadooda29_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
        Haburabadooda30_MouseUp 0, 0, 0, 0
        Crop = False
        THBImage1.RegionStartRectangle
        Limit = True
        Haburabadooda30.Visible = True
        THBImage1.PopupMenu = False
End Sub

Private Sub Haburabadooda30_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
        Crop = False
        THBImage1.RegionClear
        THBImage1.Redraw
        Limit = False
        mnuResize.Visible = True
        mnuCrop.Visible = True
        mnuCropCapture.Visible = True
        mnuScan.Visible = True
        mnuRecycle.Visible = True
        RegionsFlag = False
        Haburabadooda30.Visible = False
End Sub

Private Sub Haburabadooda5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
Clipboard.Clear
On Error Resume Next
butClipboardCopyTo_Click
If RegionsFlag = False Then Undoit

End Sub

Private Sub Haburabadooda7_Click()

End Sub

Private Sub ieevents_ThreadReady(ByVal strError As String)
    If Not ie.ThreadRunning Then Set THBImage1.Picture = ie.THBStdPicture
End Sub
Private Sub Form_Activate()
Dim Filepath As String, VbPaint As String
On Error Resume Next
''clist4.setfocus
rtn = FindWindow("Shell_traywnd", "") 'get the Window
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
TickTock = 0
Set THBImage1.Picture = ie.THBStdPicture
If Captured = True Then
    Captured = False
    LoadFileAndUpdateDisplay App.path & "\Captured.bmp"
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
    Exit Sub
End If
THBImage1.ZoomFit
THBImage1.DropFiles = True
'clist4.setfocus
File1.Pattern = "*.psd;*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga"
Label14 = "Number " & CList4.ListIndex & " of " & CList4.ListCount & " total files"
'Me.Height = 8355
'Me.Width = 9285
End Sub
Private Sub Form_Load()
Dim Filepath As String
Dim i As Integer
Dim t As Single
Dim rtn As Long
Dim Holder As String
Dim Holder1 As String
Dim Please As String
Dim aPlease As String
Dim Pleased As Boolean
On Error Resume Next
MkDir "c:\1down"
File2.path = App.path & "\Images\"
Set clsWinApi = New CWinAPI 'This makes a new instance of the class
If Dir(App.path & "\FirstTime") = "" Then
   Open App.path & "\FirstTime" For Output As #1
   Close #1
End If
szzFile = "0"
Step = False
PictureExt(0) = "gif"
PictureExt(1) = "jpeg"
PictureExt(2) = "jpg"
PictureExt(3) = "ico"
PictureExt(4) = "cur"
PictureExt(5) = "wmf"
PictureExt(6) = "emf"
PictureExt(7) = "bmp"
PictureExt(8) = "pcx"
PictureExt(9) = "tif"
PictureExt(10) = "tiff"
PictureExt(11) = "png"
PictureExt(12) = "tga"
PictureExt(13) = "psd"

VideoExt(0) = "mpg"
VideoExt(1) = "mpeg"
VideoExt(2) = "m2v"
VideoExt(3) = "avi"
VideoExt(4) = "asf"
VideoExt(5) = "mov"
VideoExt(6) = "wmv"
VideoExt(7) = "mp3"
VideoExt(8) = "wma"
VideoExt(9) = "wav"
VideoExt(10) = "m1v"

TextExt(0) = "mht"
TextExt(1) = "html"
TextExt(2) = "htm"
TextExt(3) = "pdf"

TextExt1(0) = "txt"
TextExt1(1) = "rtf"



ScanFlag = False
MoveFlag = False
RegionsFlag = False
Limit = False
Crap = False
Mht = False
RTF = False
Effect = False
Pleased = False
Thumb = False
Captured = False
Iconic = False
ShowOn = False
Crop = False
LoadRegion = False
Painting = False
AsBut = False
PickFlag = False
ReplaceFlag = False
Schnorbel = False
SlideShowFlag = False
Scanned = False
MoveMe = False
SlideSHowFlagget = False
Overlaid = False
ReNameMoveFile = ""

Set ie1 = CreateTHBImageEdit()
Set ieevents = ie1
Set ie2 = CreateTHBImageEdit()
Set ieevents = ie2
Set ie = CreateTHBImageEdit()
Set ieevents = ie

THBImage1.PopupMenu = True
THBImage1.PositionPic = thbPosCC
THBImage1.StretchPic = thbStretchNone
THBImage1.KeepAspect = True
THBImage1.Scrolling = True
THBImage1.PreviewScrollWindow = True
THBImage1.Clipboard = True
THBImage1.DropFiles = True 'Enable Ole Drag&Drop
THBImage1.RegionDetectTolerance = 20
THBImage1.RegionPenColor = RGB(0, 0, 255)
THBImage1.RegionPenStyle = thbPS_SOLID
THBImage1.RegionPenWidth = 2
ie.DrawQualityHigh = True
ie.CacheOn = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MouseWheel1.WheelDisconnect
'MouseWheel1.WheelConnect Me.hWnd

    If AsBut = False Then
        If Trim(CList4.Text) = "" Then CList4.Text = "Drag & Drop on me now, if you dare": WebBrowser1.Navigate2 ("about:blank")
    End If
    Command7.ToolTipText = "Open Text Documents. Presently: " & File1.path
    Command8.ToolTipText = "Open Graphic Files. Presently: " & File1.path
    Command22.ToolTipText = "Open HTM Documents. Presently: " & File1.path
    Haburabadooda26.ToolTipText = "Save:" & CList4.Text
    Haburabadooda1.ToolTipText = "Thumbnails for Selected Directory:" & File1.path
    'Haburabadooda27.ToolTipText = "Open Application Directory: " & App.path
    'AnyShape26.ToolTipText = "Browse for Folder, Presently: " & File1.path
    If CList4.Text <> "Drag & Drop on me now, if you dare" Then
        Haburabadooda26.ToolTipText = "Save: " & CList4.Text
    Else
        Haburabadooda26.ToolTipText = "Save Actual File"
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next: 'clist4.setfocus
If Button = 2 Then
    If Mht = True Or RTF = True Then Exit Sub
    PopupMenu mnuFile
End If
On Error Resume Next: 'clist4.setfocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    'Dim hWnd1 As Long
    'hWnd1 = FindWindow("Shell_traywnd", "")
    'Call SetWindowPos(hWnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    UnloadAll
    Unload Angst
    Set Angst = Nothing
    Unload Regions
    Set Regions = Nothing
    CloseAll
End Sub

Sub CloseAll()
    On Error Resume Next
    Dim intFrmNum As Integer
    intFrmNum = Forms.Count


    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        Set Forms(intFrmNum - 1) = Nothing
        intFrmNum = intFrmNum - 1
    Loop
End Sub


Private Sub FxS_Change(NewValue As Double)

End Sub

Private Sub FxS_Click()

End Sub

Private Sub Haburabadooda1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
On Error Resume Next
If RegionsFlag = False Then Undoit
Dim Filepath As String
On Error Resume Next
Open App.path & "\LastPath" For Input As #1
    Line Input #1, Filepath
Close #1
Load frmThmb
frmThmb.Show
If Trim(Filepath) <> "" Then
    frmThmb.Dir1.path = Filepath
Else
    frmThmb.Dir1.path = App.path
End If
frmThmb.Dir1.Refresh

Exit Sub
    hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
    With Angst
        .CmDlg.InitialDir = App.path & "\SaveBin"
        '.CmDlg.CancelError = True 'Set cancel error to true
        .CmDlg.MultiSelect = False   'True 'Allow multi select
        .CmDlg.DialogTitle = "Select file (s) to open" 'Set dialog title
        '.CmDlg.Filter = "All Files (*.*)|*.*"

'FileDialog.sFilter = "All Graphic Files (*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga)|*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga|Word Files (*.doc)|*.doc|Html Files (*.htm;*.html)|*.htm;*.html|Batch Files (*.bat)|*.bat|INI Files (*.ini)|*.ini|All Files (*.*)|*.*|"
        .CmDlg.Filter = "All Graphic Files (*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga;*.psd)" & Chr$(0) & "*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga;*.psd" & Chr$(0) _
        & "Gif Files (*.gif)" & Chr$(0) & "*.gif" & Chr$(0) & "Jpeg Files (*.jpeg;*.jpg)" & Chr$(0) & "*.jpeg;*.jpg" & Chr$(0) & "Icon/Cursor (*.ico;*.cur)" & Chr$(0) & "*.ico;*.cur" & Chr$(0) & "BMP Files (*.bmp)" & Chr$(0) & "*.bmp" & Chr$(0) & "Meta Files (*.wmf;*.emf)" & Chr$(0) & "*.wmf;*.emf" & Chr$(0) & "PCX Files (*.pcx)" & Chr$(0) & "*.pcx" & Chr$(0) & "TIF Files (*.tif;*.tiff)" & Chr$(0) & "*.tif;*.tiff" _
        & Chr$(0) & "PNG Files (*.png)" & Chr$(0) & "*.png" & Chr$(0) & "TGA Files (*.tga)" & Chr$(0) & "*.tga" & Chr$(0) & "Photoshop Files (*.psd)" & Chr$(0) & "*.psd" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*"


        '"Psd Files (*.psd)|*.psd; *.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga;*.txt
        .CmDlg.FilterIndex = 1 'Set filter index
        .CmDlg.ShowOpen 'Show open dialog
        If hHook Then UnhookWindowsHookEx hHook
        ThePicture = CmDlg.cFileName(1)
        CList4.Clear
        CList4.AddItem CmDlg.cFileName(1)
        CList4.Text = CList4.List(0)
        LoadFileAndUpdateDisplay CmDlg.cFileName(1)
        CList4.ListIndex = 0
    End With

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = False
Image4.Visible = True

End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Visible = True
Image4.Visible = False
Shell App.path & "\VBCPUID.exe", vbNormalFocus
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Haburabadooda10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
On Error Resume Next
If RegionsFlag = False Then Undoit
CList4.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
SendKeys "{ESC}", True
''SetTopMostWindow Me.hWnd, False
'Ontop.Visible = False
'OnBottom.Visible = True
Me.WindowState = 1

Load frmCapture
frmCapture.Show

End Sub


Private Sub Haburabadooda12_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
Me.WindowState = 1
End Sub

Private Sub Haburabadooda13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If RegionsFlag = False Then Undoit
'LoadFileAndUpdateDisplay App.path & "\" & "Munc.BMP"
'ThePicture = App.path & "\" & "Munc.BMP"
'Set THBImage1.Picture = ie.THBStdPicture
Set ie.Picture = Picture1.Picture
Set THBImage1.Picture = ie.THBStdPicture
'clist4.setfocus

End Sub

Private Sub Haburabadooda14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to Delete All Files in the List?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Delete All Files in the List"   ' Define title.
Ctxt = 1000   ' Define topic
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Mht = False And RTF = False Then
    If RegionsFlag = False Then Undoit
End If
If Response = vbYes Then   ' User chose Yes.
'DoEvents
    Dim ActionFlag As Long
    Dim i As Integer
    On Error Resume Next
    If CList4.Text <> "Drag & Drop on me now, if you dare" And CList4.Text <> "" Then
        If Dir(CList4.Text) <> "" Then
            Screen.MousePointer = 11
            ActionFlag = FOF_ALLOWUNDO
            For i = 0 To CList4.ListCount - 1
                If Exists(CList4.List(i)) = True Then
                    ShellDeleteOne CList4.List(i), ActionFlag
                End If
            Next
            File1.Refresh
            Screen.MousePointer = 0
        End If
    End If
End If
CList4.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
LoadFileAndUpdateDisplay ""
'Command8_Click
End Sub

Private Sub Haburabadooda15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
If Thumb = False Then
    Thumb = True
    Load frmThumbs
    frmThumbs.Show
End If

End Sub

Private Sub Haburabadooda16_MouseEnter()
'Tips.Visible = True
'Tips.Caption = "Video Library"

End Sub

Private Sub Haburabadooda16_MouseExit()
'Tips.Visible = False

End Sub

Public Sub Haburabadooda16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
'If List1.ListCount = 0 Then Exit Sub
MouseWheel1.WheelDisconnect
Me.Hide
Load VideoLibrary
'VideoLibrary.Show
If List1.ListCount <> 0 Then
    VideoLibrary.Moviee.Clear
    
    For i = 0 To List1.ListCount - 1
            VideoLibrary.Moviee.AddItem List1.List(i)
            VideoLibrary.Moviee.Text = VideoLibrary.Moviee.List(0)
    Next
End If


End Sub

Private Sub Haburabadooda17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Haburabadooda18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
On Error Resume Next
Dim Filepath As String
Dim ii As Long, NewMic As String, i As Long
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to DELETE These Files?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Permanently Delete Files Up to Present"   ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then
Screen.MousePointer = 11
List5.Clear
    For ii = 0 To CList4.ListCount - 1
        If ii < CList4.ListIndex Then
            Kill CList4.List(ii)
        Else
            List5.AddItem CList4.List(ii)
        End If
    Next ii
    CList4.Clear
    For ii = 0 To List5.ListCount - 1
            CList4.AddItem List5.List(ii)
    Next
    CList4.Text = CList4.List(0)
    ThePicture = CList4.List(0)
    If Mht = False And RTF = False Then
        RichTextBox1.Visible = False
        WebBrowser1.Visible = False
        THBImage1.Visible = True
        LoadFileAndUpdateDisplay CList4.List(0)
    End If
    If Mht = True And RTF = False Then
        RichTextBox1.Visible = False
        WebBrowser1.Visible = True
        THBImage1.Visible = False
        WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & Trim(ThePicture) & "'></img></p></body></html>"
        WebBrowser2.Top = THBImage1.Top
        WebBrowser2.Left = THBImage1.Left
        WebBrowser2.Height = THBImage1.Height
        WebBrowser2.Width = THBImage1.Width
    End If
    If Mht = False And RTF = True Then
        RichTextBox1.Visible = True
        WebBrowser1.Visible = False
        THBImage1.Visible = False
        RichTextBox1.LoadFile CList4.List(0)
    End If
    'clist4.setfocus
End If
Screen.MousePointer = 0
End Sub

Private Sub Haburabadooda2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
If WebBrowser2.Visible = True Then
    WebBrowser2.Visible = False
    THBImage1.Visible = True
Else
    WebBrowser2.Visible = True
    WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & Trim(CList4.Text) & "'></img></p></body></html>"
    'AniGIF1.ReadGIF Trim(CList4.Text)
    WebBrowser2.Top = THBImage1.Top
    WebBrowser2.Left = THBImage1.Left
    WebBrowser2.Height = THBImage1.Height
    WebBrowser2.Width = THBImage1.Width
    THBImage1.Visible = False
    WebBrowser2.Visible = True
End If
'clist4.setfocus
End Sub

Public Sub Haburabadooda3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Mht = True Or RTF = True Then Exit Sub
    Dim lngPrintWidth  As Long
    Dim lngPrintHeight As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    CmDlg.ShowPrinter
    Printer.Print Space(1)
    If Limit = False Then
        'Calculate destination rect
        'lngPrintWidth = 1.65 * Printer.Width
        'lngPrintHeight = 1.65 * Printer.Height
        'lngLeft = 300   'lngPrintWidth / 4
        'lngTop = 500   'lngPrintHeight / 4
        'lngRight = lngLeft + lngPrintWidth / 2
        'lngBottom = lngTop + lngPrintHeight / 2
        
        lngPrintWidth = Printer.Width
        lngPrintHeight = Printer.Height
        lngLeft = 0
        lngTop = 0
        lngRight = lngPrintWidth
        lngBottom = lngPrintHeight
        
        ie.PrintPicAligned Printer.hDC, lngLeft, lngTop, lngRight, lngBottom, thbguTwips, True, thbPosCC, thbStretchBoth
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        'Calculate destination rect
        lngPrintWidth = Printer.Width
        lngPrintHeight = Printer.Height
        lngLeft = 0
        lngTop = 0
        lngRight = lngPrintWidth
        lngBottom = lngPrintHeight
        ieLimit.PrintPicAligned Printer.hDC, lngLeft, lngTop, lngRight, lngBottom, thbguTwips, True, thbPosCC, thbStretchBoth
    End If
    Printer.EndDoc
End Sub
Public Sub PrintNot()

If Mht = True Or RTF = True Then Exit Sub
    Dim lngPrintWidth  As Long
    Dim lngPrintHeight As Long
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    'CmDlg.ShowPrinter
    Printer.Print Space(1)
    If Limit = False Then
        'Calculate destination rect
        'lngPrintWidth = 1.65 * Printer.Width
        'lngPrintHeight = 1.65 * Printer.Height
        'lngLeft = 300   'lngPrintWidth / 4
        'lngTop = 500   'lngPrintHeight / 4
        'lngRight = lngLeft + lngPrintWidth / 2
        'lngBottom = lngTop + lngPrintHeight / 2
        
        lngPrintWidth = Printer.Width
        lngPrintHeight = Printer.Height
        lngLeft = 0
        lngTop = 0
        lngRight = lngPrintWidth
        lngBottom = lngPrintHeight
        
        ie.PrintPicAligned Printer.hDC, lngLeft, lngTop, lngRight, lngBottom, thbguTwips, True, thbPosCC, thbStretchBoth
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        'Calculate destination rect
        lngPrintWidth = 1.65 * Printer.Width
        lngPrintHeight = 1.65 * Printer.Height
        lngLeft = 300 'lngPrintWidth / 4
        lngTop = 500  'lngPrintHeight / 4
        lngRight = lngLeft + lngPrintWidth / 2
        lngBottom = lngTop + lngPrintHeight / 2
        ieLimit.PrintPicAligned Printer.hDC, lngLeft, lngTop, lngRight, lngBottom, thbguTwips, True, thbPosCC, thbStretchBoth
    End If
    Printer.EndDoc
End Sub

Private Sub Haburabadooda4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
On Error Resume Next
butClipboardPasteFrom_Click
If RegionsFlag = False Then Undoit

End Sub

Private Sub Haburabadooda6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Mht = True Or RTF = True Then Exit Sub
On Error Resume Next
If RegionsFlag = False Then Undoit
''SetTopMostWindow Me.hWnd, False
'Ontop.Visible = False
'OnBottom.Visible = True
CList4.Clear
CList4.Text = "Drag & Drop on me now, if you dare"
WebBrowser1.Navigate2 ("about:blank")
'Dim RetVal
'RetVal = Shell(App.path & "\Snap-It.exe", 1)

End Sub

Private Sub Haburabadooda8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MoveMe = True Then
        MoveMe = False
    Else
        MoveMe = True
    End If
End Sub

Private Sub Haburabadooda9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseWheel1.WheelDisconnect
'clist4.setfocus
    Shell App.path & "\AVICreator6.exe", vbNormalFocus
    Call Images1
End Sub


Private Sub Label14_Change()
Label14.Caption = Replace(Label14.Caption, "-1", "0")
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Label30_Click()

End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Mht = True Or RTF = True Then Exit Sub
If Button = 2 Then PopupMenu mnuFile
End Sub


Private Sub List2_Click()
Dim MM As String
Dim i, J, k As Long

CList4.Clear
For i = 0 To List2.ListCount - 1
    J = InStr(List2.List(i), "*")
    k = Len(List2.List(i)) - J
        CList4.AddItem Right(List2.List(i), k)
Next

End Sub

Private Sub mn1024_Click()
Resizette = 1024
Command30_Click
End Sub

Private Sub mn640_Click()
Resizette = 640
Command30_Click
End Sub

Private Sub mn800_Click()
Resizette = 800
Command30_Click

End Sub

Private Sub mnu1024_Click()

If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    On Error Resume Next
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 1024)
    lngNewWidth = CLng(1024)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End Sub

Private Sub mnu180_Click()
    Dim dAngle As Double
    On Error Resume Next
    dAngle = CDbl(180)
If RegionsFlag = False Then Undoit
If Limit = False Then
    ie.Rotate dAngle
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.Rotate dAngle
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub

Private Sub mnu16_Click()
If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    On Error Resume Next
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 16)
    lngNewWidth = CLng(16)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture
End Sub

Private Sub mnu32_Click()
If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    On Error Resume Next
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 32)
    lngNewWidth = CLng(32)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture

End Sub

Private Sub mnu320_Click()
If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    On Error Resume Next
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 320)
    lngNewWidth = CLng(320)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture
End Sub

Private Sub mnu640_Click()
    If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long, i As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    On Error Resume Next
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 640)
    lngNewWidth = CLng(640)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
    'ie.SavePictureToFile "c:\1down\test\" & i & ".jpg", thbifJPG
    If Check2.Value = 1 Then
       For i = 0 To CList4.ListCount - 1
            LoadFile CList4.List(i), ie
            Ration = ie.Height / ie.Width
            Ration = Int(Ration * 640)
            lngNewWidth = CLng(640)
            lngNewHeight = CLng(Ration)
            ie.Resize lngNewWidth, lngNewHeight, 1
            ie.SavePictureToFile "c:\1down\test\" & i & ".jpg", thbifJPG
        Next i
    End If
End Sub

Private Sub mnu65_Click()
If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    On Error Resume Next
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 65)
    lngNewWidth = CLng(65)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture
End Sub

Private Sub mnu800_Click()
If RegionsFlag = False Then Undoit
    Dim lngNewWidth As Long
    Dim lngNewHeight As Long
    Dim Ration As Single
    On Error Resume Next
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 800)
    lngNewWidth = CLng(800)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture
End Sub

Private Sub mnu90CCW_Click()
    Dim dAngle As Double
    On Error Resume Next
    dAngle = CDbl(270)
If RegionsFlag = False Then Undoit
If Limit = False Then
    ie.Rotate dAngle
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.Rotate dAngle
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub

Private Sub mnu90CW_Click()
    Dim dAngle As Double
    On Error Resume Next
    dAngle = CDbl(90)
If RegionsFlag = False Then Undoit
If Limit = False Then
    ie.Rotate dAngle
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.Rotate dAngle
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If

End Sub

Private Sub mnuantialias_Click()
If RegionsFlag = False Then Undoit
    If Limit = False Then
        Set ie2.Picture = ie.Picture
        Set THBImage3.Picture = ie2.THBStdPicture
        THBImage3.ZoomFit
        ie2.FilterAntialias
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.FilterAntialias
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
End Sub

Private Sub mnuAssociate_Click()
Dim PathToApp As String
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "This will associate jpg, gif, jpeg, wmf," _
& vbCrLf & "bmp, pcx, tif, tiff, png and tgs with this program!" _
& vbCrLf & "Your Icon Cache will also be Rebuilt!" _
& vbCrLf & "This may reset the positions of your Desktop Icons!" _
& vbCrLf & "OK?"   ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2   ' Define buttons.
Title = "Associate and Rebuild Icons"   ' Define title.
Help = "DEMO.HLP"   ' Define Help file.
Ctxt = 1000   ' Define topic
      ' context.
      ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbNo Then   ' User chose Yes.
    Exit Sub
End If

PathToApp = App.path
MakeFileAssociation "jpg", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "gif", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "jpeg", PathToApp, "Angst.exe", "Graphic Solution"
'MakeFileAssociation "ico", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "wmf", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "bmp", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "pcx", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "tif", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "tiff", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "png", PathToApp, "Angst.exe", "Graphic Solution"
MakeFileAssociation "tga", PathToApp, "Angst.exe", "Graphic Solution"
'MakeFileAssociation "cur", PathToApp, "Angst.exe", "Graphic Solution"
'MakeFileAssociation "psd", PathToApp, "Angst.exe", "Graphic Solution"
Shell App.path & "\IconX.exe", vbHide
End Sub

Private Sub mnuAtal_Click()
Dim RetVal
RetVal = Shell(App.path & "\ImgX.exe", 1)
'RetVal = Shell(App.path & "\Vbpaint.exe " & CList4.Text, 1)

End Sub

Private Sub mnuBlackWhite_Click()
If RegionsFlag = False Then Undoit

    If Limit = False Then
        Set ie2.Picture = ie.Picture
        Set THBImage3.Picture = ie2.THBStdPicture
        THBImage3.ZoomFit
        ie2.ConvertToBPP thbbppBW, thbDitherFS, True
        ie2.PaletteSetToWhiteBlack
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ie.ConvertToBPP thbbppBW, thbDitherFS, True
        ie.PaletteSetToWhiteBlack
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
    mnuInvert_Click
Exit Sub
    
    
    If Limit = False Then
        Set ie2.Picture = ie.Picture
        Set THBImage3.Picture = ie2.THBStdPicture
        THBImage3.ZoomFit
        ie2.ConvertToBlackWhite
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.ConvertToBlackWhite
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If

End Sub

Private Sub mnuBlurru_Click()
    Slider1.Visible = True
    Slider1.Min = 0
    Slider1.Max = 100
    Slider1.Value = 0
    FxLabel.Caption = "Blur"
    FxLabel.Visible = True
    Command4.Visible = True
    Fx = 1

End Sub

Private Sub mnuBmp_Click()
On Error Resume Next
ie.BMPUseRLE = True
ie.SavePictureToFile App.path & "\" & Format(Now, "ddmmyyhhmmss") & ".bmp", thbifBMP
End Sub

Private Sub mnuBW_Click()
If RegionsFlag = False Then Undoit
If Limit = False Then
    On Error Resume Next
    ie.ConvertToBlackWhite
    Set THBImage1.Picture = ie.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.ConvertToBlackWhite
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
End Sub

Private Sub mnuChangeDown_Click()
On Error Resume Next
 'Timer2.Interval = 3000
End Sub

Private Sub mnuBright_Click()
    Slider1.Visible = True
    Slider1.Min = -100
    Slider1.Max = 100
    Slider1.Value = 0
    FxLabel.Caption = "Brightness"
    FxLabel.Visible = True
    Slider2.Visible = True
    Slider2.Min = -100
    Slider2.Max = 100
    Slider2.Value = 0
    FxLabel1.Caption = "Contrast"
    FxLabel1.Visible = True
    Command4.Visible = True
    Command3.Visible = True
    Fx = 2

End Sub



Private Sub mnuCrop_Click()
If RegionsFlag = False Then Undoit
Crop = True
    On Error Resume Next
    gnMode = MODE_PULLQUOTE
    THBImage1.RegionStartRectangle
End Sub

Private Sub mnuCropCapture_Click()
If RegionsFlag = False Then Undoit
Crop = True
    On Error Resume Next
    gnMode = MODE_PULLQUOTE
    THBImage1.RegionStartRectangle

End Sub

Private Sub mnuCustom_Click()
Frame3.Visible = True
Picture4.Visible = False
On Error Resume Next
    Dim varImageProperties As Variant
    Dim i As Long: Dim SSTT As String
    List4.Clear
    List4.AddItem "  "
    varImageProperties = ie.ImageProperties
    If Not IsEmpty(varImageProperties) Then
        For i = 0 To UBound(varImageProperties) Step 2
            If InStr(varImageProperties(i), "WIDTH") <> 0 Or InStr(varImageProperties(i), "HEIGHT") <> 0 Then
                List4.AddItem "   " & varImageProperties(i) + ": " & varImageProperties(i + 1)
                Select Case varImageProperties(i)
                Case "WIDTH"
                     Spinner1.Text = varImageProperties(i + 1)
                     AspectW = Val(varImageProperties(i + 1))
                     VScroll1.Value = -1 * AspectW
                Case "HEIGHT"
                    Spinner2.Text = varImageProperties(i + 1)
                    AspectH = Val(varImageProperties(i + 1))
                    VScroll2.Value = -1 * AspectH
                End Select
            End If
        Next i
    End If

Open App.path & "\Aspect" For Input As #1
Line Input #1, SSTT
Check3.Value = Val(SSTT)
Close #1
End Sub

Private Sub mnuDeleteRB_Click()
Dim ActionFlag As Long
Dim i As Integer
On Error Resume Next
If CList4.Text <> "Drag & Drop on me now, if you dare" And CList4.Text <> "" Then
ActionFlag = FOF_ALLOWUNDO
ShellDeleteOne CList4.Text, ActionFlag
            For i = 0 To CList4.ListCount - 1
                If Dir(CList4.List(i)) = "" Then CList4.RemoveItem i
            Next
            File1.Refresh
            CList4.Refresh
            CList4.Text = CList4.List(CList4.ListIndex)
                LoadFileAndUpdateDisplay CList4.List(CList4.ListIndex + 1)
            'clist4.setfocus
End If

End Sub

Private Sub mnuDetect1_Click()
If RegionsFlag = False Then Undoit
If Limit = False Then
   Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    ie2.FilterEdgeDetection 50
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterEdgeDetection 50
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If

End Sub

Private Sub mnuDetect2_Click()
    ASpinnerBox1.Value = -1
    ASpinnerBox2.Value = 0
    ASpinnerBox3.Value = 1
    ASpinnerBox4.Value = -1
    ASpinnerBox5.Value = 0
    ASpinnerBox6.Value = 1
    ASpinnerBox7.Value = -1
    ASpinnerBox8.Value = 0
    ASpinnerBox9.Value = 1
If RegionsFlag = False Then Undoit
Command11_Click

End Sub

Private Sub mnuDetect3_Click()
    ASpinnerBox1.Value = 2
    ASpinnerBox2.Value = 2
    ASpinnerBox3.Value = 3
    ASpinnerBox4.Value = 2
    ASpinnerBox5.Value = -16
    ASpinnerBox6.Value = 2
    ASpinnerBox7.Value = 2
    ASpinnerBox8.Value = 2
    ASpinnerBox9.Value = 2
    ASpinnerBox10.Value = 1
If RegionsFlag = False Then Undoit
Command11_Click
End Sub

Private Sub mnuDetect4_Click()
    ASpinnerBox1.Value = -1
    ASpinnerBox2.Value = -1
    ASpinnerBox3.Value = 0
    ASpinnerBox4.Value = -1
    ASpinnerBox5.Value = 1
    ASpinnerBox6.Value = -1
    ASpinnerBox7.Value = -1
    ASpinnerBox8.Value = -1
    ASpinnerBox9.Value = -1
If RegionsFlag = False Then Undoit
Command11_Click

End Sub
Private Sub mnuDetect5_Click()
'Emboss Filter
' -2 -1  0
' -1  1  1
'  0  1  2
ReDim arMatrix(0 To 24) As Long
On Error Resume Next
arMatrix(0) = -1: arMatrix(1) = -1: arMatrix(2) = -1: arMatrix(3) = 1: arMatrix(4) = 1
arMatrix(5) = -1: arMatrix(6) = -1: arMatrix(7) = -1: arMatrix(8) = -1: arMatrix(9) = 1
arMatrix(10) = 2: arMatrix(11) = 0: arMatrix(12) = -1: arMatrix(13) = -1: arMatrix(14) = -1
arMatrix(15) = 4: arMatrix(16) = 2: arMatrix(17) = -1: arMatrix(18) = -1: arMatrix(19) = -1
arMatrix(20) = 1: arMatrix(21) = 1: arMatrix(22) = 1: arMatrix(23) = -1: arMatrix(24) = -1
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    ie2.FilterUserDefined arMatrix, 5, 5, 5, 100 ', False, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterUserDefined arMatrix, 5, 5, 5, 100 ', False, True
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub
Private Sub mnuDetect6_Click()
'Emboss Filter
' -2 -1  0
' -1  1  1
'  0  1  2
ReDim arMatrix(0 To 24) As Long
On Error Resume Next
arMatrix(0) = -1: arMatrix(1) = -1: arMatrix(2) = -1: arMatrix(3) = -1: arMatrix(4) = -1
arMatrix(5) = -1: arMatrix(6) = -1: arMatrix(7) = -1: arMatrix(8) = -1: arMatrix(9) = -1
arMatrix(10) = -1: arMatrix(11) = -1: arMatrix(12) = 24: arMatrix(13) = -1: arMatrix(14) = -1
arMatrix(15) = -1: arMatrix(16) = -1: arMatrix(17) = -1: arMatrix(18) = -1: arMatrix(19) = -1
arMatrix(20) = -1: arMatrix(21) = -1: arMatrix(22) = -1: arMatrix(23) = -1: arMatrix(24) = -1
If RegionsFlag = False Then Undoit
    
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    ie2.FilterUserDefined arMatrix, 5, 5, 5, 100 ', False, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterUserDefined arMatrix, 5, 5, 5, 100 ', False, True
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub
Private Sub mnuDropshadow_Click()
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
    ie2.DropShadow 6, 6, 25
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.DropShadow 6, 6, 25
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
Exit Sub
    Slider1.Min = 0
    Slider1.Max = 100
    Slider1.Value = 0
    Slider2.Min = 0
    Slider2.Max = 10
    Slider2.Value = 0
    Slider3.Min = 0
    Slider3.Max = 10
    Slider3.Value = 0
    FxLabel.Caption = "%"
    FxLabel1.Caption = "Y"
    FxLabel2.Caption = "X"
    Slider1.Visible = True
    Slider2.Visible = True
    Slider3.Visible = True
    FxLabel.Visible = True
    FxLabel1.Visible = True
    FxLabel2.Visible = True
    Command3.Visible = True
    Command4.Visible = True
    Command9.Visible = True
    Fx = 4

End Sub

Private Sub mnuEmboss1_Click()
    ReDim arMatrix(0 To 8) As Long
    arMatrix(0) = -2: arMatrix(1) = -1: arMatrix(2) = 0
    arMatrix(3) = -1: arMatrix(4) = 1: arMatrix(5) = 1
    arMatrix(6) = 0: arMatrix(7) = 1: arMatrix(8) = 2
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    'Emboss Filter
    ' -2 -1  0
    ' -1  1  1
    '  0  1  2
    ie2.FilterUserDefined arMatrix, 1, 3, 3, 100 ', False, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterUserDefined arMatrix, 1, 3, 3, 100 ', False, True
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If

End Sub

Private Sub mnuEmboss2_Click()
    ASpinnerBox1.Value = 1
    ASpinnerBox2.Value = 0
    ASpinnerBox3.Value = 0
    ASpinnerBox4.Value = 0
    ASpinnerBox5.Value = 0
    ASpinnerBox6.Value = 0
    ASpinnerBox7.Value = 0
    ASpinnerBox8.Value = 0
    ASpinnerBox9.Value = -1
    ASpinnerBox10.Value = 1
If RegionsFlag = False Then Undoit
If Limit = False Then
    Command11_Click
    ie.BrightnessAndContrast 30, 20
    Set THBImage1.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    Command11_Click
    ieLimit.BrightnessAndContrast 30, 20
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub

Private Sub mnuEmboss3_Click()
Dim i As Integer
    ReDim arMatrix(0 To 8) As Long
    arMatrix(0) = -1: arMatrix(1) = -1: arMatrix(2) = -1
    arMatrix(3) = 0: arMatrix(4) = 1: arMatrix(5) = 0
    arMatrix(6) = 1: arMatrix(7) = 1: arMatrix(8) = 1
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    ie2.FilterUserDefined arMatrix, 1, 3, 3, 100  ', False, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterUserDefined arMatrix, 1, 3, 3, 100
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub
Private Sub mnuEmboss4_Click()
Dim i As Integer
    ReDim arMatrix(0 To 8) As Long
    arMatrix(0) = -1: arMatrix(1) = -1: arMatrix(2) = 0
    arMatrix(3) = -1: arMatrix(4) = 1: arMatrix(5) = 1
    arMatrix(6) = 0: arMatrix(7) = 1: arMatrix(8) = 1
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    ie2.FilterUserDefined arMatrix, 1, 3, 3, 100  ', False, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterUserDefined arMatrix, 1, 3, 3, 100  ', False, True
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub
Private Sub mnuEmboss5_Click()
Dim i As Integer
    ReDim arMatrix(0 To 8) As Long
    arMatrix(0) = -1: arMatrix(1) = 0: arMatrix(2) = 1
    arMatrix(3) = -1: arMatrix(4) = 1: arMatrix(5) = 1
    arMatrix(6) = -1: arMatrix(7) = 0: arMatrix(8) = 1
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    ie2.FilterUserDefined arMatrix, 1, 3, 3, 100  ', False, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterUserDefined arMatrix, 1, 3, 3, 100  ', False, True
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub

Private Sub mnuEmptyRB_Click()
Dim RetVal As Long
RetVal = SHEmptyRecycleBin(Angst.hwnd, "", SHERB_NOPROGRESSUI)

End Sub

Private Sub mnuExplore_Click()
clsWinApi.ShowRecycleBin 0
End Sub

Private Sub mnuGauss_Click()
    ASpinnerBox1.Value = 1
    ASpinnerBox2.Value = 2
    ASpinnerBox3.Value = 1
    ASpinnerBox4.Value = 2
    ASpinnerBox5.Value = 4
    ASpinnerBox6.Value = 2
    ASpinnerBox7.Value = 1
    ASpinnerBox8.Value = 2
    ASpinnerBox9.Value = 1
    ASpinnerBox10.Value = 16
If RegionsFlag = False Then Undoit
Command11_Click

End Sub

Private Sub mnuGif_Click()
'If ThePicture = "" Then Exit Sub
Dim errorstring As String
Dim sourcebmp, targetgif As String
On Error Resume Next
ie.ConvertToBPP thbbpp8Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
ie.ConvertToBPP thbbpp24Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
ie.BMPUseRLE = True
ie.SavePictureToFile App.path & "\" & "gif.bmp", thbifBMP
sourcebmp = App.path & "\gif.bmp"
targetgif = App.path & "\" & Format(Now, "ddmmyyhhmmss") & ".gif"
'BMP2GIF1.BMP2GIF sourcebmp, targetgif, False
'If BMP2GIF1.IsError Then      ' in case of error on converting
    'Debug.Print BMP2GIF1.errorstring
'End If
If RegionsFlag = False Then Undoit

End Sub

Private Sub mnuGetBMP_Click()
Shell App.path & "\BmpExt.exe", vbNormalFocus
End Sub

Private Sub mnuGray_Click()
If RegionsFlag = False Then Undoit
If Limit = False Then
    On Error Resume Next
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
    If ie2.BitsPerPixel <> thbbppBW Or ie2.BitsPerPixel <> thbbpp8Bit Then
        ie2.Grayscale
        'ie2.ConvertToBPP thbbpp8Bit, thbDitherFS, True
        'ie2.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    Else
        ie2.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    End If
    
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie.THBStdPicture
    
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.Grayscale
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub

Private Sub mnuHoriz_Click()
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set THBImage1.Picture = ie.THBStdPicture
    On Error Resume Next
    ie.MirrorHorizontal
    UpdatePicInfo
    
    Set THBImage1.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.MirrorHorizontal
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If

End Sub

Private Sub mnuHSV_Click()
    Slider1.Min = -100
    Slider1.Max = 100
    Slider1.Value = 0
    Slider2.Min = -100
    Slider2.Max = 100
    Slider2.Value = 0
    Slider3.Min = -180
    Slider3.Max = 180
    Slider3.Value = 0
    FxLabel.Caption = "V"
    FxLabel1.Caption = "S"
    FxLabel2.Caption = "H"
    Slider1.Visible = True
    Slider2.Visible = True
    Slider3.Visible = True
    FxLabel.Visible = True
    FxLabel1.Visible = True
    FxLabel2.Visible = True
    Command3.Visible = True
    Command4.Visible = True
    Command9.Visible = True
    Fx = 3

End Sub

Private Sub mnuIcon_Click()
If RegionsFlag = False Then Undoit
    Iconic = True
    ie.BMPUseRLE = True
    ie.SavePictureToFile App.path & "\SaveBin\" & "Iconest.bmp", thbifBMP
    ie.ConvertToBPP thbbpp8Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 25)
    lngNewWidth = CLng(32)
    lngNewHeight = CLng(32)   'Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    ie.BMPUseRLE = True
    ie.SavePictureToFile App.path & "\" & "Icon.bmp", thbifBMP
    RetVal = Shell(App.path & "\Icon.exe", 1)
    Unload Me
End Sub

Private Sub mnuInvert_Click()
If RegionsFlag = False Then
    Undoit
End If
If Limit = False Then
    On Error Resume Next
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
    If ie2.BitsPerPixel = thbbppBW Or ie2.BitsPerPixel = thbbpp8Bit Then
        ie2.InvertPalette
    Else
        ie2.Invert
    End If
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie.THBStdPicture
    
        
    
Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.Invert
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub

Private Sub mnuJpg_Click()
'If ThePicture = "" Then Exit Sub
On Error Resume Next
ie.JPGProgressive = False
ie.JPGQuality = 60
ie.JPGGrayscale = False
ie.BMPUseRLE = True
ie.SavePictureToFile App.path & "\" & Format(Now, "ddmmyyhhmmss") & ".jpg", thbifJPG
End Sub

Private Sub mnuMeanRemoval_Click()
    ASpinnerBox1.Value = 0
    ASpinnerBox2.Value = -2
    ASpinnerBox3.Value = 0
    ASpinnerBox4.Value = -2
    ASpinnerBox5.Value = 11
    ASpinnerBox6.Value = -2
    ASpinnerBox7.Value = 0
    ASpinnerBox8.Value = -2
    ASpinnerBox9.Value = 0
    ASpinnerBox10.Value = 3
If RegionsFlag = False Then Undoit
Command11_Click

End Sub

Private Sub mnuMediumBlur_Click()
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
    ie2.FilterBlurMedian
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.FilterBlurMedian
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub


Private Sub mnuMzoom_Click()
Command5.Visible = True
Slider4.Visible = True

End Sub

Private Sub mnuPickColor_Click()
PickFlag = True
ReplaceFlag = False
Command12.Visible = True
End Sub

Private Sub mnuPickReplaceColor_Click()
ReplaceFlag = True
PickFlag = False
Command12_Click

End Sub

Private Sub mnuPsd_Click()
'If ThePicture = "" Then Exit Sub
On Error Resume Next
ie.SavePictureToFile App.path & "\" & Format(Now, "ddmmyyhhmmss") & ".psd", thbifPSD
End Sub

Private Sub mnuReplaceNow_Click()
    Slider1.Visible = True
    Slider1.Min = 0
    Slider1.Max = 1000
    Slider1.Value = 0
    FxLabel.Caption = "Tolerance"
    FxLabel.Visible = True
    Command4.Visible = True
    Fx = 5
    
End Sub

Private Sub mnuRandomFilter_Click()
Dim MyValue
Randomize   ' Initialize random-number generator.

'MyValue = Int((1 * Rnd) + 0)   ' Generate random value between 1 and 6.

    ASpinnerBox1.Value = Int((1 * Rnd) + 0)
    ASpinnerBox1.Value = Int((1 * Rnd) + 0)
    ASpinnerBox3.Value = Int((1 * Rnd) + 0)
    ASpinnerBox4.Value = Int((1 * Rnd) + 0)
    ASpinnerBox5.Value = Int((5 * Rnd) - 5)
    ASpinnerBox6.Value = Int((1 * Rnd) + 0)
    ASpinnerBox7.Value = Int((1 * Rnd) + 0)
    ASpinnerBox8.Value = Int((1 * Rnd) + 0)
    ASpinnerBox9.Value = Int((1 * Rnd) + 0)
    ASpinnerBox10.Value = 1
    Fx = 1000
If RegionsFlag = False Then Undoit
Frame1.Visible = True: Picture4.Visible = False

    
End Sub

Private Sub mnurCustom_Click()
Dim message, Title, Default, MyValue
Dim metop As Long, meleft As Long
message = "Enter a value between 360 and -360"   ' Set prompt.
Title = "Custom Rotation"   ' Set title.
Default = "45"   ' Set default.
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
MyValue = InputBox(message, Title, Default, metop, meleft)
Dim dAngle As Double
On Error Resume Next
dAngle = CDbl(Val(MyValue))
If RegionsFlag = False Then Undoit

If Limit = False Then
    ie.Rotate dAngle
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.Rotate dAngle
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If

End Sub

Private Sub mnuReset_Click()
j1 = 0
'ShowOn = False
'Timer2.Enabled = False
 'mnuReverse.Checked = False
 'Timer2.Interval = 1000
 'Timer2.Enabled = False
End Sub

Private Sub mnuReverse_Click()
If mnuReverse.Checked = False Then
    mnuReverse.Checked = True
Else
    mnuReverse.Checked = False
End If
End Sub

Private Sub mnuRGB_Click()
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
    ie2.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub

Private Sub mnuSaveAs_Click()
Call Haburabadooda19_MouseUp(0, 0, 0, 0)
End Sub

Private Sub mnuScan_Click()
On Error Resume Next
Load frmTest
frmTest.Show
frmTest.mnuScanSelectSource_Click
frmTest.mnuScanAcquire_Click

Exit Sub
'If RegionsFlag = False Then Undoit
''SetTopMostWindow Me.hWnd, False
'Screen.MousePointer = 11
'TWAIN_SelectImageSource (Me.hwnd)
Dim PictureFileName As String
Dim ReturnValue As Long
PictureFileName = App.path + "\Temp.bmp"
'ReturnValue = TWAIN_AcquireToFilename(Me.hwnd, PictureFileName)
    
If ReturnValue = 0 Then
        LoadFileAndUpdateDisplay PictureFileName
        Kill PictureFileName
Else
        'GoTo errHandler
End If

Screen.MousePointer = 0
'''SetTopMostWindow Me.hWnd, True
End Sub

Private Sub mnuSharpenMatrix1_Click()
    ASpinnerBox1.Value = 1
    ASpinnerBox2.Value = 1
    ASpinnerBox3.Value = 0
    ASpinnerBox4.Value = 1
    ASpinnerBox5.Value = -8
    ASpinnerBox6.Value = 1
    ASpinnerBox7.Value = 1
    ASpinnerBox8.Value = 1
    ASpinnerBox9.Value = 1
If RegionsFlag = False Then Undoit
Command11_Click
End Sub

Private Sub mnuSharpenScr1_Click()
Slider1.Min = 0
Slider1.Max = 100
Slider1.Value = 0
Slider1.Visible = True
FxLabel.Caption = "Sharpen"
FxLabel.Visible = True
Command4.Visible = True
Fx = 0

End Sub

Private Sub mnuSharpenScr2_Click()
    ASpinnerBox1.Value = -1 '* Slider1.Value
    ASpinnerBox2.Value = -1 '* Slider1.Value
    ASpinnerBox3.Value = -1 '* Slider1.Value
    ASpinnerBox4.Value = -1 '* Slider1.Value
    ASpinnerBox5.Value = 9 '* Slider1.Value + 1
    ASpinnerBox6.Value = -1 '* Slider1.Value
    ASpinnerBox7.Value = -1 '* Slider1.Value
    ASpinnerBox8.Value = -1 '* Slider1.Value
    ASpinnerBox9.Value = -1 '* Slider1.Value
If RegionsFlag = False Then Undoit
Command11_Click


Exit Sub
Slider1.Min = 0
Slider1.Max = 4
Slider1.Value = 0
Slider1.Visible = True
FxLabel.Caption = "Mean Removal"
FxLabel.Visible = True
Command4.Visible = True
Fx = 6

End Sub
Private Sub mnuSharpenScr3_Click()
    ASpinnerBox1.Value = 0
    ASpinnerBox2.Value = -2
    ASpinnerBox3.Value = 0
    ASpinnerBox4.Value = -2
    ASpinnerBox5.Value = 11
    ASpinnerBox6.Value = -2
    ASpinnerBox7.Value = 0
    ASpinnerBox8.Value = -2
    ASpinnerBox9.Value = 0
    ASpinnerBox10.Value = 3
If RegionsFlag = False Then Undoit
Command11_Click

End Sub

Private Sub mnuStretch_Click()
On Error Resume Next
Me.Hide
If RegionsFlag = False Then Undoit
Stretchly.Picture1.Picture = THBImage1.Picture
Load Stretchly
Stretchly.Show

End Sub

Private Sub mnuUndo_Click()
If Mht = True Or RTF = True Then Exit Sub
Set ie2.Picture = ie.Picture
Set ie.Picture = ie1.Picture
Set THBImage2.Picture = ie2.THBStdPicture
Set THBImage1.Picture = ie.THBStdPicture
THBImage1.ZoomFit
THBImage2.ZoomFit
UpdatePicInfo
End Sub

Private Sub mnuUserDefined_Click()
If RegionsFlag = False Then Undoit
Frame1.Visible = True: Picture4.Visible = False

End Sub

Private Sub mnuVertical_Click()
    On Error Resume Next
If RegionsFlag = False Then Undoit
If Limit = False Then
    ie.MirrorVertical
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
Else
    Set rg = THBImage1.RegionGetByIndex(0)
    rg.GetPoint 0, lngLeft, lngTop
    rg.GetPoint 2, lngRight, lngBottom
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ieLimit.MirrorVertical
    ie.Overlay ieLimit, lngLeft, lngTop, 100, False
    UpdatePicInfo
    Set THBImage1.Picture = ie.THBStdPicture
End If
End Sub


Private Sub No_Click()
'Yes.Visible = False
'No.Visible = False
'Tips.Visible = False
End Sub


Private Sub Option2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Option3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

End Sub

Private Sub Option4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Option4(Index).Value = False

End Sub

Private Sub Picture2_Change()
If RegionsFlag = False Then Undoit
Set ie2.Picture = Picture2.Picture
Set THBImage3.Picture = ie2.THBStdPicture
On Error Resume Next: 'clist4.setfocus

End Sub

Public Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture4.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub


Private Sub Picture8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub Picture9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub
Private Sub Option1_Click(Index As Integer)

Select Case Index

Case 0
    Patherino = "c:\1down\1\Gif\"
Case 1
    Patherino = "c:\1down\1\Legal Security Government\"
Case 2
    Patherino = "c:\1down\1\Economic Financial\"
Case 3
    Patherino = "c:\1down\"
Case 4
    Patherino = "c:\1down\1\Microphones\"
Case 5
    Patherino = "c:\1down\1\Programming\"
Case 6
    Patherino = "c:\1down\1\Guitar\"
Case 7
    Patherino = "c:\1down\1\Medical\"
Case 8
    Patherino = "c:\1down\1\Utilities Cracks\"
Case 9
    Patherino = "c:\1down\1\Personals\"
Case 10
    Patherino = "c:\1down\1\Multimedia & Graphic\"
Case 11
    Patherino = "c:\1down\1\Telecom Internet FTP Networking\"
Case 12
    Patherino = "c:\1down\1\Science Education\"
Case 13
    Patherino = "c:\1down\1\Wordpro Spreadsheet Datapro\"
Case 14
    Patherino = "c:\1down\1\Multimedia & Graphic\"
Case 15
    Patherino = "c:\1down\My Pictures\"
Case 16
    Patherino = "c:\1down\1\Language Chinese, Russian etc\"
Case 17
    Patherino = "c:\1down\1\MP3\"
Case 18
    Patherino = "c:\1down\1\Travel & Scuba\"
Case 19
    Patherino = "c:\1down\1\Literature\"
    
End Select
Yes_Click

'Yes.Visible = True
'No.Visible = True
'Tips.Caption = "Move to " & Patherino
'Tips.Visible = True



End Sub

Private Sub Option1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Option1(Index).Value = False
'clist4.setfocus

End Sub

Private Sub Option2_Click()
Dim i As Long
CList4.Clear
For i = 0 To List3.ListCount - 1
        CList4.AddItem List3.List(i)
Next
CList4.Refresh
CList4.Text = CList4.List(0)
LoadFileAndUpdateDisplay CList4.Text
End Sub

Private Sub Option3_Click()
Dim i, J, k As Long
CList4.Clear
For i = List2.ListCount - 1 To 0 Step -1
    J = InStr(List2.List(i), "*")
    k = Len(List2.List(i)) - J
        CList4.AddItem Right(List2.List(i), k)
Next
CList4.Text = CList4.List(0)
CList4.Refresh
CList4_Click
LoadFileAndUpdateDisplay CList4.Text
End Sub

Private Sub Option4_Click(Index As Integer)
Dim RetVal As Long
If Mht = True Or RTF = True Then Exit Sub
Dim DesktopPath As String
DesktopPath = GetShellFolderPath(&H0)
FormatNow = DesktopPath & "\" & Format(Now, "ddmmyyhhmmss")
MouseWheel1.WheelDisconnect

Select Case Option4(Index).Caption

Case "Gif"  '
        ie.ConvertToBPP thbbpp8Bit, thbDitherFS, True
        Dim arComments() As String
        Dim varComments As Variant
        ReDim arComments(0 To 3)
        arComments(0) = "Angst Draw"
        arComments(1) = "Warren S. Goff, D.O."
        arComments(2) = ""
        arComments(3) = ""
        varComments = arComments
        targetgif = FormatNow & ".gif"
        ie.GIFSettings thbGIFComp_LZW, varComments, -1, 100, 0, 0, 0, 1
        ie.SavePictureToFile targetgif, thbifGIF
    'clist4.setfocus
Case "Bmp"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.BMPUseRLE = True
        ie.SavePictureToFile FormatNow & ".bmp", thbifBMP
    'clist4.setfocus
Case "Tif"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.TIFFCompression = thbTIFFCompPACKBITS
        OcrTiff = FormatNow & ".tif"
        ie.SavePictureToFile OcrTiff, thbifTIF
    'clist4.setfocus
    Option4(Index).Value = False
Case "Jpg"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
    ie.JPGProgressive = False
    ie.JPGQuality = 60  'Compression quality from 1-100
    ie.JPGGrayscale = False 'Export to 8bit grayscale JPEG
        ie.SavePictureToFile FormatNow & ".jpg", thbifJPG
    'clist4.setfocus
Case "Psd"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.SavePictureToFile FormatNow & ".psd", thbifPSD
    'clist4.setfocus
Case "Icon"
    Iconic = True
    Me.WindowState = 1
    ie.BMPUseRLE = True
    ie.SavePictureToFile App.path & "\" & "Iconest.bmp", thbifBMP
    ie.ConvertToBPP thbbpp8Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 25)
    lngNewWidth = CLng(32)
    lngNewHeight = CLng(32)   'Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    ie.BMPUseRLE = True
    ie.SavePictureToFile App.path & "\" & "Icon.bmp", thbifBMP
    RetVal = Shell(App.path & "\Icon.exe", 1)
    
    CList4.AddItem App.path & "\Icon.bmp"
    'Unload Me
    Haburabadooda13_MouseUp 0, 0, 0, 0
    Exit Sub
Case "Pdf"
    ie.ConvertToBPP thbbpp24Bit, thbDitherFS, True
        ie.PDFSettings "AngstArt", Format(Now), _
            "MooseNoseInc", "AngstImage", _
            "", "", "PDF, Picture"
        If ie.BitsPerPixel = thbbppBW Then
            ie.TIFFCompression = thbTIFFCompCCITTFAX4
            ie.TIFFFaxMode = thbTIFFFaxModeClassF
            ie.TIFFGroup4Options = thbTIFFGroup4None
        Else
            ie.TIFFCompression = thbTIFFCompPACKBITS
        End If
        ie.SavePictureToFile FormatNow & ".pdf", thbifPDF
    'clist4.setfocus
Case "All"
    On Error Resume Next
    ie.SavePictureToFile App.path & "\SaveBin\" & ".bmp", thbifBMP
    ie.JPGProgressive = False
    ie.JPGQuality = 60
    ie.JPGGrayscale = False
    ie.BMPUseRLE = True
    'ie.SavePictureToFile App.path & "\SaveBin\"  & ".jpg", thbifJPG
    'ie.SavePictureToFile App.path & "\SaveBin\"  & ".psd", thbifPSD
    ie.ConvertToBPP thbbpp8Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
    ie.ConvertToBPP thbbpp24Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
    ie.BMPUseRLE = True
    ie.SavePictureToFile App.path & "\" & "gif.bmp", thbifBMP
    sourcebmp = App.path & "\gif.bmp"
    targetgif = App.path & "\SaveBin\" & ".gif"
    
    Ration = ie.Height / ie.Width
    Ration = Int(Ration * 25)
    lngNewWidth = CLng(25)
    lngNewHeight = CLng(Ration)
    ie.Resize lngNewWidth, lngNewHeight, 1
    ie.ConvertToBPP thbbpp8Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
    ie.ConvertToBPP thbbpp24Bit, thbDitherNone, True 'thbbpp8Bit, thbDitherFS, True
    ie.BMPUseRLE = True
    ie.SavePictureToFile App.path & "\bmp2ico.bmp", thbifBMP
    RetVal = Shell(App.path & "\Icon.exe", 1)
    'clist4.setfocus
End Select
    CList4.AddItem FormatNow & "." & Option4(Index).Caption
End Sub

Private Sub Option4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Option4(Index).Value = False
'clist4.setfocus

End Sub

Private Sub Option5_Click()

End Sub



Private Sub Slide_Show_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Mht = True Or RTF = True Then Exit Sub
If ShowOn = False Then
    ShowOn = True
     'Timer2.Enabled = True
Else
    ShowOn = False
     'Timer2.Enabled = False
    'Slide_Show.Value = False
    'Option1.Value = False
End If
End Sub

Private Sub Picture3_Change()
If RegionsFlag = False Then Undoit
Set ie.Picture = Picture3.Picture
Set THBImage1.Picture = ie.THBStdPicture
On Error Resume Next: 'clist4.setfocus

End Sub

Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
dPercent = CDbl(Slider1.Value)
dPercent1 = CDbl(Slider2.Value)
dPercent2 = CDbl(Slider3.Value)
nHue = CDbl(Slider3.Value)
nSat = CDbl(Slider2.Value)
nValue = CDbl(Slider1.Value)
nDeltaX = CLng(Slider3.Value)
nDeltaY = CLng(Slider2.Value)
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
End If

Select Case Fx
Case 0
Case 1
Case 2
    If Limit = False Then
        ie2.BrightnessAndContrast dPercent, dPercent1
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.BrightnessAndContrast dPercent, dPercent1
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
Case 3
    If Limit = False Then
        ie2.HSVAdjustment nHue, nSat, nValue
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.HSVAdjustment nHue, nSat, nValue
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
Case 4
    If Limit = False Then
        ie.DropShadow nDeltaX, nDeltaY, dPercent
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.DropShadow nDeltaX, nDeltaY, dPercent
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
End Select
Slider2.Value = 0
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus

dPercent = CDbl(Slider1.Value)
dPercent1 = CDbl(Slider2.Value)
dPercent2 = CDbl(Slider3.Value)
nHue = CDbl(Slider3.Value)
nSat = CDbl(Slider2.Value)
nValue = CDbl(Slider1.Value)
nDeltaX = CLng(Slider3.Value)
nDeltaY = CLng(Slider2.Value)
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
End If
Label8(7).Visible = True
Select Case Fx
Case 0
    If Limit = False Then
        ie2.FilterSharpen dPercent
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.FilterSharpen dPercent
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
Case 1
    If Limit = False Then
        ie2.FilterBlur dPercent
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.FilterBlur dPercent
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
Case 2
    If Limit = False Then
        ie2.BrightnessAndContrast dPercent, dPercent1
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.BrightnessAndContrast dPercent, dPercent1
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
Case 3
    If Limit = False Then
        ie2.HSVAdjustment nHue, nSat, nValue
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.HSVAdjustment nHue, nSat, nValue
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
Case 4
    If Limit = False Then
        ie.DropShadow 6, 6, dPercent
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.DropShadow 6, 6, dPercent
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
    Command4_Click
Case 5
    Dim lngTolerance As Long
    Dim colOld As Long
    Dim colNew As Long
    colOld = CDec(PickColor)
    colNew = CDec(ReplaceColor)
    lngTolerance = CLng(Slider1.Value)
    If Limit = False Then
        ie.ReplaceColor colOld, colNew, lngTolerance
        If Not ie.ThreadRunning Then Set THBImage1.Picture = ie.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.ReplaceColor colOld, colNew, lngTolerance
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
Case 6
    ASpinnerBox1.Value = -1 * Slider1.Value
    ASpinnerBox2.Value = -1 * Slider1.Value
    ASpinnerBox3.Value = -1 * Slider1.Value
    ASpinnerBox4.Value = -1 * Slider1.Value
    ASpinnerBox5.Value = 9 * Slider1.Value + 1
    ASpinnerBox6.Value = -1 * Slider1.Value
    ASpinnerBox7.Value = -1 * Slider1.Value
    ASpinnerBox8.Value = -1 * Slider1.Value
    ASpinnerBox9.Value = -1 * Slider1.Value
    Command11_Click
End Select

Slider1.Value = 0

End Sub

Private Sub Slider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
dPercent = CDbl(Slider1.Value)
dPercent1 = CDbl(Slider2.Value)
dPercent2 = CDbl(Slider3.Value)
nHue = CDbl(Slider3.Value)
nSat = CDbl(Slider2.Value)
nValue = CDbl(Slider1.Value)
nDeltaX = CLng(Slider3.Value)
nDeltaY = CLng(Slider2.Value)
If RegionsFlag = False Then Undoit
If Limit = False Then
    Set ie2.Picture = ie.Picture
    Set THBImage3.Picture = ie2.THBStdPicture
    THBImage3.ZoomFit
End If
Label8(7).Visible = True

Select Case Fx
Case 0
Case 1
Case 2
Case 3
    If Limit = False Then
        ie2.HSVAdjustment nHue, nSat, nValue
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.HSVAdjustment nHue, nSat, nValue
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        'ie.OverlayWithTransparency ieLimit, 0, 0, 44
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If

Case 4
    If Limit = False Then
        ie.DropShadow nDeltaX, nDeltaY, dPercent
        If Not ie2.ThreadRunning Then Set THBImage3.Picture = ie2.THBStdPicture
    Else
        Set rg = THBImage1.RegionGetByIndex(0)
        rg.GetPoint 0, lngLeft, lngTop
        rg.GetPoint 2, lngRight, lngBottom
        ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
        ieLimit.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
        ieLimit.DropShadow nDeltaX, nDeltaY, dPercent
        ie.Overlay ieLimit, lngLeft, lngTop, 100, False
        UpdatePicInfo
        Set THBImage1.Picture = ie.THBStdPicture
    End If
End Select
Slider3.Value = 0
End Sub


Private Sub Slider4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'clist4.setfocus
If Slider4.Value < 0 Then
    THBImage1.ZoomMinus Abs(Slider4.Value)
Else
    THBImage1.ZoomPlus Abs(Slider4.Value)
End If
THBImage1.Redraw
Slider4.Value = 0
End Sub

Private Sub Spinner1_LostFocus()
 On Error Resume Next
If Check3.Value = 1 Then
        VScroll1.Value = -1 * Int(Spinner1.Text)
        VScroll2.Value = -1 * Int(Spinner1.Text * AspectH / AspectW)
Else
        VScroll1.Value = -1 * Int(Spinner1.Text)
End If

End Sub


Private Sub Spinner2_LostFocus()
 On Error Resume Next
If Check3.Value = 1 Then
        VScroll2.Value = -1 * Int(Spinner2.Text)
        VScroll1.Value = -1 * Int(Spinner2.Text * AspectW / AspectH)
Else
        VScroll2.Value = -1 * Int(Spinner2.Text)
End If
End Sub

Private Sub CleanTxt1()
Dim Moose, TempClip
    Moose = Trim(Text1.Text)
    TempClip = ""
    Do While TempClip <> Moose
        TempClip = Moose
        For i = 33 To 47
            Moose = Replace(Moose, Chr(i), "")
        Next
        For i = 58 To 64
            Moose = Replace(Moose, Chr(i), "")
        Next
        For i = 91 To 94
            Moose = Replace(Moose, Chr(i), "")
        Next
        Moose = Replace(Moose, vbCrLf, " ")
    Loop
Text1.Text = Moose & " " & Format(Now, "ddmmyyhhmmss")

End Sub

Private Sub Text12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub


Private Sub Text13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub

Private Sub THBImage1_DropFile(ByVal strFileName As String)
Dim Please As String
On Error Resume Next
    'CList4.Clear
    'Option2.Value = True        'Sort by Name
    'List1.Clear 'video
    'List2.Clear 'pics
    'List3.Clear 'pics
    'List6.Clear 'mht, htm, pdf
    'List7.Clear 'txt, rtf
    PicThere = False
    TxtThere = False
    Txt1There = False
    VideoThere = False
    
    File1.path = GetFilePath(strFileName)
    File1.Refresh

    'set file path to whereever drag from
    'Open App.path & "\LastPath" For Output As #1
        'Print #1, File1.path
    'Close #1
    
    'loading lists and clist4
        Please = strFileName
If Trim(Please) = "" Then GoTo EmptyOne     'Don't add empty strings
        For i = 0 To 13 'pics
            If LCase(GetFileExtension(Please)) = PictureExt(i) Then
                    PicThere = True
                    CList4.AddItem Please
                    List3.AddItem Please    'sort name
                    List2.AddItem Format(FileDateTime(Please), "YYYYMMDDHHMMSS") & "*" & Please 'sort date
                Exit For
            End If
        Next
        For i = 0 To 3  'htm pdf
            If LCase(GetFileExtension(Please)) = TextExt(i) Then
                    TxtThere = True
                    List6.AddItem Please
                    Exit For
            End If
        Next
        For i = 0 To 1  'txt rtf
            If LCase(GetFileExtension(Please)) = TextExt1(i) Then
                    Txt1There = True
                    List7.AddItem Please
                    Exit For
            End If
        Next
        For i = 0 To 10   'videos
            If LCase(GetFileExtension(Please)) = VideoExt(i) Then
                VideoThere = True
                    List1.AddItem Please
                    Exit For
            End If
        Next
EmptyOne:
    
    If PicThere = True Then
        Undoit
        WebBrowser1.Visible = False
        RichTextBox1.Visible = False
        THBImage1.Visible = True
        Mht = False
        RTF = False
        File1.Pattern = "*.psd;*.gif;*.jpeg;*.jpg;*.ico;*.cur;*.wmf;*.emf;*.bmp;*.pcx;*.tif;*.tiff;*.png;*.tga"
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        Label4.Caption = CList4.Text
        LoadFileAndUpdateDisplay CList4.List(0)
        Exit Sub
    End If
    If TxtThere = True Then
        WebBrowser1.Visible = True      'use browser
        RichTextBox1.Visible = False
        THBImage1.Visible = False
        File1.Pattern = "*.mht;*.htm;*.html;*.pdf"
        File1.Refresh
        Mht = True
        RTF = False
        For i = 0 To List6.ListCount - 1
                CList4.AddItem List6.List(i)
        Next
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        WebBrowser2.Navigate "about:<html><body bgcolor=" & Chr(34) & "Blue" & Chr(34) & _
            " scroll='no'><p align=" & Chr(34) & "center" & Chr(34) & "><img src='" & _
            Trim(ThePicture) & "'></img></p></body></html>"
            
        WebBrowser2.Top = THBImage1.Top
        WebBrowser2.Left = THBImage1.Left
        WebBrowser2.Height = THBImage1.Height
        WebBrowser2.Width = THBImage1.Width
        Exit Sub
    End If
    If Txt1There = True Then
        Haburabadooda22.Visible = True
        THBImage1.Visible = False
        WebBrowser1.Visible = False
        Mht = False
        RTF = True
        For i = 0 To List7.ListCount - 1
                CList4.AddItem List7.List(i)
        Next
        File1.Pattern = "*.txt;*.rtf"
        File1.Refresh
        CList4.Refresh
        CList4.Text = CList4.List(0)
        ThePicture = CList4.List(0)
        RichTextBox1.Visible = True
        RichTextBox1.LoadFile CList4.List(0)
        Exit Sub
    End If
    If VideoThere = True Then
        Me.Hide
        Load VideoLibrary
        VideoLibrary.Moviee.Clear
        For i = 0 To List1.ListCount - 1
                VideoLibrary.Moviee.AddItem List1.List(i)
        Next
        VideoLibrary.Moviee.Text = VideoLibrary.Moviee.List(0)
        VideoLibrary.Show
        Exit Sub
     End If
     CList4.ListIndex = CList4.ListCount - 1
     
End Sub

Private Sub THBImage1_LButtonDblClk(ByVal Keys As THBImageLibCtl.thbMouseKeys, ByVal X As Long, ByVal Y As Long)
If Mht = True Or RTF = True Then Exit Sub
Dim i As Long
Angst.Visible = False
MouseWheel1.WheelDisconnect
Load Form2
Form2.Show
Form2.CList4.Clear
SlideShowFlag = True
For i = 0 To Angst.CList4.ListCount - 1
        Form2.CList4.AddItem Angst.CList4.List(i)
Next
End Sub

Private Sub THBImage1_LButtonDown(ByVal Keys As THBImageLibCtl.thbMouseKeys, ByVal X As Long, ByVal Y As Long)
If RegionsFlag = True Then
   XM = X
   YM = Y
End If
End Sub



Private Sub THBImage1_LButtonUp(ByVal Keys As THBImageLibCtl.thbMouseKeys, ByVal X As Long, ByVal Y As Long)
If MoveFlag = True Then RegionsFlag = False

End Sub

Private Sub THBImage1_MouseMove(ByVal Keys As THBImageLibCtl.thbMouseKeys, ByVal X As Long, ByVal Y As Long)
On Error Resume Next
THBImage1.SetFocus
MouseWheel1.WheelConnect THBImage1.hwnd
'If MoveMe = True Then
'    Haburabadooda8.Top = Y + 5
'    Haburabadooda8.Left = x + 5
'    Exit Sub
'End If
''clist4.setfocus
If RegionsFlag = True Then
    MoveFlag = True
    Dim Region As THBImageLibCtl.THBRegion
    Set Region = THBImage1.RegionGetById(0)
    Xd = X - XM
    Yd = Y - YM
    Dim i As Long
    Dim X1 As Long
    Dim Y1 As Long
    For i = 0 To Region.NumPoints - 1
        Region.GetPoint i, X1, Y1
        X1 = X1 + Xd: Y1 = Y1 + Yd
        Region.SetPoint i, X1, Y1
    Next i
    X = (X1 / 2) + 50
    Y = Y1 / 2
    XM = X
    YM = Y
    MoveFlag = True
    THBImage1.Redraw
End If
End Sub

Private Sub THBImage1_NewRegion(ByVal Region As THBImageLibCtl.THBRegion)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long
    Dim Ration As Single
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    On Error Resume Next
      Set Region.Picture = ImageList1.ListImages(nMarkerPictureCounter + 1).Picture
      Set Region.PictureHighlighted = ImageList1.ListImages(nMarkerPictureCounter + 2).Picture
      nMarkerPictureCounter = nMarkerPictureCounter + 2
      If nMarkerPictureCounter > 5 Then nMarkerPictureCounter = 0
If Crop = False Then
    If Region.NumPoints = 1 Then
        Region.PenWidth = 10
        Region.PenColor = RGB(128, 0, 0)
    'Standard Regionstyle
    Else
        Region.PenStyle = thbPS_SOLID
        Region.PenColor = RGB(0, 0, 255)
        Region.PenWidth = 2
    End If
    
    'Apply Fillstyle
    'If cbFill.Value = 1 Then
        'Region.FillStyle = thbFS_CROSS
        'Region.FillBkMode = thbFM_OPAQUE
        'Region.FillForeColor = RGB(255, 255, 128)
        'Region.FillBackColor = RGB(0, 128, 0)
    'End If
    'Transparent Fillstyle
    'If cbFillTransparent.Value = 1 Then
        'Region.FillStyle = thbFS_CROSS
        'Region.FillBkMode = thbFM_TRANSPARENT
        'Region.FillForeColor = RGB(255, 255, 128)
    'End If
        
    'UseMarkerImage
    'Assign the Url to the region
    'If Len(tbRegionUrl.Text) > 0 Then
        'Region.HyperlinkURL = tbRegionUrl.Text
    'End If
    'Region.AsString = "0;0|319;0|319;199|0;199|0;0|"
    THBImage1.RegionAdd Region
    Region.Id = 0
    Stringy = ""
        For i = 0 To Region.NumPoints - 1
            Region.GetPoint i, X, Y
            Stringy = Stringy & X & ";" & Y & "|"
        Next i
        RegionsFlag = True
    Exit Sub
Else
    Crop = False
    'We need rectangle region
    If Region.NumPoints <> 5 Then
       'MsgBox "Invalid Fence!"
        Exit Sub
    End If
    
    'Define the part we are interested in
    Region.GetPoint 0, lngLeft, lngTop
    Region.GetPoint 2, lngRight, lngBottom
    If lngLeft > lngRight Then
        Region.GetPoint 0, lngRight, lngBottom
        Region.GetPoint 2, lngLeft, lngTop
    End If
    If Crap = True Then
        If lngRight > lngBottom Then
            lngBottom = ((480 * (lngBottom - lngTop)) / 640) + lngTop '
            PullQuote lngLeft, lngTop, lngRight, lngBottom
        Else
            lngRight = ((640 * lngRight - lngLeft) / 480) + lngLeft
            PullQuote lngLeft, lngTop, lngRight, lngBottom
        End If
        Exit Sub
    End If
    PullQuote lngLeft, lngTop, lngRight, lngBottom

End If

End Sub
Private Sub PullQuote(lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long)
    Dim ienew As New THBImageEdit
    Dim pBytes As Long
    Dim lngSizeBytes As Long
    'Dim frmSV As frmSimpleView

    On Error GoTo ErrHandler
    
    ie.CropIntoMemory lngLeft, lngTop, lngRight, lngBottom, pBytes, lngSizeBytes
    'Create a new image from the cropped part
    'ie.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    ie.CreateFromMemory pBytes, lngSizeBytes, lngRight - lngLeft + 1, lngBottom - lngTop + 1, ie.BitsPerPixel
    Set THBImage1.Picture = ie.THBStdPicture
    'Just necessary because the modal form prevents some mouse events
    THBImage1.RegionCancel
    THBImage1.ZoomFit
    ie.BMPUseRLE = True
    ie.SavePictureToFile "c:\1down\" & Format(Now, "ddmmyyhhmmss") & ".bmp", thbifBMP    'App.Path & "\" & Format(Now, "ddmmyyhhmmss") & ".bmp", thbifBMP
    butClipboardCopyTo_Click
    
    'Display the cropped part in a new form
    'Set frmSV = New frmSimpleView
    'frmSV.Init ieNew
    'frmSV.Show vbModal, Me
    'Me.THBImage1.Picture = frmSV.THBImage1.Picture
    Exit Sub
    
ErrHandler:
    'MsgBox err.Description
End Sub
Public Sub Undoit()
On Error Resume Next
Set ie1.Picture = ie.Picture
Set THBImage2.Picture = ie1.THBStdPicture
THBImage2.ZoomFit
ie.BMPUseRLE = True
ie.SavePictureToFile App.path & "\Undo\" & Format(Now, "ddmmyyhhmmss") & ".bmp", thbifBMP

End Sub

Private Sub THBImage1_OverRegion(ByVal lngRegionIdSelected As Long, ByVal X As Long, ByVal Y As Long)
If Mht = True Or RTF = True Then Exit Sub
    mnuResize.Visible = False
    mnuCrop.Visible = False
    mnuCropCapture.Visible = False
    mnuScan.Visible = False
    mnuRecycle.Visible = False
    PopupMenu mnuFile

End Sub
Private Sub RegionOffset(Region As THBImageLibCtl.THBRegion, XOffset As Long, YOffset As Long)
On Error Resume Next
Dim i As Long
    Dim X As Long
    Dim Y As Long
    For i = 0 To Region.NumPoints - 1
        Region.GetPoint i, X, Y
        X = X + XOffset: Y = Y + YOffset
        Region.SetPoint i, X, Y
    Next i
End Sub

Private Sub THBImage1_PictureChanged()

THBImage1.ZoomFit
Propertiz
If CList4.ListIndex = -1 Then
    Label4.Caption = CList4.List(CList4.ListIndex)
    Label5.Caption = CList4.ListIndex
End If
AsBut = False
If ScanFlag = True Then
    ScanFlag = False
    Call Haburabadooda3_MouseUp(0, 0, 0, 0)
End If
End Sub


Private Sub THBImage1_RButtonDown(ByVal Keys As THBImageLibCtl.thbMouseKeys, ByVal X As Long, ByVal Y As Long)


Exit Sub
THBImage1.PopupMenu = False
Dim Region As THBImageLibCtl.THBRegion
Set Region = THBImage1.RegionGetById(0)
    Dim i As Long
    Dim X1 As Long
    Dim Y1 As Long
    For i = 0 To Region.NumPoints - 1
        Region.GetPoint i, X1, Y1
        If X1 - X < 0 Then
            X1 = X1 + X
        Else
            X1 = X1 - X
        End If
        If Y1 - Y < 0 Then
            Y1 = Y1 + Y
        Else
            Y1 = Y1 - Y
        End If
        Region.SetPoint i, X1, Y1
    Next i
THBImage1.Redraw
End Sub

Private Sub THBImage1_RButtonUp(ByVal Keys As THBImageLibCtl.thbMouseKeys, ByVal X As Long, ByVal Y As Long)
MoveMe = False
End Sub

Private Sub THBImage2_LButtonUp(ByVal Keys As THBImageLibCtl.thbMouseKeys, ByVal X As Long, ByVal Y As Long)
'clist4.setfocus
If Mht = True Or RTF = True Then Exit Sub
mnuUndo_Click
End Sub

Private Sub THBImage2_PictureChanged()
     THBImage1.ZoomFit

End Sub

Private Sub THBImage3_PictureChanged()
Set ie.Picture = ie2.Picture
Set THBImage1.Picture = ie.THBStdPicture
'THBImage1.ZoomFit
'UpdatePicInfo

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If j1 <> CList4.ListCount - 1 And CList4.List(CList4.ListIndex) <> "Drag & Drop on me now, if you dare" Then
    If SlideShowFlag = True Then
        Form2.LoadFileAndUpdateDisplay CList4.List(j1)
    Else
        LoadFileAndUpdateDisplay CList4.List(j1)
    End If
    CList4.Text = CList4.List(j1)
    If mnuReverse.Checked = False Then
        j1 = j1 + 1
    Else
        j1 = j1 - 1
    End If
    Propertiz
    'Label14 = "Number " & j1 & " of " & CList4.ListCount
    CList4.ListIndex = j1
Else
    'j1 = 0
End If

End Sub


Private Sub Tips_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub Undoer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Thumb = False Then
    Thumb = True
    'Load frmThumbs
    'frmThumbs.Show
End If
End Sub

Private Sub VScroll1_Change()
 On Error Resume Next
Spinner1.Text = Abs(VScroll1.Value)
If Check3.Value = 1 Then
        Spinner2.Text = Int(Spinner1.Text * AspectH / AspectW)
End If

End Sub

Private Sub VScroll2_Change()
 On Error Resume Next
Spinner2.Text = Abs(VScroll2.Value)
If Check3.Value = 1 Then
        Spinner1.Text = Int(Spinner2.Text * AspectW / AspectH)
End If
End Sub

Private Sub WebBrowser1_DownloadComplete()
On Error Resume Next
WebBrowser1.Silent = True
'clist4.setfocus
End Sub

Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
'clist4.setfocus
End Sub

Public Sub Yes_Click()
Dim Moove As Boolean
Dim J, k As Long
Dim ToFileName As String
On Error Resume Next
k = CList4.ListIndex
Moove = False
    If Check4.Value = 0 Then
        ToFileName = GetFileTitle(CList4.Text)
    Else
        ToFileName = GetFileTitle(ReNameMoveFile)
    End If
    Do While Moove = False
        If Dir(Patherino & ToFileName) = "" Then
            If Dir(CList4.Text) <> "" Then
                'MsgBox CList4.Text & " to " & Patherino & ToFileName
                'MsgBox ReNameMoveFile
                'MsgBox ReNameMoveFile & "  to  " & Patherino & ToFileName
                If Check4.Value = 0 Then
                    FileOPS1.MoveFile CList4.Text, Patherino & ToFileName
                Else
                    FileOPS1.MoveFile ReNameMoveFile, Patherino & ToFileName
                End If
            End If
            LoadFileAndUpdateDisplay CList4.Text
            If LCase(Right(ToFileName, 3)) = "gif" Then
                LoadFileAndUpdateDisplay ""
            End If
            Moove = True
        Else
            ToFileName = ToFileName & Str(J)
            J = J + 1
        End If
        If Check4.Value = 0 Then
            Kill CList4.Text
            CList4.RemoveItem CList4.ListIndex
            CList4.ListIndex = k
            CList4.Text = CList4.List(k)
            CList4_Click
        End If
    Loop
'Tips.Visible = False
End Sub



