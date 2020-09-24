VERSION 5.00
Begin VB.Form Sender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sending app"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2550
   Icon            =   "Sender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   2550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Captured.bmp"
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Enter text to send:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Sender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'Send string project by Peter Hebels, Website "www.phsoft.nl"                             *
'I'am not responsible for any damages may caused by this project                           *
'******************************************************************************************
      
'To test this project you have to run both project's from the 'Send' And
''Recieve' directory's

'This is the sending project

      'Memory copy data structure
      Private Type COPYDATASTRUCT
              dwData As Long
              cbData As Long
              lpData As Long
      End Type

      'Copy memory data
      Private Const WM_COPYDATA = &H4A

      'Find a window on the desktop
      Private Declare Function FindWindow Lib "user32" Alias _
         "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
         As String) As Long

      'Used to send messages between app's
      Private Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal _
         wParam As Long, lParam As Any) As Long

      'Copies a block of memory from one location to another.
      Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

      Public Sub Command1_Click()
          Dim cds As COPYDATASTRUCT
          Dim ThWnd As Long
          Dim buf(1 To 255) As Byte

      'Check if text is entered
      If Text1.Text = "" Then Exit Sub
            
      ' Get the hWnd of the target application
         ThWnd = FindWindow(vbNullString, "Testapp")
      'Text to send
         A$ = Text1.Text
      ' Copy the string into a byte array, converting it to ASCII
         Call CopyMemory(buf(1), ByVal A$, Len(A$))
          cds.dwData = 3
          cds.cbData = Len(A$) + 1
          cds.lpData = VarPtr(buf(1))
      'Send the string to the other app
         i = SendMessage(ThWnd, WM_COPYDATA, Me.hwnd, cds)
      End Sub
     

