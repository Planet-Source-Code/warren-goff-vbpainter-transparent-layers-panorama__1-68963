VERSION 5.00
Begin VB.UserControl TransPicBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   MaskColor       =   &H00FFFFFF&
   MaskPicture     =   "TransPicBox.ctx":0000
   ScaleHeight     =   4080
   ScaleWidth      =   6450
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   0
      ScaleHeight     =   1800
      ScaleWidth      =   2250
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   2250
   End
End
Attribute VB_Name = "TransPicBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_AutoSize = True

Enum pbBackStyle
    [Opaque] = 1
    [Transparent] = 0
End Enum

Enum pbBorderstyle
    [No Border] = 0
    [Fixed Single] = 1
End Enum

Enum pbAppearance
    [3D] = 1
    [Flat] = 0
End Enum
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim m_AutoSize As Boolean
Dim pHwnd As Long


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As pbAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As pbAppearance)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As pbBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As pbBackStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As pbBorderstyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As pbBorderstyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    Set UserControl.MaskPicture = New_Picture
    Set Picture1.Picture = New_Picture
    If m_AutoSize Then
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
      
    PropertyChanged "Picture"
End Property

Private Sub UserControl_Initialize()
    pHwnd = hWnd                        ' Window handle of control

End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", 16777215)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
End Sub

Private Sub UserControl_Resize()
    If m_AutoSize Then
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, 16777215)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
End Sub
'


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AutoSize = m_def_AutoSize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    If m_AutoSize Then
        UserControl.Width = Picture1.Width
        UserControl.Height = Picture1.Height
    End If
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub
