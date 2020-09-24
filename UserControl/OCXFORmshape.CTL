VERSION 5.00
Begin VB.UserControl ocxFormShape 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   Enabled         =   0   'False
   FillStyle       =   0  'Solid
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   510
   ScaleWidth      =   420
End
Attribute VB_Name = "ocxFormShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'************************************************************************************************
'   ocxFormShape                    Version 2
'
'   Programmer:
'       Michael Schofield           July 10, 2001
'
'   Functionality:
'       This ActiveX component allows the user to reshape the form where it resides.
'       Along with adding mouse movement to the form.
'       API Calls are used to create a region of a specific shape, and then set on the
'       form.
'
'   Properties:
'
'       Shape =         The user must pick from a list of different shapes to make the form.
'       CornerCurve =   Allows the user to specify how round they would like the corners to be
'                       on shapes with rounded corners.
'       Inverted =      Allow the use to turn the shape upside down.
'       MouseMovable =  Allows the user to speciefy if the form should move when the user
'                       clicks and drags the mouse while on the form. Aswell, it allows the
'                       user to resize the form. This property is used when the form has no
'                       borderStyle.
'
'
'   Problems/Enhancements
'       There seems to be a problem Unhooking the SystemMessages.
'       I have placed the Unhook method in the Control_unload event,
'       but it does not seem to get fired before the form is unload and it causes
'       my application to crash VB...
'           - Must fix this problem.
'           - FIXED June 26.
'               - the problem was that the CallWindowProc function call at the end
'                 of the WindowProc function was being passed 0 as the m_lngPrevWndProc
'                 value. this was crashing the system because that is supposed to be
'                 a pointer to a form, or object. I simply put in a check to make sure
'                 it was not 0.
'
'       I would like to be able to add a form caption, that takes up the upper part
'       of the new shape that the form in is, aswell as a border.
'
'       If possible I would like to generate a shape from the points given by the user,
'       instead of hard coding ones that can be selected.
'
'************************************************************************************************

Option Explicit


Private Declare Function API_CreateRectRgn Lib "gdi32" Alias "CreateRectRgn" _
                                            (ByVal X1 As Long, ByVal Y1 As Long, _
                                             ByVal X2 As Long, ByVal Y2 As Long) As Long
                                                    
Private Declare Function API_CreateRoundRectRgn Lib "gdi32" Alias "CreateRoundRectRgn" _
                                            (ByVal X1 As Long, ByVal Y1 As Long, _
                                             ByVal X2 As Long, ByVal Y2 As Long, _
                                             ByVal X3 As Long, ByVal Y3 As Long) As Long
                                                         
Private Declare Function API_CreateEllipticRgn Lib "gdi32" Alias "CreateEllipticRgn" _
                                            (ByVal X1 As Long, ByVal Y1 As Long, _
                                             ByVal X2 As Long, ByVal Y2 As Long) As Long
                                                        
Private Declare Function API_CreatePolygonRgn Lib "gdi32" Alias "CreatePolygonRgn" _
                                            (lpPoint As POINTAPI, ByVal nCount As Long, _
                                             ByVal nPolyFillMode As Long) As Long
                                                       
Private Declare Function API_SetWindowRgn Lib "user32" Alias "SetWindowRgn" _
                                            (ByVal hwnd As Long, ByVal hRgn As Long, _
                                             ByVal bRedraw As Boolean) As Long
                                             
Private Declare Function API_DeleteObject Lib "gdi32" Alias "DeleteObject" _
                                            (ByVal hObject As Long) As Long
                                             
                                                    

' Property constants
Private Const m_strSHAPECaption As String = "Shape"
Private Const m_strROUNDCaption As String = "CornerCurve"
Private Const m_strInvertedCaption As String = "Inverted"
Private Const m_strMouseMoveableCaption As String = "MouseMoveable"
Private Const m_ImgControlPicture As String = "Form.ico"

Private Const m_lngSHAPEValue As Long = 1
Private Const m_intROUNDValue As Integer = 40
Private Const m_bolInvertedValue As Boolean = False
Private Const m_bolMouseMoveableValue As Boolean = False

' Variables used in the subClassing Event
'Private WithEvents objGrabWindowMessages As GetWindowMessages.clsGetWinMessage

Private Const MK_LBUTTON = &H1              ' value of Left mouse button
Private Const WM_MOUSEMOVE = &H200          ' value when mouse moves
Private Const WM_LBUTTONDOWN = &H201        ' value when left mouse button pressed
Private Const m_lngBorderWidth As Long = 30 ' the width of the difference between the parent width
                                            ' and the regions width
Private m_lngRegionWidth                    ' difference between the region and the parent width
Private m_lngRegionHeight                   ' difference between the region and the parent Height
Private m_lngOriginalX As Long              ' X poistion of the mouse when the user first clicks
                                            ' the left mouse button.
Private m_lngOriginalY As Long              ' Y poistion of the mouse when the user first clicks
                                            ' the left mouse button.
Private m_lngPointer As Long                ' Mouse Pointer



Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum MyShape
    frmSquare = 0
    frmRectangle = 1
    frmOval = 2
    frmCircle = 3
    frmRoundRectangle = 4
    frmRoundSquare = 5
    frmTriangle = 6
    frmOctagon = 7
    frmDiamond = 8
    frmPentagon = 9
    frmHexagon = 10
End Enum


Private m_lngFrmShape As MyShape
Private m_intCornerValue As Integer
Private m_bolInverted As Boolean
Private m_bolMouseMovable As Integer
Private m_bolHooked As Boolean          ' Am I currently Being Spied on?

Private m_ErrNum As Long                  ' Error Checking
Private m_ErrDesc As String
Private m_ErrSource As String

Public Property Get shape() As MyShape

    shape = m_lngFrmShape

End Property

Public Property Let shape(ByVal lngNewShape As MyShape)
'*****
'   When the user changes the shape he wants to use
'   we are to redraw the shape on his form.
'****
    
    If (lngNewShape <= 10) And (lngNewShape >= 0) Then
        m_lngFrmShape = lngNewShape
    Else
        ' if it is not the right type, make it the default.
        m_lngFrmShape = m_lngSHAPEValue
    End If
    
    Call DrawForm
    
End Property

Public Property Let CornerCurve(ByVal intValue As Integer)
'*****
'   Allows the user to Set the amount the corners should
'   be rounded in the squared, and rectangle objects
'   the corners can not be greater than 1/2 the form width
'*****
    
    If intValue > (0.5 * Parent.Width) Then
        m_intCornerValue = (0.5 * Parent.Width)
    Else
        m_intCornerValue = intValue
    End If
    
    Call DrawForm
    
End Property

Public Property Get CornerCurve() As Integer
'*****
'   Return to the user the size of the corners of our Form.
'*****
    
    CornerCurve = m_intCornerValue

End Property

Public Property Let MouseMoveable(ByVal bolMoveable As Boolean)
'*****
'   Do we want the form to move when the user clicks and drags the mouse?
'*****
    
    m_bolMouseMovable = bolMoveable

End Property

Public Property Get MouseMoveable() As Boolean
    
    MouseMoveable = m_bolMouseMovable

End Property

Public Property Get Inverted() As Boolean
Attribute Inverted.VB_Description = "Returns/Sets the inverted value of the form."
    
    Inverted = m_bolInverted

End Property

Public Property Let Inverted(ByVal m_bolInvertedValue As Boolean)
'*****
'   allows the form to be inverted (upside down)
'*****
    
    m_bolInverted = m_bolInvertedValue
    
    Call DrawForm
    
End Property


Private Sub UserControl_InitProperties()
'*****
'   When the control is added to the form, this sub procedure is executed.
'   in here, draw the shape on the form.
'*****

    ' check to see if our parent is a form.
    ' if it is not send an error message to the user.
    
    If Not (TypeOf Parent Is Form) Then
        On Error GoTo 0
        err.Raise vbObjectError + 1025, "InitProperties", _
                 "UserControl must be placed on a form."
    Else
        shape = m_lngSHAPEValue
        CornerCurve = m_intROUNDValue
        Inverted = m_bolInvertedValue
        MouseMoveable = m_bolMouseMoveableValue
    End If
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'*****
'   Read the property of Shape and put it in our bag
'*****

    
       
    With PropBag
        shape = .ReadProperty(m_strSHAPECaption, m_lngSHAPEValue)
        CornerCurve = .ReadProperty(m_strROUNDCaption, m_intROUNDValue)
        Inverted = .ReadProperty(m_strInvertedCaption, m_bolInvertedValue)
        MouseMoveable = .ReadProperty(m_strMouseMoveableCaption, m_bolMouseMoveableValue)
    End With
        
    
    If Ambient.UserMode And MouseMoveable Then
    ' only hook the form when the form is in runtime mode (ocx in UserMode).
    ' not when the ocx is in AUthor, or Developer mode.
        'Set objGrabWindowMessages = New GetWindowMessages.clsGetWinMessage
        
        'Call objGrabWindowMessages.Hook(Parent)
        m_bolHooked = True
    End If
    
    Call DrawForm
    
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'*****
'   Write the property of Shape to our bag.
'*****

    With PropBag
        .WriteProperty m_strSHAPECaption, shape, m_lngSHAPEValue
        .WriteProperty m_strROUNDCaption, CornerCurve, m_intROUNDValue
        .WriteProperty m_strInvertedCaption, Inverted, m_bolInvertedValue
        .WriteProperty m_strMouseMoveableCaption, MouseMoveable, m_bolMouseMoveableValue
    End With
    
End Sub

Private Sub UserControl_Terminate()
'*****
'   Release the form
'   If the user has selected Moveable = True, we will need to UnHook the form
'*****
   If m_bolHooked Then
        'Call objGrabWindowMessages.Unhook
        m_bolHooked = False
   End If
   
   'Set objGrabWindowMessages = Nothing
   
End Sub

Private Sub UserControl_Resize()
'*****
'   Dont allow the user to resize me.
'*****
       
    ' I put this here instead of initialize incase it fails.
    On Error Resume Next
    UserControl.Picture = LoadPicture(App.Path & "\" & m_ImgControlPicture)
    
    With UserControl
        .Height = 500
        .Width = 450
    End With
    
End Sub

Public Sub Redraw()
'*****
'
'*****
    Call DrawForm
End Sub

Private Sub DrawForm()
'*****
'   Create a region, of specified shape, and apply it to the users Form.
'*****

    Dim mChildFormRegion As Long                ' Pointer to form region.
    Dim arrPointsTriangle() As POINTAPI         ' array of pointsAPI
    Dim lngCircleValue As Long                  ' value of the points in a circle

    Dim lngParentHeight As Long                 ' Height of form (in twips)
    Dim lngParentWidth As Long                  ' width of form  (in twips)
    
    
    lngParentWidth = Parent.Width / Screen.TwipsPerPixelX
    lngParentHeight = Parent.Height / Screen.TwipsPerPixelY
    
    
    Select Case m_lngFrmShape
    
        Case frmRoundRectangle
        
            mChildFormRegion = API_CreateRoundRectRgn(0, 0, _
                                                 lngParentWidth, lngParentHeight, _
                                                 CornerCurve, CornerCurve)
            
           
        Case frmRectangle
            ' do nothing, that is the shape of the form as it is.
            
        Case frmOval
        
            mChildFormRegion = API_CreateEllipticRgn(0, 0, lngParentWidth, lngParentHeight)
    
        Case frmCircle
            
            lngCircleValue = (lngParentWidth - lngParentHeight) / 2
            mChildFormRegion = API_CreateEllipticRgn(lngCircleValue, 0, _
                                                 lngCircleValue + lngParentHeight, _
                                                 lngParentHeight)
            
        Case frmRoundSquare
        
            mChildFormRegion = API_CreateRoundRectRgn(20, 0, (lngParentWidth) - 20, _
                                                 lngParentHeight, CornerCurve, CornerCurve)
                            
        Case frmSquare
        
            mChildFormRegion = API_CreateRectRgn(20, 0, (lngParentWidth) - 20, lngParentHeight)
                
        Case frmTriangle

            ReDim arrPointsTriangle(3)
                    

            If Inverted Then
                arrPointsTriangle(0).X = 0
                arrPointsTriangle(0).Y = 0
                arrPointsTriangle(1).X = 0.5 * lngParentWidth
                arrPointsTriangle(1).Y = lngParentHeight
                arrPointsTriangle(2).X = lngParentWidth
                arrPointsTriangle(2).Y = 0
            Else
                arrPointsTriangle(0).X = 0.5 * lngParentWidth
                arrPointsTriangle(0).Y = 0
                arrPointsTriangle(1).X = 0
                arrPointsTriangle(1).Y = lngParentHeight
                arrPointsTriangle(2).X = lngParentWidth
                arrPointsTriangle(2).Y = lngParentHeight
            End If
            

            mChildFormRegion = API_CreatePolygonRgn(arrPointsTriangle(0), 3, 1)
        
        Case frmOctagon
     
            ReDim arrPointsTriangle(8)
            
            arrPointsTriangle(0).X = 0.25 * lngParentWidth
            arrPointsTriangle(0).Y = 0
            arrPointsTriangle(1).X = 0
            arrPointsTriangle(1).Y = 0.25 * lngParentHeight
            arrPointsTriangle(2).X = 0
            arrPointsTriangle(2).Y = lngParentHeight - (0.25 * lngParentHeight)
            arrPointsTriangle(3).X = 0.25 * lngParentWidth
            arrPointsTriangle(3).Y = lngParentHeight
            arrPointsTriangle(4).X = lngParentWidth - (0.25 * lngParentWidth)
            arrPointsTriangle(4).Y = lngParentHeight
            arrPointsTriangle(5).X = lngParentWidth
            arrPointsTriangle(5).Y = lngParentHeight - (0.25 * lngParentHeight)
            arrPointsTriangle(6).X = lngParentWidth
            arrPointsTriangle(6).Y = 0.25 * lngParentHeight
            arrPointsTriangle(7).X = lngParentWidth - (0.25 * lngParentWidth)
            arrPointsTriangle(7).Y = 0
    
            mChildFormRegion = API_CreatePolygonRgn(arrPointsTriangle(0), 8, 1)
            
        Case frmDiamond
            
            ReDim arrPointsTriangle(4)
            
            arrPointsTriangle(0).X = 0.5 * lngParentWidth
            arrPointsTriangle(0).Y = 0
            arrPointsTriangle(1).X = 0
            arrPointsTriangle(1).Y = 0.5 * lngParentHeight
            arrPointsTriangle(2).X = 0.5 * lngParentWidth
            arrPointsTriangle(2).Y = lngParentHeight
            arrPointsTriangle(3).X = lngParentWidth
            arrPointsTriangle(3).Y = 0.5 * lngParentHeight
            
            mChildFormRegion = API_CreatePolygonRgn(arrPointsTriangle(0), 4, 1)
        
        Case frmPentagon
            
            ReDim arrPointsTriangle(5)
            
            If Inverted Then
                arrPointsTriangle(0).X = 0.25 * lngParentWidth
                arrPointsTriangle(0).Y = 0
                arrPointsTriangle(1).X = 0
                arrPointsTriangle(1).Y = 0.5 * lngParentHeight
                arrPointsTriangle(2).X = 0.5 * lngParentWidth
                arrPointsTriangle(2).Y = lngParentHeight
                arrPointsTriangle(3).X = lngParentWidth
                arrPointsTriangle(3).Y = 0.5 * lngParentHeight
                arrPointsTriangle(4).X = 0.75 * lngParentWidth
                arrPointsTriangle(4).Y = 0
            Else
                arrPointsTriangle(0).X = 0.5 * lngParentWidth
                arrPointsTriangle(0).Y = 0
                arrPointsTriangle(1).X = 0
                arrPointsTriangle(1).Y = 0.5 * lngParentHeight
                arrPointsTriangle(2).X = 0.25 * lngParentWidth
                arrPointsTriangle(2).Y = lngParentHeight
                arrPointsTriangle(3).X = lngParentWidth - (0.25 * lngParentWidth)
                arrPointsTriangle(3).Y = lngParentHeight
                arrPointsTriangle(4).X = lngParentWidth
                arrPointsTriangle(4).Y = 0.5 * lngParentHeight
            End If
            
            mChildFormRegion = API_CreatePolygonRgn(arrPointsTriangle(0), 5, 1)
        
        Case frmHexagon
            
            ReDim arrPointsTriangle(6)
            
            If Inverted Then
                arrPointsTriangle(0).X = 0.5 * lngParentWidth
                arrPointsTriangle(0).Y = 0
                arrPointsTriangle(1).X = 0
                arrPointsTriangle(1).Y = 0.25 * lngParentHeight
                arrPointsTriangle(2).X = 0
                arrPointsTriangle(2).Y = 0.75 * lngParentHeight
                arrPointsTriangle(3).X = 0.5 * lngParentWidth
                arrPointsTriangle(3).Y = lngParentHeight
                arrPointsTriangle(4).X = lngParentWidth
                arrPointsTriangle(4).Y = 0.75 * lngParentHeight
                arrPointsTriangle(5).X = lngParentWidth
                arrPointsTriangle(5).Y = 0.25 * lngParentHeight
            Else
                arrPointsTriangle(0).X = 0.25 * lngParentWidth
                arrPointsTriangle(0).Y = 0
                arrPointsTriangle(1).X = 0
                arrPointsTriangle(1).Y = 0.5 * lngParentHeight
                arrPointsTriangle(2).X = 0.25 * lngParentWidth
                arrPointsTriangle(2).Y = lngParentHeight
                arrPointsTriangle(3).X = 0.75 * lngParentWidth
                arrPointsTriangle(3).Y = lngParentHeight
                arrPointsTriangle(4).X = lngParentWidth
                arrPointsTriangle(4).Y = 0.5 * lngParentHeight
                arrPointsTriangle(5).X = 0.75 * lngParentWidth
                arrPointsTriangle(5).Y = 0
            End If
            
            mChildFormRegion = API_CreatePolygonRgn(arrPointsTriangle(0), 6, 1)
            
        Case Else
            ' this should not happen, since our property let checks to make
        ' sure the value is of correct type.
    End Select
    
    ' Assign the region to take over the form.
    
    API_SetWindowRgn Parent.hwnd, mChildFormRegion, True
    
    API_DeleteObject mChildFormRegion
    
End Sub

'Private Sub objGrabWindowMessages_GetWindowMessage(ByVal hw As Long, _
                                                   ByVal uMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   ByVal lParam As Long)
'*****
'   Event triggered from the GetWindowMessages Dll to send information back to me about the message
'   grabbed.
'   it is here that I will process the information. I am looking for mouse move and mouse down events.
'*****
    
    'On Error GoTo ErrHandler
    
    'If Not (Parent Is Nothing) Then
       'Call MouseMessages(hw, uMsg, wParam, lParam)
    'End If
    
    'Exit Sub
    
'ErrHandler:
    ' dont do anything, if an error occurs here, it is becuase the Parent could not be found.
    ' that means we are unloading our control. just ignore it until I can figure out what is the
    ' proper thing to do.
    
    'err.Clear
    
'End Sub

Private Sub MouseMessages(ByVal hw As Long, _
                          ByVal uMsg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long)
'*****
'   This is what we would like to happen if we trap a mouse Message.
'*****

    On Error GoTo ErrHandler
    
    Dim xPos As Long                ' current X position of the mouse
    Dim yPos As Long                ' current Y position of the mouse
    
    If uMsg = WM_LBUTTONDOWN Then
            ' if the user presses the Left Mouse button, Store the position of the
            ' mouse on the form, when the user moves the mouse, we will need to know the starting
            ' position.
            
            Call GetHiLoWord(lParam, m_lngOriginalX, m_lngOriginalY)
    
            m_lngOriginalX = Parent.ScaleX(m_lngOriginalX, vbPixels, vbTwips)
            m_lngOriginalY = Parent.ScaleY(m_lngOriginalY, vbPixels, vbTwips)
            
            m_lngPointer = Parent.MousePointer
        
        End If
        
        If uMsg = WM_MOUSEMOVE And Not (wParam = MK_LBUTTON) Then
        ' change the mouse pointer if the position of hte mouse is along the outer
        ' region of the form.
        ' dont do this, if the mouse button is being held down.
            
            Dim lngX As Long
            Dim lngY As Long
            
            Call GetHiLoWord(lParam, lngX, lngY)
            
            lngX = Parent.ScaleX(lngX, vbPixels, vbTwips)
            lngY = Parent.ScaleY(lngY, vbPixels, vbTwips)
            
            m_lngRegionWidth = Parent.Width - m_lngBorderWidth
            m_lngRegionHeight = Parent.Height - m_lngBorderWidth
            
            ' the border of the form will take up 30 twips.
            ' and since that is hiden we will never get that pos
            ' so the form width is actually the form width - boder size, 30.
            '***
            ' This does not seem to work properly, the border width seems to
            ' fluctuate???
            
            If (lngX = 0 Or _
                lngX = m_lngRegionWidth Or _
                lngY = 0 Or _
                lngY = m_lngRegionHeight) Then
                
                Parent.MousePointer = vbSizePointer
            Else
                Parent.MousePointer = vbDefault
            End If
        
        End If
        
        If uMsg = WM_MOUSEMOVE And wParam = MK_LBUTTON Then
            
            Call GetHiLoWord(lParam, xPos, yPos)
           
            ' let xpos and ypos be the change in x/y co-ordinates.
            xPos = Parent.ScaleX(xPos, vbPixels, vbTwips) - m_lngOriginalX
            yPos = Parent.ScaleY(yPos, vbPixels, vbTwips) - m_lngOriginalY
                        
            If m_lngPointer = vbSizePointer Then
            ' the user is trying to resize the form.
            ' if the cursor is on any part of the outer limit of the form region.
                
                With Parent
                    
                    Select Case m_lngOriginalX
                        Case 0
                            .Left = .Left + xPos
                            .Width = .Width + (xPos * -1)
                        Case m_lngRegionWidth
                            .Width = .Width + xPos
                            m_lngOriginalX = m_lngOriginalX + xPos
                            m_lngRegionWidth = m_lngOriginalX
                            ' reset the original pos to be the new width, so we get a loop.
                    End Select
                
                    Select Case m_lngOriginalY
                        Case 0
                            .Top = .Top + yPos
                            .Height = .Height + (yPos * -1)
                
                        Case m_lngRegionHeight
                            .Height = .Height + yPos
                            m_lngOriginalY = m_lngOriginalY + yPos
                            m_lngRegionHeight = m_lngOriginalY
                    End Select
                
                    .Refresh
                    
                End With
            Else
            ' we are inside the form region just move the form.
                Parent.Move Parent.Left + xPos, Parent.Top + yPos
                Parent.Refresh
            End If
        End If
        
        Exit Sub
        
ErrHandler:
    ' Dont send error to user, just dont do anything.
    
    err.Clear
    'Call objGrabWindowMessages.Unhook

End Sub

Private Sub GetHiLoWord(ByVal r_lngWord As Long, _
                          ByRef r_lngLowOrder As Long, _
                          ByRef r_lngHighOrder As Long)
                          
'**************************************************************************************
'*
'* Purpose:  Get the Low-order and High-order of a word
'*
'**************************************************************************************
    On Error GoTo ErrHandler
    
    'Get Low word
    r_lngLowOrder = r_lngWord And &HFFFF&
    
    If r_lngLowOrder > Abs(&H8000&) Then
    ' we need to see if it is signed. if the value is greater then 32000 it
    ' must be a signed number.
    ' subtract that number from 62000 and multiply by -1 to find out by how much
    ' it has changed.
        r_lngLowOrder = (&HFFFF& - r_lngLowOrder) * -1
    End If
              
    'Get high word
    r_lngHighOrder = r_lngWord \ &H10000 And &HFFFF&
    
    If r_lngHighOrder > Abs(&H8000&) Then
        r_lngHighOrder = (&HFFFF& - r_lngHighOrder) * -1
    End If

    Exit Sub
    
ErrHandler:
    With err
        m_ErrNum = err.Number
        m_ErrDesc = err.Description
        m_ErrSource = "GetLowWord"
    End With
    
    On Error GoTo 0
    err.Raise m_ErrNum, m_ErrSource, m_ErrDesc
    
End Sub

