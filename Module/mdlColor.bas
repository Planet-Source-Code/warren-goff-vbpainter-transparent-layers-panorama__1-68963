Attribute VB_Name = "mdlColor"
Option Explicit
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type

'Author: Andrew Gray
'Date: 9/10/2001 8:21:44 PM
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=927
Public Type HSL
    Hue As Integer
    Saturation As Integer
    Luminance As Integer
End Type

'Author: Andrew Gray
'Date: 9/10/2001 8:18:23 PM
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=926
Public Type RGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Public Const HueMAX = 239, SatMAX = 240, LumMAX = 240

Public Function GetColorAtCursor() As Long
Dim P As POINTAPI, C As Long
    GetCursorPos P
    C = GetPixel(GetWindowDC(GetDesktopWindow), P.X, P.Y)
    GetColorAtCursor = C
End Function

Public Function HSL(ByVal Hue As Integer, _
                         ByVal Saturation As Integer, _
                         ByVal Luminance As Integer) As Long
Dim RGBis As RGB
    RGBis = HSLtoRGB(Hue, Saturation, Luminance)
    If RGBis.Red < 0 Then RGBis.Red = 0
    If RGBis.Green < 0 Then RGBis.Green = 0
    If RGBis.Blue < 0 Then RGBis.Blue = 0
    HSL = RGB(RGBis.Red, RGBis.Green, RGBis.Blue)
End Function

Public Function HSLtoRGB(ByVal Hue As Integer, _
                         ByVal Saturation As Integer, _
                         ByVal Luminance As Integer) As RGB
'Author: Andrew Gray
'Date: 9/10/2001 8:18:23 PM
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=926
'Modified by SB

    Dim pHue As Single
    Dim pSat As Single
    Dim pLum As Single
    Dim RetVal As RGB
    Dim pRed As Single
    Dim pGreen As Single
    Dim pBlue As Single
    Dim temp2 As Single
    Dim temp3() As Single
    Dim temp1 As Single
    Dim N As Integer

   ReDim temp3(0 To 2)
   
   pHue = Hue / HueMAX '239
   pSat = Saturation / SatMAX '239
   pLum = Luminance / LumMAX '239

   If pSat = 0 Then
      pRed = pLum!
      pGreen = pLum
      pBlue = pLum
   Else
      If pLum < 0.5 Then
         temp2 = pLum * (1 + pSat)
      Else
         temp2 = pLum + pSat - pLum * pSat
      End If
      temp1! = 2 * pLum! - temp2!
   
      temp3(0) = pHue + 1 / 3
      temp3(1) = pHue
      temp3(2) = pHue - 1 / 3
      
      For N = 0 To 2
         If temp3(N) < 0 Then temp3(N) = temp3(N) + 1
         If temp3(N) > 1 Then temp3(N) = temp3(N) - 1
      
         If 6 * temp3(N) < 1 Then
            temp3(N) = temp1 + (temp2 - temp1) * 6 * temp3(N)
         Else
            If 2 * temp3(N) < 1 Then
               temp3(N) = temp2
            Else
               If 3 * temp3(N%) < 2 Then
                  temp3(N%) = temp1 + (temp2 - temp1) _
                        * ((2 / 3) - temp3(N%)) * 6
               Else
                  temp3(N%) = temp1
                End If
             End If
          End If
       Next N%

       pRed = temp3(0)
       pGreen = temp3(1)
       pBlue = temp3(2)
    End If

    RetVal.Red = Int(pRed * 255)
    RetVal.Green = Int(pGreen * 255)
    RetVal.Blue = Int(pBlue * 255)
    
    HSLtoRGB = RetVal
End Function


Public Function RGBtoHSL(ByVal Red As Integer, _
                         ByVal Green As Integer, _
                         ByVal Blue As Integer) As HSL
'Author: Andrew Gray
'Date: 9/10/2001 8:21:44 PM
'Link: http://abstractvb.com/code.asp?F=50&P=1&A=927
'Modified by SB

    Dim pRed As Single
    Dim pGreen As Single
    Dim pBlue As Single
    Dim RetVal As HSL
    Dim pMax As Single
    Dim pMin As Single
    Dim pLum As Single
    Dim pSat As Single
    Dim pHue As Single
    
    pRed = Red / 255
    pGreen = Green / 255
    pBlue = Blue / 255
   
    If pRed > pGreen Then
       If pRed > pBlue Then
          pMax = pRed
       Else
          pMax = pBlue
       End If
    ElseIf pGreen > pBlue Then
        pMax = pGreen
    Else
        pMax = pBlue
    End If

    If pRed < pGreen Then
        If pRed < pBlue Then
            pMin = pRed
        Else
            pMin = pBlue
        End If
    ElseIf pGreen < pBlue Then
        pMin = pGreen
    Else
        pMin = pBlue
    End If

    pLum = (pMax + pMin) / 2
   
    If pMax = pMin Then
        pSat = 0
        pHue = 0
    Else
        If pLum < 0.5 Then
            pSat = (pMax - pMin) / (pMax + pMin)
        Else
            pSat = (pMax - pMin) / (2 - pMax - pMin)
        End If
        
        Select Case pMax!
            Case pRed
                pHue = (pGreen - pBlue) / (pMax - pMin)
            Case pGreen
                pHue = 2 + (pBlue - pRed) / (pMax - pMin)
            Case pBlue
                pHue = 4 + (pRed - pGreen) / (pMax - pMin)
        End Select
    End If

    RetVal.Hue = pHue * HueMAX \ 6
    If RetVal.Hue < 0 Then RetVal.Hue = RetVal.Hue + HueMAX + 1
    
    RetVal.Saturation = Int(pSat * SatMAX)
    RetVal.Luminance = Int(pLum * LumMAX)
    
    RGBtoHSL = RetVal
End Function

Public Sub DrawColorSquare(pctBox As PictureBox, XYZ As String, ZVal As Integer)
Dim I As Integer, J As Integer, LM As Integer
Dim X As Single, Y As Single, XMAX As Integer, YMAX As Integer
Dim StepLenX As Integer, StepLenY As Integer
Dim ColorStepLenX As Single, ColorStepLenY As Single
Dim LenMultX As Integer, LenMultY As Integer
    pctBox.Cls
    pctBox.Refresh
    Select Case XYZ
        Case "RGB", "RBG", "GRB", "GBR", "BRG", "BGR"
            XMAX = 255
            YMAX = 255
        Case "HSL"
            XMAX = HueMAX
            YMAX = SatMAX
        Case "HLS"
            XMAX = HueMAX
            YMAX = LumMAX
        Case "SHL"
            XMAX = SatMAX
            YMAX = HueMAX
        Case "SLH"
            XMAX = SatMAX
            YMAX = LumMAX
        Case "LHS"
            XMAX = LumMAX
            YMAX = HueMAX
        Case "LSH"
            XMAX = LumMAX
            YMAX = SatMAX
        Case Else
            Exit Sub
    End Select
    LM = ZVal
    
    X = 0
    Y = 0
    StepLenX = Screen.TwipsPerPixelX
    StepLenY = Screen.TwipsPerPixelY
    LenMultX = (pctBox.Width - 30) / XMAX
    'If LenMultX > StepLenX Then LenMultX = StepLenX
    LenMultY = (pctBox.Height - 30) / YMAX
    'If LenMultY > StepLenY Then LenMultY = StepLenY
    ColorStepLenX = StepLenX / LenMultX
    ColorStepLenY = StepLenY / LenMultY
    For I = 0 To (pctBox.Width - 30) Step StepLenX
        For J = 0 To (pctBox.Height - 30) Step StepLenY
            Select Case XYZ
                Case "RGB": pctBox.PSet (I, J), RGB(CInt(X), CInt(Y), ZVal)
                Case "RBG": pctBox.PSet (I, J), RGB(CInt(X), ZVal, CInt(Y))
                Case "GRB": pctBox.PSet (I, J), RGB(CInt(Y), CInt(X), ZVal)
                Case "GBR": pctBox.PSet (I, J), RGB(ZVal, CInt(X), CInt(Y))
                Case "BRG": pctBox.PSet (I, J), RGB(CInt(Y), ZVal, CInt(X))
                Case "BGR": pctBox.PSet (I, J), RGB(ZVal, CInt(Y), CInt(X))
                Case "HSL": pctBox.PSet (I, J), HSL(CInt(X), CInt(Y), ZVal)
                Case "HLS": pctBox.PSet (I, J), HSL(CInt(X), ZVal, CInt(Y))
                Case "SHL": pctBox.PSet (I, J), HSL(CInt(Y), CInt(X), ZVal)
                Case "SLH": pctBox.PSet (I, J), HSL(ZVal, CInt(X), CInt(Y))
                Case "LHS": pctBox.PSet (I, J), HSL(CInt(Y), ZVal, CInt(X))
                Case "LSH": pctBox.PSet (I, J), HSL(ZVal, CInt(Y), CInt(X))
            End Select
            Y = Y + ColorStepLenY
        Next
        X = X + ColorStepLenX
        Y = 0
    Next
End Sub


