Attribute VB_Name = "MApi"
Option Explicit
Private Const VK_MENU = &H12
Private Const VK_SNAPSHOT = &H2C
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long


Sub TimerProc(ByVal hwnd As Long)
  Dim sPic As IPictureDisp
   On Error GoTo Errhandler
   frmPaint.Picture2.Picture = Clipboard.GetData(0)
   Clipboard.Clear
   keybd_event VK_MENU, 0, 0, 0
   DoEvents
   keybd_event VK_SNAPSHOT, 1, 0, 0
   DoEvents
   keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
   DoEvents
   'Image is now in clipboard, save it
   Set sPic = Clipboard.GetData(0)
   SavePicture sPic, App.path & "\Merged.bmp"
   Clipboard.Clear
   Clipboard.SetData frmPaint.Picture2.Picture
   Set sPic = Nothing
   Exit Sub
Errhandler:
'exit the sub for now
'trowhing away the error
Resume Exit_from_Here
Exit_from_Here:
End Sub
 Private Function retrieveFilePath() As String
   Dim thefileName As String
   
   'path= application working path
   retrieveFilePath = App.path
   'be sure path ends with a "\"
   If Right(retrieveFilePath, 1) <> "\" Then
      retrieveFilePath = retrieveFilePath & "\"
   End If
   'now find a valid filename
   thefileName = Now()
   thefileName = Replace(thefileName, "/", "_")
   thefileName = Replace(thefileName, "\", "_")
   thefileName = Replace(thefileName, ":", "_")
   thefileName = Replace(thefileName, ".", "_")
   'add an extension
   thefileName = thefileName & ".bmp"
   retrieveFilePath = retrieveFilePath & thefileName
 End Function

