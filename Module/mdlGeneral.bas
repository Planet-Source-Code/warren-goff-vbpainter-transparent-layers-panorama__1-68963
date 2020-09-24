Attribute VB_Name = "mdlGeneral"
'*******************************************************************************
'** File Name  : mdlGeneral.bas                                               **
'** Language   : Visual Basic 6.0                                             **
'** Reference  : Microsoft Scripting Runtime (only for ForceSave sub)         **
'** Components : -                                                            **
'** Modules    : -                                                            **
'** Developer  : Theo Zacharias (theo_yz@yahoo.com)                           **
'** Description: A modul to handle other public operations                    **
'** Last modified on August 15, 2003                                          **
'*******************************************************************************

Option Explicit

Public Enum enmError
  conErrWrite = 1
  conErrPrint = 2
  conErrReadImage = 3
  conErrDrawing = 4
  conErrPermission = 70
  conErrCancel = 32755
  conErrOthers = 0
End Enum

' Purpose    : Show error message intErr
' Assumptions: -
' Effect     : The error message has just been showed
' Inputs     : * intErr (error number)
'              * strMessage (for intErr = conErrOthers)
' Returns    : -
Public Sub ShowErrMessage(intErr As enmError, Optional strErrMessage As String)
Exit Sub
  Select Case intErr
    Case conErrWrite
      MsgBox "Cannot write to the disk." & vbNewLine & vbNewLine & _
               "Make sure the disk is not full or write-protected.", _
             vbOKOnly + vbCritical
    Case conErrPrint
      MsgBox "Cannot print the file." & vbNewLine & vbNewLine & _
               "Make sure the print is ready.", vbOKOnly + vbCritical
    Case conErrReadImage
      MsgBox "Cannot open the file." & vbNewLine & vbNewLine & _
               "The file may be corrupt or not a valid picture file.", _
             vbOKOnly + vbCritical
    Case conErrDrawing
      MsgBox "Cannot drawing using the selected tool." & _
               vbNewLine & vbNewLine & _
               "The needed file may be missing.", _
             vbOKOnly + vbCritical
    Case conErrOthers
      MsgBox strErrMessage, vbOKOnly + vbCritical
  End Select
End Sub


' Purpose    : Return varTrue if blnCondition = true, or varFalse otherwise
' Assumptions: -
' Effects    : -
' Inputs     : blnCondition, varTrue, varFalse
' Returns    : As specified
Public Function varIIf(blnCondition As Boolean, _
                        varTrue As Variant, varFalse As Variant) As Variant
  'On Error GoTo ErrorHandler
  
  If blnCondition Then
    varIIf = varTrue
  Else
    varIIf = varFalse
  End If
  Exit Function

ErrorHandler:
  ShowErrMessage intErr:=conErrOthers, strErrMessage:=Err.Description
End Function
