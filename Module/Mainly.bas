Attribute VB_Name = "Mainly"
Sub Main()
On Error Resume Next
If App.PrevInstance Then
    End
Else
    Load frmPaint
    frmPaint.Show
End If
End Sub

