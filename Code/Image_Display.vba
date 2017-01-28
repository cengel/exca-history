Option Compare Database
Option Explicit
Private Sub Form_Current()
On Error GoTo err_Current
Dim newStr, newstr2
newStr = Replace(Me![Path], ":", "\")
newstr2 = newStr
Me!Image145.Picture = newstr2
Exit Sub
err_Current:
    If Err.Number = 2220 Then
        If Dir(newstr2) = "" Then
            MsgBox "The directory where images are supposed to be stored cannot be found. Please contact the database administrator"
        Else
            MsgBox "The image file cannot be found - check the file exists"
            DoCmd.GoToControl "txtSketch"
        End If
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub
