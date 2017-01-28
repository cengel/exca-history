Option Compare Database
Option Explicit
Private Sub close_Click()
On Error GoTo Err_close_Click
    DoCmd.Close
Exit_close_Click:
    Exit Sub
Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
End Sub
