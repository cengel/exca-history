Option Compare Database
Option Explicit
Sub close_Click()
On Error GoTo Err_close_Click
    DoCmd.Close
Exit_close_Click:
    Exit Sub
Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
End Sub
Sub prevrep_Click()
On Error GoTo Err_prevrep_Click
    Dim stDocName As String
    stDocName = "Exca: Priority Units"
    DoCmd.OpenReport stDocName, acPreview
Exit_prevrep_Click:
    Exit Sub
Err_prevrep_Click:
    MsgBox Err.Description
    Resume Exit_prevrep_Click
End Sub
