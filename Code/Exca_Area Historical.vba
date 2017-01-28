Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdClose_Click()
    Excavation_Click
End Sub
Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click
    DoCmd.Close acForm, "Exca: Area Historical"
Exit Sub
err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
