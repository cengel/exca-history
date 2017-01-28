Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub Button13_Click()
interpret_Click
End Sub
Private Sub cmdClose_Click()
On Error GoTo err_cmdClose_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Excavation:ListsMenu"
Exit Sub
err_cmdClose_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Command18_Click()
Feature_types_Click
End Sub
Private Sub Command24_Click()
cmdClose_Click
End Sub
Sub Feature_types_Click()
On Error GoTo Err_Feature_types_Click
    Dim stDocName As String
    stDocName = "Exca: Feature Types"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
Exit_Feature_types_Click:
    Exit Sub
Err_Feature_types_Click:
    Call General_Error_Trap
    Resume Exit_Feature_types_Click
End Sub
Sub interpret_Click()
On Error GoTo Err_interpret_Click
    Dim stDocName As String
    stDocName = "Exca: List Interpretive Categories"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
Exit_interpret_Click:
    Exit Sub
Err_interpret_Click:
     Call General_Error_Trap
    Resume Exit_interpret_Click
End Sub
