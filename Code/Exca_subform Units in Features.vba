Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
End Sub
Private Sub Unit_AfterUpdate()
End Sub
Sub Command5_Click()
End Sub
Sub go_to_unit_Click()
On Error GoTo Err_go_to_unit_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Unit Sheet"
    stLinkCriteria = "[Unit Number]=" & Me![Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormReadOnly
    Exit Sub
Err_go_to_unit_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
