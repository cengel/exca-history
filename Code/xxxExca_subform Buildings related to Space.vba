Option Compare Database
Option Explicit
Private Sub cmdGoToBuilding_Click()
On Error GoTo Err_cmdGoToBuilding_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Building Sheet"
    stLinkCriteria = "[Number]= " & Me![Number]
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, acDialog
    Exit Sub
Err_cmdGoToBuilding_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
