Option Compare Database
Option Explicit
Private Sub cmdGoToSpace_Click()
On Error GoTo Err_cmdGoToSpace_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Space Sheet"
    stLinkCriteria = "[Space Number]= " & Me![Space number]
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly ', acDialog
    Exit Sub
Err_cmdGoToSpace_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
