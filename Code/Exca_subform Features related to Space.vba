Option Compare Database
Option Explicit
Private Sub cmdGoToFeature_Click()
On Error GoTo Err_cmdGoToFeature_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Feature Sheet"
    stLinkCriteria = "[Feature Number]= " & Me![Feature Number]
    DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly ', acDialog
    Exit Sub
Err_cmdGoToFeature_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
