Option Compare Database
Option Explicit
Sub copy_cut_Click()
On Error GoTo Err_copy_cut_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Copy Cut description"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_copy_cut_Click:
    Exit Sub
Err_copy_cut_Click:
    MsgBox Err.Description
    Resume Exit_copy_cut_Click
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
Dim permiss
permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
        Me![copy cut].Enabled = False
        If Me.AllowAdditions = False Then Me.AllowAdditions = True
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
