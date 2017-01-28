Option Compare Database
Option Explicit
Private Sub cboFindPriority_AfterUpdate()
On Error GoTo err_cboFindPriority_AfterUpdate
    If Me![cboFindPriority] <> "" Then
        If Me![Unit Number].Enabled = False Then Me![Unit Number].Enabled = True
        DoCmd.GoToControl "Unit Number"
        DoCmd.FindRecord Me![cboFindPriority]
        If Me![Short Description].Enabled = False Then Me![Short Description].Enabled = True
        DoCmd.GoToControl "Short Description"
        Me![Unit Number].Enabled = False
        Me![cboFindPriority] = ""
    End If
Exit Sub
err_cboFindPriority_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdClose_Click()
On Error GoTo Err_close_Click
    DoCmd.Close acForm, "Exca: Priority Detail", acSaveYes
Exit_close_Click:
    Exit Sub
Err_close_Click:
    Call General_Error_Trap
    Resume Exit_close_Click
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Then
        Me![Priority].Locked = False
        Me![Priority].Enabled = True
        Me![Priority].BackColor = 16777215
        Me![Discussion].Locked = False
        Me![Discussion].Enabled = True
        Me![Discussion].BackColor = 16777215
        Me![Short Description].Locked = False
        Me![Short Description].Enabled = True
        Me![Short Description].BackColor = 16777215
    Else
        Me![Priority].Locked = True
        Me![Priority].Enabled = False
        Me![Priority].BackColor = Me.Section(0).BackColor
        Me![Discussion].Locked = True
        Me![Discussion].Enabled = False
        Me![Discussion].BackColor = Me.Section(0).BackColor
        Me![Short Description].Locked = True
        Me![Short Description].Enabled = False
        Me![Short Description].BackColor = Me.Section(0).BackColor
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Sub prevrep_Click()
On Error GoTo Err_prevrep_Click
    Dim stDocName As String
    stDocName = "Exca: Priority Units"
    DoCmd.OpenReport stDocName, acPreview
Exit_prevrep_Click:
    Exit Sub
Err_prevrep_Click:
    Call General_Error_Trap
    Resume Exit_prevrep_Click
End Sub
