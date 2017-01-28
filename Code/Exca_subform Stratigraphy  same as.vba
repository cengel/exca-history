Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
