Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    Dim resp
    resp = InputBox("Please enter the name of the new Hodder Phase:", "New Phase")
    If resp <> "" Then
        Me.AllowAdditions = True
        DoCmd.RunCommand acCmdRecordsGoToNew
        Me![txtHPhase] = resp
        DoCmd.GoToControl "txtDisplayOrder"
        Me.AllowAdditions = False
    End If
Exit Sub
err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Excavation_Click()
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub cmdEdit_Click()
On Error GoTo Err_cmdEdit_Click
Dim checkValidAction, checkValidAction2, checkValidAction3, retval
    checkValidAction = CheckIfLOVValueUsed("Exca:HodderPhaseLOV", "HodderPhase", Me![txtHPhase], "Exca: Space Sheet", "Space Number", "HodderPhase", "edit")
    If checkValidAction = "ok" Then
                retval = InputBox("No records refer to this Phase (" & Me![txtHPhase] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Phase that you wish to replace this entry with:", "Enter edited Phase")
                If retval <> "" Then
                    Me![txtHPhase] = retval
                End If
    ElseIf checkValidAction = "fail" Then
        MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
    Else
        MsgBox checkValidAction, vbExclamation, "Action Report"
    End If
Exit_cmdEdit_Click:
    Exit Sub
Err_cmdEdit_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdDelete_Click()
On Error GoTo Err_cmdDelete_Click
Dim checkValidAction, checkValidAction2, checkValidAction3, retval
    checkValidAction = CheckIfLOVValueUsed("Exca:HodderPhaseLOV", "HodderPhase", Me![txtHPhase], "Exca: Space Sheet", "Space Number", "HodderPhase", "delete")
    If checkValidAction = "ok" Then
                retval = MsgBox("No records refer to this Phase (" & Me![txtHPhase] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtHPhase] & " from the list of available phases?", vbExclamation + vbYesNo, "Confirm Deletion")
                If retval = vbYes Then
                    Me.AllowDeletions = True
                    DoCmd.RunCommand acCmdDeleteRecord
                    Me.AllowDeletions = False
                End If
    ElseIf checkValidAction = "fail" Then
        MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
    Else
        MsgBox checkValidAction, vbExclamation, "Action Report"
    End If
Exit_cmdDelete_Click:
    Exit Sub
Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss <> "ADMIN" Then
        MsgBox "Sorry but only Administrators have access to this form"
        DoCmd.Close acForm, Me.Name
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
