Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.RunCommand acCmdRecordsGoToNew
Exit Sub
err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFind_Click()
On Error GoTo err_cboFind
    If Me![cboFind] <> "" Then
        DoCmd.GoToControl "txtBuildingNumber"
        DoCmd.FindRecord Me![cboFind]
    End If
Exit Sub
err_cboFind:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Excavation_Click()
    DoCmd.Close acForm, Me.Name
End Sub
Private Sub cmdEdit_Click()
On Error GoTo Err_cmdEdit_Click
Dim checkValidAction, checkValidAction2, checkValidAction3, retVal
    checkValidAction = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "Level", "edit")
    If checkValidAction = "ok" Then
        checkValidAction2 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelStart", "edit")
        If checkValidAction2 = "ok" Then
            checkValidAction3 = CheckIfLOVValueUsed("Exca:LevelLOV", "Level", Me![txtLevel], "Exca: Space Sheet", "Space Number", "UncertainLevelEnd", "edit")
            If checkValidAction3 = "ok" Then
                retVal = InputBox("No records refer to this Level (" & Me![txtLevel] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Level that you wish to replace this entry with:", "Enter edited Level")
                If retVal <> "" Then
                    Me![txtLevel] = retVal
                End If
            ElseIf checkValidAction3 = "fail" Then
                MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
            Else
                MsgBox checkValidAction3, vbExclamation, "Action Report"
            End If
        ElseIf checkValidAction2 = "fail" Then
            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
        Else
            MsgBox checkValidAction2, vbExclamation, "Action Report"
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
Dim BUnit, BFeature
Dim retVal, msg, msg1
retVal = MsgBox("You have selected to delete Building number: " & Me![txtBuildingNumber] & ". The system will now check what additional data exists for this Building and will prompt you again before deleting it." & Chr(13) & Chr(13) & "Are you sure you want to continue?", vbCritical + vbYesNo, "Confirm Action")
If retVal = vbYes Then
    BUnit = AdminDeletionCheck("Exca: Units in Buildings", "In_Building", Me![txtBuildingNumber], "Related to Unit", "Unit")
    BFeature = AdminDeletionCheck("Exca: Features in Buildings", "In_Building", Me![txtBuildingNumber], "Related to Feature", "Feature")
    If BUnit <> "" Then msg = msg & BUnit & "; "
    If BFeature <> "" Then msg = msg & BFeature & "; "
    If msg = "" Then
        msg = "This Building can safely be deleted."
    Else
        msg1 = "This Building has the following relationships that will also be removed by the deletion - " & Chr(13) & Chr(13)
        msg = msg1 & msg
    End If
    msg = msg & Chr(13) & Chr(13) & "Are you quite sure that you want to permanently delete Building " & Me![txtBuildingNumber] & "?"
    retVal = MsgBox(msg, vbCritical + vbYesNoCancel, "Confirm Permanent Deletion")
    If retVal = vbYes Then
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        wrkdefault.BeginTrans
        If BUnit <> "" Then Call DeleteARecord("Exca: Units in Buildings", "In_Building", Me![txtBuildingNumber], False, mydb)
        If BFeature <> "" Then Call DeleteARecord("Exca: Features in Buildings", "In_Building", Me![txtBuildingNumber], False, mydb)
        Call DeleteARecord("Exca: Building Details", "Number", Me![txtBuildingNumber], False, mydb)
        If Err.Number = 0 Then
            wrkdefault.CommitTrans
            MsgBox "Deletion has been successful"
            Me.Requery
            Me![cboFind].Requery
        Else
            wrkdefault.Rollback
            MsgBox "A problem has occured and the deletion has been cancelled. The error message is: " & Err.Description
        End If
        mydb.Close
        Set mydb = Nothing
        wrkdefault.Close
        Set wrkdefault = Nothing
    Else
        MsgBox "Deletion cancelled", vbInformation, "Action Cancelled"
    End If
End If
Exit_cmdDelete_Click:
    Exit Sub
Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Delete(Cancel As Integer)
Call cmdDelete_Click
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
