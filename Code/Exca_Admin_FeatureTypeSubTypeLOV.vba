Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cboFindFeature_AfterUpdate()
On Error GoTo err_find
    If Me![cboFindFeature] <> "" Then
        DoCmd.GoToControl "txtFeatureType"
        DoCmd.FindRecord Me![cboFindFeature]
    End If
    Me.AllowEdits = False
Exit Sub
err_find:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindFeature_GotFocus()
Me.AllowEdits = True
End Sub
Private Sub cboFindFeature_LostFocus()
Me.AllowEdits = False
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.RunCommand acCmdRecordsGoToNew
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
Dim checkValidAction, retval
    checkValidAction = CheckIfLOVValueUsed("Exca:FeatureTypeLOV", "FeatureType", Me![txtFeatureType], "Exca: Features", "Feature Number", "Feature Type", "edit")
    If checkValidAction = "ok" Then
        retval = InputBox("No records refer to this Feature Type (" & Me![txtFeatureType] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Feature Type that you wish to replace this entry with:", "Enter edited Feature Type")
        If retval <> "" Then
            Me![txtFeatureType] = retval
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
Dim checkValidAction, retval
    If Not IsNull(Me![Exca: Admin_Subform_FeatureSubType].Form![FeatureTypeID]) Then
        MsgBox "You must delete the Sub types associated with this feature first", vbInformation, "Invalid Action"
    Else
        checkValidAction = CheckIfLOVValueUsed("Exca:FeatureTypeLOV", "FeatureType", Me![txtFeatureType], "Exca: Features", "Feature Number", "Feature Type", "delete")
        If checkValidAction = "ok" Then
            retval = MsgBox("No records refer to this Feature Type (" & Me![txtFeatureType] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtFeatureType] & " from the list of available Feature Types?", vbExclamation + vbYesNo, "Confirm Deletion")
            If retval = vbYes Then
                Me.AllowDeletions = True
                DoCmd.RunCommand acCmdDeleteRecord
                Me.AllowDeletions = False
            End If
        ElseIf checkValidAction = "fail" Then
            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
        Else
            If Not IsEmpty(checkValidAction) Then MsgBox checkValidAction, vbExclamation, "Action Report"
        End If
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
