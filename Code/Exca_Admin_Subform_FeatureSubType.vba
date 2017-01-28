Option Compare Database
Option Explicit
Private Sub cmdDelete_Click()
On Error GoTo Err_cmdDelete_Click
Dim checkValidAction, retVal
        checkValidAction = CheckIfLOVValueUsed("Exca:SubFeatureTypeLOV", "FeatureSubType", Me![txtFeatureSubType], "Exca: Features", "Feature Number", "FeatureSubType", "delete", " AND [Feature Type] = '" & Forms![Exca: Admin_FeatureTypeSubTypeLOV]![txtFeatureType] & "'")
        If checkValidAction = "ok" Then
            retVal = MsgBox("No records refer to this Feature SubType (" & Me![txtFeatureSubType] & ") so deletion is allowed." & Chr(13) & Chr(13) & "Are you sure you want to delete " & Me![txtFeatureSubType] & " from the list of available Feature Subtypes?", vbExclamation + vbYesNo, "Confirm Deletion")
            If retVal = vbYes Then
                Me.AllowDeletions = True
                DoCmd.RunCommand acCmdDeleteRecord
                Me.AllowDeletions = False
            End If
        ElseIf checkValidAction = "fail" Then
            MsgBox "Sorry but the system has been unable to check whether this value is used by any dependant tables, please contact the DBA", vbCritical, "Integrity Check Failed"
        Else
            If Not IsEmpty(checkValidAction) Then MsgBox checkValidAction, vbExclamation, "Action Report"
        End If
Exit_cmdDelete_Click:
    Exit Sub
Err_cmdDelete_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdEdit_Click()
On Error GoTo Err_cmdEdit_Click
Dim checkValidAction, retVal
    checkValidAction = CheckIfLOVValueUsed("Exca:SubFeatureTypeLOV", "FeatureSubType", Me![txtFeatureSubType], "Exca: Features", "Feature Number", "FeatureSubType", "edit", " AND [Feature Type] = '" & Forms![Exca: Admin_FeatureTypeSubTypeLOV]![txtFeatureType] & "'")
    If checkValidAction = "ok" Then
        retVal = InputBox("No records refer to this Feature sub type (" & Me![txtFeatureSubType] & ") so an edit is allowed." & Chr(13) & Chr(13) & "Please enter the edited Feature Sub Type that you wish to replace this entry with:", "Enter edited Feature Sub Type")
        If retVal <> "" Then
            Me![txtFeatureSubType] = retVal
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
Private Sub cmdNewSubType_Click()
On Error GoTo err_cmdNewSubType_Click
    If Forms![Exca: Admin_FeatureTypeSubTypeLOV]![FeatureTypeID] <> "" Then
        Dim sql, retVal
        retVal = InputBox("Please enter the new subtype for the feature type '" & Forms![Exca: Admin_FeatureTypeSubTypeLOV]![FeatureType] & "': ", "Enter new subtype")
        If retVal <> "" Then
            sql = "INSERT INTO [Exca:FeatureSubTypeLOV] ([FeatureTypeID], [FeatureSubType]) VALUES (" & Forms![Exca: Admin_FeatureTypeSubTypeLOV]![FeatureTypeID] & ", '" & retVal & "');"
            DoCmd.RunSQL sql
            Me.Requery
        End If
    Else
        MsgBox "Sorry not all the data necessary to make a new subtype is available.", vbExclamation, "Invalid Action"
    End If
Exit Sub
err_cmdNewSubType_Click:
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
