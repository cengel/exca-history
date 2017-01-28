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
        DoCmd.GoToControl "txtUnitNumber"
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
Dim UFeature, USpace, UBuilding, UIntCat, UDataCat, UDim, UCatSpecific, UStrat, USkelSame, UGrap, USamp, UXfind
Dim retVal, msg, msg1
retVal = MsgBox("You have selected to delete Unit number: " & Me![txtUnitNumber] & ". The system will now check what additional data exists for this Unit and will prompt you again before deleting it." & Chr(13) & Chr(13) & "Are you sure you want to continue?", vbCritical + vbYesNo, "Confirm Action")
If retVal = vbYes Then
    UFeature = AdminDeletionCheck("Exca: Units in Features", "Unit", Me![txtUnitNumber], "Related to Feature", "In_Feature")
    USpace = AdminDeletionCheck("Exca: Units in Spaces", "Unit", Me![txtUnitNumber], "Related to Space", "In_Space")
    UBuilding = AdminDeletionCheck("Exca: Units in Buildings", "Unit", Me![txtUnitNumber], "Related to Building", "In_Building")
    UIntCat = AdminDeletionCheck("Exca: Unit Interpretive Categories", "Unit Number", Me![txtUnitNumber], "Interpretive Categories", "Interpretive Category")
    UDataCat = AdminDeletionCheck("Exca: Unit Data Categories", "Unit Number", Me![txtUnitNumber], "Data Categories", "Data Category")
    UDim = AdminDeletionCheck("Exca: Dimensions", "Unit Number", Me![txtUnitNumber], "Dimensions", "Length")
    If LCase(Me![txtCategory]) = "skeleton" Then
        UCatSpecific = AdminDeletionCheck("Exca: skeleton data", "Unit Number", Me![txtUnitNumber], "Skeleton", "Target A - X")
        USkelSame = AdminDeletionCheck("Exca: skeletons same as", "skell_unit", Me![txtUnitNumber], "Skeleton", "To_Unit")
        USkelSame = USkelSame & AdminDeletionCheck("Exca: skeletons same as", "to_unit", Me![txtUnitNumber], "Skeleton related", "To_Unit")
    ElseIf LCase(Me![txtCategory]) = "cut" Then
        UCatSpecific = AdminDeletionCheck("Exca: descriptions cut", "Unit Number", Me![txtUnitNumber], "Cut Description", "Shape")
    Else
        UCatSpecific = AdminDeletionCheck("Exca: descriptions layer", "Unit Number", Me![txtUnitNumber], "Description", "Consistency")
    End If
    UStrat = AdminDeletionCheck("Exca: stratigraphy", "Unit", Me![txtUnitNumber], "Stratigraphy", "To_Units")
    UStrat = UStrat & AdminDeletionCheck("Exca: stratigraphy", "To_Units", Me![txtUnitNumber], "Stratigraphy", "Unit")
    UGrap = AdminDeletionCheck("Exca: graphics list", "Unit", Me![txtUnitNumber], "Graphics", "Type")
    USamp = AdminDeletionCheck("Exca: samples", "Unit number", Me![txtUnitNumber], "Samples", "Sample Number")
    UXfind = AdminDeletionCheck("Exca: X-Finds: Basic data", "Unit number", Me![txtUnitNumber], "X Finds", "GID Number")
    If UFeature <> "" Then msg = msg & UFeature & "; "
    If USpace <> "" Then msg = msg & USpace & "; "
    If UBuilding <> "" Then msg = msg & UBuilding & "; "
    If UIntCat <> "" Then msg = msg & UIntCat & "; "
    If UDataCat <> "" Then msg = msg & UDataCat & "; "
    If UDim <> "" Then msg = msg & UDim & "; "
    If UCatSpecific <> "" Then msg = msg & UCatSpecific & "; "
    If UStrat <> "" Then msg = msg & UStrat & "; "
    If LCase(Me![txtCategory]) = "skeleton" Then
        If USkelSame <> "" Then msg = msg & USkelSame & "; "
    End If
    If UGrap <> "" Then msg = msg & UGrap & "; "
    If USamp <> "" Then msg = msg & USamp & "; "
    If UXfind <> "" Then msg = msg & UXfind & "; "
    If msg = "" Then
        msg = "This Unit can safely be deleted."
    Else
        msg1 = "This Unit has the following relationships that will also be removed by the deletion - " & Chr(13) & Chr(13)
        msg = msg1 & msg
    End If
    msg = msg & Chr(13) & Chr(13) & "Are you quite sure that you want to permanently delete Unit " & Me![txtUnitNumber] & "?"
    retVal = MsgBox(msg, vbCritical + vbYesNoCancel, "Confirm Permanent Deletion")
    If retVal = vbYes Then
        On Error Resume Next
        Dim mydb As DAO.Database, wrkdefault As Workspace
        Set wrkdefault = DBEngine.Workspaces(0)
        Set mydb = CurrentDb
        wrkdefault.BeginTrans
        If UFeature <> "" Then Call DeleteARecord("Exca: Units in Features", "Unit", Me![txtUnitNumber], False, mydb)
        If USpace <> "" Then Call DeleteARecord("Exca: Units in Spaces", "Unit", Me![txtUnitNumber], False, mydb)
        If UBuilding <> "" Then Call DeleteARecord("Exca: Units in Buildings", "Unit", Me![txtUnitNumber], False, mydb)
        If UIntCat <> "" Then Call DeleteARecord("Exca: Unit Interpretive Categories", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UDataCat <> "" Then Call DeleteARecord("Exca: Unit Data Categories", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UDim <> "" Then Call DeleteARecord("Exca: Dimensions", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UCatSpecific <> "" Then
            If LCase(Me![txtCategory]) = "skeleton" Then
                Call DeleteARecord("Exca: Skeleton data", "Unit Number", Me![txtUnitNumber], False, mydb)
                Call DeleteARecord("Exca: skeletons same as", "skell_unit", Me![txtUnitNumber], False, mydb)
                Call DeleteARecord("Exca: skeletons same as", "to_unit", Me![txtUnitNumber], False, mydb)
            ElseIf LCase(Me![txtCategory]) = "cut" Then
                Call DeleteARecord("Exca: description cut", "Unit Number", Me![txtUnitNumber], False, mydb)
            Else
                Call DeleteARecord("Exca: descriptions layer", "Unit Number", Me![txtUnitNumber], False, mydb)
            End If
        End If
        If UStrat <> "" Then
            Call DeleteARecord("Exca: stratigraphy", "Unit", Me![txtUnitNumber], False, mydb)
            Call DeleteARecord("Exca: stratigraphy", "to_Units", Me![txtUnitNumber], True, mydb)
        End If
        If UGrap <> "" Then Call DeleteARecord("Exca: graphics list", "Unit", Me![txtUnitNumber], False, mydb)
        If USamp <> "" Then Call DeleteARecord("Exca: samples", "Unit Number", Me![txtUnitNumber], False, mydb)
        If UXfind <> "" Then Call DeleteARecord("Exca: X-Finds: Basic data", "Unit Number", Me![txtUnitNumber], False, mydb)
        Call DeleteARecord("Exca: Unit Sheet", "Unit Number", Me![txtUnitNumber], False, mydb)
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
