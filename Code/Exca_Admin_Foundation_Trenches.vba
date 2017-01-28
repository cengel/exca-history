Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cboFindFT_AfterUpdate()
On Error GoTo err_cboFindFT_AfterUpdate
    If Me![cboFindFT] <> "" Then
        If Me![FTName].Enabled = False Then Me![FTName].Enabled = True
        DoCmd.GoToControl "FTName"
        DoCmd.FindRecord Me![cboFindFT]
        Me![cboFindFT] = ""
        DoCmd.GoToControl "cboFindFT"
    End If
Exit Sub
err_cboFindFT_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Excavation"
    DoCmd.Close acForm, Me.Name
Exit Sub
err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Current()
If Me![FTName] <> "" Then
        Me![FTName].Locked = True
        Me![FTName].Enabled = False
        Me![FTName].BackColor = Me.Section(0).BackColor
    Else
        Me![FTName].Locked = False
        Me![FTName].Enabled = True
        Me![FTName].BackColor = 16777215
        Me![FTName].SetFocus
    End If
    If Me![LevelCertain] = True Then
        Me![Level].Enabled = True
        Me![cboUncertainLevelStart].Enabled = False
        Me![cboUnCertainLevelEnd].Enabled = False
    Else
        Me![Level].Enabled = False
        Me![cboUncertainLevelStart].Enabled = True
        Me![cboUnCertainLevelEnd].Enabled = True
    End If
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
If Me.FilterOn = True Or Me.AllowEdits = False Then
    Me![cboFindFT].Enabled = False
    Me.AllowAdditions = False
Else
    DoCmd.GoToControl "cboFindFT"
End If
Dim permiss
permiss = GetGeneralPermissions
If (permiss = "ADMIN") And (Me.AllowAdditions = True Or Me.AllowDeletions = True Or Me.AllowEdits = True) Then
    ToggleFormReadOnly Me, False, "NoDeletions"
Else
    ToggleFormReadOnly Me, True
End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub frmLevelCertain_AfterUpdate()
On Error GoTo err_frmLevelCertain_AfterUpdate
Dim retval
If Me![frmLevelCertain] = -1 Then
    If Me![cboUncertainLevelStart] <> "" And Me![cboUnCertainLevelEnd] <> "" Then
        retval = MsgBox("Do you want the Start Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then
            Me![Level] = Me![cboUncertainLevelStart]
        Else
            retval = MsgBox("Do you want the End Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
            If retval = vbYes Then
                Me![Level] = Me![cboUnCertainLevelEnd]
            Else
                retval = MsgBox("The start and end level fields will now be cleared and you will have to select the Certain level from that list. Are you sure you want to continue?", vbQuestion + vbYesNo, "Uncertain Levels will be cleared")
                If retval = vbYes Then
                    Me![cboUncertainLevelStart] = ""
                    Me![cboUnCertainLevelEnd] = ""
                Else
                    Me![frmLevelCertain] = 0
                End If
            End If
        End If
    ElseIf Me![cboUncertainLevelStart] <> "" Then
        retval = MsgBox("Do you want the Start Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then Me![Level] = Me![cboUncertainLevelStart]
        Me![cboUncertainLevelStart] = ""
    ElseIf Me![cboUnCertainLevelEnd] <> "" Then
        retval = MsgBox("Do you want the End Level to become the certain level for this FT?", vbQuestion + vbYesNo, "Set Level")
        If retval = vbYes Then Me![Level] = Me![cboUnCertainLevelEnd]
        Me![cboUnCertainLevelEnd] = ""
    End If
    If Me![frmLevelCertain] = -1 Then 'they have decide not to change their mind
        Me![Level].Enabled = True
        Me![cboUncertainLevelStart].Enabled = False
        Me![cboUnCertainLevelEnd].Enabled = False
    End If
Else
    Me![Level].Enabled = False
    If Me![Level] <> "" Then
        Me![cboUncertainLevelStart] = Me![Level]
        Me![Level] = ""
    End If
    Me![cboUncertainLevelStart].Enabled = True
    Me![cboUnCertainLevelEnd].Enabled = True
End If
Exit Sub
err_frmLevelCertain_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub FTName_AfterUpdate()
On Error GoTo err_FTNAME
Dim resp
If Not IsNull(Me!FTName.OldValue) Then
    resp = DLookup("[FoundationTrench]", "[Exca: Unit Sheet]", "[FoundationTrench] = '" & Me!FTName.OldValue & "' AND [Area] = '" & Me![cboArea] & "'")
    If Not IsNull(resp) Then
        MsgBox "This FT is assocated with a Unit so the name cannot be altered. Please enter this change as a new FT name and then re-allocate the units to the new record", vbExclamation, "Changed Cancelled"
        Me!FTName = Me!FTName.OldValue
    End If
End If
Exit Sub
err_FTNAME:
    Call General_Error_Trap
    Exit Sub
End Sub
