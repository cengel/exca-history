Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_Form_BeforeUpdate
    Me![Date changed] = Now()
Exit Sub
err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Unit_AfterUpdate()
On Error GoTo err_Unit_AfterUpdate
Dim checknum, msg, retVal, checknum2
If Me![Unit] <> "" Then
    If IsNumeric(Me![Unit]) Then
        checknum = DLookup("[Unit Number]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![Unit])
        If IsNull(checknum) Then
            msg = "This Unit Number DOES NOT EXIST in the database, it cannot be used here until it has been entered."
            MsgBox msg, vbInformation, "Unit Number does not exist"
            If Not IsNull(Me![Unit].OldValue) Then
                Me![Unit] = Me![Unit].OldValue
            Else
                Me.Undo
            End If
            DoCmd.GoToControl "Unit"
        Else
            checknum2 = DLookup("[Data Category]", "[Exca: Unit Data Categories]", "[Unit Number] = " & Me![Unit])
                If Not IsNull(checknum2) Then 'there is a space for this related feature
                    If UCase(checknum2) <> "FLOORS (USE)" And UCase(checknum2) <> "CONSTRUCTION/MAKE-UP/PACKING" Then
                        msg = "This entry is not allowed:  Unit (" & Me![Unit] & ")"
                        msg = msg & " has the data category " & checknum2 & ", only Units with the category 'Floor(use)' or 'construction/make-up/packing' are valid here."
                        msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please double check this issue with your Supervisor."
                        MsgBox msg, vbExclamation, "Data Category problem"
                        If Not IsNull(Me![Unit].OldValue) Then
                            Me![Unit] = Me![Unit].OldValue
                        Else
                            Me.Undo
                        End If
                        DoCmd.GoToControl "Unit"
                    End If
                Else
                    msg = "This entry is not allowed as Unit (" & Me![Unit] & ")"
                    msg = msg & " has no data category entered" & checknum2 & ", only Units with the category 'Floor(use)' or 'construction/make-up/packing' are valid here."
                    msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please update the Unit record first."
                    MsgBox msg, vbExclamation, "No Data Category"
                    If Not IsNull(Me![Unit].OldValue) Then
                        Me![Unit] = Me![Unit].OldValue
                    Else
                        Me.Undo
                    End If
                    DoCmd.GoToControl "Unit"
                End If
        End If
    Else
        MsgBox "The Unit number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If
Exit Sub
err_Unit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Sub Command5_Click()
End Sub
Sub go_to_unit_Click()
On Error GoTo Err_go_to_unit_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Unit Sheet"
    If Me![Unit] <> "" Then
        stLinkCriteria = "[Unit Number]=" & Me![Unit]
        DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormReadOnly
    Else
        MsgBox "No Unit number to show", vbInformation, "No Unit Number"
    End If
Exit_go_to_unit_Click:
    Exit Sub
Err_go_to_unit_Click:
    Call General_Error_Trap
    Resume Exit_go_to_unit_Click
End Sub
