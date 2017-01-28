Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
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
Sub open_skell_Click()
On Error GoTo Err_open_skell_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Skeleton Sheet"
    stLinkCriteria = "[Unit Number]=" & Me![To_Unit]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_open_skell_Click:
    Exit Sub
Err_open_skell_Click:
    MsgBox Err.Description
    Resume Exit_open_skell_Click
End Sub
Private Sub To_Unit_AfterUpdate()
On Error GoTo err_To_Unit_AfterUpdate
Dim checknum, msg, retVal, checknum2
If Me![To_Unit] <> "" Then
    If IsNumeric(Me![To_Unit]) Then
        checknum = DLookup("[Unit Number]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![To_Unit])
        If IsNull(checknum) Then
            msg = "This Unit Number DOES NOT EXIST in the database yet, please ensure it is entered soon."
            MsgBox msg, vbInformation, "Unit Number does not exist yet"
           DoCmd.GoToControl "To_Unit"
        Else
            checknum2 = DLookup("[Category]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![To_Unit])
                If Not IsNull(checknum2) Then 'category found this unit
                    If UCase(checknum2) <> "SKELETON" Then
                        msg = "This entry is not allowed:  Unit (" & Me![To_Unit] & ")"
                        msg = msg & " has the category " & checknum2 & ", only Units with the category 'Skeleton' are valid here."
                        msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please double check this issue with your Supervisor."
                        MsgBox msg, vbExclamation, "Category problem"
                        If Not IsNull(Me![To_Unit].OldValue) Then
                            Me![To_Unit] = Me![To_Unit].OldValue
                        Else
                            Me.Undo
                        End If
                        DoCmd.GoToControl "To_Unit"
                    End If
                Else
                    msg = "The Unit (" & Me![To_Unit] & ")"
                    msg = msg & " has no category entered yet. Please correct this as soon as possible"
                    MsgBox msg, vbInformation, "Category Missing"
                    DoCmd.GoToControl "To_Unit"
                End If
        End If
    Else
        MsgBox "The Unit number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If
Exit Sub
err_To_Unit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
