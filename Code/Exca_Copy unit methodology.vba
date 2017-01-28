Option Compare Database
Option Explicit
Private Sub cboFindUnitToCopy_AfterUpdate()
On Error GoTo err_cboFindUnitToCopy_AfterUpdate
If Me![cboFindUnitToCopy] <> "" Then
    Me.RecordSource = "SELECT * FROM [Exca: Unit Sheet] WHERE [Unit Number] = " & Me![cboFindUnitToCopy]
    Me![Unit Number].ControlSource = "Unit Number"
    Me![Recognition].ControlSource = "Recognition"
    Me![Definition].ControlSource = "Definition"
    Me![Execution].ControlSource = "Execution"
    Me![Condition].ControlSource = "Condition"
    Me![copy data].Enabled = True
End If
Exit Sub
err_cboFindUnitToCopy_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub copy_data_Click()
On Error GoTo Err_copy_data_Click
Dim msg, Style, Title, Response
msg = "This action will replace the unit sheet (" & Me![Text17] & ") "
msg = msg & "fields: Recognition, Definition, Execution and Definition with those of unit " & Me![Unit Number] & " shown here." & Chr(13) & Chr(13)
msg = msg & "This action cannot be undone. Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
Title = "Overwriting Records"  ' Define title.
Response = MsgBox(msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
    Forms![Exca: Unit Sheet]![Recognition] = Me![Recognition]
    Forms![Exca: Unit Sheet]![Definition] = Me![Definition]
    Forms![Exca: Unit Sheet]![Execution] = Me![Execution]
    Forms![Exca: Unit Sheet]![Condition] = Me![Condition]
Else    ' User chose No.
End If
Exit_copy_data_Click:
    Exit Sub
Err_copy_data_Click:
    Call General_Error_Trap
    Resume Exit_copy_data_Click
End Sub
Sub close_Click()
On Error GoTo Err_close_Click
    DoCmd.Close acForm, "Exca: copy unit methodology"
Exit_close_Click:
    Exit Sub
Err_close_Click:
    Call General_Error_Trap
    Resume Exit_close_Click
End Sub
