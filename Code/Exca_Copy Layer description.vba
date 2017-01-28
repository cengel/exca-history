Option Compare Database
Option Explicit
Private Sub cboFindUnitToCopy_AfterUpdate()
On Error GoTo err_cboFindUnitToCopy_AfterUpdate
If Me![cboFindUnitToCopy] <> "" Then
    Me.RecordSource = "SELECT * FROM [Exca: Descriptions Layer] WHERE [Unit Number] = " & Me![cboFindUnitToCopy]
    Me![Unit Number].ControlSource = "Unit Number"
    Me![Consistency].ControlSource = "Consistency"
    Me![Colour].ControlSource = "Colour"
    Me![Texture].ControlSource = "Texture"
    Me![Bedding].ControlSource = "Bedding"
    Me![Inclusions].ControlSource = "Inclusions"
    Me![Post-depositional Features].ControlSource = "Post-depositional Features"
    Me![Basal Boundary].ControlSource = "Basal Boundary"
    Me![copy data].Enabled = True
End If
Exit Sub
err_cboFindUnitToCopy_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub copy_data_Click()
On Error GoTo Err_copy_data_Click
Dim msg, Style, Title, response
msg = "This action will replace the unit sheet (" & Me![Text17] & ") "
msg = msg & "data with with that of unit " & Me![Unit Number] & " shown here." & Chr(13) & Chr(13)
msg = msg & "This action cannot be undone. Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
Title = "Overwriting Records"  ' Define title.
response = MsgBox(msg, Style, Title)
If response = vbYes Then    ' User chose Yes.
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Consistency] = Me![Consistency]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Colour] = Me![Colour]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Texture] = Me![Texture]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Bedding] = Me![Bedding]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Inclusions] = Me![Inclusions]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Post-depositional Features] = Me![Post-depositional Features]
    Forms![Exca: Unit Sheet]![Exca: Subform Layer descr]![Basal Boundary] = Me![Basal Boundary]
Else    ' User chose No.
End If
Exit_copy_data_Click:
    Exit Sub
Err_copy_data_Click:
    Call General_Error_Trap
    Resume Exit_copy_data_Click
End Sub
Sub find_unit_Click()
End Sub
Sub close_Click()
On Error GoTo Err_close_Click
    DoCmd.Close acForm, "Exca: Copy layer description"
Exit_close_Click:
    Exit Sub
Err_close_Click:
    Call General_Error_Trap
    Resume Exit_close_Click
End Sub
