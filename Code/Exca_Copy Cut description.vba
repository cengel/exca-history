Option Compare Database
Option Explicit
Private Sub copy_data_Click()
On Error GoTo Err_copy_data_Click
Dim Msg, Style, Title, Response
Msg = "This action will replace the unit sheet contents, and cannot be undone. Do you want to continue?"   ' Define message.
Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
Title = "Overwriting Records"  ' Define title.
Response = MsgBox(Msg, Style, Title)
If Response = vbYes Then    ' User chose Yes.
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Shape] = Me![Shape]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Corners] = Me![Corners]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Top Break] = Me![Top Break]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Sides] = Me![Sides]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Base Break] = Me![Base Break]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Base] = Me![Base]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![Orientation] = Me![Orientation]
    Forms![Exca: Unit Sheet]![Exca: Subform Cut descr]![All Layers within] = Me![All Layers within]
Else    ' User chose No.
End If
Exit_copy_data_Click:
    Exit Sub
Err_copy_data_Click:
    MsgBox Err.Description
    Resume Exit_copy_data_Click
End Sub
Sub find_unit_Click()
On Error GoTo Err_find_unit_Click
    Screen.PreviousControl.SetFocus
     Unit_Number.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70
Exit_find_unit_Click:
    Exit Sub
Err_find_unit_Click:
    MsgBox Err.Description
    Resume Exit_find_unit_Click
End Sub
Sub Command13_Click()
On Error GoTo Err_Command13_Click
    Screen.PreviousControl.SetFocus
    DoCmd.FindNext
Exit_Command13_Click:
    Exit Sub
Err_Command13_Click:
    MsgBox Err.Description
    Resume Exit_Command13_Click
End Sub
Sub Command14_Click()
On Error GoTo Err_Command14_Click
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
Exit_Command14_Click:
    Exit Sub
Err_Command14_Click:
    MsgBox Err.Description
    Resume Exit_Command14_Click
End Sub
Sub close_Click()
On Error GoTo Err_close_Click
    DoCmd.Close
Exit_close_Click:
    Exit Sub
Err_close_Click:
    MsgBox Err.Description
    Resume Exit_close_Click
End Sub
