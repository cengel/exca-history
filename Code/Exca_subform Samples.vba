Option Compare Database
Private Sub Form_BeforeUpdate(Cancel As Integer)
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
Private Sub SampleType_NotInList(NewData As String, response As Integer)
On Error GoTo err_Sampletype_NotInList
Dim retVal, sql
retVal = MsgBox("This Sample Type does not appear in the pre-defined list. Have you checked the list to make sure there is no match?", vbQuestion + vbYesNo, "New Sample Type")
If retVal = vbYes Then
    MsgBox "Ok this sample type will now be added to the list", vbInformation, "New Sample Type Allowed"
     response = acDataErrAdded
    sql = "INSERT INTO [Exca:SampleTypeLOV] ([SampleType]) VALUES ('" & NewData & "');"
    DoCmd.RunSQL sql
Else
    response = acDataErrContinue
End If
Exit Sub
err_Sampletype_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub
