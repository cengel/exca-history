Option Compare Database
Option Explicit
Private Sub Amount_AfterUpdate()
On Error GoTo err_Amount
    If Not IsNumeric(Me!Amount) Then
        MsgBox Me!Amount & " is not a numeric amount, please enter the amount if Litres but as a number only", vbInformation, "Invalid Amount"
        DoCmd.GoToControl "SampleType"
        Me!Amount.SetFocus
    End If
Exit Sub
err_Amount:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub SampleType_AfterUpdate()
On Error GoTo err_sampletype
    If Me![SampleType].Column(1) <> "" Then
        If Me![Amount] <> "" Then
            Dim response
            response = MsgBox("There is a default amount for this sample type of " & Me![SampleType].Column(1) & ". Do you wish to overwrite the current amount?", vbYesNo + vbQuestion, "Amount?")
            If response = vbYes Then Me![Amount] = Me![SampleType].Column(1)
        Else
            Me![Amount] = Me![SampleType].Column(1)
        End If
    End If
    If InStr(Me![SampleType], "subsample") > 0 Then MsgBox "You must write the original sample number from which you are taking the sample in the Comment field as well as the details of the purpose of the sample", vbExclamation, "Sub sample requirements"
    If Me![SampleType] = "" Or IsNull(Me![SampleType]) Then MsgBox "YOU MUST ENTER A SAMPLE TYPE", vbExclamation, "Missing Sample Type"
Exit Sub
err_sampletype:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub SampleType_NotInList(NewData As String, response As Integer)
On Error GoTo err_Sampletype_NotInList
MsgBox "This Sample Type is not found in the current list, look carefully and consult the Type list via the button above. " & Chr(13) & Chr(13) & "There is a new format for sample types. This is: main type-subtype " & Chr(13) & Chr(13) & "eg: Flotation-routine" & Chr(13) & Chr(13) & "If you really cannot find your sample type then please use: Other and write specific details in the comment field. Then tell your Supervisor who will inform the project team.", vbExclamation, "Sample Types"
SendKeys "{ESC}{ESC}"
Exit Sub
err_Sampletype_NotInList:
    Call General_Error_Trap
    Exit Sub
End Sub
