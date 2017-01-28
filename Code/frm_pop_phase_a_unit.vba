Option Compare Database
Option Explicit
Private Sub cmdCancel_Click()
On Error GoTo err_cancel
    DoCmd.Close acForm, Me.Name
Exit Sub
err_cancel:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdOK_Click()
On Error GoTo err_cmdOK
    Dim Unit, getBuildingorSpace, getDivider
    If Me![cboSelect] <> "" Then
        Unit = Forms![Exca: Unit Sheet]![Unit Number]
        getDivider = InStr(Me!cboSelect, ".") 'format is B42.A or Sp115.1 etc etc
        getBuildingorSpace = Left(Me!cboSelect, getDivider - 1)
        Dim checkRec, sql
        checkRec = DLookup("OccupationPhase", "[Exca: Units Occupation Phase]", "[Unit] = " & Unit & " AND [OccupationPhase] like '" & getBuildingorSpace & "%'")
        If IsNull(checkRec) Then
            sql = "INSERT INTO [Exca: Units Occupation Phase] ([Unit], [OccupationPhase]) VALUES (" & Unit & ",'" & Me!cboSelect & "');"
            DoCmd.RunSQL sql
        Else
            sql = "UPDATE [Exca: Units Occupation Phase] SET [OccupationPhase] = '" & Me![cboSelect] & "' WHERE Unit = " & Unit & " AND [OccupationPhase] = '" & checkRec & "';"
            DoCmd.RunSQL sql
        End If
        DoCmd.Close acForm, Me.Name
    Else
        MsgBox "You must select a phase from the list or press cancel to leave this form", vbInformation, "No Phase Selected"
    End If
Exit Sub
err_cmdOK:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdRemove_Click()
On Error GoTo err_cmdRemove
    Dim Unit, getEquals, Phase, getNumber, sql, resp
    Unit = Forms![Exca: Unit Sheet]![Unit Number]
 resp = MsgBox("This will remove all the phasing associated with Unit " & Unit & " - ARE YOU SURE?" & Chr(13) & Chr(13) & "To remove one phase item only: on the main unit sheet click over the arrow to the right of the specific phase and press delete.", vbCritical + vbYesNo, "Confirm Action")
 If resp = vbYes Then
    getEquals = InStr(Me!cboSelect.RowSource, "=") 'format is =Sp115.1 or B42. etc etc
    getNumber = Mid(Me!cboSelect.RowSource, getEquals + 1, (Len(Me!cboSelect.RowSource) - 1) - getEquals)
    If InStr(Me!cboSelect.RowSource, "Space") > 0 Then
        Phase = "Sp" & getNumber & "."
    Else
        Phase = "B" & getNumber & "."
    End If
    sql = "DELETE FROM [Exca: Units Occupation Phase] WHERE Unit = " & Unit & ";"
    DoCmd.RunSQL sql
End If
    DoCmd.Close acForm, Me.Name
Exit Sub
err_cmdRemove:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
    If Not IsNull(Me.OpenArgs) Then
        Me!cboSelect.RowSource = Me.OpenArgs
        Me!cboSelect.Requery
        Dim Unit, getEquals, getNumber, Phase, phasedalready, sql
        Unit = Forms![Exca: Unit Sheet]![Unit Number]
        getEquals = InStr(Me!cboSelect.RowSource, "=") 'format is =Sp115.1 or B42. etc etc
        getNumber = Mid(Me!cboSelect.RowSource, getEquals + 1, (Len(Me!cboSelect.RowSource) - 1) - getEquals)
        If InStr(Me!cboSelect.RowSource, "Space") > 0 Then
            Phase = "Sp" & getNumber & "."
        Else
            Phase = "B" & getNumber & "."
        End If
        phasedalready = DCount("[OccupationPhase]", "[Exca: Units Occupation Phase]", "[OccupationPhase] like '" & Phase & "%'")
        Me!cmdRemove.Caption = "Remove Unit from Phasing of " & Phase
        If phasedalready >= 1 Then
            Me!cmdRemove.Enabled = True
        Else
            Me!cmdRemove.Enabled = False
        End If
    Else
        MsgBox "Form opened with no parametres. Invalid action. The form will now close.", vbInformation, "No OpenArgs"
        DoCmd.Close acForm, Me.Name
    End If
Exit Sub
err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
