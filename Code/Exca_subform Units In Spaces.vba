Option Compare Database
Option Explicit
Private Sub cmdGoToSpace_Click()
On Error GoTo Err_cmdGoToSpace_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retval, sql, insertArea, permiss
    stDocName = "Exca: Space Sheet"
    If Not IsNull(Me![txtIn_Space]) Or Me![txtIn_Space] <> "" Then
        checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = " & Me![txtIn_Space])
        If IsNull(checknum) Then
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
                msg = "This Space Number DOES NOT EXIST in the database."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
                If retval = vbNo Then
                    MsgBox "No space record to view, please alert the your team leader about this.", vbExclamation, "Missing Space Record"
                Else
                    If Forms![Exca: Unit Sheet]![Area] <> "" Then
                        insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                    Else
                        insertArea = Null
                    End If
                    sql = "INSERT INTO [Exca: Space Sheet] ([Space Number], [Area]) VALUES (" & Me![txtIn_Space] & ", " & insertArea & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = " & Me![txtIn_Space], acFormEdit, acDialog
                End If
            Else
                MsgBox "Sorry but this space record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Space Record"
            End If
        Else
            stLinkCriteria = "[Space Number]=" & Me![txtIn_Space]
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        End If
    Else
        MsgBox "No Space number to view", vbInformation, "No Space Number"
    End If
Exit_cmdGoToSpace_Click:
    Exit Sub
Err_cmdGoToSpace_Click:
    Call General_Error_Trap
    Resume Exit_cmdGoToSpace_Click
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
End Sub
Private Sub Form_Current()
On Error GoTo err_Current
    If Me![txtIn_Space] = "" Or IsNull(Me![txtIn_Space]) Then
        Me![cmdGoToSpace].Enabled = False
    Else
        Me![cmdGoToSpace].Enabled = True
    End If
Exit Sub
err_Current:
    Call General_Error_Trap
    Exit Sub
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
Private Sub txtIn_Space_AfterUpdate()
On Error GoTo err_txtIn_Space_AfterUpdate
Dim checknum, msg, retval, sql, insertArea
If Me![txtIn_Space] <> "" Then
    If IsNumeric(Me![txtIn_Space]) Then
        checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = " & Me![txtIn_Space])
        If IsNull(checknum) Then
            msg = "This Space Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
            If retval = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                If Forms![Exca: Unit Sheet]![Area] <> "" Then
                    insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                Else
                    insertArea = Null
                End If
                sql = "INSERT INTO [Exca: Space Sheet] ([Space Number], [Area]) VALUES (" & Me![txtIn_Space] & ", " & insertArea & ");"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = " & Me![txtIn_Space], acFormEdit, acDialog
            End If
        Else
            Me![cmdGoToSpace].Enabled = True
        End If
    Else
        MsgBox "The Space number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If
Exit Sub
err_txtIn_Space_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub txtIn_Space_BeforeUpdate(Cancel As Integer)
On Error GoTo err_spacebefore
If Me![txtIn_Space] = 0 Then
        MsgBox "Space 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
        Cancel = True
        SendKeys "{ESC}" 'seems to need it done 3x
        SendKeys "{ESC}"
        SendKeys "{ESC}"
End If
Exit Sub
err_spacebefore:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub txtIn_Space_LostFocus()
On Error GoTo err_lost
    Forms![Exca: Unit Sheet]![Exca: subform Units  in Buildings].Form.Requery
Exit Sub
err_lost:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Unit_AfterUpdate()
Me.Requery
DoCmd.GoToRecord , , acLast
End Sub
