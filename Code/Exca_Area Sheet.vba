Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub Area_name_AfterUpdate()
On Error GoTo err_Area_name_afterupdate
Dim msg, retval
    If Not IsNull(Me![Area Name].OldValue) Or (Me![Area Name].OldValue <> Me![Area Name]) Then
        msg = "Sorry but edits to the Area name are not allowed. Area names are stored in many different tables "
        msg = msg & "and this name may have already been used." & Chr(13) & Chr(13)
        msg = msg & "It is possible to archive this as an old area name and add it to the list of Historical area names if you wish. This would "
        msg = msg & " take the format of:" & Chr(13) & Chr(13) & "Old Area name: " & Me![Area Name].OldValue & " now equates to " & Me![Area Name]
        msg = msg & Chr(13) & Chr(13) & "Press Cancel to return to the original Area name"
        msg = msg & Chr(13) & "or "
        msg = msg & "Press OK to change this area name and add the old one to the historical list. "
        retval = MsgBox(msg, vbExclamation + vbOKCancel + vbDefaultButton2, "Stop: Area names cannot just be altered")
        If retval = vbCancel Then
            Me![Area Name] = Me![Area Name].OldValue 'reset to oldval
        ElseIf retval = vbOK Then
            Dim sql, sql2, sql3, newAreaNum
            sql = "INSERT INTO [Exca: Area Sheet] ([Area name], [Mound], [Description]) VALUES ('" & Me![Area Name] & "','" & Me![Mound] & "'," & IIf(IsNull(Me![Description]), "null", "'" & Me![Description] & "'") & ");"
            DoCmd.RunSQL sql
            newAreaNum = DLookup("[Area Number]", "Exca: Area Sheet", "[Area Name] = '" & Me![Area Name] & "'")
            sql2 = "INSERT INTO [Exca: Area_Historical_Names] (CurrentAreaNumber, CurrentAreaName, OldAreaNumber, OldAreaName, OldMound, OldDescription)"
            sql2 = sql2 & " VALUES (" & newAreaNum & ", '" & Me![Area Name] & "', " & Me![Area number] & ", '" & Me![Area Name].OldValue & "', '" & Me![Mound] & "', '" & Me![Description] & "');"
            DoCmd.RunSQL sql2
            DoCmd.RunCommand acCmdDeleteRecord
            Me.Requery 'get updated RS
            DoCmd.GoToRecord acActiveDataObject, , acLast
        End If
    End If
Exit Sub
err_Area_name_afterupdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindArea_AfterUpdate()
On Error GoTo err_cboFindArea_AfterUpdate
    If Me![cboFindArea] <> "" Then
        If Me![Area Name].Enabled = False Then Me![Area Name].Enabled = True
        DoCmd.GoToControl "Area Name"
        DoCmd.FindRecord Me![cboFindArea]
        Me![cboFindArea] = ""
    End If
Exit Sub
err_cboFindArea_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindArea_NotInList(NewData As String, response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "Sorry this area cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    Me![cboFind].Undo
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Area name"
Exit Sub
err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdViewHistorical_Click()
On Error GoTo err_cmdViewHistorical_Click
    DoCmd.OpenForm "Exca: Area Historical", acNormal, , "[CurrentAreaNumber] = " & Me![Area number], acFormReadOnly, acDialog
Exit Sub
err_cmdViewHistorical_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Area Sheet"
Exit Sub
err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Current()
On Error GoTo err_Form_Current
If IsNull(Me![Area number]) Then
    Me![Field24].Visible = True
    Me![txtMound].Visible = False
Else
    Me![Field24].Visible = False
    Me![txtMound].Visible = True
    Me![txtMound].Locked = True
End If
Dim historical
historical = Null
If Not IsNull(Me![Area number]) Then
    historical = DLookup("[CurrentAreaNumber]", "[Exca: Area_Historical_Names]", "[CurrentAreaNumber] = " & Me![Area number])
End If
If Not IsNull(historical) Then
    Me![cmdViewHistorical].Enabled = True
Else
    Me![cmdViewHistorical].Enabled = False
End If
Exit Sub
err_Form_Current:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
    If GetGeneralPermissions = "ADMIN" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
