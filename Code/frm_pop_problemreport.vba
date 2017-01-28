Option Compare Database
Option Explicit
Dim toShow, entitynum
Private Sub cmdCancel_Click()
On Error GoTo err_cmdCancel
    DoCmd.Close acForm, "frm_pop_problemreport"
Exit Sub
err_cmdCancel:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdClear_Click()
On Error GoTo err_cmdClear
    Me![txtToFind] = ""
    Me![cboSelect] = ""
Exit Sub
err_cmdClear:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdOK_Click()
On Error GoTo err_cmdOK
    If (Me![Comment] = "" Or Me![ReportersName] = "") Or (IsNull(Me![Comment]) Or IsNull(Me![ReportersName])) Then
        MsgBox "Please enter both your comment and your name, otherwise cancel the report", vbInformation, "Insufficient Info"
        Exit Sub
    Else
        Dim sql, strcomment
        strcomment = Replace(Me![Comment], "'", "''") 'bug fix july 2009 on site
        If spString <> "" Then
            Dim mydb As DAO.Database
            Dim myq1 As QueryDef
            Set mydb = CurrentDb
            Set myq1 = mydb.CreateQueryDef("")
            myq1.Connect = spString
            myq1.ReturnsRecords = False
            myq1.sql = "sp_Excavation_Add_Problem_Report_Entry " & entitynum & ", '" & toShow & "','" & strcomment & "','" & Me![ReportersName] & "','" & Format(Date, "dd/mm/yyyy") & "'"
            myq1.Execute
            myq1.Close
            Set myq1 = Nothing
            mydb.Close
            Set mydb = Nothing
            MsgBox "Thank you, your report has been saved for the Administrator to check", vbInformation, "Done"
        Else
            MsgBox "Sorry but this comment cannot be inserted at this time, please restart the database and try again", vbCritical, "Error"
        End If
        DoCmd.Close acForm, "frm_pop_problemreport"
    End If
Exit Sub
err_cmdOK:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
Dim colonpos
    If Not IsNull(Me.OpenArgs) Then
        Me![cboSelect].Visible = False
        toShow = LCase(Me.OpenArgs)
        colonpos = InStr(toShow, ";")
        If colonpos > 0 Then
            entitynum = right(toShow, Len(toShow) - colonpos)
            toShow = Left(toShow, colonpos - 1)
        End If
        Select Case toShow
        Case "building"
            Me![lblTitle].Caption = "Report a Building Record Problem"
            Me![cboSelect].RowSource = "Select [Number] from [Exca: Building Details];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Building Number: " & entitynum
        Case "space"
            Me![lblTitle].Caption = "Report a Space Record Problem"
            Me![cboSelect].RowSource = "Select [Space Number] from [Exca: Space Sheet];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Space Number: " & entitynum
        Case "feature number"
            Me![lblTitle].Caption = "Report a Feature Record Problem"
            Me![cboSelect].RowSource = "Select [Feature Number] from [Exca: Features];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Feature Number: " & entitynum
        Case "unit number"
            Me![lblTitle].Caption = "Report a Unit Record Problem"
            Me![cboSelect].RowSource = "Select [unit number] from [Exca: Unit Sheet] ORDER BY [unit number];"
            If entitynum <> "" Then Me![lblEntity].Caption = "Unit Number: " & entitynum
        End Select
        Me.Refresh
Else
    Me![lblTitle].Visible = False
    Me![lblEntity].Visible = False
    Me![cboSelect].Visible = True
End If
Exit Sub
err_open:
    Call General_Error_Trap
    Exit Sub
End Sub
