Option Compare Database
Option Explicit
Dim g_reportfilter
Private Sub Close_Click()
On Error GoTo err_close_Click
    DoCmd.Close acForm, Me.Name
    Exit Sub
err_close_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdBuildSQL_Click()
On Error GoTo err_buildsql
Dim selectsql, wheresql, orderbysql, fullsql
selectsql = "SELECT * FROM [Exca: Unit Sheet with Relationships] "
wheresql = ""
If Me![txtBuildingNumbers] <> "" Then
    wheresql = wheresql & "(" & Me![txtBuildingNumbers] & ")"
End If
If Me![txtSpaceNumbers] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtSpaceNumbers] & ")"
End If
If Me![txtFeatureNumbers] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtFeatureNumbers] & ")"
End If
If Me![txtLevels] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtLevels] & ")"
End If
If Me![txtCategory] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Category] like '%" & Me![txtCategory] & "%'"
End If
If Me![cboArea] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Area] = '" & Me![cboArea] & "'"
End If
If Me![cboYear] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Year] = " & Me![cboYear]
End If
If Me![txtUnitNumbers] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "(" & Me![txtUnitNumbers] & ")"
End If
If Me![txtText] <> "" Then
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "([Discussion] like '%" & Me![txtText] & "%' OR [Exca: Unit Sheet with Relationships].[Description] like '%" & Me![txtText] & "%')"
End If
If Me![cboDataCategory] <> "" Then
    selectsql = "SELECT [Exca: Unit Sheet with Relationships].[Unit Number], [Exca: Unit Sheet with Relationships].Year, " & _
                "[Exca: Unit Sheet with Relationships].Area, [Exca: Unit Sheet with Relationships].Category, " & _
                "[Exca: Unit Sheet with Relationships].[Grid X], [Exca: Unit Sheet with Relationships].[Grid Y], " & _
                "[Exca: Unit Sheet with Relationships].Description, [Exca: Unit Sheet with Relationships].Discussion, [Exca: Unit Sheet with Relationships].[Priority Unit], " & _
                "[Exca: Unit Sheet with Relationships].ExcavationStatus, [Exca: Unit Sheet with Relationships].Levels, " & _
                "[Exca: Unit Sheet with Relationships].Building, [Exca: Unit Sheet with Relationships].Space, [Exca: Unit Sheet with Relationships].Feature, " & _
                "[Exca: Unit Sheet with Relationships].TimePeriod, [Exca: Unit Data Categories].[Data Category]" & _
                " FROM [Exca: Unit Sheet with Relationships] INNER JOIN [Exca: Unit Data Categories] ON [Exca: Unit Sheet with Relationships].[Unit Number] = [Exca: Unit Data Categories].[Unit Number]"
    If wheresql <> "" Then wheresql = wheresql & " AND "
    wheresql = wheresql & "[Data Category] = '" & Me![cboDataCategory] & "'"
End If
If wheresql <> "" Then selectsql = selectsql & " WHERE "
orderbysql = " ORDER BY [Exca: Unit Sheet with Relationships].[Unit Number];"
fullsql = selectsql & wheresql & orderbysql
g_reportfilter = wheresql
Me!txtSQL = fullsql
Me![frm_subSearch].Form.RecordSource = fullsql
If Me![frm_subSearch].Form.RecordsetClone.RecordCount = 0 Then
    MsgBox "No records match the criteria you entered.", 48, "No Records Found"
    Me![cmdClearSQL].SetFocus
End If
Exit Sub
err_buildsql:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdClear_Click()
On Error GoTo err_clear
Me![txtBuildingNumbers] = ""
Exit Sub
err_clear:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdClearFeature_Click()
On Error GoTo err_feature
Me![txtFeatureNumbers] = ""
Exit Sub
err_feature:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdCLearLevel_Click()
On Error GoTo err_level
Me![txtLevels] = ""
Exit Sub
err_level:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdClearSpace_Click()
On Error GoTo err_space
Me![txtSpaceNumbers] = ""
Exit Sub
err_space:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdClearSQL_Click()
On Error GoTo err_clearsql
Dim sql
Me![txtBuildingNumbers] = ""
Me![txtSpaceNumbers] = ""
Me![txtFeatureNumbers] = ""
Me![txtLevels] = ""
Me![txtCategory] = ""
Me![cboArea] = ""
Me![cboYear] = ""
Me![txtUnitNumbers] = ""
Me![txtText] = ""
Me![cboDataCategory] = ""
sql = "SELECT * FROM [Exca: Unit Sheet with Relationships] ORDER BY [Unit Number];"
Me!txtSQL = sql
Me![frm_subSearch].Form.RecordSource = sql
Exit Sub
err_clearsql:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdClearUnit_Click()
On Error GoTo err_unit
Me![txtUnitNumbers] = ""
Exit Sub
err_unit:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdEnterBuilding_Click()
On Error GoTo err_building
Dim openarg
openarg = "Building"
If Me![txtBuildingNumbers] <> "" Then openarg = "Building;" & Me![txtBuildingNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_building:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdEnterFeature_Click()
On Error GoTo err_enterfeature
Dim openarg
openarg = "Feature"
If Me![txtFeatureNumbers] <> "" Then openarg = "Feature;" & Me![txtFeatureNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterfeature:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdEnterLevel_Click()
On Error GoTo err_enterlevel
Dim openarg
openarg = "Levels"
If Me![txtLevels] <> "" Then openarg = "Levels;" & Me![txtLevels]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterlevel:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdEnterSpace_Click()
On Error GoTo err_enterspace
Dim openarg
openarg = "Space"
If Me![txtSpaceNumbers] <> "" Then openarg = "Space;" & Me![txtSpaceNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_enterspace:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdPrint_Click()
On Error GoTo err_cmdPrint
    Call cmdBuildSQL_Click
    If Me![frm_subSearch].Form.RecordsetClone.RecordCount = 0 Then
        Me![cmdClearSQL].SetFocus
        Exit Sub
    Else
        DoCmd.OpenReport "R_unit_search_report", acViewPreview
        If Not IsNull(g_reportfilter) Then
            Reports![R_unit_search_report].FilterOn = True
            Reports![R_unit_search_report].Filter = g_reportfilter
        End If
    End If
Exit Sub
err_cmdPrint:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdUnit_Click()
On Error GoTo err_unitclick
Dim openarg
openarg = "unit number"
If Me![txtUnitNumbers] <> "" Then openarg = "unit number;" & Me![txtUnitNumbers]
DoCmd.OpenForm "frm_popsearch", , , , , acDialog, openarg
Exit Sub
err_unitclick:
    Call General_Error_Trap
    Exit Sub
End Sub
