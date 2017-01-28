Option Compare Database
Option Explicit
Sub Close_Feature_Sheet_Click()
End Sub
Private Sub Building_AfterUpdate()
On Error GoTo err_Building_AfterUpdate
Dim checknum, msg, retVal
If Me![Building] <> "" Then
    If IsNumeric(Me![Building]) Then
        checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
        If IsNull(checknum) Then
            msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
            If retVal = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Combo27]
            End If
        Else
            Me![cmdGoToBuilding].Enabled = True
        End If
    Else
        MsgBox "The Building number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If
Exit Sub
err_Building_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindFeature_AfterUpdate()
On Error GoTo err_cboFindFeature_AfterUpdate
    If Me![cboFindFeature] <> "" Then
        If Me![Feature Number].Enabled = False Then Me![Feature Number].Enabled = True
        DoCmd.GoToControl "Feature Number"
        DoCmd.FindRecord Me![cboFindFeature]
        Me![cboFindFeature] = ""
    End If
Exit Sub
err_cboFindFeature_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindFeature_NotInList(NewData As String, response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "Sorry this Feature cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    Me![cboFindFeature].Undo
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Feature Number"
Exit Sub
err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdGoToBuilding_Click()
On Error GoTo Err_cmdGoToBuilding_Click
Dim checknum, msg, retVal, permiss
If Not IsNull(Me![Building]) Or Me![Building] <> "" Then
    checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
    If IsNull(checknum) Then
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
            msg = "This Building Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
            If retVal = vbNo Then
                MsgBox "No building record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
            Else
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Combo27]
            End If
        Else
            MsgBox "Sorry but this building record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Building Record"
        End If
    Else
        Dim stDocName As String
        Dim stLinkCriteria As String
        stDocName = "Exca: Building Sheet"
        stLinkCriteria = "[Number]= " & Me![Building]
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, , "FILTER"
    End If
End If
Exit Sub
Err_cmdGoToBuilding_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdGoToImage_Click()
On Error GoTo err_cmdGoToImage_Click
Dim mydb As DAO.Database
Dim tmptable As TableDef, tblConn, I, msg, fldid
Set mydb = CurrentDb
    Dim myq1 As QueryDef, connStr
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = mydb.TableDefs(0).Connect & ";UID=portfolio;PWD=portfolio"
    myq1.ReturnsRecords = True
    myq1.sql = "sp_Portfolio_GetFeatureFieldID_2009 " & Me![Year]
    Dim myrs As Recordset
    Set myrs = myq1.OpenRecordset
    If myrs.Fields(0).Value = "" Or myrs.Fields(0).Value = 0 Then
        fldid = 0
    Else
        fldid = myrs.Fields(0).Value
    End If
    myrs.Close
    Set myrs = Nothing
    myq1.Close
    Set myq1 = Nothing
    For I = 0 To mydb.TableDefs.Count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
    If tmptable.Connect <> "" Then
        tblConn = tmptable.Connect
        Exit For
    End If
    Next I
    If tblConn <> "" Then
        If InStr(tblConn, "catalsql") = 0 Then
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Feature Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, Me![Year]
        Else
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=feature&id=" & Me![Feature Number])
        End If
    Else
    End If
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
Exit Sub
err_cmdGoToImage_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdPrintFeatureSheet_Click()
On Error GoTo err_print
    If Me![Feature Number] <> "" Then
        DoCmd.OpenReport "R_Feature_Sheet", acViewPreview, , "[feature number] = " & Me![Feature Number]
    End If
Exit Sub
err_print:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdReportProblem_Click()
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "feature number;" & Me![Feature Number]
Exit Sub
err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Combo27_AfterUpdate()
On Error GoTo err_Combo27_AfterUpdate
If Me![Combo27].Column(1) <> "" Then
    Me![Mound] = Me![Combo27].Column(1)
End If
Exit Sub
err_Combo27_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    DoCmd.Close acForm, "Exca: Feature Sheet"
Exit_Excavation_Click:
    Exit Sub
err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub
Private Sub Feature_Number_AfterUpdate()
On Error GoTo err_Feature_Number_AfterUpdate
Dim checknum
If Me![Feature Number] <> "" Then
    checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![Feature Number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but the Feature Number " & Me![Feature Number] & " already exists, please enter another number.", vbInformation, "Duplicate Feature Number"
        If Not IsNull(Me![Feature Number].OldValue) Then
            Me![Feature Number] = Me![Feature Number].OldValue
        Else
            DoCmd.GoToControl "Year"
            DoCmd.GoToControl "Feature Number"
            Me![Feature Number].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        ToggleFormReadOnly Me, False
    End If
End If
If Me![Feature Number] <> "" Then Me![lblMsg].Visible = False
Exit Sub
err_Feature_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Feature_Number_Exit(Cancel As Integer)
End Sub
Private Sub Feature_Type_AfterUpdate()
On Error GoTo err_Feature_Type
If Me![Feature Type] <> "" Then
    Me![cboFeatureSubType] = ""
    Me![cboFeatureSubType].RowSource = "SELECT [Exca:FeatureSubTypeLOV].FeatureSubType FROM [Exca:FeatureTypeLOV] INNER JOIN [Exca:FeatureSubTypeLOV] ON [Exca:FeatureTypeLOV].FeatureTypeID = [Exca:FeatureSubTypeLOV].FeatureTypeID WHERE ((([Exca:FeatureTypeLOV].FeatureType)='" & Me![Feature Type] & "')) ORDER BY [Exca:FeatureSubTypeLOV].FeatureSubType; "
    Me![cboFeatureSubType].Requery
End If
If LCase(Me![Feature Type]) = "burial" Then
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Then
        Me!txtBurialMNI.Enabled = True
        Me!txtBurialMNI.Enabled = False
    Else
        Me!txtBurialMNI.Enabled = False
        Me!txtBurialMNI.Enabled = True
    End If
End If
Exit Sub
err_Feature_Type:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_Form_BeforeUpdate
If IsNull(Me![Exca: subform Feature Plans].Form![Graphic Number]) Then
    If Me.ActiveControl.Name <> "Dimensions" And Me.ActiveControl.Name <> "Description" Then
        MsgBox "There is no Plan number entered for this Feature. Please can you enter one soon", vbInformation, "What is the Plan Number?"
    End If
End If
Me![Date changed] = Now()
Exit Sub
err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Current()
On Error GoTo err_Form_Current
Dim permiss
permiss = GetGeneralPermissions
If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
    If IsNull(Me![Feature Number]) Or Me![Feature Number] = "" Then
        ToggleFormReadOnly Me, True, "Additions" 'code in GeneralProcedures-shared
        Me![lblMsg].Visible = True
        Me![Feature Number].Locked = False
        Me![Feature Number].Enabled = True
        Me![Feature Number].BackColor = 16777215
        Me![Feature Number].SetFocus
    Else
        If Me.FilterOn = True And Me.AllowEdits = False Then
            ToggleFormReadOnly Me, True, "NoAdditions"
        Else
            If Me.FilterOn Then
                ToggleFormReadOnly Me, False, "NoAdditions"
            Else
                ToggleFormReadOnly Me, False, "NoDeletions"
            End If
            Me![Year].SetFocus
            Me![Feature Number].Locked = True
            Me![Feature Number].Enabled = False
            Me![Feature Number].BackColor = Me.Section(0).BackColor
        End If
        Me![lblMsg].Visible = False
    End If
End If
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        Me![cboFindFeature].Enabled = False
        Me![cmdAddNew].Enabled = False
    Else
        If Me![cboFindFeature].Enabled Then DoCmd.GoToControl "cboFindFeature"
    End If
Me![cboFeatureSubType].RowSource = "SELECT [Exca:FeatureSubTypeLOV].FeatureSubType FROM [Exca:FeatureTypeLOV] INNER JOIN [Exca:FeatureSubTypeLOV] ON [Exca:FeatureTypeLOV].FeatureTypeID = [Exca:FeatureSubTypeLOV].FeatureTypeID WHERE ((([Exca:FeatureTypeLOV].FeatureType)='" & Me![Feature Type] & "')) ORDER BY [Exca:FeatureSubTypeLOV].FeatureSubType; "
Dim imageCount, Imgcaption
backhere:
Imgcaption = "Images of Feature"
Me![cmdGoToImage].Caption = Imgcaption
Me![cmdGoToImage].Enabled = True
If permiss = "ADMIN" And LCase(Me![Feature Type]) = "burial" Then
    Me!txtBurialMNI.Enabled = True
    Me!txtBurialMNI.Locked = False
Else
    Me!txtBurialMNI.Enabled = False
    Me!txtBurialMNI.Locked = True
End If
Exit Sub
err_Form_Current:
    If Err.Number = 3146 Then 'odbc call failed, crops up every so often on all
        imageCount = "?"
        GoTo backhere
    Else
        Call General_Error_Trap
        Exit Sub
    End If
End Sub
Private Sub Form_Error(DataErr As Integer, response As Integer)
Dim msg
If DataErr = 3162 Then
    msg = "An error has occurred: invalid entry in the current field, probably a null value." & Chr(13) & Chr(13)
    msg = msg & "The system will attempt to resolve this, please re-try the action, but if you continue to get an error press the ESC key."
    MsgBox msg, vbInformation, "Error encountered"
    response = acDataErrContinue
    SendKeys "{ESC}"
    SendKeys "{ESC}"
ElseIf DataErr = 3146 Then
    DoCmd.RunCommand acCmdUndo
    response = acDataErrContinue
End If
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
    If Not IsNull(Me.OpenArgs) Then
        Dim getArgs, whatTodo, NumKnown, AreaKnown
        Dim firstcomma, Action
        getArgs = Me.OpenArgs
        If Len(getArgs) > 0 Then
            firstcomma = InStr(getArgs, ",")
            If firstcomma <> 0 Then
                Action = Left(getArgs, firstcomma - 1)
                If UCase(Action) = "NEW" Then DoCmd.GoToRecord acActiveDataObject, , acNewRec
                NumKnown = InStr(UCase(getArgs), "NUM:")
                If NumKnown <> 0 Then
                    NumKnown = Mid(getArgs, NumKnown + 4, InStr(NumKnown, getArgs, ",") - (NumKnown + 4))
                    Me![Feature Number] = NumKnown 'add it to the number fld
                    Me![Feature Number].Locked = True 'lock the number field
                End If
                AreaKnown = InStr(UCase(getArgs), "AREA:")
                If AreaKnown <> 0 Then
                    AreaKnown = Mid(getArgs, AreaKnown + 5, Len(getArgs))
                    Me![Combo27] = AreaKnown 'add it to the area fld
                    Me![Combo27].Locked = True
                End If
            End If
            Me![cboFindFeature].Enabled = False
            Me![cmdAddNew].Enabled = False
            ToggleFormReadOnly Me, False
            Me.AllowAdditions = False
            Me![lblMsg].Visible = False
        End If
    Else
    End If
    Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
    Else
        ToggleFormReadOnly Me, True
        Me![cmdAddNew].Enabled = False
        Me![Feature Number].BackColor = Me.Section(0).BackColor
        Me![Feature Number].Locked = True
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub go_next_Click()
On Error GoTo Err_go_next_Click
    DoCmd.GoToRecord , , acNext
Exit_go_next_Click:
    Exit Sub
Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
End Sub
Private Sub go_previous_Click()
On Error GoTo Err_go_previous_Click
    DoCmd.GoToRecord , , acPrevious
Exit_go_previous_Click:
    Exit Sub
Err_go_previous_Click:
    MsgBox Err.Description
    Resume Exit_go_previous_Click
End Sub
Private Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click
    DoCmd.GoToRecord , , acFirst
Exit_go_to_first_Click:
    Exit Sub
Err_go_to_first_Click:
    MsgBox Err.Description
    Resume Exit_go_to_first_Click
End Sub
Private Sub go_to_last_Click()
On Error GoTo Err_go_last_Click
    DoCmd.GoToRecord , , acLast
Exit_go_last_Click:
    Exit Sub
Err_go_last_Click:
    MsgBox Err.Description
    Resume Exit_go_last_Click
End Sub
Private Sub Master_Control_Click()
End Sub
Private Sub New_entry_Click()
End Sub
Sub find_feature_Click()
End Sub
