Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub cboFindBuilding_AfterUpdate()
On Error GoTo err_cboFindBuilding_AfterUpdate
    If Me![cboFindBuilding] <> "" Then
        If Me![Number].Enabled = False Then Me![Number].Enabled = True
        DoCmd.GoToControl "Number"
        DoCmd.FindRecord Me![cboFindBuilding]
        Me![cboFindBuilding] = ""
    End If
Exit Sub
err_cboFindBuilding_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindBuilding_NotInList(NewData As String, response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "Sorry this Building cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    Me![cboFindBuilding].Undo
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Number"
Exit Sub
err_cmdAddNew_Click:
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
    myq1.sql = "sp_Portfolio_GetBuildingFieldID_2009 2009"
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
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, 2009
        Else
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=feature&id=" & Me![Number])
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
Private Sub cmdPrintBuildingSheet_Click()
On Error GoTo err_cmdBuilding
    Dim both
    both = MsgBox("Do you want a list of the associated Units as well?", vbQuestion + vbYesNo, "Units?")
        DoCmd.OpenReport "R_BuildingSheet", acViewPreview, , "[Number] = " & Me![Number]
        If both = vbYes Then DoCmd.OpenReport "R_Units_in_Buildings", acViewPreview, , "[In_Building] = " & Me![Number]
Exit Sub
err_cmdBuilding:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdReportProblem_Click()
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "building;" & Me![Number]
Exit Sub
err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    DoCmd.Close acForm, "Exca: Building Sheet"
Exit Sub
err_Excavation_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_Form_BeforeUpdate
If IsNull(Me![Number] And (Not IsNull(Me![Field24]) Or Not IsNull(Me![Location]) Or (Me![Description] <> "" And Not IsNull(Me![Description])))) Then
    MsgBox "You must enter a building number otherwise the record cannot be saved." & Chr(13) & Chr(13) & "If you wish to cancel this record being entered and start again completely press ESC", vbInformation, "Incomplete data"
    Cancel = True
    DoCmd.GoToControl "Number"
ElseIf IsNull(Me![Number]) And IsNull(Me![Field24]) And IsNull(Me![Location]) And (IsNull(Me![Description]) Or Me![Description] = "") Then
    Me.Undo
End If
If Me.Dirty And (IsNull(Me![LastUpdatedBy]) Or Me![LastUpdatedBy] = "") Then
    MsgBox "You must enter your name in the Last Updated By field", vbInformation, "Last Updated By Field not filled out"
    Cancel = True
    DoCmd.GoToControl "LastUpdatedBy"
End If
Me![LastUpdatedDate] = Date
Exit Sub
err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Current()
On Error GoTo err_Form_Open
   If Me![Number] <> "" Then
        Me![Number].Locked = True
        Me![Number].Enabled = False
        Me![Number].BackColor = Me.Section(0).BackColor
        Me![Location].SetFocus
    Else
        Me![Number].Locked = False
        Me![Number].Enabled = True
        Me![Number].BackColor = 16777215
        Me![Number].SetFocus
    End If
Dim imageCount, Imgcaption
backhere:
Imgcaption = "Images of Building"
    Me![cmdGoToImage].Caption = Imgcaption
    Me![cmdGoToImage].Enabled = True
Exit Sub
err_Form_Open:
    If Err.Number = 3146 Then 'odbc call failed, crops up every so often on all
        imageCount = "?"
        GoTo backhere
    Else
        Call General_Error_Trap
    End If
    Exit Sub
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
                    Me![Number] = NumKnown 'add it to the number fld
                    Me![Number].Locked = True 'lock the number field
                End If
                AreaKnown = InStr(UCase(getArgs), "AREA:")
                If AreaKnown <> 0 Then
                    AreaKnown = Mid(getArgs, AreaKnown + 5, Len(getArgs))
                    Me![Field24] = AreaKnown 'add it to the area fld
                    Me![Field24].Locked = True
                End If
            End If
            Me![cboFindBuilding].Enabled = False
            Me![cmdAddNew].Enabled = False
            Me.AllowAdditions = False
        End If
    End If
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        Me![cboFindBuilding].Enabled = False
        Me![cmdAddNew].Enabled = False
        Me.AllowAdditions = False
    Else
        DoCmd.GoToControl "cboFindBuilding"
    End If
    Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper") And (Me.AllowAdditions = True Or Me.AllowDeletions = True Or Me.AllowEdits = True) Then
        ToggleFormReadOnly Me, False, "NoDeletions"
    Else
        ToggleFormReadOnly Me, True
        Me![cmdAddNew].Enabled = False
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Number_AfterUpdate()
On Error GoTo err_Number_AfterUpdate
Dim checknum
If Me![Number] <> "" Then
    checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but this Building Number already exists, please enter another number.", vbInformation, "Duplicate Building Number"
        If Not IsNull(Me![Number].OldValue) Then
            Me![Number] = Me![Number].OldValue
        Else
            Dim currloc, currarea, currdesc
            currloc = Me![Location]
            currarea = Me![Field24]
            currdesc = Me![Description]
            DoCmd.GoToControl "Number"
            Me.Undo
            If Not IsNull(currloc) Then Me![Location] = currloc
            If Not IsNull(currarea) Then Me![Field24] = currarea
            If Not IsNull(currdesc) Then Me![Description] = currdesc
            DoCmd.GoToControl "Description"
            DoCmd.GoToControl "Number"
        End If
    End If
End If
Exit Sub
err_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
