Option Compare Database   'Use database order for string comparisons
Option Explicit
Private Sub Building_AfterUpdate()
On Error GoTo err_Building_AfterUpdate
Dim checknum, msg, retVal
If Me![Building] <> "" Then
    If IsNumeric(Me![Building]) Then
        If Me![Building] = 0 Then
            MsgBox "Building 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
            Me![Building] = ""
        Else
            checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
            If IsNull(checknum) Then
                msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retVal = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
                If retVal = vbNo Then
                    MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
                Else
                    DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Field26]
                End If
            Else
                Me![cmdGoToBuilding].Enabled = True
            End If
            Me![chkExternal] = False
        End If
    Else
        MsgBox "The Building number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
Else
    Me![chkExternal] = True
End If
Exit Sub
err_Building_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindSpace_AfterUpdate()
On Error GoTo err_cboFindSpace_AfterUpdate
    If Me![cboFindSpace] <> "" Then
        If Me![Space number].Enabled = False Then Me![Space number].Enabled = True
        DoCmd.GoToControl "Space Number"
        DoCmd.FindRecord Me![cboFindSpace]
        Me![cboFindSpace] = ""
    End If
Exit Sub
err_cboFindSpace_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindSpace_NotInList(NewData As String, response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "Sorry this Space cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    Me![cboFindSpace].Undo
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboHodderPhase_AfterUpdate()
On Error GoTo err_phase
    If Me![cboHodderPhase] <> "" Then
        Me![txtHodderPhase] = Me![Area] & "." & Me![cboHodderPhase]
    Else
        Dim response
        response = MsgBox("Do you wish the Hodder Phase field to be blank?", vbYesNo + vbQuestion, "Action confirmation")
        If response = vbYes Then
            Me![txtHodderPhase] = ""
        End If
    End If
    Me![cboHodderPhase] = ""
Exit Sub
err_phase:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub chkExternal_AfterUpdate()
On Error GoTo err_chk
If Me!chkExternal = True Then
    Me![Exca: subform Phases related to Space].Enabled = True
    Me![Exca: subform Phases related to Space].Locked = False
Else
    Me![Exca: subform Phases related to Space].Enabled = False
    Me![Exca: subform Phases related to Space].Locked = True
End If
Exit Sub
err_chk:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Space Number"
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
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , , acFormAdd, acDialog, "NEW,Num:" & Me![Building] & ",Area:" & Me![Field26]
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
    myq1.sql = "sp_Portfolio_GetSpaceFieldID_2009 2009"
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
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Space number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, 2009
        Else
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=feature&id=" & Me![Space number])
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
Private Sub cmdPrintSpaceSheet_Click()
On Error GoTo err_cmdSpace
    Dim resp, both
    both = MsgBox("Do you want a list of the associated Units as well?", vbQuestion + vbYesNo, "Units?")
        DoCmd.OpenReport "R_SpaceSheet", acViewPreview, , "[Space Number] = " & Me![Space number]
        If both = vbYes Then DoCmd.OpenReport "R_Units_in_Spaces", acViewPreview, , "[In_Space] = " & Me![Space number]
Exit Sub
err_cmdSpace:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdReportProblem_Click()
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "space;" & Me![Space number]
Exit Sub
err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Excavation_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String
    DoCmd.Close acForm, "Exca: Space Sheet"
End Sub
Private Sub Field26_AfterUpdate()
Me![Mound] = Me!Field26.Column(1)
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_Form_BeforeUpdate
If IsNull(Me![Space number] And (Not IsNull(Me![Field26]) Or Not IsNull(Me![Building]) Or Not IsNull(Me![Level]) Or (Not IsNull(Me![Description]) And Me![Description] <> ""))) Then
    MsgBox "You must enter a space number otherwise the record cannot be saved." & Chr(13) & Chr(13) & "If you wish to cancel this record being entered and start again completely press ESC", vbInformation, "Incomplete data"
    Cancel = True
    DoCmd.GoToControl "Space Number"
ElseIf IsNull(Me![Space number]) And IsNull(Me![Field26]) And IsNull(Me![Building]) And IsNull(Me![Level]) And (IsNull(Me![Description]) Or Me![Description] = "") Then
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
On Error GoTo err_Form_Current
    If Me![Building] = "" Or IsNull(Me![Building]) Then
        Me![cmdGoToBuilding].Enabled = False
    Else
        Me![cmdGoToBuilding].Enabled = True
    End If
    If Me![Space number] <> "" Then
        Me![Space number].Locked = True
        Me![Space number].Enabled = False
        Me![Space number].BackColor = Me.Section(0).BackColor
        Me![Building].SetFocus
    Else
        Me![Space number].Locked = False
        Me![Space number].Enabled = True
        Me![Space number].BackColor = 16777215
        Me![Space number].SetFocus
    End If
    If Me![MellaartLevelCertain] = True Then
        Me![Level].Enabled = True
        Me![cboUncertainLevelStart].Enabled = False
        Me![cboUnCertainLevelEnd].Enabled = False
    Else
        Me![Level].Enabled = False
        Me![cboUncertainLevelStart].Enabled = True
        Me![cboUnCertainLevelEnd].Enabled = True
    End If
Dim imageCount, Imgcaption
backhere:
Imgcaption = "Images of Space"
Me![cmdGoToImage].Caption = Imgcaption
Me![cmdGoToImage].Enabled = True
If Me!chkExternal = True Then
    Me![Exca: subform Phases related to Space].Enabled = True
    Me![Exca: subform Phases related to Space].Locked = False
Else
    Me![Exca: subform Phases related to Space].Enabled = False
    Me![Exca: subform Phases related to Space].Locked = True
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
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open
If Me.FilterOn = True Or Me.AllowEdits = False Then
    Me![cboFindSpace].Enabled = False
    Me![cmdAddNew].Enabled = False
    Me.AllowAdditions = False
Else
    DoCmd.GoToControl "cboFindSpace"
End If
Dim permiss
permiss = GetGeneralPermissions
If (permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper") And (Me.AllowAdditions = True Or Me.AllowDeletions = True Or Me.AllowEdits = True) Then
    ToggleFormReadOnly Me, False, "NoDeletions"
Else
    ToggleFormReadOnly Me, True
    Me![cmdAddNew].Enabled = False
End If
If Me!chkExternal = True Then
    Me![Exca: subform Phases related to Space].Enabled = True
    Me![Exca: subform Phases related to Space].Locked = False
Else
    Me![Exca: subform Phases related to Space].Enabled = False
    Me![Exca: subform Phases related to Space].Locked = True
End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub frmLevelCertain_AfterUpdate()
On Error GoTo err_frmLevelCertain_AfterUpdate
Dim retVal
If Me![frmLevelCertain] = -1 Then
    If Me![cboUncertainLevelStart] <> "" And Me![cboUnCertainLevelEnd] <> "" Then
        retVal = MsgBox("Do you want the Start Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
        If retVal = vbYes Then
            Me![Level] = Me![cboUncertainLevelStart]
        Else
            retVal = MsgBox("Do you want the End Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
            If retVal = vbYes Then
                Me![Level] = Me![cboUnCertainLevelEnd]
            Else
                retVal = MsgBox("The start and end level fields will now be cleared and you will have to select the Certain level from that list. Are you sure you want to continue?", vbQuestion + vbYesNo, "Uncertain Levels will be cleared")
                If retVal = vbYes Then
                    Me![cboUncertainLevelStart] = ""
                    Me![cboUnCertainLevelEnd] = ""
                Else
                    Me![frmLevelCertain] = 0
                End If
            End If
        End If
    ElseIf Me![cboUncertainLevelStart] <> "" Then
        retVal = MsgBox("Do you want the Start Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
        If retVal = vbYes Then Me![Level] = Me![cboUncertainLevelStart]
        Me![cboUncertainLevelStart] = ""
    ElseIf Me![cboUnCertainLevelEnd] <> "" Then
        retVal = MsgBox("Do you want the End Level to become the certain level for this Space?", vbQuestion + vbYesNo, "Set Level")
        If retVal = vbYes Then Me![Level] = Me![cboUnCertainLevelEnd]
        Me![cboUnCertainLevelEnd] = ""
    End If
    If Me![frmLevelCertain] = -1 Then 'they have decide not to change their mind
        Me![Level].Enabled = True
        Me![cboUncertainLevelStart].Enabled = False
        Me![cboUnCertainLevelEnd].Enabled = False
    End If
Else
    Me![Level].Enabled = False
    If Me![Level] <> "" Then
        Me![cboUncertainLevelStart] = Me![Level]
        Me![Level] = ""
    End If
    Me![cboUncertainLevelStart].Enabled = True
    Me![cboUnCertainLevelEnd].Enabled = True
End If
Exit Sub
err_frmLevelCertain_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Level_NotInList(NewData As String, response As Integer)
End Sub
Private Sub Space_number_AfterUpdate()
On Error GoTo err_Space_Number_AfterUpdate
Dim checknum
If Me![Space number] <> "" Then
    checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = " & Me![Space number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but this Space Number already exists, please enter another number.", vbInformation, "Duplicate Space Number"
        If Not IsNull(Me![Space number].OldValue) Then
            Me![Space number] = Me![Space number].OldValue
        Else
            Dim currBuild, currarea, currdesc, currLevel
            currBuild = Me![Building]
            currarea = Me![Field26]
            currLevel = Me![Level]
            currdesc = Me![Description]
            DoCmd.GoToControl "Space Number"
            Me.Undo
            If Not IsNull(currBuild) Then Me![Building] = currBuild
            If Not IsNull(currarea) Then Me![Field26] = currarea
            If Not IsNull(currLevel) Then Me![Level] = currLevel
            If Not IsNull(currdesc) Then Me![Description] = currdesc
            DoCmd.GoToControl "Description"
            DoCmd.GoToControl "Space Number"
        End If
    End If
End If
Exit Sub
err_Space_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
