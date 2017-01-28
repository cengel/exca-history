Option Explicit
Option Compare Database   'Use database order for string comparisons
Private Sub Area_AfterUpdate()
On Error GoTo err_Area_AfterUpdate
If Me![Area].Column(1) <> "" Then
    Me![Mound] = Me![Area].Column(1)
End If
Exit Sub
err_Area_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Building_AfterUpdate()
On Error GoTo err_Building_AfterUpdate
Dim checknum, msg, retVal, sql
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
                sql = "INSERT INTO [Exca: Building Details] ([Number]) VALUES (" & Me![Building] & ");"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![Building], acFormEdit, acDialog
            End If
        Else
            Me![cmdGoToBuilding].Enabled = True
        End If
    Else
        MsgBox "This Building number is not numeric, it cannot be checked for validity", vbInformation, "Non numeric Entry"
    End If
End If
Exit Sub
err_Building_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Category_AfterUpdate()
On Error GoTo Err_Category_AfterUpdate
Select Case Me.Category
Case "cut"
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = True
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = True
    Me![Exca: Unit Data Categories CUT subform]![Data Category] = "cut"
    Me![Exca: Unit Data Categories CUT subform]![In Situ] = ""
    Me![Exca: Unit Data Categories CUT subform]![Location] = ""
    Me![Exca: Unit Data Categories CUT subform]![Description] = ""
    Me![Exca: Unit Data Categories CUT subform]![Material] = ""
    Me![Exca: Unit Data Categories CUT subform]![Deposition] = ""
    Me![Exca: Unit Data Categories CUT subform]![basal spit] = ""
    Me.Refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "layer"
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    Me![Exca: Unit Data Categories LAYER subform]![Data Category] = ""
    Me![Exca: Unit Data Categories LAYER subform]![In Situ] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Location] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Description] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Material] = ""
    Me![Exca: Unit Data Categories LAYER subform]![Deposition] = ""
    Me![Exca: Unit Data Categories LAYER subform]![basal spit] = ""
    Me.Refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "cluster"
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
    Me![Exca: Unit Data Categories CLUSTER subform]![Data Category] = "cluster"
    Me![Exca: Unit Data Categories CLUSTER subform]![In Situ] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Location] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Description] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Material] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![Deposition] = ""
    Me![Exca: Unit Data Categories CLUSTER subform]![basal spit] = ""
    Me.Refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "skeleton"
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories SKELL subform]![Data Category] = "skeleton"
    Me![Exca: Unit Data Categories SKELL subform]![In Situ] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Location] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Description] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Material] = ""
    Me![Exca: Unit Data Categories SKELL subform]![Deposition] = ""
    Me![Exca: Unit Data Categories SKELL subform]![basal spit] = ""
    Me.Refresh
    Me![Exca: subform Skeleton Sheet].Visible = True
    Me![subform Unit: stratigraphy  same as].Visible = False
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: subform Skeletons same as].Visible = True
    Me![Exca: Unit Data Categories SKELL subform].Visible = True
End Select
Exit Sub
Err_Category_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindUnit_AfterUpdate()
On Error GoTo err_cboFindUnit_AfterUpdate
    If Me![cboFindUnit] <> "" Then
        If Me![Unit Number].Enabled = False Then Me![Unit Number].Enabled = True
        DoCmd.GoToControl "Unit Number"
        DoCmd.FindRecord Me![cboFindUnit]
        Me![cboFindUnit] = ""
    End If
Exit Sub
err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    DoCmd.GoToControl "Unit Number"
Exit Sub
err_cmdAddNew_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdGoToBuilding_Click()
On Error GoTo Err_cmdGoToBuilding_Click
Dim checknum, msg, retVal, sql, permiss
If Not IsNull(Me![Building]) Or Me![Building] <> "" Then
    checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
    If IsNull(checknum) Then
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Then
            msg = "This Building Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
            If retVal = vbNo Then
                MsgBox "No Building record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
            Else
                sql = "INSERT INTO [Exca: Building Details] ([Number]) VALUES (" & Me![Building] & ");"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![Building], acFormEdit, acDialog
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
Dim tmptable As TableDef, tblConn, I, msg
Set mydb = CurrentDb
    For I = 0 To mydb.TableDefs.count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
    If tmptable.Connect <> "" Then
        tblConn = tmptable.Connect
        Exit For
    End If
    Next I
    If tblConn <> "" Then
        If InStr(tblConn, "catalsql") = 0 Then
            DoCmd.OpenForm "Image_Display", acNormal, , "[Unit] = '" & Me![Unit Number] & "'", acFormReadOnly, acDialog
        Else
            msg = "As you are working remotely the system will have to display the images in a web browser." & Chr(13) & Chr(13)
            msg = msg & "At present this part of the website is secure, you must enter following details to gain access:" & Chr(13) & Chr(13)
            msg = msg & "Username: catalhoyuk" & Chr(13)
            msg = msg & "Password: SiteDatabase1" & Chr(13) & Chr(13)
            msg = msg & "When you have finished viewing the images close your browser to return to the database."
            MsgBox msg, vbInformation, "Photo Web Link"
            Application.FollowHyperlink (ImageLocationOnWeb & "?field=unit&id=" & Me![Unit Number])
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
Private Sub cmdGoToSpace_Click()
On Error GoTo Err_cmdGoToSpace_Click
Dim checknum, msg, retVal, sql, permiss
If Not IsNull(Me![Space]) Or Me![Space] <> "" Then
    checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = '" & Me![Space] & "'")
    If IsNull(checknum) Then
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Then
            msg = "This Space Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
            If retVal = vbNo Then
                MsgBox "No Space record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
            Else
                sql = "INSERT INTO [Exca: Space Sheet] ([Space Number]) VALUES ('" & Me![Space] & "');"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = '" & Me![Space] & "'", acFormEdit, acDialog
            End If
        Else
            MsgBox "Sorry but this space record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Space Record"
        End If
    Else
        Dim stDocName As String
        Dim stLinkCriteria As String
        stDocName = "Exca: Space Sheet"
        stLinkCriteria = "[Space Number]= '" & Me![Space] & "'"
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly, , "FILTER"
    End If
End If
Exit Sub
Err_cmdGoToSpace_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub copy_method_Click()
On Error GoTo Err_copy_method_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Copy unit methodology"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_copy_method_Click:
    Exit Sub
Err_copy_method_Click:
    MsgBox Err.Description
    Resume Exit_copy_method_Click
End Sub
Private Sub Excavation_Click()
On Error GoTo err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Unit Sheet"
Exit_Excavation_Click:
    Exit Sub
err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub
Private Sub FastTrack_Click()
On Error GoTo err_FastTrack_Click
    If Me![FastTrack] = True Then
        Me![NotExcavated] = False
    End If
Exit Sub
err_FastTrack_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Sub find_unit_Click()
On Error GoTo Err_find_unit_Click
    Screen.PreviousControl.SetFocus
    Unit_Number.SetFocus
    DoCmd.DoMenuItem acFormBar, acEditMenu, 10, , acMenuVer70
Exit_find_unit_Click:
    Exit Sub
Err_find_unit_Click:
    MsgBox Err.Description
    Resume Exit_find_unit_Click
End Sub
Private Sub Form_AfterInsert()
On Error GoTo err_Form_AfterInsert
Me![Date changed] = Now()
Exit Sub
err_Form_AfterInsert:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_AfterUpdate()
On Error GoTo err_Form_AfterUpdate
Me![Date changed] = Now()
Exit Sub
err_Form_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
On Error GoTo err_Form_BeforeUpdate
If IsNull(Me![Exca: subform Unit Plans].Form![Graphic Number]) Then
    If Me.ActiveControl.Name = "Discussion" Or Me.ActiveControl.Name = "Checked By" Or Me.ActiveControl.Name = "Date Checked" Or Me.ActiveControl.Name = "Phase" Then
        MsgBox "There is no Plan number entered for this Unit. Please can you enter one soon", vbInformation, "What is the Plan Number?"
    End If
End If
Me![Date changed] = Now()
Exit Sub
err_Form_BeforeUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Current()
Dim stDocName As String
Dim stLinkCriteria As String
On Error GoTo err_Form_Current
Dim permiss
permiss = GetGeneralPermissions
If permiss = "ADMIN" Or permiss = "RW" Then
    If IsNull(Me![Unit Number]) Or Me![Unit Number] = "" Then 'make rest of fields read only
        ToggleFormReadOnly Me, True, "Additions" 'code in GeneralProcedures-shared
        Me![lblMsg].Visible = True
        Me![Unit Number].Locked = False
        Me![Unit Number].Enabled = True
        Me![Unit Number].BackColor = 16777215
        Me![Unit Number].SetFocus
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
            Me![Unit Number].Locked = True
            Me![Unit Number].Enabled = False
            Me![Unit Number].BackColor = Me.Section(0).BackColor
        End If
        Me![lblMsg].Visible = False
    End If
End If
Me![Text407].Locked = True
If Me![Priority Unit] = True Then
    Me![Open Priority].Enabled = True
Else
    Me![Open Priority].Enabled = False
End If
Me![Exca: Unit Data Categories CUT subform].Visible = False
Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
Me![Exca: Unit Data Categories LAYER subform].Visible = False
Me![Exca: Unit Data Categories SKELL subform].Visible = False
Me![Description].TabStop = True
Me![Recognition].TabStop = True
Me![Definition].TabStop = True
Me![Execution].TabStop = True
Me![Condition].TabStop = True
Dim imageCount, Imgcaption
imageCount = DCount("[Unit]", "Image_Metadata_Units", "[Unit] = '" & Me![Unit Number] & "'")
If imageCount > 0 Then
    Imgcaption = imageCount
    If imageCount = 1 Then
        Imgcaption = Imgcaption & " Image to Display"
    Else
        Imgcaption = Imgcaption & " Images to Display"
    End If
    Me![cmdGoToImage].Caption = Imgcaption
    Me![cmdGoToImage].Enabled = True
Else
    Me![cmdGoToImage].Caption = "No Image to Display"
    Me![cmdGoToImage].Enabled = False
End If
Select Case Me.Category
Case "layer"
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "cut"
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = True
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = True
    Me.Refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "cluster"
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
    Me.Refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "skeleton"
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me.Refresh
    Me![Exca: subform Skeleton Sheet].Visible = True
    Me![subform Unit: stratigraphy  same as].Visible = False
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: subform Skeletons same as].Visible = True
    Me![Exca: Unit Data Categories SKELL subform].Visible = True
    Me![Description].TabStop = False
    Me![Recognition].TabStop = False
    Me![Definition].TabStop = False
    Me![Execution].TabStop = False
    Me![Condition].TabStop = False
Case Else
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
End Select
Exit Sub
err_Form_Current: 'SAJ
    General_Error_Trap 'sub in generalprocedures module
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open:
Dim permiss
    permiss = GetGeneralPermissions
    If permiss = "ADMIN" Or permiss = "RW" Then
    Else
        ToggleFormReadOnly Me, True
        Me![cmdAddNew].Enabled = False
        Me![Unit Number].BackColor = Me.Section(0).BackColor
        Me![Unit Number].Locked = True
        Me![copy_method].Enabled = False
    End If
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        Me![cboFindUnit].Enabled = False
        Me![cmdAddNew].Enabled = False
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Sub go_next_Click()
On Error GoTo Err_go_next_Click
    DoCmd.GoToRecord , , acNext
Exit_go_next_Click:
    Exit Sub
Err_go_next_Click:
    MsgBox Err.Description
    Resume Exit_go_next_Click
End Sub
Sub go_to_first_Click()
On Error GoTo Err_go_to_first_Click
    DoCmd.GoToRecord , , acFirst
Exit_go_to_first_Click:
    Exit Sub
Err_go_to_first_Click:
    MsgBox Err.Description
    Resume Exit_go_to_first_Click
End Sub
Sub go_to_last_Click()
On Error GoTo Err_go_last_Click
    DoCmd.GoToRecord , , acLast
Exit_go_last_Click:
    Exit Sub
Err_go_last_Click:
    MsgBox Err.Description
    Resume Exit_go_last_Click
End Sub
Sub go_previous2_Click()
On Error GoTo Err_go_previous2_Click
    DoCmd.GoToRecord , , acPrevious
Exit_go_previous2_Click:
    Exit Sub
Err_go_previous2_Click:
    MsgBox Err.Description
    Resume Exit_go_previous2_Click
End Sub
Private Sub Master_Control_Click()
On Error GoTo Err_Master_Control_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Catal Data Entry"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Unit Sheet"
Exit_Master_Control_Click:
    Exit Sub
Err_Master_Control_Click:
    MsgBox Err.Description
    Resume Exit_Master_Control_Click
End Sub
Sub New_entry_Click()
End Sub
Sub interpretation_Click()
On Error GoTo Err_interpretation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
    stDocName = "Interpret: Unit Sheet"
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_interpretation_Click:
    Exit Sub
Err_interpretation_Click:
    MsgBox Err.Description
    Resume Exit_interpretation_Click
End Sub
Sub Command466_Click()
On Error GoTo Err_Command466_Click
    DoCmd.DoMenuItem acFormBar, acRecordsMenu, 5, , acMenuVer70
Exit_Command466_Click:
    Exit Sub
Err_Command466_Click:
    MsgBox Err.Description
    Resume Exit_Command466_Click
End Sub
Private Sub NotExcavated_Click()
On Error GoTo err_NotExcavated_Click
Dim checknum, checknum1, sql1
    If Me![NotExcavated] = True Then
        If Me![Priority Unit] = True Then
            checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number])
            If Not IsNull(checknum) Then
                checknum1 = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number] & " AND [Exca: Priority Detail].Priority =1 AND [Exca: Priority Detail].Comment Is Null AND [Exca: Priority Detail].Discussion Is Null")
                If IsNull(checknum1) Then
                    MsgBox "Sorry there is information relating to this Unit as a Priority, you cannot change this Unit to Not Excavated", vbExclamation, "Priority Information"
                    Me![NotExcavated] = False
                Else
                    sql1 = "DELETE * FROM [Exca: Priority Detail] WHERE [Unit number] = " & Me![Unit Number] & ";"
                    DoCmd.RunSQL sql1
                    MsgBox "This Unit is no longer marked a Priority. This action has been allowed because it had no Priority specific information entered.", vbExclamation, "Priority change"
                    GoTo allow_check
                End If
            Else
                GoTo allow_check
            End If
        Else
            GoTo allow_check
        End If
    End If
Exit Sub
allow_check:
    Me![FastTrack] = False
    Me![Priority Unit] = False
Exit Sub
err_NotExcavated_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Sub Open_priority_Click()
On Error GoTo Err_Open_priority_Click
    DoCmd.RunCommand acCmdSaveRecord
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, sql, permiss
    checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Unit Number] = " & Me![Unit Number])
    If IsNull(checknum) Then
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Then
            sql = "INSERT INTO [Exca: Priority Detail] ([Unit Number], [DateSet]) VALUES (" & Me![Unit Number] & ", #" & Date & "#);"
            DoCmd.RunSQL sql
        Else
            MsgBox "Sorry but this unit record has not been added to the priority detail table yet, there is no record to view.", vbInformation, "Missing Priority Record"
        End If
    End If
    stDocName = "Exca: Priority Detail"
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_Open_priority_Click:
    Exit Sub
Err_Open_priority_Click:
    Call General_Error_Trap
    Resume Exit_Open_priority_Click
End Sub
Sub go_feature_Click()
On Error GoTo Err_go_feature_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Feature Sheet"
    stLinkCriteria = "[Feature Number]=" & Forms![Exca: Unit Sheet]![Exca: subform Features for Units]![In_feature]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_go_feature_Click:
    Exit Sub
Err_go_feature_Click:
    MsgBox Err.Description
    Resume Exit_go_feature_Click
End Sub
Sub close_Click()
On Error GoTo err_Excavation_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    DoCmd.Close acForm, "Exca: Unit Sheet"
Exit_Excavation_Click:
    Exit Sub
err_Excavation_Click:
    MsgBox Err.Description
    Resume Exit_Excavation_Click
End Sub
Sub open_copy_details_Click()
On Error GoTo Err_open_copy_details_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Copy unit details form"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_open_copy_details_Click:
    Exit Sub
Err_open_copy_details_Click:
    MsgBox Err.Description
    Resume Exit_open_copy_details_Click
End Sub
Private Sub Priority_Unit_Click()
On Error GoTo err_Priority_Unit_Click
Dim checknum, checknum1, sql, sql1
    If Me![Priority Unit] = True Then
        Me![NotExcavated] = False
        checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Unit Number] = " & Me![Unit Number])
        If IsNull(checknum) Then
            sql = "INSERT INTO [Exca: Priority Detail] ([Unit Number], [DateSet]) VALUES (" & Me![Unit Number] & ", #" & Date & "#);"
            DoCmd.RunSQL sql
        End If
        Me![Open Priority].Enabled = True
    Else
        checknum = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number])
        If Not IsNull(checknum) Then
            checknum1 = DLookup("[Unit Number]", "[Exca: Priority Detail]", "[Exca: Priority Detail].[unit Number] = " & Me![Unit Number] & " AND [Exca: Priority Detail].Priority =1 AND [Exca: Priority Detail].Comment Is Null AND [Exca: Priority Detail].Discussion Is Null")
            If IsNull(checknum1) Then
                MsgBox "Sorry there is information relating to this Unit as a Priority, you cannot uncheck it", vbExclamation, "Priority Information"
                Me![Priority Unit] = True
            Else
                sql1 = "DELETE * FROM [Exca: Priority Detail] WHERE [Unit number] = " & Me![Unit Number] & ";"
                DoCmd.RunSQL sql1
                Me![Open Priority].Enabled = False
            End If
        Else
            Me![Open Priority].Enabled = False
        End If
    End If
Exit Sub
err_Priority_Unit_Click:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Space_AfterUpdate()
On Error GoTo err_Space_AfterUpdate
Dim checknum, msg, retVal, sql
If Me![Space] <> "" Then
    If IsNumeric(Me![Space]) Then
        checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = '" & Me![Space] & "'")
        If IsNull(checknum) Then
            msg = "This Space Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
            If retVal = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                sql = "INSERT INTO [Exca: Space Sheet] ([Space Number]) VALUES ('" & Me![Space] & "');"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Space Sheet", acNormal, , "[Space Number] = '" & Me![Space] & "'", acFormEdit, acDialog
            End If
        Else
            Me![cmdGoToSpace].Enabled = True
        End If
    Else
        MsgBox "This Space number is not numeric, it cannot be checked for validity", vbInformation, "Non numeric Entry"
    End If
End If
Exit Sub
err_Space_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Unit_Number_AfterUpdate()
On Error GoTo err_Unit_Number_AfterUpdate
Dim checknum
If Me![Unit Number] <> "" Then
    checknum = DLookup("[Unit Number]", "[Exca: Unit Sheet]", "[Unit Number] = " & Me![Unit Number])
    If Not IsNull(checknum) Then
        MsgBox "Sorry but the Unit Number " & Me![Unit Number] & " already exists, please enter another number.", vbInformation, "Duplicate Unit Number"
        If Not IsNull(Me![Unit Number].OldValue) Then
            Me![Unit Number] = Me![Unit Number].OldValue
        Else
            DoCmd.GoToControl "Year"
            DoCmd.GoToControl "Unit Number"
            Me![Unit Number].SetFocus
            DoCmd.RunCommand acCmdUndo
        End If
    Else
        ToggleFormReadOnly Me, False
    End If
End If
If Me![Unit Number] <> "" Then Me![lblMsg].Visible = False
Exit Sub
err_Unit_Number_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Unit_number_Exit(Cancel As Integer)
End Sub
Sub Command497_Click()
On Error GoTo Err_Command497_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Skeleton Sheet"
    stLinkCriteria = "[Exca: Unit Sheet.Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_Command497_Click:
    Exit Sub
Err_Command497_Click:
    MsgBox Err.Description
    Resume Exit_Command497_Click
End Sub
Sub go_skell_Click()
On Error GoTo Err_go_skell_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Skeleton Sheet"
    stLinkCriteria = "[Unit Number]=" & Me![Unit Number]
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_go_skell_Click:
    Exit Sub
Err_go_skell_Click:
    MsgBox Err.Description
    Resume Exit_go_skell_Click
End Sub
