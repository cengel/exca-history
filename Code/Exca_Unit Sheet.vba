Option Explicit
Option Compare Database   'Use database order for string comparisons
Sub UpdateDataCategory()
On Error GoTo err_updatedatacategory
    Dim sql1
    If spString <> "" Then
        Dim mydb As DAO.Database
        Dim myq1 As QueryDef
        Set mydb = CurrentDb
        Set myq1 = mydb.CreateQueryDef("")
        myq1.Connect = spString
            myq1.ReturnsRecords = False
            myq1.sql = "sp_Excavation_Delete_DataCategory_Entry " & Me![Unit Number]
            myq1.Execute
        myq1.Close
        Set myq1 = Nothing
        mydb.Close
        Set mydb = Nothing
    Else
        MsgBox "The data category record has not been deleted, please update it manually.", vbCritical, "Error"
    End If
    If Me![cboExcavationStatus] = "void" Then
        sql1 = "INSERT INTO [Exca: Unit Data Categories] ([Unit Number], [Data Category], Description, [in situ], location, material, deposition) VALUES (" & Me![Unit Number] & ", 'arbitrary', 'void (unused unit no)', '','','','');"
        DoCmd.RunSQL sql1
    ElseIf Me![cboExcavationStatus] = "natural" Then
        sql1 = "INSERT INTO [Exca: Unit Data Categories] ([Unit Number], [Data Category], Description, [in situ], location, material, deposition) VALUES (" & Me![Unit Number] & ", 'natural', '', '','','','');"
        DoCmd.RunSQL sql1
    ElseIf Me![cboExcavationStatus] = "unstratified" Then
        sql1 = "INSERT INTO [Exca: Unit Data Categories] ([Unit Number], [Data Category], Description, [in situ], location, material, deposition) VALUES (" & Me![Unit Number] & ", 'arbitrary', 'unstratified', '','','','');"
        DoCmd.RunSQL sql1
    End If
    Me![Exca: Unit Data Categories LAYER subform].Requery
    Me![Exca: Unit Data Categories CLUSTER subform].Requery
    Me![Exca: Unit Data Categories CUT subform].Requery
    Me![Exca: Unit Data Categories SKELL subform].Requery
    Me![Category] = Me![cboExcavationStatus]
    Call Form_Current 'update screen correctly
Exit Sub
err_updatedatacategory:
    Call General_Error_Trap
    Exit Sub
End Sub
Sub Delete_Category_SubTable_Entry(deleteFrom, Unit)
On Error GoTo err_delete_cat
If spString <> "" Then
    Dim mydb As DAO.Database
    Dim myq1 As QueryDef
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = spString
        myq1.ReturnsRecords = False
        myq1.sql = "sp_Excavation_Delete_Category_SubTable_Entry " & Unit & ", '" & deleteFrom & "'"
        myq1.Execute
    myq1.Close
    Set myq1 = Nothing
    mydb.Close
    Set mydb = Nothing
Else
    MsgBox "The " & deleteFrom & " record cannot be deleted, please restart the database, set this unit back to " & deleteFrom & " and try this change again", vbCritical, "Error"
End If
Exit Sub
err_delete_cat:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Area_AfterUpdate()
On Error GoTo err_Area_AfterUpdate
If Me![Area].Column(1) <> "" Then
    Me![Mound] = Me![Area].Column(1)
    If IsNull(Me![TimePeriod]) Then
        If Me![Mound] = "West" Then
            Me![TimePeriod] = "Chalcolithic"
        ElseIf Me![Mound] = "Off-Site" Then
            Me![TimePeriod] = "Unknown"
        Else
            Me![TimePeriod] = "Neolithic"
        End If
    Else
        Dim response
        If Me![Mound] = "West" And Me![TimePeriod] <> "Chalcolithic" Then
            response = MsgBox("A timeperiod " & Me![TimePeriod] & " has previously been set for this unit. The latest change means the system think it should now be set to Chalcolithic, is this right?", vbQuestion + vbYesNo, "Timeperiod check")
            If response = vbYes Then
                Me![TimePeriod] = "Chalcolithic"
            Else
                MsgBox "The timeperiod has been left as " & Me![TimePeriod] & ". Please let your supervisor know if this is incorrect.", vbInformation, "Timeperiod"
            End If
        ElseIf Me![Mound] = "East" And Me![TimePeriod] <> "Neolithic" Then
            response = MsgBox("A timeperiod " & Me![TimePeriod] & " has previously been set for this unit. The latest change means the system think it should now be set to Neolithic, is this right?", vbQuestion + vbYesNo, "Timeperiod check")
            If response = vbYes Then
                Me![TimePeriod] = "Neolithic"
            Else
                MsgBox "The timeperiod has been left as " & Me![TimePeriod] & ". Please let your supervisor know if this is incorrect.", vbInformation, "Timeperiod"
            End If
        ElseIf Me![Mound] = "Off-Site" And Me![TimePeriod] <> "Unknown" Then
            response = MsgBox("A timeperiod " & Me![TimePeriod] & " has previously been set for this unit. The latest change means the system think it should now be set to Unknown, is this right?", vbQuestion + vbYesNo, "Timeperiod check")
            If response = vbYes Then
                Me![TimePeriod] = "Unknown"
            Else
                MsgBox "The timeperiod has been left as " & Me![TimePeriod] & ". Please let your supervisor know if this is incorrect.", vbInformation, "Timeperiod"
            End If
        End If
    End If
End If
If Me![Area] <> "" Then
    Me![cboFT].RowSource = "SELECT [Exca: Foundation Trench Description].FTName, [Exca: Foundation Trench Description].Description, [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder FROM [Exca: Foundation Trench Description] WHERE [Area] = '" & Me![Area] & "' ORDER BY [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder;"
End If
Exit Sub
err_Area_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Building_AfterUpdate()
On Error GoTo err_Building_AfterUpdate
Dim checknum, msg, retval, sql
If Me![Building] <> "" Then
    If IsNumeric(Me![Building]) Then
        checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
        If IsNull(checknum) Then
            msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
            If retval = vbNo Then
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
If Me![Category].OldValue <> "" Or Not IsNull(Me![Category]) Then
    If Not ((Me![Category].OldValue = "cluster" Or Me![Category].OldValue = "layer") And (Me![Category] = "cluster" Or Me![Category] = "layer")) Then
        Dim checkit
        checkit = Null
        If Me![Category].OldValue = "cut" Then 'check for cut info
            checkit = DLookup("[Unit Number]", "[Exca: Descriptions Cut]", "[Unit Number] = " & Me![Unit Number])
        ElseIf Me![Category].OldValue = "layer" Or Me![Category].OldValue = "cluster" Then
            checkit = DLookup("[Unit Number]", "[Exca: Descriptions Layer]", "[Unit Number] = " & Me![Unit Number])
        ElseIf Me![Category].OldValue = "skeleton" Then
            checkit = DLookup("[Unit Number]", "[Exca: Skeleton Data]", "[Unit Number] = " & Me![Unit Number])
        End If
        If Not IsNull(checkit) Then
            Dim resp, sql
            resp = MsgBox("By changing the category of this Unit you will lose the " & Me![Category].OldValue & " specific data (if any). Do you still want to change the category?", vbQuestion + vbYesNo, "Confirm Action")
            If resp = vbNo Then
                Me![Category] = Me![Category].OldValue
            ElseIf resp = vbYes Then
                If Me![Category].OldValue = "layer" Or Me![Category].OldValue = "cluster" Then
                    Call Delete_Category_SubTable_Entry("layer", Me![Unit Number])
                ElseIf Me![Category].OldValue = "cut" Then
                    Call Delete_Category_SubTable_Entry("cut", Me![Unit Number])
                ElseIf Me![Category].OldValue = "skeleton" Then
                    Call Delete_Category_SubTable_Entry("skeleton", Me![Unit Number])
                End If
            End If
        End If
    End If
End If
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
    Me.refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![Exca: subform Skeleton Sheet 2013].Visible = False
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
    Me.refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![Exca: subform Skeleton Sheet 2013].Visible = False
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
    Me.refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![Exca: subform Skeleton Sheet 2013].Visible = False
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "skeleton"
    If Forms![Exca: Unit Sheet]!Year < 2013 Then
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
        Me.refresh
        Me![Exca: subform Skeleton Sheet].Visible = True
        Me![Exca: subform Skeleton Sheet 2013].Visible = False
        Me![subform Unit: stratigraphy  same as].Visible = False
        Me![Exca: Subform Layer descr].Visible = False
        Me![Exca: Subform Cut descr].Visible = False
        Me![Exca: subform Skeletons same as].Visible = True
        Me![Exca: Unit Data Categories SKELL subform].Visible = True
    Else
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
        Me.refresh
        Me![Exca: subform Skeleton Sheet].Visible = False
        Me![Exca: subform Skeleton Sheet 2013].Visible = True
        Me![subform Unit: stratigraphy  same as].Visible = False
        Me![Exca: Subform Layer descr].Visible = False
        Me![Exca: Subform Cut descr].Visible = False
        Me![Exca: subform Skeletons same as].Visible = True
        Me![Exca: Unit Data Categories SKELL subform].Visible = True
    End If
End Select
Exit Sub
Err_Category_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboExcavationStatus_AfterUpdate()
On Error GoTo err_cboExcaStatus
    If Me![cboExcavationStatus] <> "excavated" And Me![cboExcavationStatus] <> "not excavated" And Me![cboExcavationStatus] <> "partially excavated" Then
        If Me![Category] = "" Or IsNull(Me![Category]) Then
            Me![Category] = Me![cboExcavationStatus]
            Me![Category].Locked = True
            Me![Category].Enabled = False
        Else
            If Me![Category] = "cut" Or Me![Category] = "skeleton" Then
                Dim checkit
                checkit = Null
                If Me![Category] = "cut" Then 'check for cut info
                    checkit = DLookup("[Unit Number]", "[Exca: Descriptions Cut]", "[Unit Number] = " & Me![Unit Number])
                ElseIf Me![Category].OldValue = "skeleton" Then
                    checkit = DLookup("[Unit Number]", "[Exca: Skeleton Data]", "[Unit Number] = " & Me![Unit Number])
                End If
                If Not IsNull(checkit) Then
                    Dim resp, sql
                    resp = MsgBox("By changing the status of this Unit you will lose the " & Me![Category].OldValue & " specific data (if any). Do you still want to change the status?", vbQuestion + vbYesNo, "Confirm Action")
                    If resp = vbNo Then
                        Me![cboExcavationStatus] = Me![cboExcavationStatus].OldValue
                    ElseIf resp = vbYes Then
                        If Me![Category] = "cut" Then
                            Call Delete_Category_SubTable_Entry("cut", Me![Unit Number])
                            Call UpdateDataCategory 'local sub
                        ElseIf Me![Category] = "skeleton" Then
                            Call Delete_Category_SubTable_Entry("skeleton", Me![Unit Number])
                            Call UpdateDataCategory 'local sub
                        End If
                    End If
                Else
                    Call UpdateDataCategory 'local sub
                End If
                Me![Category].Locked = True
                Me![Category].Enabled = False
            Else 'If Me![Category] = "void" Or Me![Category] = "natural" Or Me![Category] = "unstratified" Then
                Call UpdateDataCategory 'local sub
            End If
        End If
    Else
        If Me![cboExcavationStatus].OldValue <> "Excavated" And Me![cboExcavationStatus].OldValue <> "Not Excavated" Then
            Call UpdateDataCategory
            Me![Category] = "Layer" 'set default to layer
            Me![Category].Locked = False
            Me![Category].Enabled = True
        Else
        End If
    End If
Exit Sub
err_cboExcaStatus:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindUnit_AfterUpdate()
On Error GoTo err_cboFindUnit_AfterUpdate
    If Me![cboFindUnit] <> "" Then
        If Me![Unit Number].Enabled = False Then Me![Unit Number].Enabled = True
        DoCmd.GoToControl "Unit Number"
        DoCmd.FindRecord Me![cboFindUnit]
        DoCmd.GoToControl "cboFindUnit"
        Me![cboFindUnit] = ""
    End If
Exit Sub
err_cboFindUnit_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFindUnit_NotInList(NewData As String, response As Integer)
On Error GoTo err_cbofindNot
    MsgBox "Sorry this Unit cannot be found in the list", vbInformation, "No Match"
    response = acDataErrContinue
    Me![cboFindUnit].Undo
    SendKeys "{ESC}"
Exit Sub
err_cbofindNot:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cboFT_AfterUpdate()
If Me![cboFT] <> "" Then
    Me![cmdGoToFT].Enabled = True
End If
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
Private Sub cmdEditPhase_Click()
On Error GoTo err_edit
If Not IsNull(Me![Unit Number]) Then
    Dim checkB, checkSp, getB, getSp, counter, response
    checkB = DCount("[In_Building]", "[Exca: Units in Buildings]", "[Unit] = " & Me![Unit Number])
    If checkB = 0 Then
        checkSp = DCount("[In_Space]", "[Exca: Units in Spaces]", "[Unit] = " & Me![Unit Number])
        If checkSp = 0 Then
            MsgBox "This unit is not associated with a Building or a Space so it cannot be phased in this way", vbInformation, "Nothing to Phase"
            Exit Sub
        Else
            If checkSp > 1 Then
                counter = 1
                Me![Exca: subform  Features in Spaces].Form.RecordsetClone.MoveFirst
                Do Until counter > checkSp
                    response = MsgBox("Do you want to phase this unit to Space " & Me![Exca: subform  Features in Spaces].Form.RecordsetClone(1).Value & "?" & _
                                Chr(13) & Chr(13) & "Clicking No will prompt the question for the next Space in the list if there are more.", vbQuestion + vbYesNoCancel, "Which Space to phase now?")
                    If response = vbYes Then
                        DoCmd.OpenForm "frm_pop_phase_a_unit", acNormal, , , acFormPropertySettings, acDialog, "SELECT [Exca: SpacePhases].SpacePhase FROM [Exca: SpacePhases] WHERE [Exca: SpacePhases].SpaceNumber=" & Me![Exca: subform  Features in Spaces].Form.RecordsetClone(1).Value & ";" 'open form with space number
                        Exit Do
                    ElseIf response = vbCancel Then
                        Exit Do
                    End If
                    Me![Exca: subform  Features in Spaces].Form.RecordsetClone.MoveNext
                    counter = counter + 1
                Loop
            Else
                getSp = DLookup("[In_Space]", "[Exca: Units in Spaces]", "[Unit] = " & Me![Unit Number])
                DoCmd.OpenForm "frm_pop_phase_a_unit", acNormal, , , acFormPropertySettings, acDialog, "SELECT [Exca: SpacePhases].SpacePhase FROM [Exca: SpacePhases] WHERE [Exca: SpacePhases].SpaceNumber=" & getSp & ";" 'open form with space number
            End If
        End If
    Else
        If checkB > 1 Then
            MsgBox "This unit is associated with more than one Building. Currently the system does not support phasing a unit to more than one building. Please discuss this with Shahina.", vbInformation, "Multiple Building Numbers"
            Exit Sub
        Else
            getB = DLookup("[In_Building]", "[Exca: Units in Buildings]", "[Unit] = " & Me![Unit Number])
            DoCmd.OpenForm "frm_pop_phase_a_unit", acNormal, , , acFormPropertySettings, acDialog, "SELECT [Exca: BuildingPhases].BuildingPhase FROM [Exca: BuildingPhases] WHERE [Exca: BuildingPhases].BuildingNumber=" & getB & ";" 'open form with building number
        End If
    End If
    Me![Exca: subform Units Occupation Phase].Requery
End If
Exit Sub
err_edit:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdGoToBuilding_Click()
On Error GoTo Err_cmdGoToBuilding_Click
Dim checknum, msg, retval, sql, permiss
If Not IsNull(Me![Building]) Or Me![Building] <> "" Then
    checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![Building])
    If IsNull(checknum) Then
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
            msg = "This Building Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
            If retval = vbNo Then
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
Private Sub cmdGoToFT_Click()
On Error GoTo Err_cmdGoToFT_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retval, sql, insertArea, permiss
    stDocName = "Exca: Admin_Foundation_Trenches"
    If Not IsNull(Me![cboFT]) Or Me![cboFT] <> "" Then
        stLinkCriteria = "[FTName]='" & Me![cboFT] & "' AND [Area] = '" & Me![Area] & "'"
        DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
    Else
        MsgBox "No FT record to view", vbInformation, "No FT Name"
    End If
Exit_cmdGoToFT_Click:
    Exit Sub
Err_cmdGoToFT_Click:
    Call General_Error_Trap
    Resume Exit_cmdGoToFT_Click
End Sub
Private Sub cmdGoToImage_Click()
On Error GoTo err_cmdGoToImage_Click
Dim mydb As DAO.Database
Dim tmptable As TableDef, tblConn, I, msg, fldid
Set mydb = CurrentDb
    Dim myq1 As QueryDef, connStr
    Set mydb = CurrentDb
    Set myq1 = mydb.CreateQueryDef("")
    myq1.Connect = mydb.TableDefs("view_Portfolio_Previews_2008").Connect & ";UID=portfolio;PWD=portfolio"
    myq1.ReturnsRecords = True
    myq1.sql = "sp_Portfolio_GetUnitFieldID_2008 " & Me![Year]
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
            DoCmd.OpenForm "Image_Display", acNormal, , "[IntValue] = " & Me![Unit Number] & " AND [Field_ID] = " & fldid, acFormReadOnly, acDialog, Me![Year]
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
Dim checknum, msg, retval, sql, permiss
If Not IsNull(Me![Space]) Or Me![Space] <> "" Then
    checknum = DLookup("[Space Number]", "[Exca: Space Sheet]", "[Space Number] = '" & Me![Space] & "'")
    If IsNull(checknum) Then
        permiss = GetGeneralPermissions
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
            msg = "This Space Number DOES NOT EXIST in the database."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
            If retval = vbNo Then
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
Private Sub cmdPrintUnitSheet_Click()
On Error GoTo err_print
    If LCase(Me![Category]) = "layer" Or LCase(Me![Category]) = "cluster" Then
        DoCmd.OpenReport "R_Unit_Sheet_layercluster", acViewPreview, , "[unit number] = " & Me![Unit Number]
    ElseIf LCase(Me![Category]) = "cut" Then
        DoCmd.OpenReport "R_Unit_Sheet_cut", acViewPreview, , "[unit number] = " & Me![Unit Number]
    ElseIf LCase(Me![Category]) = "skeleton" Then
        DoCmd.OpenReport "R_Unit_Sheet_skeleton", acViewPreview, , "[unit number] = " & Me![Unit Number]
    End If
Exit Sub
err_print:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdReportProblem_Click()
On Error GoTo err_reportprob
    DoCmd.OpenForm "frm_pop_problemreport", , , , acFormAdd, acDialog, "unit number;" & Me![Unit Number]
Exit Sub
err_reportprob:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdTypeLOV_Click()
On Error GoTo err_typeLOV
    DoCmd.OpenForm "Frm_subform_sampletypeLOV", acNormal, , , acFormReadOnly, acDialog
Exit Sub
err_typeLOV:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub cmdViewSketch_Click()
On Error GoTo err_opensketch
    Dim Path
    Dim fname
    If Me![Year] < 2015 Then
    Path = sketchpath
    Path = Path & Me![Unit Number] & ".jpg"
    Else
    Path = sketchpath2015
    Path = Path & "units\sketches\" & "U" & Me![Unit Number] & "*" & ".jpg"
    fname = Dir(Path & "*", vbNormal)
    While fname <> ""
        Debug.Print fname
        fname = Dir()
    Wend
    Path = sketchpath2015 & "units\sketches\" & fname
    End If
    If Dir(Path) = "" Then
        MsgBox "The sketch plan of this unit has not been scanned in yet.", vbInformation, "No Sketch available to view"
    Else
        DoCmd.OpenForm "frm_pop_graphic", acNormal, , , acFormReadOnly, , Me![Unit Number]
    End If
Exit Sub
err_opensketch:
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
Private Sub Exca__subform_Skeleton_Sheet_Enter()
End Sub
Private Sub FastTrack_Click()
On Error GoTo err_FastTrack_Click
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
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
Exit Sub
err_Form_AfterInsert:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_AfterUpdate()
On Error GoTo err_Form_AfterUpdate
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
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
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
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
If (permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper") And ([Unit Number] <> 0 Or IsNull([Unit Number])) Then
    If IsNull(Me![Unit Number]) Or Me![Unit Number] = "" Then
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
    Imgcaption = "Images of Unit"
    Me![cmdGoToImage].Caption = Imgcaption
    Me![cmdGoToImage].Enabled = True
Dim Path
Path = sketchpath
Path = Path & Me![Unit Number] & ".jpg"
        Me!cmdViewSketch.Enabled = True
Select Case Me.Category
Case "layer"
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "cut"
    Me![Exca: Subform Layer descr].Visible = False
    Me![Exca: Subform Cut descr].Visible = True
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = True
    Me.refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "cluster"
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = True
    Me.refresh
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
Case "skeleton"
    If Me![Year] < 2013 Then
        Me![Exca: Unit Data Categories CUT subform].Visible = False
        Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
        Me![Exca: Unit Data Categories LAYER subform].Visible = False
        Me.refresh
        Me![Exca: subform Skeleton Sheet].Visible = True
        Me![Exca: subform Skeleton Sheet 2013].Visible = False
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
    Else
        Me![Exca: Unit Data Categories CUT subform].Visible = False
        Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
        Me![Exca: Unit Data Categories LAYER subform].Visible = False
        Me.refresh
        Me![Exca: subform Skeleton Sheet].Visible = False
        Me![Exca: subform Skeleton Sheet 2013].Visible = True
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
    End If
Case Else
    Me![Exca: Subform Layer descr].Visible = True
    Me![Exca: Subform Cut descr].Visible = False
    Me![Exca: Unit Data Categories CUT subform].Visible = False
    Me![Exca: Unit Data Categories CLUSTER subform].Visible = False
    Me![Exca: Unit Data Categories LAYER subform].Visible = True
    Me![Exca: subform Skeleton Sheet].Visible = False
    Me![Exca: subform Skeleton Sheet 2013].Visible = False 'skelli 2013
    Me![subform Unit: stratigraphy  same as].Visible = True
    Me![Exca: subform Skeletons same as].Visible = False
    Me![Exca: Unit Data Categories SKELL subform].Visible = False
End Select
If Me![Area] <> "" Then
    Me![cboFT].RowSource = "SELECT [Exca: Foundation Trench Description].FTName, [Exca: Foundation Trench Description].Description, [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder FROM [Exca: Foundation Trench Description] WHERE [Area] = '" & Me![Area] & "' ORDER BY [Exca: Foundation Trench Description].Area, [Exca: Foundation Trench Description].DisplayOrder;"
End If
If Me![cboFT] <> "" Then
    Me![cmdGoToFT].Enabled = True
Else
    Me![cmdGoToFT].Enabled = False
End If
If (permiss = "ADMIN" Or permiss = "exsuper") Then
    Me![cboFT].Locked = False
    Me![cboFT].Enabled = True
    Me![cmdEditPhase].Enabled = True
    Me![cboTimePeriod].Locked = False
    Me![cboTimePeriod].Enabled = True
Else
    Me![cboFT].Locked = True
    Me![cboFT].Enabled = False
    Me![cmdEditPhase].Enabled = False
    Me![cboTimePeriod].Locked = True
    Me![cboTimePeriod].Enabled = False
End If
If Me![cboExcavationStatus] <> "excavated" And Me![cboExcavationStatus] <> "not excavated" Then
    Me![Category].Enabled = False
Else
    Me![Category].Enabled = True
End If
If permiss <> "ADMIN" And (Me![Year] >= 2003 And Me![Year] <= 2008 And (Me![Area] = "4040" Or Me![Area] = "South")) Then
    Me![TotalSampleAmount].Enabled = False
    Me![Dry sieve volume].Enabled = False
    Me![RemainingVolume].Enabled = False
    Me![TotalDepositVolume].Enabled = False
    Me![HowVolumeCalc].Enabled = False
Else
    Me![TotalSampleAmount].Enabled = True
    Me![Dry sieve volume].Enabled = True
    Me![TotalDepositVolume].Enabled = True
    Me![HowVolumeCalc].Enabled = True
        Me![RemainingVolume].Enabled = True
End If
Exit Sub
err_Form_Current: 'SAJ
    General_Error_Trap 'sub in generalprocedures module
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_Form_Open:
Dim permiss
    permiss = GetGeneralPermissions
    If (permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper") And ([Unit Number] <> 0) Then
    Else
        ToggleFormReadOnly Me, True
        If permiss <> "ADMIN" And permiss <> "RW" And permiss <> "exsuper" Then
            Me![cmdAddNew].Enabled = False
        ElseIf permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
            Me.AllowAdditions = True 'this ensures rw can add record right from start
        End If
        Me![Unit Number].BackColor = Me.Section(0).BackColor
        Me![Unit Number].Locked = True
        Me![copy_method].Enabled = False
    End If
    If Me.FilterOn = True Or Me.AllowEdits = False Then
        Me![cboFindUnit].Enabled = False
        Me![cmdAddNew].Enabled = False
    Else
        DoCmd.GoToControl "cboFindUnit"
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
        If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
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
Sub Close_Click()
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
Private Sub print_bulk_Click()
On Error GoTo Err_print_bulk_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "print_bulk_units"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_print_bulk_Click:
    Exit Sub
Err_print_bulk_Click:
    Call General_Error_Trap
    Resume Exit_print_bulk_Click
End Sub
Private Sub Priority_Unit_Click()
On Error GoTo err_Priority_Unit_Click
Dim checknum, checknum1, sql, sql1
    If Me![Priority Unit] = True Then
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
Dim checknum, msg, retval, sql
If Me![Space] <> "" Then
    If IsNumeric(Me![Space]) Then
        checknum = DLookup("[Category]", "[Exca: Space Sheet]", "[Space Number] = '" & Me![Space] & "'")
        If IsNull(checknum) Then
            msg = "This Space Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Space Number does not exist")
            If retval = vbNo Then
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
Private Sub Year_AfterUpdate()
End Sub
Private Sub Year_LostFocus()
On Error GoTo err_Year
    If IsNull(Me![cboFindUnit]) Then
        If IsNull(Me![Year]) Or Me![Year] = "" Then
            MsgBox "You must enter the year this unit number was excavated, or allocated if not excavated yet", vbInformation, "Invalid Year"
            DoCmd.GoToControl "Area"
            DoCmd.GoToControl "Year"
            Me![Year].SetFocus
        ElseIf Me![Year] < 1993 Or Me![Year] > ThisYear Then
            MsgBox Me![Year] & " is not a valid Year please try again", vbInformation, "Invalid Year"
            DoCmd.GoToControl "Area"
            DoCmd.GoToControl "Year"
            Me![Year].SetFocus
        End If
    End If
Exit Sub
err_Year:
    Exit Sub
End Sub
