Option Compare Database
Option Explicit
Private Sub cmdGoToBuilding_Click()
On Error GoTo Err_cmdGoToBuilding_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retval, sql, insertArea, permiss
    stDocName = "Exca: Building Sheet"
    If Not IsNull(Me![txtIn_Building]) Or Me![txtIn_Building] <> "" Then
        checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![txtIn_Building])
        If IsNull(checknum) Then
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Then
                msg = "This Building Number DOES NOT EXIST in the database."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
                If retval = vbNo Then
                    MsgBox "No Building record to view, please alert the your team leader about this.", vbExclamation, "Missing Building Record"
                Else
                    If Forms![Exca: Unit Sheet]![Area] <> "" Then
                        insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                    Else
                        insertArea = Null
                    End If
                    sql = "INSERT INTO [Exca: Building Details] ([Number], [Area]) VALUES (" & Me![txtIn_Building] & ", " & insertArea & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![txtIn_Building], acFormEdit, acDialog
                End If
            Else
                MsgBox "Sorry but this Building record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Building Record"
            End If
        Else
            stLinkCriteria = "[Number]=" & Me![txtIn_Building]
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        End If
    Else
        MsgBox "No Building number to view", vbInformation, "No Building Number"
    End If
Exit_cmdGoToBuilding_Click:
    Exit Sub
Err_cmdGoToBuilding_Click:
    Call General_Error_Trap
    Resume Exit_cmdGoToBuilding_Click
End Sub
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
End Sub
Private Sub Form_Current()
On Error GoTo err_Current
    If Me![txtIn_Building] = "" Or IsNull(Me![txtIn_Building]) Then
        Me![cmdGoToBuilding].Enabled = False
    Else
        Me![cmdGoToBuilding].Enabled = True
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
    If permiss = "ADMIN" Or permiss = "RW" Then
        ToggleFormReadOnly Me, False
    Else
        ToggleFormReadOnly Me, True
    End If
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub txtIn_Building_AfterUpdate()
On Error GoTo err_txtIn_Building_AfterUpdate
Dim checknum, msg, retval, sql, insertArea
If Me![txtIn_Building] <> "" Then
    If IsNumeric(Me![txtIn_Building]) Then
        If Me![txtIn_Building] = 0 Then
            MsgBox "Building 0 is invalid, please edit this entry", vbInformation, "Invalid Entry"
        Else
            checknum = DLookup("[Number]", "[Exca: Building Details]", "[Number] = " & Me![txtIn_Building])
            If IsNull(checknum) Then
                msg = "This Building Number DOES NOT EXIST in the database, you must remember to enter it."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retval = MsgBox(msg, vbInformation + vbYesNo, "Building Number does not exist")
                If retval = vbNo Then
                    MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
                Else
                    If Forms![Exca: Unit Sheet]![Area] <> "" Then
                        insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                    Else
                        insertArea = Null
                    End If
                    sql = "INSERT INTO [Exca: Building Details] ([Number], [Area]) VALUES (" & Me![txtIn_Building] & ", " & insertArea & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm "Exca: Building Sheet", acNormal, , "[Number] = " & Me![txtIn_Building], acFormEdit, acDialog
                End If
            Else
                Me![cmdGoToBuilding].Enabled = True
            End If
        End If
    Else
        MsgBox "The Building number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If
Exit Sub
err_txtIn_Building_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub txtIn_Building_BeforeUpdate(Cancel As Integer)
On Error GoTo err_buildingbefore
If Me![txtIn_Building] = 0 Then
        MsgBox "Building 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
        Cancel = True
        SendKeys "{ESC}" 'seems to need it done 3x
        SendKeys "{ESC}"
        SendKeys "{ESC}"
End If
Exit Sub
err_buildingbefore:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Unit_AfterUpdate()
Me.Requery
DoCmd.GoToRecord , , acLast
End Sub
