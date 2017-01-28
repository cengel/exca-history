Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Unit Sheet]![dbo_Exca: UnitHistory].Form![lastmodify].Value = Now()
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
Private Sub go_to_feature_Click()
On Error GoTo Err_go_to_feature_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim checknum, msg, retVal, sql, insertArea, permiss
    stDocName = "Exca: Feature Sheet"
    If Not IsNull(Me![In_feature]) Or Me![In_feature] <> "" Then
        checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![In_feature])
        If IsNull(checknum) Then
            permiss = GetGeneralPermissions
            If permiss = "ADMIN" Or permiss = "RW" Or permiss = "exsuper" Then
                msg = "This Feature Number DOES NOT EXIST in the database."
                msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
                retVal = MsgBox(msg, vbInformation + vbYesNo, "Feature Number does not exist")
                If retVal = vbNo Then
                    MsgBox "No feature record to view, please alert the your team leader about this.", vbExclamation, "Missing Feature Record"
                Else
                    If Forms![Exca: Unit Sheet]![Area] <> "" Then
                        insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                    Else
                        insertArea = Null
                    End If
                    sql = "INSERT INTO [Exca: Features] ([Feature Number], [Area]) VALUES (" & Me![In_feature] & ", " & insertArea & ");"
                    DoCmd.RunSQL sql
                    DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , "[Feature Number] = " & Me![In_feature], acFormEdit, acDialog
                End If
            Else
                MsgBox "Sorry but this feature record has not been added to the system yet, there is no record to view.", vbInformation, "Missing Feature Record"
            End If
        Else
            stLinkCriteria = "[Feature Number]=" & Me![In_feature]
            DoCmd.OpenForm stDocName, acNormal, , stLinkCriteria, acFormReadOnly
        End If
    Else
        MsgBox "No Feature number to view", vbInformation, "No Feature Number"
    End If
Exit_go_to_feature_Click:
    Exit Sub
Err_go_to_feature_Click:
    Call General_Error_Trap
    Resume Exit_go_to_feature_Click
End Sub
Private Sub In_feature_AfterUpdate()
On Error GoTo err_In_feature_AfterUpdate
Dim checknum, msg, retVal, sql, insertArea
If Me![In_feature] <> "" Then
    If IsNumeric(Me![In_feature]) Then
        checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![In_feature])
        If IsNull(checknum) Then
            msg = "This Feature Number DOES NOT EXIST in the database, you must remember to enter it."
            msg = msg & Chr(13) & Chr(13) & "Would you like to enter it now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Feature Number does not exist")
            If retVal = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                If Forms![Exca: Unit Sheet]![Area] <> "" Then
                    insertArea = "'" & Forms![Exca: Unit Sheet]![Area] & "'"
                Else
                    insertArea = Null
                End If
                sql = "INSERT INTO [Exca: Features] ([Feature Number], [Area]) VALUES (" & Me![In_feature] & ", " & insertArea & ");"
                DoCmd.RunSQL sql
                DoCmd.OpenForm "Exca: Feature Sheet", acNormal, , "[Feature Number] = " & Me![In_feature], acFormEdit, acDialog
            End If
        Else
            Me![go to feature].Enabled = True
        End If
    Else
        MsgBox "The Feature number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If
Exit Sub
err_In_feature_AfterUpdate:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub In_feature_BeforeUpdate(Cancel As Integer)
On Error GoTo err_featurebefore
If Me![In_feature] = 0 Then
        MsgBox "Feature 0 is invalid, this entry will be removed", vbInformation, "Invalid Entry"
        Cancel = True
        SendKeys "{ESC}" 'seems to need it done 3x
        SendKeys "{ESC}"
        SendKeys "{ESC}"
End If
Exit Sub
err_featurebefore:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Unit_AfterUpdate()
Me.Requery
DoCmd.GoToRecord , , acLast
End Sub
Sub Command5_Click()
On Error GoTo Err_Command5_Click
    DoCmd.GoToRecord , , acLast
Exit_Command5_Click:
    Exit Sub
Err_Command5_Click:
    MsgBox Err.Description
    Resume Exit_Command5_Click
End Sub
