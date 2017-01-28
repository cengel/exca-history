Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
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
Private Sub To_feature_AfterUpdate()
On Error GoTo err_To_feature_AfterUpdate
Dim checknum, msg, retVal, sql, currentFeature, checknum2, featureRel
If Me![To_feature] <> "" Then
    If IsNumeric(Me![To_feature]) Then
        checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![To_feature])
        If IsNull(checknum) Then
            msg = "The Feature Number " & Me![To_feature] & " DOES NOT EXIST in the database. The system can enter it for you ready for you to update later."
            msg = msg & Chr(13) & Chr(13) & "Would you like the system to create this feature number now?"
            retVal = MsgBox(msg, vbInformation + vbYesNo, "Feature Number does not exist")
            If retVal = vbNo Then
                MsgBox "Ok, but you must remember to enter it soon otherwise you'll be chased!", vbExclamation, "Remember!"
            Else
                currentFeature = Me![Feature Number]
                sql = "INSERT INTO [Exca: Features] ([Feature Number]) VALUES (" & Me![To_feature] & ");"
                DoCmd.RunSQL sql
                MsgBox "Feature " & Me![To_feature] & " has been created in the database. This screen will now refresh itself.", vbInformation, "System updating"
                DoCmd.Hourglass True
                Forms![Exca: Feature Sheet].Requery
                DoCmd.GoToControl Forms![Exca: Feature Sheet]![Feature Number].Name 'goto main forms feature num
                DoCmd.FindRecord currentFeature 'find the number user was editing before
                DoCmd.Hourglass False
            End If
        Else
            If Not IsNull(Forms![Exca: Feature Sheet]![Space]) Or Forms![Exca: Feature Sheet]![Space] <> "" Then
                checknum2 = DLookup("[Space]", "[Exca: Features]", "[Feature Number] = " & Me![To_feature])
                If Not IsNull(checknum2) Then 'there is a space for this related feature
                    If checknum2 <> Forms![Exca: Feature Sheet]![Space] Then 'do not allow entry if space numbers differ
                        msg = "This entry is not allowed:  feature (" & Me![To_feature] & ")"
                        msg = msg & " is in Space " & checknum2 & " but Feature " & Forms![Exca: Feature Sheet]![Feature Number]
                        msg = msg & " is in Space " & Forms![Exca: Feature Sheet]![Space]
                        msg = msg & Chr(13) & Chr(13) & "This entry cannot be allowed. Please double check this issue with your Supervisor."
                        MsgBox msg, vbExclamation, "Space mis-match"
                        MsgBox "To remove this relationship completely press ESC", vbInformation, "Help Tip"
                        If Not IsNull(Me![To_feature].OldValue) Then
                            Me![To_feature] = Me![To_feature].OldValue
                        Else
                            featureRel = Me![Relation]
                            Me.Undo
                            Me![Relation] = featureRel
                        End If
                        DoCmd.GoToControl "Feature Type"
                        DoCmd.GoToControl "To_Feature"
                    End If
                End If
            End If
        End If
    Else
        MsgBox "The Feature number is invalid, please enter a numeric value only", vbInformation, "Invalid Entry"
    End If
End If
Exit Sub
err_To_feature_AfterUpdate:
    Call General_Error_Trap
    DoCmd.Hourglass False
    Exit Sub
End Sub
