Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
Forms![Exca: Feature Sheet]![dbo_Exca: FeatureHistory].Form![lastmodify].Value = Now()
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
    Me![Feature Type].Locked = True
    Me![Feature Type].Enabled = False
    Me![FeatureSubType].Locked = True
    Me![FeatureSubType].Enabled = False
Exit Sub
err_Form_Open:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub To_feature_AfterUpdate()
On Error GoTo err_To_feature_AfterUpdate
Dim checknum, msg, retval, sql, currentFeature, checknum2, featureRel, checknum3, myrs As DAO.Recordset, mydb As DAO.Database
If Me![To_feature] <> "" Then
    If IsNumeric(Me![To_feature]) Then
        checknum = DLookup("[Feature Number]", "[Exca: Features]", "[Feature Number] = " & Me![To_feature])
        If IsNull(checknum) Then
            msg = "The Feature Number " & Me![To_feature] & " DOES NOT EXIST in the database. The system can enter it for you ready for you to update later."
            msg = msg & Chr(13) & Chr(13) & "Would you like the system to create this feature number now?"
            retval = MsgBox(msg, vbInformation + vbYesNo, "Feature Number does not exist")
            If retval = vbNo Then
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
            checknum2 = DLookup("[In_Space]", "[Exca: Features in Spaces]", "[Feature] = " & Me![Feature Number])
            If Not IsNull(checknum2) Then 'there is a space for main feature
                checknum3 = DLookup("[In_Space]", "[Exca: Features in Spaces]", "[Feature] = " & Me![To_feature])
                If Not IsNull(checknum3) Then 'there is a space for related feature
                    sql = "SELECT [Exca: Features in Spaces].Feature, [Exca: Features in Spaces].In_Space, [Exca: Features in Spaces_1].Feature, [Exca: Features in Spaces_1].In_Space" & _
                            " FROM [Exca: Features in Spaces] INNER JOIN [Exca: Features in Spaces] AS [Exca: Features in Spaces_1] ON [Exca: Features in Spaces].In_Space = [Exca: Features in Spaces_1].In_Space " & _
                            " WHERE ([Exca: Features in Spaces].Feature =" & Me![Feature Number] & ")  AND ([Exca: Features in Spaces_1].Feature=" & Me![To_feature] & ");"
                    Set mydb = CurrentDb
                    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
                    If myrs.EOF And myrs.BOF Then
                        Dim response
                        msg = "This entry is not allowed because these two features are not currently in the same Space. They must be in the same space to create a relationship."
                        msg = msg & Chr(13) & Chr(13) & "Are you sure that " & Parent![Feature Number] & " is " & Me![Relation] & " " & Me![To_feature] & "?"
                        response = MsgBox(msg, vbYesNo + vbQuestion, "Space mis-match")
                        If response = vbYes Then
                        Else
                        If Not IsNull(Me![To_feature].OldValue) Then
                            Me![To_feature] = Me![To_feature].OldValue
                            DoCmd.GoToControl "To_Feature"
                        Else
                            Me.Undo
                             DoCmd.GoToControl "Relation"
                        End If
                    End If
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
