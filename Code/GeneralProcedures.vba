Option Compare Database
Option Explicit
Function StartUp()
On Error GoTo err_startup
DoCmd.RunCommand acCmdWindowHide 'hide the DB window from prying eyes
DoCmd.OpenForm "Excavation_Login", acNormal, , , acFormEdit, acDialog
    Application.SetOption "Confirm Action Queries", False  'this will hide behind the scences sql actions
SetCurrentVersion
DoCmd.OpenForm "Excavation", acNormal, , , acFormReadOnly 'open main menu
Forms![Excavation].Refresh
Exit Function
err_startup:
    Call General_Error_Trap
End Function
Function CheckIfLOVValueUsed(LOVName, LOVField, LOVValue, CheckTable, CheckTableKeyField, CheckTableField, task, Optional extracrit)
On Error GoTo err_CheckIFLOVValueUsed
If LOVName <> "" And LOVField <> "" And LOVValue <> "" And CheckTable <> "" And CheckTableKeyField <> "" And CheckTableField <> "" And task <> "" Then
    Dim mydb As Database, myrs As Recordset, sql As String, msg As String, msg1 As String, keyfld As Field, count As Integer
    Set mydb = CurrentDb
    If Not IsMissing(extracrit) Then
        sql = "SELECT [" & CheckTableKeyField & "], [" & CheckTableField & "] FROM [" & CheckTable & "] WHERE [" & CheckTableField & "] = '" & LOVValue & "' " & extracrit & " ORDER BY [" & CheckTableKeyField & "];"
    Else
        sql = "SELECT [" & CheckTableKeyField & "], [" & CheckTableField & "] FROM [" & CheckTable & "] WHERE [" & CheckTableField & "] = '" & LOVValue & "' ORDER BY [" & CheckTableKeyField & "];"
    End If
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
    If myrs.BOF And myrs.EOF Then
        msg = "ok"
    Else
        myrs.MoveFirst
        count = 0
        msg = "You cannot " & task & " this " & LOVField & " because the following records in the table " & CheckTable & " use it: "
        msg1 = ""
        Do Until myrs.EOF
            Set keyfld = myrs.Fields(CheckTableKeyField)
            If msg1 <> "" Then msg1 = msg1 & ", "
            msg1 = msg1 & keyfld
            count = count + 1
            If count > 50 Then
                msg1 = msg1 & ".....etc"
                Exit Do
            End If
        myrs.MoveNext
        Loop
        msg = msg & Chr(13) & Chr(13) & CheckTableKeyField & ": " & msg1
        If task = "edit" Then
            msg = msg & Chr(13) & Chr(13) & "It is suggested you add a new " & LOVField & " to the list and then change all records that refer to '"
            msg = msg & LOVValue & "' to your new " & LOVField & ". You will then be able to delete it from the list."
        ElseIf task = "delete" Then
             msg = msg & Chr(13) & Chr(13) & "You must change all records that refer to this " & LOVField
            msg = msg & " '" & LOVValue & "' before you will be able to delete it from the list."
        End If
    End If
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
    CheckIfLOVValueUsed = msg
Else
    CheckIfLOVValueUsed = "fail"
End If
Exit Function
err_CheckIFLOVValueUsed:
    Call General_Error_Trap
    Exit Function
End Function
Function AdminDeletionCheck(CheckTable, CheckField, CheckVal, Term, retField)
On Error GoTo err_AdminDeletionCheck
If CheckTable <> "" And CheckField <> "" And CheckVal <> "" And Term <> "" Then
    Dim mydb As Database, myrs As Recordset, sql As String, msg As String, msg1 As String, keyfld As Field, count As Integer
    Set mydb = CurrentDb
    If CheckTable = "Exca: stratigraphy" And CheckField = "To_units" Then
        sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [" & CheckField & "] = '" & CheckVal & "';"
    ElseIf CheckTable = "Exca: graphics list" Then 'graphics needs to define if feature num or unit
        If CheckField = "Unit" Then
           sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [Unit/feature number] = " & CheckVal & " AND lcase([Feature/Unit]) =  'u';"
        ElseIf CheckField = "Feature" Then
            sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [Unit/feature number] = " & CheckVal & " AND lcase([Feature/Unit]) =  'f';"
        End If
    Else
        sql = "SELECT [" & retField & "] FROM [" & CheckTable & "] WHERE [" & CheckField & "] = " & CheckVal & ";"
    End If
    Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
    If myrs.BOF And myrs.EOF Then
        msg = ""
    Else
        myrs.MoveFirst
        count = 0
        msg = Term & ": "
        msg1 = ""
        Do Until myrs.EOF
            Set keyfld = myrs.Fields(retField)
            If msg1 <> "" Then msg1 = msg1 & ", "
            msg1 = msg1 & keyfld
            count = count + 1
            If count > 50 Then
                msg1 = msg1 & ".....etc"
                Exit Do
            End If
        myrs.MoveNext
        Loop
        msg = msg & msg1
    End If
    myrs.Close
    Set myrs = Nothing
    mydb.Close
    Set mydb = Nothing
    AdminDeletionCheck = msg
Else
    AdminDeletionCheck = ""
End If
Exit Function
err_AdminDeletionCheck:
    Call General_Error_Trap
    Exit Function
End Function
Sub DeleteARecord(FromTable, FieldName, FieldValue, Text, mydb)
Dim sql, myq As QueryDef
Set myq = mydb.CreateQueryDef("")
        If Text = False Then
            If FromTable = "Exca: graphics list" Then 'graphics needs to define if feature num or unit
                If FieldName = "Unit" Then
                    sql = "DELETE FROM [" & FromTable & "] WHERE [Unit/feature number] = " & FieldValue & " AND lcase([Feature/Unit]) =  'u';"
                ElseIf FieldName = "Feature" Then
                    sql = "DELETE FROM [" & FromTable & "] WHERE [Unit/feature number] = " & FieldValue & " AND lcase([Feature/Unit]) =  'f';"
                End If
            Else
                sql = "DELETE FROM [" & FromTable & "] WHERE [" & FieldName & "] = " & FieldValue & ";"
            End If
        Else
            sql = "DELETE FROM [" & FromTable & "] WHERE [" & FieldName & "] = '" & FieldValue & "';"
        End If
        myq.sql = sql
        myq.Execute
myq.Close
Set myq = Nothing
End Sub
