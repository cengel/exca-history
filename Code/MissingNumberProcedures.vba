Option Compare Database
Option Explicit
Function FindMissingNumbers(tbl, fld) As Boolean
On Error GoTo err_nums
Dim mydb As DAO.Database, myrs As DAO.Recordset
Dim sql As String, sql1 As String, val As Field, holdval1 As Long, holdval2 As Long
Dim response As Integer
MsgBox "The first thing this code must do is retrieve the whole dataset. If your connection is slow it may time out but it will give you a message if this happens. Starting now......", vbInformation, "Start Procedure"
Set mydb = CurrentDb
sql = "SELECT [" & fld & "] FROM [" & tbl & "] ORDER BY [" & fld & "];"
Set myrs = mydb.OpenRecordset(sql, dbOpenSnapshot)
myrs.MoveLast
Set val = myrs.Fields(fld)
holdval1 = val
myrs.MovePrevious
holdval2 = val
response = MsgBox("The last two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
If response = vbNo Then
    holdval1 = holdval2
    myrs.MovePrevious
    holdval2 = val
    response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
    If response = vbNo Then
        holdval1 = holdval2
        myrs.MovePrevious
        holdval2 = val
        response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
        If response = vbNo Then
            holdval1 = holdval2
            myrs.MovePrevious
            holdval2 = val
            response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
            If response = vbNo Then
                holdval1 = holdval2
                myrs.MovePrevious
                holdval2 = val
                response = MsgBox("The next two values in the " & fld & " field are: " & holdval1 & " and " & holdval2 & ". Do you want the check to run upto " & holdval1, vbYesNo + vbQuestion, "Run end point")
                If response = vbNo Then
                    MsgBox "Please clean up the last " & fld & " values and run this procedure again"
                    FindMissingNumbers = False
                    Exit Function
                Else
                    GoTo cont
                End If
            Else
                GoTo cont
            End If
        Else
            GoTo cont
        End If
    Else
        GoTo cont
    End If
Else
    GoTo cont
End If
cont:
    MsgBox "The code will now run to compile the list of missinig numbers up to: " & holdval1 & ". It may be quite slow , a report will appear when complete so you know it has finished"
    sql1 = "DELETE * FROM LocalMissingNumbers;"
    DoCmd.RunSQL sql1
    sql1 = ""
    myrs.MoveFirst
    Dim counter As Long, checknum
    counter = 0
    Do Until counter = holdval1
        checknum = DLookup("[" & fld & "]", "[" & tbl & "]", "[" & fld & "] = " & counter)
        If IsNull(checknum) Then
            sql1 = "INSERT INTO [LocalMissingNumbers] (MissingNumber) VALUES (" & counter & ");"
            DoCmd.RunSQL sql1
        End If
        counter = counter + 1
    Loop
myrs.Close
Set myrs = Nothing
mydb.Close
Set mydb = Nothing
FindMissingNumbers = True
Exit Function
err_nums:
    Call General_Error_Trap
    Exit Function
End Function
