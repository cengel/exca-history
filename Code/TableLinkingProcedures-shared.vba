Option Compare Database
Option Explicit
Function LogUserIn_OLD()
On Error GoTo err_LogUserIn_OLD
Dim username, pwd, retVal
getuser:
    username = InputBox("Please enter your database LOGIN NAME:", "Login Name")
    If username = "" Then 'either the entered blank or pressed Cancel
        username = InputBox("The system cannot continue without your database login name. " & Chr(13) & Chr(13) & "Please enter your database LOGIN NAME below:", "Login Name")
        If username = "" Then 'again no entry
            retVal = MsgBox("Sorry but the system cannot continue without a LOGIN NAME. Do you want to try again?", vbCritical + vbYesNo, "Login required")
            If retVal = vbYes Then 'try again, loop back up
                GoTo getuser
            Else 'no, don't try again so quit system
                MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
                DoCmd.Quit acQuitSaveAll
            End If
        End If
    End If
getpwd:
    pwd = InputBox("Please enter your database PASSWORD:", "Password")
    If pwd = "" Then 'either the entered blank or pressed Cancel
        pwd = InputBox("The system cannot continue without your database password. " & Chr(13) & Chr(13) & "Please enter your database PASSWORD below:", "Password")
        If pwd = "" Then 'again no entry
            retVal = MsgBox("Sorry but the system cannot continue without a PASSWORD. Do you want to try again?", vbCritical + vbYesNo, "Password required")
            If retVal = vbYes Then 'try again, loop back up
                GoTo getpwd
            Else 'no, don't try again so quit system
                MsgBox "The system will now quit", vbCritical + vbOKOnly, "Invalid Login"
                DoCmd.Quit acQuitSaveAll
            End If
        End If
    End If
Dim mydb As DAO.Database, I
Dim tmptable As TableDef
Set mydb = CurrentDb
For I = 0 To mydb.TableDefs.count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
    If tmptable.Connect <> "" Then
        tmptable.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
        tmptable.RefreshLink
        Exit For 'only necessary for one table for Access to set up the correct link to SQL Server
    End If
Next I
cleanup:
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
LogUserIn_OLD = True
Exit Function
err_LogUserIn_OLD:
    If Err.Number = 3059 Then
        retVal = MsgBox("Sorry but the login you have given is incorrect or the database/internet connection is not available. You cannot connect to the database. Do you wish to try logging in again?", vbCritical + vbYesNo, "Login Failure")
        If retVal = vbYes Then Resume
    ElseIf Err.Number = 3151 Then
        AlterODBC
    Else
        MsgBox Err.Description & Chr(13) & Chr(13) & "The system will now quit", vbCritical, "Login Failure"
    End If
    LogUserIn_OLD = False
    DoCmd.Quit
End Function
Function AlterODBC()
Dim startstr, endstr, namestr
    If Err.Number = 3151 Then
        startstr = InStr(Err.Description, "'")
        endstr = InStr(startstr + 1, Err.Description, "'")
        namestr = Mid(Err.Description, startstr + 1, endstr - startstr)
        MsgBox "This system requires the ODBC connection: " & namestr & Chr(13) & Chr(13) & _
                        "The error returned is: " & Err.Description & Chr(13) & Chr(13) & "Instructions of how " & _
                        "to setup ths DSN can be found on the Web at http://catalsql.arch.cam.ac.uk/database/odbc.html" & _
                        "", vbCritical, "The system cannot start"
    End If
Exit Function
Dim username, pwd, tblName, rstemp
Dim mydb As DAO.Database, I
Dim tmptable As TableDef
Set mydb = CurrentDb
For I = 0 To mydb.TableDefs.count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
    If tmptable.Connect <> "" Then
        tmptable.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
        tmptable.RefreshLink
        Exit For 'only necessary for one table for Access to set up the correct link to SQL Server
    End If
Next I
cleanup:
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
End Function
Function LogUserIn(username As String, pwd As String)
On Error GoTo err_LogUserIn
Dim retVal
If username <> "" And pwd <> "" Then
    Dim mydb As DAO.Database, I, errmsg, connStr
    Dim tmptable As TableDef
    Set mydb = CurrentDb
    Dim myq As QueryDef
    Set myq = mydb.CreateQueryDef("")
    connStr = ""
    For I = 0 To mydb.TableDefs.count - 1 'loop the tables collection
         Set tmptable = mydb.TableDefs(I)
        If tmptable.Connect <> "" Then
            If connStr = "" Then connStr = tmptable.Connect
            On Error Resume Next
                myq.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
                myq.ReturnsRecords = False 'don't waste resources bringing back records
                myq.sql = "select [Unit Number] from [Exca: Unit Sheet] WHERE [Unit Number] = 1" 'this is a shared and core table so should always be avail, the record doesn't have to exist
                myq.Execute
            If Err <> 0 Then 'the login deails are incorrect
                GoTo err_LogUserIn
            Else
                On Error GoTo err_LogUserIn:
                tmptable.Connect = tmptable.Connect & ";UID=" & username & ";PWD=" & pwd
                tmptable.RefreshLink
            End If
            Exit For 'only necessary for one table for Access to set up the correct link to SQL Server
        End If
    Next I
Else
    MsgBox "Both a username and password are required to operate the system correctly. Please quit and restart the application.", vbCritical, "Login problem encountered"
End If
SetGeneralPermissions username, pwd, connStr 'requires more thought
LogUserIn = True
cleanup:
    myq.Close
    Set myq = Nothing
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
Exit Function
err_LogUserIn:
    If Err.Number = 3059 Or Err.Number = 3151 Then
        errmsg = "Sorry but the system cannot log you into the database. There are three reasons this may have occurred:" & Chr(13) & Chr(13)
        errmsg = errmsg & "1. Your login details have been entered incorrectly" & Chr(13) & Chr(13)
        errmsg = errmsg & "2. There is no ODBC connection to the database setup on this computer." & Chr(13) & "    See http://www.catalhoyuk.com/database/odbc.html for details." & Chr(13) & Chr(13)
        errmsg = errmsg & "3. Your computer is not connected to the Internet at this time." & Chr(13) & Chr(13)
        errmsg = errmsg & "Do you wish to try logging in again?"
        retVal = MsgBox(errmsg, vbCritical + vbYesNo, "Login Failure")
        If retVal = vbYes Then
            GoTo cleanup 'used to be resume before querydef intro, now just cleanup and leave so user can try again
        Else
            retVal = MsgBox("Are you really sure you want to quit and close the system?", vbCritical + vbYesNo, "Confirm System Closure")
            If retVal = vbNo Then
                GoTo cleanup 'on 2nd thoughts the user doesn't want to quit so now just cleanup and leave so user can try again
            Else
                MsgBox "The system will now quit" & Chr(13) & Chr(13) & "The error reported was: " & Err.Description, vbCritical, "Login Failure"
            End If
        End If
    Else
        MsgBox Err.Description & Chr(13) & Chr(13) & "The system will now quit", vbCritical, "Login Failure"
    End If
    LogUserIn = False
    DoCmd.Quit acQuitSaveAll
End Function
Sub WriteOutTableNames()
Dim mydb As DAO.Database, I
Dim tmptable As TableDef
Set mydb = CurrentDb
For I = 0 To mydb.TableDefs.count - 1 'loop the tables collection
    Set tmptable = mydb.TableDefs(I)
    If InStr(tmptable.Name, "MSys") = 0 Then
        Debug.Print tmptable.Name
        If tmptable.Connect <> "" Then
            Debug.Print "Linked"
        Else
            Debug.Print "Local"
        End If
    End If
Next I
cleanup:
    Set tmptable = Nothing
    mydb.Close
    Set mydb = Nothing
End Sub
