Option Compare Database
Private Sub cmdAddNew_Click()
On Error GoTo err_cmdAddNew_Click
    DoCmd.GoToRecord acActiveDataObject, , acNewRec
    Me![statusdate].Value = Now()
    DoCmd.GoToControl Me![unitstatus_determination]
Exit Sub
err_cmdAddNew_Click:
    If Err.Number = 2498 Then
        Resume Next
    Else
    Call General_Error_Trap
    End If
    Exit Sub
End Sub
