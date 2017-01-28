Option Compare Database
Option Explicit
Private Sub cmdDelete_Click()
On Error GoTo Err_cmdDelete_Click
Dim resp
resp = MsgBox("Are you sure you want to delete the phasing " & Me![txtOccupationPhase] & " from this unit?", vbQuestion + vbYesNo, "Confirm Deletion")
If resp = vbYes Then
    Me![txtOccupationPhase].Locked = False
    DoCmd.RunCommand acCmdDeleteRecord
    Me![txtOccupationPhase].Locked = True
End If
Exit_cmdDelete_Click:
    Exit Sub
Err_cmdDelete_Click:
    MsgBox Err.Description
    Resume Exit_cmdDelete_Click
End Sub
