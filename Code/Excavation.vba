Option Compare Database   'Use database order for string comparisons
Private Sub Area_Sheet_Click()
On Error GoTo Err_Area_Sheet_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Area Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_Area_Sheet_Click:
    Exit Sub
Err_Area_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Area_Sheet_Click
End Sub
Private Sub Building_Sheet_Click()
On Error GoTo Err_Building_Sheet_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Building Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_Building_Sheet_Click:
    Exit Sub
Err_Building_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Building_Sheet_Click
End Sub
Private Sub Button10_Click()
Building_Sheet_Click
End Sub
Private Sub Button11_Click()
Space_Sheet_Button_Click
End Sub
Private Sub Button12_Click()
Feature_Sheet_Button_Click
End Sub
Private Sub Button13_Click()
Unit_Sheet_Click
End Sub
Private Sub Button17_Click()
Area_Sheet_Click
End Sub
Private Sub Button9_Click()
Return_to_Master_Con_Click
End Sub
Private Sub Command18_Click()
Open_priority_Click
End Sub
Private Sub Feature_Sheet_Button_Click()
On Error GoTo Err_Feature_Sheet_Button_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Feature Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_Feature_Sheet_Button_Click:
    Exit Sub
Err_Feature_Sheet_Button_Click:
    MsgBox Err.Description
    Resume Exit_Feature_Sheet_Button_Click
End Sub
Sub Open_priority_Click()
On Error GoTo Err_Open_priority_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Priority Detail"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_Open_priority_Click:
    Exit Sub
Err_Open_priority_Click:
    MsgBox Err.Description
    Resume Exit_Open_priority_Click
End Sub
Private Sub Return_to_Master_Con_Click()
DoCmd.DoMenuItem acFormBar, acFileMenu, 14, , acMenuVer70
End Sub
Private Sub Space_Sheet_Button_Click()
On Error GoTo Err_Space_Sheet_Button_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Space Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
Exit_Space_Sheet_Button_Click:
    Exit Sub
Err_Space_Sheet_Button_Click:
    MsgBox Err.Description
    Resume Exit_Space_Sheet_Button_Click
End Sub
Private Sub Unit_Sheet_Click()
On Error GoTo Err_Unit_Sheet_Click
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Exca: Unit Sheet"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.GoToRecord acForm, stDocName, acNewRec
Exit_Unit_Sheet_Click:
    Exit Sub
Err_Unit_Sheet_Click:
    MsgBox Err.Description
    Resume Exit_Unit_Sheet_Click
End Sub
Sub Command19_Click()
On Error GoTo Err_Command19_Click
    DoCmd.Close
Exit_Command19_Click:
    Exit Sub
Err_Command19_Click:
    MsgBox Err.Description
    Resume Exit_Command19_Click
End Sub
Sub Feature_types_Click()
On Error GoTo Err_Feature_types_Click
    Dim stDocName As String
    stDocName = "Exca: Feature Types"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
Exit_Feature_types_Click:
    Exit Sub
Err_Feature_types_Click:
    MsgBox Err.Description
    Resume Exit_Feature_types_Click
End Sub
Sub interpret_Click()
On Error GoTo Err_interpret_Click
    Dim stDocName As String
    stDocName = "Exca: List Interpretive Categories"
    DoCmd.OpenQuery stDocName, acNormal, acEdit
Exit_interpret_Click:
    Exit Sub
Err_interpret_Click:
    MsgBox Err.Description
    Resume Exit_interpret_Click
End Sub
