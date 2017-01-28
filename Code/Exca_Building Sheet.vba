Option Compare Database   'Use database order for string comparisons
Private Sub Excavation_Click()
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "Excavation"
    DoCmd.OpenForm stDocName, , , stLinkCriteria
    DoCmd.Close acForm, "Exca: Building Sheet"
End Sub
