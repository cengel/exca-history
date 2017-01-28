Option Compare Database
Option Explicit
Private Sub Form_Current()
Select Case Me.Location
            Case "cut"
            Me.Description.RowSource = " ; burial; ditch; foundation cut; gully; pit; posthole; scoop; stakehole"
            Me.Description.Enabled = True
            Case "feature"
            Me.Description.RowSource = " ; basin; bin; hearth; niche; oven"
            Me.Description.Enabled = True
            Case Else
            Me.Description.RowSource = ""
            Me.Description.Enabled = False
End Select
End Sub
Private Sub Location_Change()
    Me.Description = ""
    Select Case Me.Location
        Case "cut"
        Me.Description.RowSource = " ; burial; ditch; foundation cut; gully; pit; posthole; scoop; stakehole"
        Me.Description.Enabled = True
        Case "feature"
        Me.Description.RowSource = " ; basin; bin; hearth; niche; oven"
        Me.Description.Enabled = True
        Case Else
        Me.Description.RowSource = ""
        Me.Description.Enabled = False
    End Select
End Sub
