Option Compare Database
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
End Sub
