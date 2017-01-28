Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(Cancel As Integer)
Me![Date changed] = Now()
End Sub
