Option Compare Database
Option Explicit
Private Sub Form_Current()
On Error GoTo err_Current
Dim newStr, newstr2, zeronum, FileName
newStr = Replace(Me![Path], ":", "\")
newstr2 = newStr
zeronum = 10 - (Len(Me![Record_ID]))
FileName = "p"
        Do While zeronum > 0
            FileName = FileName & "0"
            zeronum = zeronum - 1
        Loop
newstr2 = newstr2 & "\" & FileName & Me![Record_ID] & ".jpg"
Me!Image145.Picture = newstr2
Exit Sub
err_Current:
    If Err.Number = 2220 Then
        If Dir(newstr2) = "" Then
            MsgBox "The directory where images are supposed to be stored cannot be found. Please contact the database administrator"
        Else
            MsgBox "The image file cannot be found - check the file exists"
            DoCmd.GoToControl "txtSketch"
        End If
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub
Private Sub Form_Open(Cancel As Integer)
If Me.OpenArgs <> "" Then
    If Me.OpenArgs = 2007 Then
        Me.RecordSource = "Select * from view_Portfolio_2007Previews WHERE " & Me.Filter
    Else
        Me.RecordSource = "Select * from view_Portfolio_Upto2007Previews WHERE " & Me.Filter
    End If
End If
End Sub
