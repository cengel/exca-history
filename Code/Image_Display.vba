Option Compare Database
Option Explicit
Private Sub Form_Current()
On Error GoTo err_Current
Dim newStr, newstr2, zeronum, FileName
newStr = Replace(Me![Path], ":", "\")
newstr2 = newStr
Dim dirpath, breakid
zeronum = 10 - (Len(Me![Record_ID]))
FileName = "p"
        Do While zeronum > 0
            FileName = FileName & "0"
            zeronum = zeronum - 1
        Loop
FileName = FileName & Me![Record_ID]
breakid = Left(FileName, 3)
breakid = Mid(FileName, 2, 2) 'chop off leading p
dirpath = breakid & "\"
breakid = Mid(FileName, 4, 2)
dirpath = dirpath & breakid & "\"
breakid = Mid(FileName, 6, 2)
dirpath = dirpath & breakid & "\"
breakid = Mid(FileName, 8, 2)
dirpath = dirpath & breakid & "\"
newstr2 = newstr2 & "\" & dirpath & FileName & ".jpg"
Me!txtFullPath = newstr2
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
    Me.RecordSource = "Select * from view_Portfolio_Previews_2008 WHERE " & Me.Filter
    If Me.RecordsetClone.RecordCount <= 0 Then
        MsgBox "No images have been found in the Portfolio catalogue for this entity", vbInformation, "No images to display"
        DoCmd.Close acForm, Me.Name
    End If
End If
End Sub
