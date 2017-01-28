Option Compare Database
Option Explicit
Private Sub cmdLoadPicture_Click()
fLoadPicture Me.JGSForm.Form.Image1, , True
ScrollToHome Me.JGSForm.Form.Image1
End Sub
Private Sub CmdClip_Click()
With Me.JGSForm.Form.Image1
    If .ImageWidth <= Me.JGSForm.Form.Width - 200 Then
        .Width = .ImageWidth
    Else
        .Width = Me.JGSForm.Form.Width - 200
    End If
    If .ImageHeight <= Me.JGSForm.Form.Detail.Height - 200 Then
        .Height = .ImageHeight
    Else
        .Height = Me.JGSForm.Form.Detail.Height - 200
    End If
    .SizeMode = acOLESizeClip '0
End With
ScrollToHome Me.JGSForm.Form.Image1
End Sub
Private Sub cmdSave_Click()
Dim blRet As Boolean
blRet = fSaveImagetoDisk(Me.JGSForm.Form.Image1)
End Sub
Private Sub CmdStretch_Click()
With Me.JGSForm.Form.Image1
    .Width = Me.JGSForm.Form.Width - 200
    .Height = Me.JGSForm.Form.Detail.Height - 200
    .SizeMode = acOLESizeStretch '3
End With
End Sub
Private Sub CmdZoom_Click()
With Me.JGSForm.Form.Image1
    .Width = Me.JGSForm.Form.Width - 200
    .Height = Me.JGSForm.Form.Detail.Height - 200
    .SizeMode = acOLESizeZoom '1
End With
End Sub
Private Sub Command46_Click()
On Error GoTo err_cmd46
    DoCmd.Close acForm, Me.Name
Exit Sub
err_cmd46:
    Call General_Error_Trap
    Exit Sub
End Sub
Private Sub Form_Activate()
DoCmd.MoveSize 5, 5, 12500, 10000
End Sub
Private Sub Form_Load()
DoCmd.MoveSize 5, 5, 12500, 10000
End Sub
Private Sub CmdBig_Click()
Dim intWidth As Integer
Dim intHeight As Integer
With Me.JGSForm.Form.Image1
    intWidth = .Width * 1.05
    intHeight = .Height * 1.05
    If intWidth < .Parent.Width Then
        .Width = intWidth
    Else
        MsgBox "Sorry, that is as Big as you can go!", vbOKOnly, "Maximum Zoom"
        Exit Sub
    End If
    If intHeight < .Parent.Detail.Height Then
        .Height = intHeight
    Else
        MsgBox "Sorry, that is as Big as you can go!", vbOKOnly, "Maximum Zoom"
        Exit Sub
    End If
    .SizeMode = acOLESizeZoom
End With
DoEvents
End Sub
Private Sub CmdSmall_Click()
Dim intWidth As Integer
Dim intHeight As Integer
With Me.JGSForm.Form.Image1
    intWidth = .Width * 0.95
    intHeight = .Height * 0.95
    If intWidth > 200 Then
        .Width = intWidth
    Else
        MsgBox "Sorry, that is as small as you can go!", vbOKOnly, "Minimum Zoom"
        Exit Sub
    End If
    If intHeight > 200 Then
        .Height = intHeight
    Else
        MsgBox "Sorry, that is as small as you can go!", vbOKOnly, "Minimum Zoom"
        Exit Sub
    End If
    .SizeMode = acOLESizeZoom
End With
DoEvents
End Sub
Private Sub Form_Open(Cancel As Integer)
On Error GoTo err_open
Dim Path, FileName, fname, newfile
If Me.OpenArgs <> "" Then
    Path = sketchpath2015 & "spaces\sketches\"
    FileName = Me.OpenArgs
    Path = Path & "S" & FileName & "*" & ".jpg"
    fname = Dir(Path & "*", vbNormal)
    While fname <> ""
        newfile = fname
        fname = Dir()
    Wend
    Path = sketchpath2015 & "spaces\sketches\" & newfile
    Me![txtImagePath] = Path
    If Dir(Path) = "" Then
            MsgBox "The sketch cannot be found, it may not have been scanned in yet. The database is looking for: " & Path & " please check it exists."
            DoCmd.Close acForm, Me.Name
    Else
        fLoadPicture Me.JGSForm.Form.Image1, Me![txtImagePath], True
        ScrollToHome Me.JGSForm.Form.Image1
    End If
Else
    MsgBox "No image name was passed in to this form when it was opened, system does not know which image to display. Please open from Unit sheet only", vbInformation, "No image to display"
End If
Exit Sub
err_open:
    If Err.Number = 2220 Then
        If Dir(Path) = "" Then
            MsgBox "The image cannot be found. The database is looking for: " & Path & " please check it exists."
        Else
            MsgBox "The image file cannot be found - check the file exists"
        End If
    Else
        Call General_Error_Trap
    End If
    Exit Sub
End Sub
