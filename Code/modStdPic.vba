Option Compare Database
Option Explicit
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECTL
    Left As Long
    top As Long
    right As Long
    Bottom As Long
End Type
Private Type SIZEL
    cx As Long
    cy As Long
End Type
Private Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgblReterved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long 'ERGBCompression
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type
Private Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
End Type
Private Type Bitmap
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  BmBits As Long
End Type
Private Type DIBSECTION
    dsBm As Bitmap
    dsBmih As BITMAPINFOHEADER
    dsBitfields(2) As Long
    dshSection As Long
    dsOffset As Long
End Type
Private Type BITMAPFILEHEADER    '14 bytes
  bfType As Integer
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type
Private Declare Function apiGetObject Lib "gdi32" _
Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, _
lpObject As Any) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal Length As Long)
Declare Function CopyEnhMetaFile Lib "gdi32" Alias "CopyEnhMetaFileA" (ByVal hemfSrc As Long, ByVal lpszFile As String) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hEMF As Long, lpRect As RECTL) As Long
Private Declare Function apiCloseEnhMetaFile Lib "gdi32" _
Alias "CloseEnhMetaFile" (ByVal hDC As Long) As Long
Private Declare Function apiCreateEnhMetaFileRECT Lib "gdi32" _
Alias "CreateEnhMetaFileA" (ByVal hDCref As Long, _
ByVal lpFileName As String, ByRef lpRect As RECTL, ByVal lpDescription As String) As Long
Private Declare Function apiDeleteEnhMetaFile Lib "gdi32" _
Alias "DeleteEnhMetaFile" (ByVal hEMF As Long) As Long
Private Declare Function GetEnhMetaFileBits Lib "gdi32" _
(ByVal hEMF As Long, ByVal cbBuffer As Long, lpbBuffer As Byte) As Long
Private Declare Function apiSelectObject Lib "gdi32" _
 Alias "SelectObject" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function apiGetDC Lib "user32" _
  Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare Function apiReleaseDC Lib "user32" _
  Alias "ReleaseDC" (ByVal hWnd As Long, _
  ByVal hDC As Long) As Long
Private Declare Function apiDeleteObject Lib "gdi32" _
  Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, _
ByVal nStretchMode As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZEL) As Long
Private Declare Function SetViewportOrgEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetWindowOrgEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetWindowExtEx Lib "gdi32" _
(ByVal hDC As Long, ByVal nX As Long, ByVal nY As Long, lpSize As SIZEL) As Long
Private Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" _
(ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const BLACKONWHITE = 1
Private Const WHITEONBLACK = 2
Private Const COLORONCOLOR = 3
Private Const HALFTONE = 4
Private Const MAXSTRETCHBLTMODE = 4
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const SB_LINEUP = 0
Private Const SB_LINELEFT = 0
Private Const SB_LINEDOWN = 1
Private Const SB_LINERIGHT = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGELEFT = 2
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGERIGHT = 3
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_LEFT = 6
Private Const SB_BOTTOM = 7
Private Const SB_RIGHT = 7
Private Const SB_ENDSCROLL = 8
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14
Private Const MM_TEXT = 1
Private Const MM_LOMETRIC = 2
Private Const MM_HIMETRIC = 3
Private Const MM_LOENGLISH = 4
Private Const MM_HIENGLISH = 5
Private Const MM_TWIPS = 6
Private Const MM_ISOTROPIC = 7
Private Const MM_ANISOTROPIC = 8
Private Const vbPicTypeNone = 0 'Picture is empty
Private Const vbPicTypeBitmap = 1 'Bitmap (.bmpBMP files)
Private Const vbPicTypeMetafile = 2 'Metafile (.wmfWMF files)
Private Const vbPicTypeIcon = 3 'Icon (.icoICO files)
Private Const vbPicTypeEMetafile = 4 'Enhanced Metafile (.emfEMF files)
Private Const WHITE_BRUSH = 0
Private Const LTGRAY_BRUSH = 1
Private Const GRAY_BRUSH = 2
Private Const DKGRAY_BRUSH = 3
Private Const BLACK_BRUSH = 4
Private Const NULL_BRUSH = 5
Private Const HOLLOW_BRUSH = NULL_BRUSH
Private Const WHITE_PEN = 6
Private Const BLACK_PEN = 7
Private Const NULL_PEN = 8
Private Const OEM_FIXED_FONT = 10
Private Const ANSI_FIXED_FONT = 11
Private Const ANSI_VAR_FONT = 12
Private Const SYSTEM_FONT = 13
Private Const DEVICE_DEFAULT_FONT = 14
Private Const DEFAULT_PALETTE = 15
Private Const SYSTEM_FIXED_FONT = 16
Private Const STOCK_LAST = 16
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Const BKMODE_LAST = 2
Private Const HORZSIZE = 4           '  Horizontal size in millimeters
Private Const VERTSIZE = 6           '  Vertical size in millimeters
Private Const HORZRES = 8            '  Horizontal width in pixels
Private Const VERTRES = 10           '  Vertical width in pixels
Private Const LOGPIXELSY = 90
Private Const LOGPIXELSX = 88
Private Const TWIPSPERINCH = 1440
Private m_GDIpToken         As Long         ' Needed to close GDI+
Function fStdPicToImageData(hStdPic As Object, ctl As Access.Image, _
Optional FileNamePath As String = "", Optional AutoSize As Boolean = False) As Boolean
On Error GoTo ERR_SHOWPIC
Dim hDCref As Long
Dim sz As SIZEL
Dim pt As POINTAPI
Dim rc As RECTL
Dim lngRet As Long
Dim s As String
Dim hMetafile As Long
Dim hDCMeta As Long
Dim arrayMeta() As Byte
Dim sngConvertX As Single
Dim sngConvertY As Single
Dim ImageWidth As Long
Dim ImageHeight As Long
Dim Xdpi As Single
Dim Ydpi As Single
Dim TwipsPerPixelX As Single
Dim TwipsPerPixely As Single
Dim sngHORZRES As Single
Dim sngVERTRES As Single
Dim sngHORZSIZE As Single
Dim sngVERTSIZE As Single
hDCref = apiGetDC(0)
If hStdPic.Type = 0 Then
 Err.Raise vbObjectError + 523, "fStdPicToImageData.modStdPic", _
    "Sorry...This function can only read Image files." & vbCrLf & "Please Select a Valid Supported Image File"
End If
sngHORZRES = apiGetDeviceCaps(hDCref, HORZRES)
sngVERTRES = apiGetDeviceCaps(hDCref, VERTRES)
sngHORZSIZE = apiGetDeviceCaps(hDCref, HORZSIZE)
sngVERTSIZE = apiGetDeviceCaps(hDCref, VERTSIZE)
sngConvertX = (sngHORZSIZE * 0.1) / 2.54
sngConvertY = (sngVERTSIZE * 0.1) / 2.54
sngConvertX = sngHORZRES / sngConvertX
sngConvertY = sngVERTRES / sngConvertY
Xdpi = sngConvertX
Ydpi = sngConvertY
sngConvertX = hStdPic.Width * 0.001
sngConvertY = hStdPic.Height * 0.001
sngConvertX = sngConvertX / 2.54
sngConvertY = sngConvertY / 2.54
sngConvertX = sngConvertX * 1440
sngConvertY = sngConvertY * 1440
TwipsPerPixelX = TWIPSPERINCH / Xdpi
TwipsPerPixely = TWIPSPERINCH / Ydpi
ImageWidth = sngConvertX / TwipsPerPixelX
ImageHeight = sngConvertY / TwipsPerPixely
rc.right = hStdPic.Width
rc.Bottom = hStdPic.Height
s = "Stephen Lebans" & Chr(0) & Chr(0) & "www.lebans.com" & Chr(0) & Chr(0)
hDCMeta = apiCreateEnhMetaFileRECT(hDCref, vbNullString, rc, s)
If hDCMeta = 0 Then
    Err.Raise vbObjectError + 525, "fStdPicToImageData.modStdPic", _
    "Sorry...cannot Create Enhanced Metafile"
End If
lngRet = SetMapMode(hDCMeta, MM_TEXT) 'ANISOTROPIC) 'TEXT)
lngRet = SetWindowExtEx(hDCMeta, ImageWidth, ImageHeight, sz)
lngRet = SetWindowOrgEx(hDCMeta, 0&, 0&, pt)
lngRet = SetWindowExtEx(hDCMeta, ImageWidth, ImageHeight, sz)
lngRet = SetBkMode(hDCMeta, TRANSPARENT)
lngRet = apiSelectObject(hDCMeta, GetStockObject(NULL_BRUSH))
lngRet = apiSelectObject(hDCMeta, GetStockObject(NULL_PEN))
lngRet = SetStretchBltMode(hDCMeta, COLORONCOLOR)
hStdPic.Render CLng(hDCMeta), 0&, 0&, CLng(ImageWidth), CLng(ImageHeight), _
0&, hStdPic.Height, hStdPic.Width, -hStdPic.Height, vbNull
DoEvents
hMetafile = apiCloseEnhMetaFile(hDCMeta)
If hMetafile = 0 Then
    fStdPicToImageData = False
    Exit Function
End If
lngRet = GetEnhMetaFileBits(hMetafile, 0, ByVal 0&)
If lngRet = 0 Then
    fStdPicToImageData = False
    Exit Function
End If
ReDim arrayMeta((lngRet - 1) + 8)
lngRet = GetEnhMetaFileBits(hMetafile, lngRet, arrayMeta(8))
lngRet = apiDeleteEnhMetaFile(hMetafile)
arrayMeta(0) = CF_ENHMETAFILE
ctl.PictureData = arrayMeta
If AutoSize Then
    If sngConvertX < ctl.Parent.Width Then
       ctl.Width = sngConvertX '+ 15
    Else
        ctl.Width = ctl.Parent.Width - 200
    End If
    If sngConvertY < ctl.Parent.Detail.Height Then
        ctl.Height = sngConvertY '+ 15
    Else
        ctl.Height = ctl.Parent.Detail.Height - 200
    End If
     ctl.SizeMode = acOLESizeStretch
End If
EXIT_SHOWPIC:
lngRet = apiReleaseDC(0&, hDCref)
Exit Function
ERR_SHOWPIC:
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume EXIT_SHOWPIC
End Function
Public Function fLoadPicture(ctl As Access.Image, Optional strfName As String = "", Optional AutoSize As Boolean = False) As Boolean
On Error GoTo Err_fLoadPicture
Dim lngRet As Long
Dim blRet As Boolean
Dim hPic As Object
If Len(strfName & vbNullString) = 0 Then
    Dim clsDialog As Object
    Dim strTemp As String
    Set clsDialog = New clsCommonDialog
    clsDialog.Filter = "All (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "JPEG (*.JPG)" & Chr$(0) & "*.JPG" & Chr$(0)
     clsDialog.Filter = clsDialog.Filter & "Tif (*.TIF)" & Chr$(0) & "*.TIF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Gif (*.GIF)" & Chr$(0) & "*.GIF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "PNG (*.PNG)" & Chr$(0) & "*.PNG" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Bitmap (*.BMP)" & Chr$(0) & "*.BMP" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Bitmap (*.DIB)" & Chr$(0) & "*.DIB" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Enhanced Metafile (*.EMF)" & Chr$(0) & "*.EMF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Windows Metafile (*.WMF)" & Chr$(0) & "*.WMF" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Icon (*.ICO)" & Chr$(0) & "*.ICO" & Chr$(0)
    clsDialog.Filter = clsDialog.Filter & "Cursor (*.CUR)" & Chr$(0) & "*.CUR" & Chr$(0)
    clsDialog.hDC = 0
    clsDialog.MaxFileSize = 256
    clsDialog.Max = 256
    clsDialog.FileTitle = vbNullString
    clsDialog.DialogTitle = "Please Select an Image File"
    clsDialog.InitDir = vbNullString
    clsDialog.DefaultExt = vbNullString
    clsDialog.ShowOpen
    strfName = clsDialog.FileName
    If Len(strfName & vbNullString) = 0 Then
      Err.Raise vbObjectError + 513, "LoadJpegGif.modStdPic", _
      "Please Select a Valid JPEG or GIF File"
    End If
End If
Select Case right$(strfName, 3)
    Case "bmp", "dib", "Gif", "emf", "Wmf", "ico", "cur", "jpg"
    Set hPic = LoadPicture(strfName)
    Case "tif", "png"
    Dim GpInput As GdiplusStartupInput
    GpInput.GdiplusVersion = 1
    If (mGDIpEx.GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
      Call MsgBox("Error loading GDI+!", vbCritical)
      Exit Function
    End If
    Set hPic = LoadPictureEx(strfName)
    Call mGDIpEx.GdiplusShutdown(m_GDIpToken)
    Case Else
    Err.Raise vbObjectError + 518, "LoadJpegGif.modStdPic", _
    "This Image format is not supported!" & vbCrLf & strfName & vbCrLf & _
    "Please Select a Supported Image format:" & vbCrLf & _
    "JPEG, TIFF, PNG, BMP, DIB, GIF, EMF, WMF, ICO or CUR"
End Select
If hPic = 0 Then
    Err.Raise vbObjectError + 514, "LoadJpegGif.modStdPic", _
    "Please Select a Supported Image format:" & vbCrLf & _
    "JPEG, TIFF, PNG, BMP, DIB, GIF, EMF, WMF, ICO or CUR"
End If
blRet = fStdPicToImageData(hPic, ctl, , AutoSize)
fLoadPicture = True
Exit_LoadPic:
Application.Echo True
Application.Screen.MousePointer = 0
Err.Clear
Set hPic = Nothing
Set clsDialog = Nothing
Exit Function
Err_fLoadPicture:
fLoadPicture = False
Application.Screen.MousePointer = 0
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_LoadPic
End Function
Public Function fSaveImagetoDisk(ctl As Access.Image) As Boolean
Dim sName As String
Dim lngRet As Long
Dim hEMF As Long
Dim hMetafile As Long
Dim arrayMeta() As Byte
sName = fSavePicture
    If Len(sName & vbNullString) = 0 Then
    fSaveImagetoDisk = False
    Exit Function
End If
ReDim arrayMeta((LenB(ctl.PictureData) - 1))
arrayMeta = ctl.PictureData
If arrayMeta(0) <> CF_ENHMETAFILE Then
    fSaveImagetoDisk = False
    MsgBox "Sorry..not a valid Enhanced Metafile contained in the Image control"
    Exit Function
End If
CopyMem hEMF, arrayMeta(4), 4
hMetafile = CopyEnhMetaFile(hEMF, sName)
lngRet = apiDeleteEnhMetaFile(hMetafile)
End Function
Public Function fSavePicture(Optional strfName As String = "") As String
On Error GoTo Err_fSavePicture
Dim lngRet As Long
Dim blRet As Boolean
If Len(strfName & vbNullString) = 0 Then
    Dim clsDialog As Object
    Dim strTemp As String
    Set clsDialog = New clsCommonDialog
       clsDialog.Filter = clsDialog.Filter & "Enhanced Metafile (*.emf)" & Chr$(0) & "*.emf" & Chr$(0)
    clsDialog.hDC = 0
    clsDialog.MaxFileSize = 256
    clsDialog.Max = 256
    clsDialog.FileTitle = vbNullString
    clsDialog.DialogTitle = "Please Enter a Valid FileName"
    clsDialog.InitDir = vbNullString
    clsDialog.DefaultExt = ".emf" 'vbNullString
    clsDialog.ShowSave
    strfName = clsDialog.FileName
    If Len(strfName & vbNullString) = 0 Then
      Err.Raise vbObjectError + 513, "fSavePicture.modStdPic", _
      "Please Enter a Valid EMF Filename"
    End If
End If
fSavePicture = strfName
Exit_SavePic:
Err.Clear
Set clsDialog = Nothing
Exit Function
Err_fSavePicture:
fSavePicture = strfName
MsgBox Err.Description, vbOKOnly, Err.Source & ":" & Err.Number
Resume Exit_SavePic
End Function
Public Sub ScrollToHome(ctl As Control)
Dim lngRet As Long
Dim lngTemp As Long
On Error Resume Next
Application.Echo False
For lngTemp = 1 To 9
lngRet = SendMessage(ctl.Parent.hWnd, WM_VSCROLL, SB_PAGEUP, 0&)
lngRet = SendMessage(ctl.Parent.hWnd, WM_HSCROLL, SB_PAGELEFT, 0&)
Next lngTemp
Application.Echo True
End Sub
