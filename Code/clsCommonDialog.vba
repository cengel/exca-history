Option Compare Database
Option Explicit
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFilename) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFilename) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Type ChooseColor
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        Flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private Type OpenFilename
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        iFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        Flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
Private iAction As Integer         'internal buffer for Action property
Private bCancelError As Boolean    'internal buffer for CancelError property
Private lColor As Long             'internal buffer for Color property
Private lCopies As Long            'internal buffer for lCopies property
Private sDefaultExt As String      'internal buffer for sDefaultExt property
Private sDialogTitle As String     'internal buffer for DialogTitle property
Private sFilename As String        'internal buffer for FileName property
Private sFileTitle As String       'internal buffer for FileTitle property
Private sFilter As String          'internal buffer for Filter property
Private iFilterIndex As Integer    'internal buffer for FilterIndex property
Private lFlags As Long             'internal buffer for Flags property
Private lhDC As Long               'internal buffer for hdc property
Private sInitDir As String         'internal buffer for InitDir property
Private lMax As Long               'internal buffer for Max property
Private lMaxFileSize As Long       'internal buffer for MaxFileSize property
Private lMin As Long               'internal buffer for Min property
Private objObject As Object        'internal buffer for Object property
Private lApiReturn As Long          'internal buffer for APIReturn property
Private lExtendedError As Long      'internal buffer for ExtendedError property
Private Const CDERR_DIALOGFAILURE = &HFFFF
Private Const CDERR_FINDRESFAILURE = &H6
Private Const CDERR_GENERALCODES = &H0
Private Const CDERR_INITIALIZATION = &H2
Private Const CDERR_LOADRESFAILURE = &H7
Private Const CDERR_LOADSTRFAILURE = &H5
Private Const CDERR_LOCKRESFAILURE = &H8
Private Const CDERR_MEMALLOCFAILURE = &H9
Private Const CDERR_MEMLOCKFAILURE = &HA
Private Const CDERR_NOHINSTANCE = &H4
Private Const CDERR_NOHOOK = &HB
Private Const CDERR_NOTEMPLATE = &H3
Private Const CDERR_REGISTERMSGFAIL = &HC
Private Const CDERR_STRUCTSIZE = &H1
Private Const FNERR_BUFFERTOOSMALL = &H3003
Private Const FNERR_FILENAMECODES = &H3000
Private Const FNERR_INVALIDFILENAME = &H3002
Private Const FNERR_SUBCLASSFAILURE = &H3001
Public Property Get Filter() As String
    Filter = sFilter
End Property
Public Sub ShowColor()
    Dim tChooseColor As ChooseColor
    Dim alCustomColors(15) As Long
    Dim lCustomColorSize As Long
    Dim lCustomColorAddress As Long
    Dim lMemHandle As Long
    Dim n As Integer
    On Error GoTo ShowColorError
    iAction = 3  'Action property - ShowColor
    lApiReturn = 0  'APIReturn property
    lExtendedError = 0  'ExtendedError property
    tChooseColor.lStructSize = Len(tChooseColor)
    tChooseColor.hwndOwner = Application.hWndAccessApp
    tChooseColor.rgbResult = lColor
    For n = 0 To UBound(alCustomColors)
        alCustomColors(n) = &HFFFFFF
    Next
    lCustomColorSize = Len(alCustomColors(0)) * 16
    lMemHandle = GlobalAlloc(GHND, lCustomColorSize)
    If lMemHandle = 0 Then
        Exit Sub
    End If
    lCustomColorAddress = GlobalLock(lMemHandle)
    If lCustomColorAddress = 0 Then
        Exit Sub
    End If
    Call CopyMemory(ByVal lCustomColorAddress, alCustomColors(0), lCustomColorSize)
    tChooseColor.lpCustColors = lCustomColorAddress
    tChooseColor.Flags = lFlags
    lApiReturn = ChooseColor(tChooseColor)
    Select Case lApiReturn
        Case 0  'user canceled
        If bCancelError = True Then
            On Error GoTo 0
            Err.Raise Number:=vbObjectError + 894, _
                Description:="Cancel Pressed"
            Exit Sub
        End If
        Case 1  'user selected a color
            lColor = tChooseColor.rgbResult
        Case Else   'an error occured
            lExtendedError = CommDlgExtendedError
    End Select
Exit Sub
ShowColorError:
    Exit Sub
End Sub
Public Sub ShowOpen()
    ShowFileDialog (1)  'Action property - ShowOpen
End Sub
Public Sub ShowSave()
    ShowFileDialog (2)  'Action property - ShowSave
End Sub
Public Property Get FileName() As String
    FileName = sFilename
End Property
Public Property Let FileName(vNewValue As String)
    sFilename = vNewValue
End Property
Public Property Let Filter(vNewValue As String)
    sFilter = vNewValue
End Property
Private Function sLeftOfNull(ByVal sIn As String)
    Dim lNullPos As Long
    sLeftOfNull = sIn
    lNullPos = InStr(sIn, Chr$(0))
    If lNullPos > 0 Then
        sLeftOfNull = Mid$(sIn, 1, lNullPos - 1)
    End If
End Function
Public Property Get Action() As Integer
    Action = iAction
End Property
Private Function sAPIFilter(sIn)
    Dim lChrNdx As Long
    Dim sOneChr As String
    Dim sOutStr As String
    For lChrNdx = 1 To Len(sIn)
        sOneChr = Mid$(sIn, lChrNdx, 1)
        If sOneChr = "|" Then
            sOutStr = sOutStr & Chr$(0)
        Else
            sOutStr = sOutStr & sOneChr
        End If
    Next
    sOutStr = sOutStr & Chr$(0)
    sAPIFilter = sOutStr
End Function
Public Property Get FilterIndex() As Integer
    FilterIndex = iFilterIndex
End Property
Public Property Let FilterIndex(vNewValue As Integer)
    iFilterIndex = vNewValue
End Property
Public Property Get CancelError() As Boolean
    CancelError = bCancelError
End Property
Public Property Let CancelError(vNewValue As Boolean)
    bCancelError = vNewValue
End Property
Public Property Get Color() As Long
    Color = lColor
End Property
Public Property Let Color(vNewValue As Long)
    lColor = vNewValue
End Property
Public Property Get DefaultExt() As String
    DefaultExt = sDefaultExt
End Property
Public Property Let DefaultExt(vNewValue As String)
    sDefaultExt = vNewValue
End Property
Public Property Get DialogTitle() As String
    DialogTitle = sDialogTitle
End Property
Public Property Let DialogTitle(vNewValue As String)
    sDialogTitle = vNewValue
End Property
Public Property Get Flags() As Long
    Flags = lFlags
End Property
Public Property Let Flags(vNewValue As Long)
    lFlags = vNewValue
End Property
Public Property Get hDC() As Long
    hDC = lhDC
End Property
Public Property Let hDC(vNewValue As Long)
    lhDC = vNewValue
End Property
Public Property Get InitDir() As String
    InitDir = sInitDir
End Property
Public Property Let InitDir(vNewValue As String)
    sInitDir = vNewValue
End Property
Public Property Get Max() As Long
    Max = lMax
End Property
Public Property Let Max(vNewValue As Long)
    lMax = vNewValue
End Property
Public Property Get MaxFileSize() As Long
    MaxFileSize = lMaxFileSize
End Property
Public Property Let MaxFileSize(vNewValue As Long)
    lMaxFileSize = vNewValue
End Property
Public Property Get Min() As Long
    Min = lMin
End Property
Public Property Let Min(vNewValue As Long)
    lMin = vNewValue
End Property
Public Property Get Object() As Object
    Object = objObject
End Property
Public Property Let Object(vNewValue As Object)
    objObject = vNewValue
End Property
Public Property Get FileTitle() As String
    FileTitle = sFileTitle
End Property
Public Property Let FileTitle(vNewValue As String)
    sFileTitle = vNewValue
End Property
Public Property Get APIReturn() As Long
    APIReturn = lApiReturn
End Property
Public Property Get ExtendedError() As Long
    ExtendedError = lExtendedError
End Property
Private Function sByteArrayToString(abBytes() As Byte) As String
    Dim lBytePoint As Long
    Dim lByteVal As Long
    Dim sOut As String
    lBytePoint = LBound(abBytes)
    While lBytePoint <= UBound(abBytes)
        lByteVal = abBytes(lBytePoint)
        If lByteVal = 0 Then
            sByteArrayToString = sOut
            Exit Function
        Else
            sOut = sOut & Chr$(lByteVal)
        End If
        lBytePoint = lBytePoint + 1
    Wend
    sByteArrayToString = sOut
End Function
Private Sub ShowFileDialog(ByVal iAction As Integer)
    Dim tOpenFile As OpenFilename
    Dim lMaxSize As Long
    Dim sFileNameBuff As String
    Dim sFileTitleBuff As String
    On Error GoTo ShowFileDialogError
    iAction = iAction  'Action property
    lApiReturn = 0  'APIReturn property
    lExtendedError = 0  'ExtendedError property
    tOpenFile.lStructSize = Len(tOpenFile)
    tOpenFile.hwndOwner = Application.hWndAccessApp
    tOpenFile.lpstrFilter = sAPIFilter(sFilter)
    tOpenFile.iFilterIndex = iFilterIndex
        If lMaxFileSize > 0 Then
            lMaxSize = lMaxFileSize
        Else
            lMaxSize = 256
        End If
        sFileNameBuff = sFilename
        While Len(sFileNameBuff) < lMaxSize - 1
            sFileNameBuff = sFileNameBuff & " "
        Wend
       sFileNameBuff = Mid$(sFileNameBuff, 1, lMaxFileSize - 1)
        sFileNameBuff = sFileNameBuff & Chr$(0)
    tOpenFile.lpstrFile = sFileNameBuff
    If lMaxFileSize <> 255 Then  'default is 255
        tOpenFile.nMaxFile = lMaxFileSize
    End If
        sFileTitleBuff = sFileTitle
        While Len(sFileTitleBuff) < lMaxSize - 1
            sFileTitleBuff = sFileTitleBuff & " "
        Wend
        sFileTitleBuff = Mid$(sFileTitleBuff, 1, lMaxFileSize - 1)
        sFileTitleBuff = sFileTitleBuff & Chr$(0)
    tOpenFile.lpstrFileTitle = sFileTitleBuff
    tOpenFile.lpstrInitialDir = sInitDir
    tOpenFile.lpstrTitle = sDialogTitle
    tOpenFile.Flags = lFlags
    tOpenFile.lpstrDefExt = sDefaultExt
    Select Case iAction
        Case 1  'ShowOpen
            lApiReturn = GetOpenFileName(tOpenFile)
        Case 2  'ShowSave
            lApiReturn = GetSaveFileName(tOpenFile)
        Case Else   'unknown action
            Exit Sub
    End Select
    Select Case lApiReturn
        Case 0  'user canceled
        If bCancelError = True Then
            Err.Raise (2001)
            Exit Sub
        End If
        Case 1  'user selected or entered a file
            sFilename = sLeftOfNull(tOpenFile.lpstrFile)
            sFileTitle = sLeftOfNull(tOpenFile.lpstrFileTitle)
        Case Else   'an error occured
            lExtendedError = CommDlgExtendedError
    End Select
Exit Sub
ShowFileDialogError:
    Exit Sub
End Sub
