Option Explicit
Public Enum GpImageFormat
    [ImageGIF] = 0
    [ImageJPEG] = 1
    [ImagePNG] = 2
    [ImageTIFF] = 3
End Enum
Public Enum GpStatus
    [OK] = 0
    [GenericError] = 1
    [InvalidParameter] = 2
    [OutOfMemory] = 3
    [ObjectBusy] = 4
    [InsufficientBuffer] = 5
    [NotImplemented] = 6
    [Win32Error] = 7
    [WrongState] = 8
    [Aborted] = 9
    [FileNotFound] = 10
    [ValueOverflow ] = 11
    [AccessDenied] = 12
    [UnknownImageFormat] = 13
    [FontFamilyNotFound] = 14
    [FontStyleNotFound] = 15
    [NotTrueTypeFont] = 16
    [UnsupportedGdiplusVersion] = 17
    [GdiplusNotInitialized ] = 18
    [PropertyNotFound] = 19
    [PropertyNotSupported] = 20
End Enum
Public Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Private Enum GpUnit
    [UnitWorld]
    [UnitDisplay]
    [UnitPixel]
    [UnitPoint]
    [UnitInch]
    [UnitDocument]
    [UnitMillimeter]
End Enum
Private Enum QualityMode
    [QualityModeInvalid] = -1
    [QualityModeDefault] = 0
    [QualityModeLow] = 1
    [QualityModeHigh] = 2
End Enum
Private Enum PixelOffsetMode
    [PixelOffsetModeInvalid] = -1
    [PixelOffsetModeDefault]
    [PixelOffsetModeHighSpeed]
    [PixelOffsetModeHighQuality]
    [PixelOffsetModeNone]
    [PixelOffsetModeHalf]
End Enum
Private Enum InterpolationMode
    [InterpolationModeInvalid] = [QualityModeInvalid]
    [InterpolationModeDefault] = [QualityModeDefault]
    [InterpolationModeLowQuality] = [QualityModeLow]
    [InterpolationModeHighQuality] = [QualityModeHigh]
    [InterpolationModeBilinear]
    [InterpolationModeBicubic]
    [InterpolationModeNearestNeighbor]
    [InterpolationModeHighQualityBilinear]
    [InterpolationModeHighQualityBicubic]
End Enum
Private Enum EncoderParameterValueType
    [EncoderParameterValueTypeByte] = 1
    [EncoderParameterValueTypeASCII] = 2
    [EncoderParameterValueTypeShort] = 3
    [EncoderParameterValueTypeLong] = 4
    [EncoderParameterValueTypeRational] = 5
    [EncoderParameterValueTypeLongRange] = 6
    [EncoderParameterValueTypeUndefined] = 7
    [EncoderParameterValueTypeRationalRange] = 8
End Enum
Private Enum EncoderValue
    [EncoderValueColorTypeCMYK] = 0
    [EncoderValueColorTypeYCCK] = 1
    [EncoderValueCompressionLZW] = 2
    [EncoderValueCompressionCCITT3] = 3
    [EncoderValueCompressionCCITT4] = 4
    [EncoderValueCompressionRle] = 5
    [EncoderValueCompressionNone] = 6
    [EncoderValueScanMethodInterlaced]
    [EncoderValueScanMethodNonInterlaced]
    [EncoderValueVersionGif87]
    [EncoderValueVersionGif89]
    [EncoderValueRenderProgressive]
    [EncoderValueRenderNonProgressive]
    [EncoderValueTransformRotate90]
    [EncoderValueTransformRotate180]
    [EncoderValueTransformRotate270]
    [EncoderValueTransformFlipHorizontal]
    [EncoderValueTransformFlipVertical]
    [EncoderValueMultiFrame]
    [EncoderValueLastFrame]
    [EncoderValueFlush]
    [EncoderValueFrameDimensionTime]
    [EncoderValueFrameDimensionResolution]
    [EncoderValueFrameDimensionPage]
End Enum
Private Type CLSID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type
Private Type ImageCodecInfo
    ClassID           As CLSID
    FormatID          As CLSID
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    Flags             As Long
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type
Private Type EncoderParameter
    GUID           As CLSID
    NumberOfValues As Long
    Type           As EncoderParameterValueType
    Value          As Long
End Type
Private Type EncoderParameters
    Count     As Long
    Parameter As EncoderParameter
End Type
Private Const EncoderCompression      As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Private Const EncoderColorDepth       As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Private Const EncoderScanMethod       As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Private Const EncoderVersion          As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Private Const EncoderRenderMethod     As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Private Const EncoderQuality          As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Private Const EncoderTransformation   As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Private Const EncoderLuminanceTable   As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Private Const EncoderChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Private Const EncoderSaveFlag         As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
Private Const CodecIImageBytes        As String = "{025D1823-6C7D-447B-BBDB-A3CBC3DFA2FC}"
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type
Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type PICTDESC
    Size       As Long
    Type       As Long
    hBmpOrIcon As Long
    hPal       As Long
End Type
Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As GpStatus
Private Declare Function GdipGetImageDecodersSize Lib "gdiplus" (numDecoders As Long, Size As Long) As GpStatus
Private Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal numDecoders As Long, ByVal Size As Long, Decoders As Any) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As GpStatus
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hBmpReturn As Long, ByVal Background As Long) As GpStatus
Private Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, gdiBitmapData As Any, Bitmap As Long) As GpStatus
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As String, hImage As Long) As GpStatus
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As String, clsidEncoder As CLSID, encoderParams As Any) As GpStatus
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal OffsetMode As PixelOffsetMode) As GpStatus
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As InterpolationMode) As GpStatus
Private Declare Function GdipDrawImageRectRect Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As CLSID) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (lpPictDesc As PICTDESC, riid As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long
Private Const vbPicTypeNone = 0 'Picture is empty
Private Const vbPicTypeBitmap = 1 'Bitmap (.bmpBMP files)
Private Const vbPicTypeMetafile = 2 'Metafile (.wmfWMF files)
Private Const vbPicTypeIcon = 3 'Icon (.icoICO files)
Private Const vbPicTypeEMetafile = 4 'Enhanced Metafile (.emfEMF files)
Public Function LoadPictureEx(ByVal sFilename As String) As StdPicture
  Dim gplRet        As Long
  Dim hImg          As Long
  Dim hBmp          As Long
  Dim uPictDesc     As PICTDESC
  Dim aGuid(0 To 3) As Long
  Dim varString As Variant
    varString = StrConv(sFilename, vbUnicode)
     Call GdipLoadImageFromFile(varString, hImg)
    gplRet = GdipCreateHBITMAPFromBitmap(hImg, hBmp, vbBlack)
    gplRet = GdipDisposeImage(hImg)
    If (gplRet = [OK]) Then
        With uPictDesc
            .Size = Len(uPictDesc)
            .Type = vbPicTypeBitmap
            .hBmpOrIcon = hBmp
            .hPal = 0
        End With
        aGuid(0) = &H7BF80980
        aGuid(1) = &H101ABF32
        aGuid(2) = &HAA00BB8B
        aGuid(3) = &HAB0C3000
        OleCreatePictureIndirect uPictDesc, aGuid(0), -1, LoadPictureEx
    End If
End Function
Private Function pvGetEncoderClsID(strMimeType As String, ClassID As CLSID) As Long
  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    pvGetEncoderClsID = -1 ' Failure flag
    Call GdipGetImageEncodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    Call GdipGetImageEncoders(Num, Size, Buffer(1))
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    For lIdx = 1 To Num
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(lIdx).ClassID ' Save the Class ID
            pvGetEncoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    Erase ICI
    Erase Buffer
End Function
Private Function pvGetDecoderClsID(strMimeType As String, ClassID As CLSID) As Long
  Dim Num      As Long
  Dim Size     As Long
  Dim lIdx     As Long
  Dim ICI()    As ImageCodecInfo
  Dim Buffer() As Byte
    pvGetDecoderClsID = -1 'Failure flag
    Call GdipGetImageDecodersSize(Num, Size)
    If (Size = 0) Then Exit Function ' Failed!
    ReDim ICI(1 To Num) As ImageCodecInfo
    ReDim Buffer(1 To Size) As Byte
    Call GdipGetImageDecoders(Num, Size, Buffer(1))
    Call CopyMemory(ICI(1), Buffer(1), (Len(ICI(1)) * Num))
    For lIdx = 1 To Num
        If (StrComp(pvPtrToStrW(ICI(lIdx).MimeType), strMimeType, vbTextCompare) = 0) Then
            ClassID = ICI(lIdx).ClassID ' Save the Class ID
            pvGetDecoderClsID = lIdx      ' Return the index number for success
            Exit For
        End If
    Next lIdx
    Erase ICI
    Erase Buffer
End Function
Private Function pvDEFINE_GUID(ByVal sGuid As String) As CLSID
    Call CLSIDFromString(StrPtr(sGuid), pvDEFINE_GUID)
End Function
Private Function pvPtrToStrW(ByVal lpsz As Long) As String
  Dim sOut As String
  Dim lLen As Long
    lLen = lstrlenW(lpsz)
    If (lLen > 0) Then
        sOut = StrConv(String$(lLen, vbNullChar), vbUnicode)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen * 2)
        pvPtrToStrW = StrConv(sOut, vbFromUnicode)
    End If
End Function
Private Function pvPtrToStrA(ByVal lpsz As Long) As String
  Dim sOut As String
  Dim lLen As Long
    lLen = lstrlenA(lpsz)
    If (lLen > 0) Then
        sOut = String$(lLen, vbNullChar)
        Call CopyMemory(ByVal sOut, ByVal lpsz, lLen)
        pvPtrToStrA = sOut
    End If
End Function
