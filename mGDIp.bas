Attribute VB_Name = "mGDIp"
'================================================
' Module:        mGDIp.bas (simplified)
' Author:        *
' Dependencies:  cDIB24.cls
' Last revision: 2003.07.05
'================================================
'
' * From original post:
'   http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=37541&lngWId=1
'   by Avery
'
'   Also, check GDI+ TypeLib TLB
'   http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=42861&lngWId=1
'   by Dana Seaman

Option Explicit

'-- GDI+ API:

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

'//

Private Enum GpUnit
    [UnitWorld]
    [UnitDisplay]
    [UnitPixel]
    [UnitPoint]
    [UnitInch]
    [UnitDocument]
    [UnitMillimeter]
End Enum

Private Enum InterpolationMode
    [InterpolationModeInvalid] = -1
    [InterpolationModeDefault]
    [InterpolationModeLowQuality]
    [InterpolationModeHighQuality]
    [InterpolationModeBilinear]
    [InterpolationModeBicubic]
    [InterpolationModeNearestNeighbor]
    [InterpolationModeHighQualityBilinear]
    [InterpolationModeHighQualityBicubic]
End Enum

Private Enum PixelOffsetMode
    [PixelOffsetModeInvalid] = -1
    [PixelOffsetModeDefault]
    [PixelOffsetModeHighSpeed]
    [PixelOffsetModeHighQuality]
    [PixelOffsetModeNone]
    [PixelOffsetModeHalf]
End Enum

Private Type RECTL
    x As Long
    y As Long
    W As Long
    H As Long
End Type

Private Const PixelFormat24bppRGB As Long = &H21808

'//

Public Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, InputBuf As GdiplusStartupInput, Optional ByVal OutputBuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus

'//

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As GpStatus
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, BITMAP As Long) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As GpStatus
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal Interpolation As InterpolationMode) As GpStatus
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal OffsetMode As PixelOffsetMode) As GpStatus
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As GpUnit, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As GpStatus

'//



'-- Public enums.
Public Enum GdipStretchQualityCts
    [Default] = [InterpolationModeDefault]
    [NoInterpolation] = [InterpolationModeNearestNeighbor]
    [LowQuality] = [InterpolationModeBilinear]
    [HighQuality] = [InterpolationModeHighQualityBicubic]
End Enum

'-- Private variables:
Private m_GdipStretchQuality As GdipStretchQualityCts

'-- Properties:
Public Property Get GdipStretchQuality() As GdipStretchQualityCts
    GdipStretchQuality = m_GdipStretchQuality
End Property

Public Property Let GdipStretchQuality(ByVal Quality As GdipStretchQualityCts)
    m_GdipStretchQuality = Quality
End Property

'-- GDI+ StrecthBlt function
Public Function GdipStretchDIB(oDIB24 As cDIB24, ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long) As Boolean

  Dim gplRet As Long
  Dim hGr    As Long
  Dim hIm    As Long
    
    If (oDIB24.hDIB) Then
        
        If (nSrcWidth <= 0) Then nSrcWidth = oDIB24.Width
        If (nSrcHeight <= 0) Then nSrcHeight = oDIB24.Height
        
        '-- Initialize Graphics object
        gplRet = GdipCreateFromHDC(hDC, hGr)
        
        '-- Create bitmap
        gplRet = GdipCreateBitmapFromScan0(oDIB24.Width, oDIB24.Height, oDIB24.BytesPerScanLine, [PixelFormat24bppRGB], ByVal oDIB24.lpBits, hIm)
         
        '-- Draw it
        gplRet = GdipSetInterpolationMode(hGr, m_GdipStretchQuality)
        gplRet = GdipSetPixelOffsetMode(hGr, [PixelOffsetModeHighQuality])
        gplRet = GdipDrawImageRectRectI(hGr, hIm, x, y, nWidth, nHeight, xSrc, ySrc, nSrcWidth, nSrcHeight, [UnitPixel])
        
        '-- Clean up!
        gplRet = GdipDisposeImage(hIm)
        gplRet = GdipDeleteGraphics(hGr)
        
        '-- Success
        GdipStretchDIB = (gplRet = [OK])
    End If
End Function
