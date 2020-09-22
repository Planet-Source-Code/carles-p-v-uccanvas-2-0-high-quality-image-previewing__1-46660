VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'================================================
' Class:         cDIB24.cls
' Author:        Carles P.V.
' Dependencies:  -
' Last revision: 2003.07.05
'================================================

Option Explicit

'-- API:

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

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    BmBits       As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const OBJ_BITMAP     As Long = 7
Private Const COLORONCOLOR   As Long = 3

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'//

'-- Private Variables:

Private m_tBIH    As BITMAPINFOHEADER
Private m_hDC     As Long
Private m_hDIB    As Long
Private m_hOldDIB As Long
Private m_lpBits  As Long



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()
    '-- Destroy current DIB
    Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function Create(ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean
    
    '-- Destroy previous
    Destroy
    
    '-- Prepare header
    With m_tBIH
        .biSize = Len(m_tBIH)
        .biPlanes = 1
        .biBitCount = 24
        .biWidth = NewWidth
        .biHeight = -NewHeight
        .biSizeImage = -4 * ((.biWidth * 24 + 31) \ 32) * .biHeight
    End With
    
    '-- Create DIB section
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        '-- Create DIB
        m_hDIB = CreateDIBSection(m_hDC, m_tBIH, DIB_RGB_COLORS, m_lpBits, 0, 0)
        If (m_hDIB <> 0) Then
            '-- Select into a DC
            m_hOldDIB = SelectObject(m_hDC, m_hDIB)
          Else
            Destroy
        End If
    End If
    
    '-- Success
    Create = (m_hDIB <> 0)
End Function

Public Function CreateFromStdPicture(Image As StdPicture) As Long

  Dim tBI      As BITMAP
  Dim lhDC     As Long
  Dim lhOldBmp As Long
    
    '-- Check object
    If (Not Image Is Nothing) Then
        '-- Check format
        If (GetObjectType(Image) = OBJ_BITMAP) Then
        
            '-- Get info
            GetObject Image.handle, Len(tBI), tBI
            '-- Create it from StdPicture
            If (Create(tBI.bmWidth, tBI.bmHeight)) Then
                lhDC = CreateCompatibleDC(0)
                If (lhDC <> 0) Then
                    lhOldBmp = SelectObject(lhDC, Image.handle)
                    
                    '-- Load Bits
                    BitBlt m_hDC, 0, 0, tBI.bmWidth, tBI.bmHeight, lhDC, 0, 0, vbSrcCopy
                    '-- Destroy temp. DC
                    SelectObject lhDC, lhOldBmp
                    DeleteDC lhDC
                    '-- Success
                    CreateFromStdPicture = (tBI.bmBitsPixel <> 0)
                End If
            End If
        End If
    End If
End Function

Public Sub Paint(ByVal hDC As Long, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, Optional ByVal ScaleFactor As Long = 1)
    
  Dim lRet As Long
  
    If (m_hDIB <> 0) Then
        With m_tBIH
            '-- BitBlt w/ scale factor
            lRet = StretchBlt(hDC, x, y, .biWidth * ScaleFactor, .biHeight * ScaleFactor, m_hDC, 0, 0, .biWidth, .biHeight, vbSrcCopy)
        End With
    End If
End Sub

Public Sub Stretch(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long)
  
  Dim lMode As Long
  Dim lRet  As Long
    
    If (m_hDIB <> 0) Then
        '-- StretchBlt
        lMode = GetStretchBltMode(hDC)
        SetStretchBltMode hDC, COLORONCOLOR
        lRet = StretchBlt(hDC, x, y, nWidth, nHeight, m_hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, vbSrcCopy)
        SetStretchBltMode hDC, lMode
    End If
End Sub

Public Sub Destroy()
    
    '-- Destroy DIB section
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            SelectObject m_hDC, m_hOldDIB
            DeleteObject m_hDIB
        End If
        DeleteDC m_hDC
    End If
    
    '-- Reset vars.
    m_hDC = 0: m_hDIB = 0: m_hOldDIB = 0: m_lpBits = 0
    ZeroMemory m_tBIH, Len(m_tBIH)
End Sub

Public Sub GetBestFitInfo(ByVal DstW As Long, ByVal DstH As Long, bfx As Long, bfy As Long, bfW As Long, bfH As Long, Optional ByVal StretchFit As Boolean = 0)
  
  Dim cW As Single
  Dim cH As Single
    
    If (m_hDIB <> 0) Then
    
        '-- Get best fit dimensions
        With m_tBIH
            If ((.biWidth > DstW Or -.biHeight > DstH) Or StretchFit) Then
                cW = DstW / .biWidth
                cH = DstH / -.biHeight
                If (cW < cH) Then
                    bfW = DstW
                    bfH = -.biHeight * cW
                  Else
                    bfH = DstH
                    bfW = .biWidth * cH
                End If
              Else
                bfW = .biWidth
                bfH = -.biHeight
            End If
        End With
        
        '-- Get best fit offsets
        bfx = (DstW - bfW) \ 2
        bfy = (DstH - bfH) \ 2
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get hDIBDC() As Long
    hDIBDC = m_hDC
End Property

Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property

Public Property Get lpBits() As Long
    lpBits = m_lpBits
End Property

Public Property Get Size() As Long
    Size = m_tBIH.biSizeImage
End Property

Public Property Get BytesPerScanLine() As Long
    BytesPerScanLine = ((m_tBIH.biWidth * 24 + 31) \ 32) * 4
End Property

Public Property Get Width() As Long
    Width = m_tBIH.biWidth
End Property

Public Property Get Height() As Long
    Height = -m_tBIH.biHeight
End Property
