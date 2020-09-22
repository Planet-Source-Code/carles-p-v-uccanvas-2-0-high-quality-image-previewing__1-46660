VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "ucCanvas 2.0 test"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7425
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   382
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   495
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucCanvas ucCanvas2 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   9763
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Open image..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuCanvasContextTop 
      Caption         =   "Canvas context"
      Visible         =   0   'False
      Begin VB.Menu mnuCanvasContextZoomTop 
         Caption         =   "Zoom"
         Begin VB.Menu mnuCanvasContextZoom 
            Caption         =   "Zoom +"
            Index           =   0
         End
         Begin VB.Menu mnuCanvasContextZoom 
            Caption         =   "Zoom -"
            Index           =   1
         End
         Begin VB.Menu mnuCanvasContextZoom 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuCanvasContextZoom 
            Caption         =   "Best fit mode"
            Index           =   3
         End
      End
      Begin VB.Menu mnuCanvasContextQualityTop 
         Caption         =   "Quality"
         Begin VB.Menu mnuCanvasContextQuality 
            Caption         =   "Default quality"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuCanvasContextQuality 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuCanvasContextQuality 
            Caption         =   "No interpolation"
            Index           =   2
         End
         Begin VB.Menu mnuCanvasContextQuality 
            Caption         =   "Low quality"
            Index           =   3
         End
         Begin VB.Menu mnuCanvasContextQuality 
            Caption         =   "High quality"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Project:       ucCanvas2.ctl test
' Author:        Carles P.V.
' Last revision: 2003.07.05
'================================================



Option Explicit

'-- GDI+
Private m_GDIpToken As Long   ' GDI+ handle
Private m_FileLoad  As String ' Temp. filename path



Private Sub Form_Load()

  Dim GpInput As GdiplusStartupInput
    
    '-- Load the GDI+ Dll
    GpInput.GdiplusVersion = 1
    If (GdiplusStartup(m_GDIpToken, GpInput) <> [OK]) Then
        MsgBox "Error loading GDI+!", vbCritical
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '-- Unload the GDI+ Dll
    GdiplusShutdown m_GDIpToken
End Sub

Private Sub Form_Resize()
    '-- Resize canvas control
    ucCanvas2.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

'//

Private Sub mnuFile_Click(Index As Integer)
    
  Dim sTmpFilename As String
    
    Select Case Index
    
        Case 0 '-- Open image...
        
            '-- Show open file dialog
            sTmpFilename = mDialogFile.GetFileName(m_FileLoad, "BMP (*.bmp)|*.bmp", , "Open image", -1)
            
            If (Len(sTmpFilename)) Then
                m_FileLoad = sTmpFilename
                
                '-- Load image...
                DoEvents
                If (Not ucCanvas2.DIB24.CreateFromStdPicture(LoadPicture(m_FileLoad))) Then
                    MsgBox "Unexpected error loading image.", vbExclamation
                  Else
                    ucCanvas2.Resize
                    ucCanvas2.Repaint
                End If
            End If
            
        Case 2 '-- Exit
            
            Unload Me
    End Select
End Sub

Private Sub mnuCanvasContextZoom_Click(Index As Integer)
'-- Change GDI+ interpolation quality
  
  Dim bValid As Boolean
    
    With ucCanvas2
    
        Select Case Index
            Case 0 ' Zoom +
                If (.Zoom < 15) Then
                    .Zoom = .Zoom + 1: bValid = -1
                End If
            Case 1 ' Zoom -
                If (.Zoom > 0) Then
                    .Zoom = .Zoom - 1: bValid = -1
                End If
            Case 3 ' Best fit mode
                .FitMode = Not .FitMode: bValid = -1
                mnuCanvasContextZoom(3).Checked = .FitMode
        End Select
    
        '-- Refresh canvas [?]
        If (bValid) Then
            .Resize
            .Repaint
        End If
    End With
End Sub

Private Sub mnuCanvasContextQuality_Click(Index As Integer)
'-- Change GDI+ interpolation quality
  
  Dim nIdx As Integer
    
    '-- Uncheck current
    For nIdx = 0 To mnuCanvasContextQuality.Count - 1
        mnuCanvasContextQuality(nIdx).Checked = 0
    Next nIdx
    mnuCanvasContextQuality(Index).Checked = -1
    
    '-- Check new selected
    Select Case Index
        Case 0: mGDIp.GdipStretchQuality = [Default]
        Case 2: mGDIp.GdipStretchQuality = [NoInterpolation]
        Case 3: mGDIp.GdipStretchQuality = [LowQuality]
        Case 4: mGDIp.GdipStretchQuality = [HighQuality]
    End Select
    
    '-- Refresh canvas
    ucCanvas2.Repaint
End Sub

Private Sub mnuHelp_Click()
    
    '-- A simple About
    MsgBox "ucCanvas 2.0 test" & vbCrLf & vbCrLf & _
           "High-quality image stretching through GDI+" & vbCrLf & _
           "Carles P.V. - 2003", , "About"
End Sub

'//

Private Sub ucCanvas2_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    
    If (Button = vbRightButton) Then
        '-- Show context menu
        PopupMenu mnuCanvasContextTop
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    With ucCanvas2
        
        '-- Key control
        Select Case KeyCode
            Case vbKeyAdd      '{NumPad +}
                mnuCanvasContextZoom_Click 0
            Case vbKeySubtract '{NumPad +}
                mnuCanvasContextZoom_Click 1
            Case vbKeyMultiply '{NumPad *}
                mnuCanvasContextZoom_Click 3
        End Select
    End With
End Sub
