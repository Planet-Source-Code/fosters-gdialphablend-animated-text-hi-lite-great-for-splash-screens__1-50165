VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   92
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   4
      Left            =   3240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picStar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   3
      Left            =   2700
      Picture         =   "Form1.frx":0972
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picStar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   2
      Left            =   2160
      Picture         =   "Form1.frx":12E4
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picStar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   1
      Left            =   1620
      Picture         =   "Form1.frx":1C56
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picStar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   0
      Left            =   1080
      Picture         =   "Form1.frx":25C8
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox PictextB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   180
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   2940
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   2520
      Top             =   3000
   End
   Begin VB.PictureBox picTextM 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   180
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   5
      Top             =   3780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   180
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picHilite 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   5940
      Picture         =   "Form1.frx":2F3A
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   1620
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picBlank 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   780
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picDisplay 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   60
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   0
      Top             =   60
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const AC_SRC_OVER As Long = &H0&
Private Const ULW_COLORKEY As Long = &H1&
Private Const ULW_ALPHA As Long = &H2&
Private Const ULW_OPAQUE As Long = &H4&
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_TRANSPARENT  As Long = &H20&
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_POPUP = &H80000000
Private Const WS_VISIBLE = &H10000000
Private Const SPI_GETSELECTIONFADE As Long = &H1014&
Private sText As String
Dim xP As Long
Dim iStar As Integer
Dim iStarDir As Integer
Dim bStarON As Boolean
Private Sub Form_Load()
    ResizePIC picBuffer
    ResizePIC picBlank
    ResizePIC picText
    ResizePIC picTextM
    ResizePIC PictextB
    'draw text mask
    sText = "Laborare non amo"
    pTextOut picTextM, sText, "Impact", 32, False, 2, 2, 0
    pTextOut PictextB, sText, "Impact", 32, False, 2, 2, vbWhite
    
    xP = picHilite.Width * -1
    Timer1.Enabled = True
    bStarON = False
End Sub
Sub ApplyHiLite(x As Long)
Dim BF As BLENDFUNCTION
Dim lBF As Long
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = WS_EX_TRANSPARENT
        .SourceConstantAlpha = 254
        .AlphaFormat = 1
    End With
    RtlMoveMemory lBF, BF, 4
    GdiAlphaBlend picText.hdc, x, 0, picHilite.Width, picHilite.Height, picHilite.hdc, 0, 0, picHilite.Width, picHilite.Height, lBF
End Sub
Sub ApplyStar(img As Integer, x As Long, y As Long)
Dim BF As BLENDFUNCTION
Dim lBF As Long
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = WS_EX_TRANSPARENT
        .SourceConstantAlpha = 254
        .AlphaFormat = 1
    End With
    RtlMoveMemory lBF, BF, 4
    GdiAlphaBlend picBuffer.hdc, x, y, picStar(img).Width, picStar(img).Height, picStar(img).hdc, 0, 0, picStar(img).Width, picStar(img).Height, lBF
End Sub

Sub ClearBuffer()
    BitBlt picBuffer.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub
Sub Cleartext()
    BitBlt picText.hdc, 0, 0, picBuffer.Width, picBuffer.Height, picBlank.hdc, 0, 0, vbSrcCopy
End Sub

Sub ResizePIC(picIN As Control)
    picIN.Move picIN.Left, picIN.Top, picDisplay.Width, picDisplay.Height
End Sub
Sub DrawFrame()

    'clear buffer
    ClearBuffer
    
    'draw black background text to pictext
    Cleartext
    pTextOut picText, sText, "Impact", 32, False, 2, 2, RGB(200, 200, 100)
    
    'apply blend hilite to black back text
    ApplyHiLite xP
    
    'apply black around text again
    BitBlt picText.hdc, 0, 0, picText.Width, picText.Height, PictextB.hdc, 0, 0, vbSrcAnd
    
    'apply black back text and white text mask to buffer
    BitBlt picBuffer.hdc, 0, 0, picText.Width, picText.Height, picTextM.hdc, 0, 0, vbSrcAnd
    BitBlt picBuffer.hdc, 0, 0, picText.Width, picText.Height, picText.hdc, 0, 0, vbSrcPaint

    xP = xP + 8
    If xP > picDisplay.Width Then xP = picHilite.Width * -1
    If xP > 250 And xP < 280 And Not bStarON Then
        bStarON = True
        iStarDir = 1
        iStar = 0
    End If
    If bStarON Then
        
        ApplyStar iStar, 287, 4
        iStar = iStar + iStarDir
        If iStarDir > 0 Then
            If iStar = 4 Then
                iStarDir = iStarDir * -1
                iStar = iStar + iStarDir
            End If
        ElseIf iStarDir < 0 And iStar < 0 Then
            bStarON = False
        End If
    End If

    'buffer to display
    BitBlt picDisplay.hdc, 0, 0, picDisplay.Width, picDisplay.Height, picBuffer.hdc, 0, 0, vbSrcCopy

End Sub
Sub pTextOut(picIN As Control, sIn As String, sFont As String, iFontSize As Integer, bFontBold As Boolean, xPos As Integer, yPos As Integer, lColor As Long)
    
    'SetTextColor picBuffer.hdc, 0
    picIN.Font = sFont
    picIN.FontSize = iFontSize
    picIN.FontBold = bFontBold
    
    'TextOut picIN.hdc, xPos + 1, yPos + 1, sIn, Len(sIn)
    'TextOut picIN.hdc, xPos - 1, yPos - 1, sIn, Len(sIn)
    'TextOut picIN.hdc, xPos - 1, yPos + 1, sIn, Len(sIn)
    'TextOut picIN.hdc, xPos + 1, yPos - 1, sIn, Len(sIn)
    'TextOut picIN.hdc, xPos - 1, yPos, sIn, Len(sIn)
    'TextOut picIN.hdc, xPos + 1, yPos, sIn, Len(sIn)
    
    SetTextColor picIN.hdc, lColor
    TextOut picIN.hdc, xPos, yPos, sIn, Len(sIn)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    DrawFrame
End Sub
