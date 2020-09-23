VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   3075
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   ShowInTaskbar   =   0   'False
   Begin VB.Timer FadeTimer 
      Left            =   15
      Top             =   435
   End
   Begin VB.Timer DisplayTimer 
      Enabled         =   0   'False
      Left            =   15
      Top             =   0
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As Long, ByVal dwFlags As Long) As Long

Private Type POINTAPI
    X                           As Long
    Y                           As Long
End Type

Private Type Size
    cX                          As Long
    cY                          As Long
End Type

Private Const AC_SRC_ALPHA      As Long = &H1&
Private Const AC_SRC_OVER       As Long = &H0&
Private Const GWL_EXSTYLE       As Long = -20
Private Const HTCAPTION         As Long = 2
Private Const HWND_NOTOPMOST    As Long = -2
Private Const HWND_TOPMOST      As Long = -1
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const ULW_COLORKEY      As Long = &H1
Private Const ULW_ALPHA         As Long = &H2
Private Const WM_NCLBUTTONDOWN  As Long = &HA1
Private Const WS_EX_LAYERED     As Long = &H80000
Private Const WS_EX_LAYOUTRTL   As Long = &H400000

Private cDC                     As Long
Private lInitialStyle           As Long
Private lBlendFunc              As Long
Private ptSrc                   As POINTAPI
Private sizeSrc                 As Size
Private TranspLevel             As Byte

Private cSplashImage            As New c32bppDIB


Private Sub DisplayTimer_Timer()

    FadeForm

End Sub

Private Sub FadeForm()

    FadeTimer.Enabled = True

End Sub

Private Sub FadeTimer_Timer()

    TranspLevel = (TranspLevel - 25)

    If TranspLevel < 25 Then
        Unload Me
        Exit Sub
    End If

    lBlendFunc = AC_SRC_OVER Or (TranspLevel * &H10000) Or (AC_SRC_ALPHA * &H1000000)

    UpdateLayeredWindow Me.hwnd, 0&, ByVal 0&, sizeSrc, cSplashImage.LoadDIBinDC(True), ptSrc, 0&, lBlendFunc, ULW_ALPHA

End Sub

Private Sub Form_Click()

    FadeForm

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    FadeForm

End Sub

Private Sub Form_Load()

    TranspLevel = 255
    FadeTimer.Interval = 100
    FadeTimer.Enabled = False

    With cSplashImage

        .LoadPicture_File App.Path & "\Splash.png"

        lBlendFunc = AC_SRC_OVER Or (TranspLevel * &H10000) Or (AC_SRC_ALPHA * &H1000000)

        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE

        lInitialStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)

        SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

        sizeSrc.cX = .Width

        sizeSrc.cY = .Height

        UpdateLayeredWindow Me.hwnd, 0&, ByVal 0&, sizeSrc, .LoadDIBinDC(True), ptSrc, 0&, lBlendFunc, ULW_ALPHA

        Me.Move ((Screen.Width / 2) - ((.Width * Screen.TwipsPerPixelX) / 2)), ((Screen.Height / 2) - ((.Height * Screen.TwipsPerPixelY) / 2))

    End With

    With DisplayTimer

        .Interval = 3000
        .Enabled = True

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set cSplashImage = Nothing

End Sub

