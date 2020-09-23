VERSION 5.00
Begin VB.Form frmDesktopDisplay 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmDesktopDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal uFormat As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As Long, ByVal dwFlags As Long) As Long

Private Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type

Private Type typeGlobalInfo
    Tooltip_Backcolor               As Long
    tooltip_Forecolor               As Long
End Type

Private Type TOOLINFO
    cbSize                          As Long
    uFlags                          As Long
    hwnd                            As Long
    uid                             As Long
    rc                              As RECT
    hinst                           As Long
    lpszText                        As String
    lParam                          As Long
End Type

Private Type typeOverlay
    Format                          As String
    Width                           As Integer
    Height                          As Integer
    Scale                           As Integer
    Opacity                         As Integer
    Colorize                        As Boolean
    Hue                             As Single
    Saturation                      As Single
    Luminosity                      As Single
    Rotate                          As Integer
    Position_X                      As Integer
    Position_Y                      As Integer
    Image                           As String
End Type

Private Type typeImages
    Format                          As String
    Width                           As Integer
    Height                          As Integer
    Opacity                         As Integer
    Colorize                        As Boolean
    Hue                             As Single
    Saturation                      As Single
    Luminosity                      As Single
    Naming                          As String
    LowestValue                     As Integer
    HighestValue                    As Integer
    Increment                       As Integer
End Type

Private Type POINTAPI
    X                               As Long
    Y                               As Long
End Type

Private Type Size
    cX                              As Long
    cY                              As Long
End Type

Private Const AC_SRC_OVER           As Long = &H0&
Private Const AC_SRC_ALPHA          As Long = &H1&

Private Const CW_USEDEFAULT         As Long = &H80000000

Private Const DT_BOTTOM             As Long = &H8
Private Const DT_CENTER             As Long = &H1
Private Const DT_LEFT               As Long = &H0
Private Const DT_RIGHT              As Long = &H2
Private Const DT_TOP                As Long = &H0
Private Const DT_VCENTER            As Long = &H4
Private Const DT_WORDBREAK          As Long = &H10
Private Const DT_TRANSPARENT        As Long = &H1

Private Const GWL_EXSTYLE           As Long = -20

Private Const HWND_NOTOPMOST        As Long = -2
Private Const HWND_TOPMOST          As Long = -1

Private Const HTCAPTION             As Long = 2

Private Const RDW_ALLCHILDREN       As Long = &H80
Private Const RDW_ERASE             As Long = &H4
Private Const RDW_FRAME             As Long = &H400
Private Const RDW_INVALIDATE        As Long = &H1

Private Const SWP_NOMOVE            As Long = &H2
Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_NOACTIVATE        As Long = &H10

Private Const TTF_IDISHWND          As Long = &H1
Private Const TTF_SUBCLASS          As Long = &H10

Private Const TTS_ALWAYSTIP         As Long = &H1
Private Const TTS_NOPREFIX          As Long = &H2
Private Const TTS_NOANIMATE         As Long = &H10
Private Const TTS_NOFADE            As Long = &H20
Private Const TTS_BALLOON           As Long = &H40
Private Const TTS_CLOSE             As Long = &H80

Private Const TTM_ACTIVATE          As Long = (&H400 + 1)
Private Const TTM_ADDTOOL           As Long = (&H400 + 4)
Private Const TTM_UPDATETIPTEXT     As Long = (&H400 + 12)
Private Const TTM_SETTIPBKCOLOR     As Long = (&H400 + 19)
Private Const TTM_SETTIPTEXTCOLOR   As Long = (&H400 + 20)
Private Const TTM_SETMAXTIPWIDTH    As Long = (&H400 + 24)
Private Const TTM_SETTITLE          As Long = (&H400 + 32)

Private Const ULW_COLORKEY          As Long = &H1
Private Const ULW_ALPHA             As Long = &H2

Private Const WM_NCLBUTTONDOWN      As Long = &HA1
Private Const WM_USER               As Long = &H400

Private Const WS_EX_LAYERED         As Long = &H80000
Private Const WS_EX_LAYOUTRTL       As Long = &H400000
Private Const WS_EX_TOPMOST         As Long = &H8&
Private Const WS_POPUP              As Long = &H80000000

Private Const TOOLTIPS_CLASSA       As String = "tooltips_class32"

Private cDC                         As Long
Private cImage                      As New c32bppDIB
Private cOverlay                    As New c32bppDIB
Private hwndTip                     As Long
Private iImageSize                  As Integer
Public iImageSizeIndex              As Integer
Private lBlendFunc                  As Long
Private oImage                      As typeImages
Private oOverlay                    As typeOverlay
Private oGlobal                     As typeGlobalInfo
Private ptSrc                       As POINTAPI
Private sizeSrc                     As Size

Private m_StyleFolder               As String
Private m_InitialStyle              As Long


Private Function BuildBatteryImage() As String

  Dim iPL As Integer

    With oImage

        iPL = Int(iPowerLevel / .Increment) * .Increment

        If iPL < .LowestValue Then iPL = .LowestValue
        If iPL > .HighestValue Then iPL = .HighestValue

        BuildBatteryImage = Replace(m_StyleFolder & "\" & oImage.Naming, "???", Format(iPL, "0#"))

    End With

End Function

Private Sub CreateToolTip()

    If hwndTip <> 0 Then
        tooltip_destroy hwndTip
    End If

    hwndTip = tooltip_create(Me.hwnd, True, True)

    If hwndTip <> 0 Then

        tooltip_setbackcolour hwndTip, oGlobal.Tooltip_Backcolor
        tooltip_settextcolour hwndTip, oGlobal.tooltip_Forecolor

        tooltip_setmaxwidth hwndTip, 0

        tooltip_activate hwndTip

        tooltip_titleadd hwndTip, "Battery Monitor 1.0.1", 1

        tooltip_addtool hwndTip, Me.hwnd, "Battery Level : " & SysPower.BatteryLifePercent & vbCrLf & _
                                          "Running On    :  " & IIf(SysPower.ACLineStatus = 0, "Battery", "A/C Power")

    End If

End Sub

Public Property Get DesktopBatteryTopMost() As Boolean

    DesktopBatteryTopMost = bDesktopBatteryTopMost

End Property

Public Property Let DesktopBatteryTopMost(ByVal bValue As Boolean)

    bDesktopBatteryTopMost = bValue

    SetWindowPos Me.hwnd, IIf(bDesktopBatteryTopMost, HWND_TOPMOST, HWND_NOTOPMOST), Me.Left / 15, Me.Top / 15, 0, 0, SWP_NOACTIVATE

    PaintImage

End Property

Private Sub Form_DblClick()

    ShowCustomSizeWindow

End Sub

Private Sub Form_Initialize()

    Call InitCommonControls

    m_StyleFolder = App.Path & "\Style1"

    ReadConfigFile

    bAttachedToAC = True

    iPowerLevel = -1

    With cImage
        .LoadPicture_File BuildBatteryImage, , , True
        .HighQualityInterpolation = True
    End With

    With cOverlay
        .LoadPicture_File m_StyleFolder & "\" & oOverlay.Image
        .HighQualityInterpolation = True
    End With

    lBlendFunc = AC_SRC_OVER Or (255& * &H10000) Or (AC_SRC_ALPHA * &H1000000)

    m_InitialStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)

    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

End Sub

Private Sub Form_Load()

    PaintImage

    CreateToolTip

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then

        ReleaseCapture

        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&

        CreateToolTip

    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbRightButton Then

        PopupMenu frmBMon.mnuPop

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    tooltip_destroy hwndTip

    Set cImage = Nothing
    Set cOverlay = Nothing

    Unload Me

End Sub

Public Property Get ImageSize() As Integer

    ImageSize = iImageSize

End Property

Public Sub PaintImage()

  Static iPrevPowerLevel    As Integer

  Dim DIBhdc                As Long
  Dim ret                   As Boolean
  Dim sBatteryImage         As String

    iImageSize = IIf(iImageSize <= 0, 128, iImageSize)

    With cImage

        If iPrevPowerLevel <> iPowerLevel Then
            ret = .LoadPicture_File(BuildBatteryImage, , , True)
         Else
            .LoadPicture_FromOrignalFormat
        End If

        If oImage.Colorize = True Then
            .Colorize oImage.Hue, oImage.Saturation, oImage.Luminosity
        End If

        If bAttachedToAC = True Then

            If oOverlay.Colorize = True Then
                cOverlay.Colorize oOverlay.Hue, oOverlay.Saturation, oOverlay.Luminosity
            End If

            If oOverlay.Rotate <> 0 Then
                cOverlay.RotateAtTopLeft .LoadDIBinDC(True), oOverlay.Rotate, oOverlay.Position_X, oOverlay.Position_Y, cOverlay.Width * (oOverlay.Scale / 100), cOverlay.Height * (oOverlay.Scale / 100), , , , , oOverlay.Opacity
             Else
                cOverlay.Render .LoadDIBinDC(True), oOverlay.Position_X, oOverlay.Position_Y, cOverlay.Width * (oOverlay.Scale / 100), cOverlay.Height * (oOverlay.Scale / 100), , , , , oOverlay.Opacity
            End If

        End If

        If .Width <> iImageSize Then
            .Resize iImageSize, iImageSize
        End If

        sizeSrc.cX = .Width
        sizeSrc.cY = .Height

        '        Dim rct As RECT
        '        rct.Left = 0
        '        rct.Top = .Height - 27
        '        rct.Right = .Width
        '        rct.Bottom = .Height
        '
        '        Me.Font.Size = 132
        '        SetBkMode .LoadDIBinDC(True), DT_TRANSPARENT
        '        SetTextColor .LoadDIBinDC(True), vbBlack
        '        DrawText .LoadDIBinDC(True), CStr(iPowerLevel) & "%", Len(CStr(iPowerLevel) & "%"), rct, DT_CENTER

        '        rct.Left = 5
        '        rct.Top = .Height - 28
        '        SetTextColor .LoadDIBinDC(True), vbRed
        '        DrawText .LoadDIBinDC(True), CStr(iPowerLevel) & "%", Len(CStr(iPowerLevel) & "%"), rct, DT_CENTER

        SetBkMode .LoadDIBinDC(True), 0&

        lBlendFunc = AC_SRC_OVER Or (bTransparency * &H10000) Or (AC_SRC_ALPHA * &H1000000)

        UpdateLayeredWindow Me.hwnd, 0&, ByVal 0&, sizeSrc, .LoadDIBinDC(True), ptSrc, 0&, lBlendFunc, ULW_ALPHA

    End With

End Sub

Private Sub ReadConfigFile()

  Dim lFileNum  As Long
  Dim sInput    As String
  Dim sKey      As String
  Dim sSection  As String
  Dim sValue    As String

    lFileNum = FreeFile
    Open m_StyleFolder & "\config.ini" For Input As lFileNum

    While Not EOF(lFileNum)

        Line Input #lFileNum, sInput

        sInput = Trim(sInput)

        If Left(sInput, 1) <> "'" Then

            If Left(sInput, 1) = "[" Then
                sSection = Replace(sInput, "[", "")
                sSection = LCase(Trim(Replace(sSection, "]", "")))

             Else

                If InStr(1, sInput, "=", vbTextCompare) > 1 Then
                    sKey = LCase(Trim(Left(sInput, InStr(1, sInput, "=", vbTextCompare) - 1)))
                    sValue = LCase(Trim(Mid(sInput, InStr(1, sInput, "=", vbTextCompare) + 1)))
                    sValue = LCase(Replace(sValue, "'", ""))

                    Select Case sSection

                     Case "overlay"

                        With oOverlay

                            Select Case sKey
                             Case "format"
                                .Format = "." & sValue

                             Case "width"
                                .Width = CInt(sValue)

                             Case "height"
                                .Height = CInt(sValue)

                             Case "scale"
                                .Scale = CInt(sValue)

                             Case "opacity"
                                .Opacity = CInt(sValue)

                             Case "colorize"
                                .Colorize = IIf(sValue = "true", True, False)

                             Case "hue"
                                .Hue = CSng(sValue)

                             Case "saturation"
                                .Saturation = CSng(sValue)

                             Case "luminosity"
                                .Luminosity = CSng(sValue)

                             Case "rotate"
                                .Rotate = CInt(sValue)

                             Case "position_x"
                                .Position_X = CInt(sValue)

                             Case "position_y"
                                .Position_Y = CInt(sValue)

                             Case "image"
                                .Image = sValue & .Format

                            End Select

                        End With

                     Case "images"

                        With oImage

                            Select Case sKey
                             Case "format"
                                .Format = "." & sValue

                             Case "width"
                                .Width = CInt(sValue)

                             Case "height"
                                .Height = CInt(sValue)

                             Case "opacity"
                                .Opacity = CInt(sValue)

                             Case "colorize"
                                .Colorize = IIf(sValue = "true", True, False)

                             Case "hue"
                                .Hue = CSng(sValue)

                             Case "saturation"
                                .Saturation = CSng(sValue)

                             Case "luminosity"
                                .Luminosity = CSng(sValue)

                             Case "naming"
                                .Naming = sValue & .Format

                             Case "lowestvalue"
                                .LowestValue = CInt(sValue)

                             Case "highestvalue"
                                .HighestValue = CInt(sValue)

                             Case "increment"
                                .Increment = CInt(sValue)

                            End Select

                        End With

                     Case "global"

                        With oGlobal

                            Select Case sKey

                             Case "backcolor"
                                .Tooltip_Backcolor = sValue

                             Case "forecolor"
                                .tooltip_Forecolor = sValue

                            End Select

                        End With

                    End Select

                End If

            End If

        End If

    Wend

    Close #lFileNum

End Sub

Public Sub ShowCustomSizeWindow()

    frmSizeIcon.Show vbModal, Me
    
    UpdateImageSize 0, iIconSize

End Sub

Private Sub tooltip_activate(hwndTip As Long)

    SendMessage hwndTip, TTM_ACTIVATE, 1, ByVal 0&

End Sub

Private Function tooltip_addtool(hwndTip As Long, hwndControl As Long, sTipText As String) As Long

  Dim ti   As TOOLINFO
  Dim rc   As RECT

    If hwndTip <> 0 Then

        GetClientRect hwndControl, rc

        With ti
            .cbSize = Len(ti)
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            .hwnd = hwndControl
            .hinst = App.hInstance
            .uid = hwndControl
            .lpszText = sTipText
            .rc = rc
        End With

        SendMessage hwndTip, TTM_ADDTOOL, 0&, ti

    End If

End Function

Private Function tooltip_create(hwndForm As Long, bTipAsBalloon As Boolean, bTipAlways As Boolean) As Long

  Dim TTS_TIPSTYLE     As Long
  Dim TTS_TIPALWAYS    As Long

    TTS_TIPSTYLE = IIf(bTipAsBalloon = True, TTS_BALLOON, 0&)
    TTS_TIPALWAYS = IIf(bTipAlways = True, TTS_ALWAYSTIP, 0&)

    hwndTip = CreateWindowEx(WS_EX_TOPMOST, _
            TOOLTIPS_CLASSA, _
            vbNullString, _
            WS_POPUP Or TTS_TIPSTYLE Or TTS_TIPALWAYS, _
            CW_USEDEFAULT, CW_USEDEFAULT, _
            CW_USEDEFAULT, CW_USEDEFAULT, _
            hwndForm, _
            0&, _
            App.hInstance, _
            ByVal 0&)

    SetWindowPos hwndTip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE

    tooltip_create = hwndTip

End Function

Private Sub tooltip_deactivate(hwndTip As Long)

    SendMessage hwndTip, TTM_ACTIVATE, 0, ByVal 0&

End Sub

Private Function tooltip_destroy(hwndTip As Long) As Long

    If hwndTip <> 0 Then

        DestroyWindow hwndTip

        hwndTip = 0

    End If

End Function

Private Sub tooltip_setbackcolour(hwndTip As Long, dwColour As Long)

    SendMessage hwndTip, TTM_SETTIPBKCOLOR, dwColour, ByVal 0&

End Sub

Private Sub tooltip_setmaxwidth(hwndTip As Long, dwNewWidth As Long)

    SendMessage hwndTip, TTM_SETMAXTIPWIDTH, 0, ByVal dwNewWidth

End Sub

Private Sub tooltip_settextcolour(hwndTip As Long, dwColour As Long)

    SendMessage hwndTip, TTM_SETTIPTEXTCOLOR, dwColour, ByVal 0&

End Sub

Private Sub tooltip_titleadd(hwndTip As Long, sTipTitle As String, dwIconId As Long)

    SendMessage hwndTip, TTM_SETTITLE, dwIconId, ByVal sTipTitle

End Sub

Private Sub tooltip_titledelete(hwndTip As Long)

    SendMessage hwndTip, TTM_SETTITLE, 0&, ByVal vbNullString

End Sub

Private Sub tooltip_updatetext(hwndTip As Long, hwndControl As Long, sNewText As String)

  Dim ti As TOOLINFO

    With ti
        .cbSize = Len(ti)
        .hwnd = hwndControl
        .uid = hwndControl
        .lpszText = sNewText & vbNullString
    End With

    SendMessage hwndTip, TTM_UPDATETIPTEXT, 0&, ti

End Sub

Public Sub UpdateImageSize(Optional ByVal iIndex As Integer, Optional ByVal iCustomSize As Integer = 128)

    Select Case iIndex

     Case 0
        iImageSize = iCustomSize

     Case 1
        iImageSize = 256

     Case 2
        iImageSize = 128

     Case 3
        iImageSize = 96

     Case 4
        iImageSize = 72

     Case 5
        iImageSize = 48

     Case 6
        iImageSize = 32

    End Select

    iImageSizeIndex = iIndex

    iIconSize = iImageSize

    If (Me.Left + (iImageSize * Screen.TwipsPerPixelX)) > Screen.Width Then
        Me.Left = Screen.Width - (ImageSize * Screen.TwipsPerPixelX)
    End If

    If (Me.Top + (ImageSize * Screen.TwipsPerPixelY)) > Screen.Height Then
        Me.Top = Screen.Height - (ImageSize * Screen.TwipsPerPixelY)
    End If

    PaintImage

End Sub

