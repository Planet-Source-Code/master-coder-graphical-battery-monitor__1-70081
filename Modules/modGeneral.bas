Attribute VB_Name = "modGeneral"
Option Explicit

Public Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const RegKey             As String = "Software\TheWWWpc\BatteryMonitor"

Public Const NIF_ICON           As Long = &H2
Public Const NIF_MESSAGE        As Long = &H1
Public Const NIF_TIP            As Long = &H4

Public Const NIM_ADD            As Long = &H0
Public Const NIM_DELETE         As Long = &H2
Public Const NIM_MODIFY         As Long = &H1

Public Const WM_LBUTTONDBLCLK   As Long = &H203
Public Const WM_LBUTTONDOWN     As Long = &H201
Public Const WM_LBUTTONUP       As Long = &H202
Public Const WM_MBUTTONDBLCLK   As Long = &H209
Public Const WM_MBUTTONDOWN     As Long = &H207
Public Const WM_MBUTTONUP       As Long = &H208
Public Const WM_MOUSEMOVE       As Long = &H200
Public Const WM_RBUTTONDBLCLK   As Long = &H206
Public Const WM_RBUTTONDOWN     As Long = &H204
Public Const WM_RBUTTONUP       As Long = &H205

Public Type NOTIFYICONDATA
    cbSize                      As Long
    hwnd                        As Long
    uid                         As Long
    uFlags                      As Long
    uCallbackMessage            As Long
    hIcon                       As Long
    szTip                       As String * 255
End Type

Public Type SYSTEM_POWER_STATUS
    ACLineStatus                As Byte
    BatteryFlag                 As Byte
    BatteryLifePercent          As Byte
    Reserved1                   As Byte
    BatteryLifeTime             As Long
    BatteryFullLifeTime         As Long
End Type

Public Animate                  As Integer
Public bAttachedToAC            As Boolean
Public bDesktopBatteryTopMost   As Boolean
Public bShowDesktopIcon         As Boolean
Public bTransparency            As Byte
Public iPowerLevel              As Integer
Public iIconSize                As Integer
Public nID                      As NOTIFYICONDATA
Public sDesktopBatteryToolTip   As String
Public SysPower                 As SYSTEM_POWER_STATUS
Public Tray                     As Boolean

