Attribute VB_Name = "modRegKeys"
Option Explicit

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Const RunKey             As String = "Software\Microsoft\Windows\CurrentVersion\Run"
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002

Const REG_SZ                    As Long = 1
Const REG_EXPAND_SZ             As Long = 2
Const REG_DWORD                 As Long = 4
Const REG_OPTION_NON_VOLATILE   As Long = 0

Const READ_CONTROL              As Long = &H20000
Const KEY_QUERY_VALUE           As Long = &H1
Const KEY_SET_VALUE             As Long = &H2
Const KEY_CREATE_SUB_KEY        As Long = &H4
Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Const KEY_NOTIFY                As Long = &H10
Const KEY_CREATE_LINK           As Long = &H20
Const KEY_READ                  As Long = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE                 As Long = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE               As Long = KEY_READ
Const KEY_ALL_ACCESS            As Long = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

Const HKEY_CLASSES_ROOT         As Long = &H80000000
Const HKEY_CURRENT_USER         As Long = &H80000001
Const HKEY_PERFORMANCE_DATA     As Long = &H80000004
Const HKEY_USERS                As Long = &H80000003

Const ERROR_ACCESS_DENIED       As Long = 8
Const ERROR_BADKEY              As Long = 2
Const ERROR_NONE                As Long = 0
Const ERROR_SUCCESS             As Long = 0

Private Type SECURITY_ATTRIBUTES
    nLength                     As Long
    lpSecurityDescriptor        As Long
    bInheritHandle              As Boolean
End Type


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String

  Dim hKey        As Long
  Dim hDepth      As Long
  Dim i           As Long
  Dim KeyValSize  As Long
  Dim lKeyValType As Long
  Dim rc          As Long
  Dim sKeyVal     As String
  Dim tmpVal      As String

    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)

    If (rc <> ERROR_SUCCESS) Then
        GoTo GetKeyError
    End If

    tmpVal = String$(1024, 0)
    KeyValSize = 1024

    rc = RegQueryValueEx(hKey, SubKeyRef, 0, lKeyValType, tmpVal, KeyValSize)

    If (rc <> ERROR_SUCCESS) Then
        GoTo GetKeyError
    End If

    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    Select Case lKeyValType
     Case REG_SZ, REG_EXPAND_SZ
        sKeyVal = tmpVal

     Case REG_DWORD

        For i = Len(tmpVal) To 1 Step -1
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next

        sKeyVal = Format$("&h" + sKeyVal)
    End Select

    GetKeyValue = sKeyVal
    rc = RegCloseKey(hKey)

    Exit Function

GetKeyError:
    GetKeyValue = vbNullString
    rc = RegCloseKey(hKey)

End Function

Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean

  Dim hDepth  As Long
  Dim hKey    As Long
  Dim lpAttr  As SECURITY_ATTRIBUTES
  Dim rc      As Long

    lpAttr.nLength = 50
    lpAttr.lpSecurityDescriptor = 0
    lpAttr.bInheritHandle = True

    rc = RegCreateKeyEx(KeyRoot, KeyName, 0, REG_SZ, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, hKey, hDepth)

    If (rc <> ERROR_SUCCESS) Then
        GoTo CreateKeyError
    End If

    If (SubKeyValue = "") Then
        SubKeyValue = " "
    End If

    rc = RegSetValueEx(hKey, SubKeyName, 0, REG_SZ, SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))

    If (rc <> ERROR_SUCCESS) Then
        GoTo CreateKeyError
    End If

    rc = RegCloseKey(hKey)

    UpdateKey = True

    Exit Function

CreateKeyError:

    UpdateKey = False
    rc = RegCloseKey(hKey)

End Function

