Attribute VB_Name = "modAPI"
Option Explicit

'
' GDI declarations
'
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

' The function name is GetObject, but VB has a simlar named
' internal function...
Declare Function GDIGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Type RECT
    left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Const IMAGE_BITMAP = 0
Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_LOADFROMFILE = &H10
Public Const SRCCOPY = &HCC0020
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'
' Misc. declarations
'
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const API_NULL_HANDLE = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SPI_GETWORKAREA = 48


'
' INI file handling. Usual stuff.
'
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
    
    
'
' API To modify Form Titlebar
'
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SWP_NOREDRAW = &H8
Public Const SWP_SHOWWINDOW = &H40
Public Const WS_CAPTION = &HC00000
Public Const WS_BORDER = &H800000
Public Const WS_SYSMENU = &H80000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_DLGFRAME = &H400000
Public Const WS_THICKFRAME = &H40000


' Read / Write INI files
Public Function INIRead(Section As String, KeyName As String, FileName As String) As String
Dim Str As String
    
    Str = String(255, Chr(0))
'    INIRead = left(Str, GetPrivateProfileString(Section, ByVal KeyName, "NO_SUCH_KEY", Str, Len(Str), FileName))
    INIRead = left(Str, GetPrivateProfileString(Section, ByVal KeyName, "", Str, Len(Str), FileName))

End Function

Public Function INIWrite(Section As String, KeyName As String, KeyValue As String, FileName As String) As Boolean
Dim Ret As Long
    
    Ret = WritePrivateProfileString(Section, KeyName, KeyValue, FileName)
    If Ret = 0 Then
        INIWrite = True
    Else
        INIWrite = False
    End If
    
End Function
