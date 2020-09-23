Attribute VB_Name = "ModMain"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Integer
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public FrameCount As Integer
Public FinishedTest As Boolean
Public StartTime As Double

Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
    End Type


Public Type SIZE
    cx As Long
    cy As Long
    End Type


Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
    End Type
    Public Const GWL_STYLE = (-16)
    Public Const GWL_EXSTYLE = (-20)
    Public Const WS_EX_LAYERED = &H80000
    Public Const ULW_COLORKEY = &H1
    Public Const ULW_ALPHA = &H2
    Public Const ULW_OPAQUE = &H4
    Public Const AC_SRC_OVER = &H0
    Public Const AC_SRC_ALPHA = &H1
    Public Const AC_SRC_NO_PREMULT_ALPHA = &H1
    Public Const AC_SRC_NO_ALPHA = &H2
    Public Const AC_DST_NO_PREMULT_ALPHA = &H10
    Public Const AC_DST_NO_ALPHA = &H20
    Public Const LWA_COLORKEY = &H1
    Public Const LWA_ALPHA = &H2

Public Sub Mache_Transparent(hWnd As Long, Rate As Byte)
    Dim WinInfo As Long
    WinInfo = GetWindowLong(hWnd, GWL_EXSTYLE)
    WinInfo = WinInfo Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, WinInfo
    SetLayeredWindowAttributes hWnd, 0, Rate, LWA_ALPHA
End Sub

