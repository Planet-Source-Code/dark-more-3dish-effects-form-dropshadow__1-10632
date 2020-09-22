Attribute VB_Name = "Blending"
Option Explicit

Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Global Inner As RECT 'Client Area Of Window
Global Outer As RECT 'Total Window Area

'For Message Hook
Global Const WM_MOVE = &H3

'For Taking A Picture Of The Screen w/o The Form In The Way
Global Const SWP_NOMOVE = &H2
Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_SHOWWINDOW = &H40

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'From BoS
Declare Function AlphaBlending Lib "Alphablending.dll" (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long
'Used For Determing The Size Of The Window Borders And Captions
'(Remove if you're only using borderless forms.)
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Sub Blend(Destination As Long, Source As Object, Amount As Integer, X, Y, X2, Y2)
AlphaBlending Destination, X, Y, X2, Y2, Source.hdc, X, Y, X2, Y2, Amount
End Sub

