Attribute VB_Name = "Module1"
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Global Const conHwndTopmost = -1
    Global Const conSwpNoActivate = &H10
    Global Const conSwpShowWindow = &H40

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public ratio As Double
Public justloaded As Boolean
Public mhour As String, mmin As String, msec As String
Public chour As String, cmin As String, csec As String
Public shour As String, smin As String, ssec As String
Type POINTAPI
        mx As Long
        my As Long
End Type




Public Sub reloadmovie()
Form2.Width = Form2.MediaPlayer1.ImageSourceWidth * 15
Form2.Height = Form2.MediaPlayer1.ImageSourceHeight * 15
Form2.MediaPlayer1.Width = Form2.Width
Form2.MediaPlayer1.Height = Form2.Height
Form2.MediaPlayer1.ClickToPlay = False
ratio = Form2.MediaPlayer1.ImageSourceWidth / Form2.MediaPlayer1.ImageSourceHeight
End Sub
Public Sub sectomin(lsecs As Long)
'Den här snutten översätter sekunder till "tt:mm:ss"
smin = Format(Fix(lsecs / 60), "#0")
ssec = Format(lsecs Mod 60, "00")
shour = Format(Fix(smin / 60), "#0")
smin = Format(smin Mod 60, "00")
End Sub

