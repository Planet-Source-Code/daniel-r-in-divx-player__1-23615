VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5610
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   975
      Top             =   3195
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   2850
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Visible         =   0   'False
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   12648384
      DisplayForeColor=   12648384
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   -1  'True
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   -1  'True
      SendMouseMoveEvents=   -1  'True
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim positioned As Boolean, down As Boolean, counting As Integer, fullscreen As Boolean
Dim oldtop As Integer, oldleft As Integer, oldwidth As Integer, oldheight As Integer, testx As Long, testy As Long, p As Integer


Private Sub Form_Activate()
SetWindowPos Form2.hwnd, conHwndTopmost, 100, 100, 400, 141, conSwpNoActivate Or conSwpShowWindow
End Sub

Private Sub Form_DblClick()
MediaPlayer1_DblClick 1, 0, 100, 100
End Sub

Private Sub Form_Initialize()
On Error Resume Next
ratio = Form2.MediaPlayer1.ImageSourceWidth / Form2.MediaPlayer1.ImageSourceHeight
Me.Width = Form2.MediaPlayer1.ImageSourceWidth * 15
Me.Height = Me.Width / ratio
MediaPlayer1.Width = Me.Width
MediaPlayer1.Height = Me.Height
testar = MediaPlayer1.FileName
'MaxMin.MinWidth = 120
'MaxMin.MinHeight = 120 / ratio
End Sub

Private Sub Form_Load()
'SetWinOnTop = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub



Private Sub Form_Resize()
On Error Resume Next
If fullscreen = False Then
counting = counting + 1
ratio = Form2.MediaPlayer1.ImageSourceWidth / Form2.MediaPlayer1.ImageSourceHeight
If counting = 4 Then Me.Width = Form2.MediaPlayer1.ImageSourceWidth * 15
Me.Height = Me.Width / ratio
If Me.Width / 15 < 150 Then Me.Width = 150 * 15
MediaPlayer1.Width = Me.Width
MediaPlayer1.Top = 0
MediaPlayer1.Height = Me.Height

Else
MediaPlayer1.Left = 0
MediaPlayer1.Width = Me.Width
MediaPlayer1.Height = Me.Height
MediaPlayer1.Top = Screen.Height / 2 - (MediaPlayer1.Height / 2)

End If
End Sub

Private Sub MediaPlayer1_DblClick(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If fullscreen = False Then
fullscreen = True
ShowCursor 0
oldtop = Me.Top
oldleft = Me.Left
oldwidth = Me.Width
oldheight = Me.Height
Me.Top = 0
Me.Left = 0
Me.Width = Screen.Width
Me.Height = Screen.Height
Else
fullscreen = False
ShowCursor 1
MediaPlayer1.Top = 0
MediaPlayer1.Left = 0
Me.Top = oldtop
Me.Left = oldleft
Me.Width = oldwidth
Me.Height = oldheight
End If
End Sub

Private Sub MediaPlayer1_MouseDown(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If positioned = False Then
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
Else

ReleaseCapture
SendMessage Me.hwnd, &H112, 61448, 0
SetCursor 2599
End If
End Sub

Private Sub MediaPlayer1_MouseMove(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
If X > MediaPlayer1.Width - 150 And Y > MediaPlayer1.Height - 150 Then
    positioned = True
    SetCursor 2599
Else
    positioned = False
    'SetCursor 2558
End If
    
End Sub

Private Sub MediaPlayer1_MouseUp(Button As Integer, ShiftState As Integer, X As Single, Y As Single)
positioned = False
SetCursor 2558
End Sub

Private Sub Timer1_Timer()
'Form1.trackbar.Value = Int(MediaPlayer1.CurrentPosition)
tempval = MediaPlayer1.CurrentPosition / MediaPlayer1.SelectionEnd
tempval = tempval * Form1.Picture3.ScaleWidth
tempval = tempval
Call BitBlt(Form1.Picture4.hdc, 0, 0, tempval, Form1.Picture3.ScaleHeight, Form1.Picture5.hdc, 0, 0, vbSrcCopy)
Call BitBlt(Form1.Picture3.hdc, 0, 0, Form1.Picture1.ScaleWidth, Form1.Picture1.ScaleHeight, Form1.Picture4.hdc, 0, 0, vbSrcCopy)
Form1.Picture3.Refresh
sectomin (Form2.MediaPlayer1.CurrentPosition)
Form1.Label2.Caption = shour & ":" & smin & ":" & ssec & " - " & mhour & ":" & mmin & ":" & msec
End Sub


