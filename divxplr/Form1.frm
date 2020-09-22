VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Divx;-) player"
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   145
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   240
      Left            =   105
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   11
      Top             =   2565
      Width           =   4035
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   240
      Left            =   90
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   10
      Top             =   2265
      Width           =   4035
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Enabled         =   0   'False
      Height          =   240
      Left            =   90
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   265
      TabIndex        =   9
      Top             =   1755
      Width           =   4035
   End
   Begin VB.PictureBox volorig 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   3885
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   7
      Top             =   1365
      Width           =   210
      Visible         =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   855
      ScaleWidth      =   4005
      TabIndex        =   3
      Top             =   360
      Width           =   4005
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   3630
         Picture         =   "Form1.frx":07A6
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   8
         Top             =   600
         Width           =   240
      End
      Begin VB.PictureBox voldisp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   650
         Left            =   3720
         ScaleHeight     =   43
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   14
         TabIndex        =   6
         Top             =   105
         Width           =   210
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "0:00:00 - 0:00:00"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2070
         TabIndex        =   5
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No file loaded"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3840
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   ":"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "9"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   375
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      X1              =   7
      X2              =   274
      Y1              =   23
      Y2              =   23
   End
   Begin VB.Image Image3 
      Height          =   150
      Left            =   3975
      Picture         =   "Form1.frx":0A88
      ToolTipText     =   "Close"
      Top             =   105
      Width           =   165
   End
   Begin VB.Image Image2 
      Height          =   150
      Left            =   3780
      Picture         =   "Form1.frx":0C32
      ToolTipText     =   "Minimize"
      Top             =   105
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Top             =   30
      Width           =   4200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   4
      X2              =   276
      Y1              =   9
      Y2              =   9
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000000&
      X1              =   5
      X2              =   277
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      X1              =   4
      X2              =   276
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      X1              =   5
      X2              =   277
      Y1              =   13
      Y2              =   13
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      X1              =   4
      X2              =   276
      Y1              =   12
      Y2              =   12
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      X1              =   5
      X2              =   277
      Y1              =   10
      Y2              =   10
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      X1              =   5
      X2              =   277
      Y1              =   7
      Y2              =   7
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   4
      X2              =   276
      Y1              =   6
      Y2              =   6
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   4
      X2              =   276
      Y1              =   3
      Y2              =   3
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   5
      X2              =   277
      Y1              =   4
      Y2              =   4
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   144
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   280
      X2              =   280
      Y1              =   0
      Y2              =   160
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   296
      Y1              =   144
      Y2              =   144
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public playing As Boolean, check As String, t As Long, volchange As Boolean, volume As Integer, currentpos As Long
Public down As Boolean


Private Sub Command4_Click()
Me.WindowState = 1
Unload Form2
End
End Sub

Private Sub Image4_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
down = True
currentpos = (X / Picture3.ScaleWidth) * 100
Picture4.Cls
Form2.MediaPlayer1.CurrentPosition = Form2.MediaPlayer1.SelectionEnd * (currentpos / 100)
Call BitBlt(Picture4.hdc, 0, 0, X, Picture3.ScaleHeight, Picture5.hdc, 0, 0, vbSrcCopy)
Call BitBlt(Picture3.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture4.hdc, 0, 0, vbSrcCopy)
Picture3.Refresh
End Sub

Private Sub Command1_Click()

If playing = False Then
playing = True
Form2.MediaPlayer1.Play
Label1.Caption = "Playing - " & check
Command1.Caption = ";"
Else
playing = False
Form2.MediaPlayer1.Pause
Label1.Caption = "Paused - " & check
Command1.Caption = "4"
End If
End Sub

Private Sub Form_Load()
Load Form2
Form2.Hide
Form1.OLEDropMode = 1
Dim g As Double
On Error Resume Next
g = 255
For t = 0 To Picture5.ScaleWidth
    g = g - 0.8
    Picture5.ForeColor = RGB(0, g - 50, g - 50)
    Picture5.Line (t, 0)-(t, Picture1.ScaleHeight)
Next
Call BitBlt(voldisp.hdc, 0, 0, 15, 43, volorig.hdc, 0, 0, vbSrcCopy)
End Sub



Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i%, tempname As String, ext As String
If Data.GetFormat(vbCFFiles) Then
        On Error Resume Next
        Form2.MediaPlayer1.Stop
        counting = 0
        Unload Form2
        Command1.Caption = "4"
        playing = False
        ext = Right(Data.Files(1), 4)
        If ext = ".peg" Or ext = "mpeg" Or ext = ".m1v" Or ext = ".mp2" _
        Or ext = ".mpa" Or ext = ".avi" Or ext = ".asf" Or ext = ".wmw" _
        Or ext = ".mpg" Then
        
        Form2.MediaPlayer1.FileName = Data.Files(1)
        Form2.MediaPlayer1.DisplaySize = mpFitToSize
        sectomin (Form2.MediaPlayer1.Duration)
        mhour = shour: mmin = smin: msec = ssec
        Label2.Caption = "00:00:00 - " & mhour & ":" & mmin & ":" & msec
        tempname = Form2.MediaPlayer1.FileName
        For t = 1 To Len(tempname)
            check = Right(tempname, t)
            If Mid(check, 1, 1) = "\" Then
            check = Right(tempname, t - 1)
            Exit For
            End If
        Label1.Caption = "Loaded - " & check
        Form2.Timer1.Enabled = True
        Next
        
        reloadmovie
        Command1.Enabled = True
        Picture3.Enabled = True
        ratio = Form2.MediaPlayer1.ImageSourceWidth / Form2.MediaPlayer1.ImageSourceHeight
        Form2.Width = Form2.MediaPlayer1.ImageSourceWidth * 15
        Form2.Height = Form2.Width / ratio
        Form2.MediaPlayer1.Width = Form2.Width
        Form2.MediaPlayer1.Height = Form2.Height
        If Form2.MediaPlayer1.CanScan = True Then
        Command2.Enabled = True
        Command3.Enabled = True
        End If
        Load Form2
        Form2.Show
        End If
        
End If

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub


Private Sub Image2_Click()
Me.WindowState = 1
End Sub

Private Sub Image3_Click()
Unload Form2
End
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub



Private Sub trackbar_Scroll()

End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim temptime As String
'sectomin (trackbar.Value)
'temptime = shour & ":" & smin & ":" & ssec
'trackbar.Text = temptime
'Form2.MediaPlayer1.CurrentPosition = trackbar.Value
down = True

regmark.Caption = currentpos
Form2.MediaPlayer1.CurrentPosition = Form2.MediaPlayer1.SelectionEnd * (currentpos / 100)
End Sub





Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If down = True Then
If X < 0 Then X = 0
If X > Picture3.ScaleWidth Then X = Picture3.ScaleWidth
currentpos = (X / Picture3.ScaleWidth) * 100
Form2.MediaPlayer1.CurrentPosition = Form2.MediaPlayer1.SelectionEnd * (currentpos / 100)
Picture4.Cls
Call BitBlt(Picture4.hdc, 0, 0, X, Picture3.ScaleHeight, Picture5.hdc, 0, 0, vbSrcCopy)
Call BitBlt(Picture3.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture4.hdc, 0, 0, vbSrcCopy)
Picture3.Refresh
End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
down = False
End Sub

Private Sub voldisp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
volchange = True
voldisp.Cls
Call BitBlt(voldisp.hdc, 0, Y, 15, 43, volorig.hdc, 0, Y, vbSrcCopy)
End Sub

Private Sub voldisp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If volchange = True Then
If Y < 0 Then Y = 0
If Y > 43 Then Y = 43
voldisp.Cls
Call BitBlt(voldisp.hdc, 0, Y, 15, 43, volorig.hdc, 0, Y, vbSrcCopy)
voltemp = Int(Y * 2.33)
vol = -85 * voltemp
Form2.MediaPlayer1.volume = vol
End If
End Sub

Private Sub voldisp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
volchange = False
End Sub

