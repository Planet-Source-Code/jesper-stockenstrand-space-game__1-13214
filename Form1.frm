VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13020
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10695
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2760
      Top             =   5760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   5160
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   1800
      ScaleHeight     =   1635
      ScaleWidth      =   13035
      TabIndex        =   0
      Top             =   9240
      Width           =   13095
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Press 8 to increase speed and 2 to decrease speed. Higher speed gives more points. Press SPACE to quit."
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   975
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "POINTS: 0"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   1680
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PRESS S TO START"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   2280
      TabIndex        =   7
      Top             =   1920
      Width           =   6435
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   0
      Left            =   8400
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1215
      Visible         =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer3 
      Height          =   375
      Left            =   12000
      TabIndex        =   6
      Top             =   10200
      Width           =   735
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
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "Laser-Public_D-33.wav"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
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
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer2 
      Height          =   375
      Left            =   12000
      TabIndex        =   5
      Top             =   9840
      Width           =   735
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
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "Big_bang-Public_D-298.wav"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
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
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   12000
      TabIndex        =   4
      Top             =   9360
      Width           =   735
      Visible         =   0   'False
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   "SpaceLoop.aiff.wav"
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
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
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Xreal(199), Yreal(199), Xrec(199), Yrec(199), Xspar(199), Yspar(199), Zspar(199), Z(199) As Double
Dim color As Integer
Dim moveX, moveY As Integer
Private Declare Function ShowCursor Lib "User32" (ByVal Show As Integer) As Integer
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uflags As Long) As Long
'Private Declare Function sndPlaySound2 Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName2 As String, ByVal uflags2 As Long) As Long

Dim NowX, NowY As Integer
Dim speed As Double
Dim i As Integer
Dim shoot(20) As Integer
Dim Xshoot(20), Yshoot(20), Zshoot(20), XshootOld(20), YshootOld(20), ZshootOld(20) As Integer
Dim shootcolor
Dim points As Integer
Dim r As Long
Dim l As Long
Dim xlaser As Integer
Dim ylaser As Integer
Dim laser As Integer
Dim bildnr As Integer
Dim bildcount As Integer
Dim stage As Integer
Dim comets As Integer







Private Sub Form_KeyPress(KeyAscii As Integer)


If KeyAscii = 56 And speed < 0.3 Then
speed = speed + 0.005
ElseIf KeyAscii = 50 And speed > 0.02 Then
speed = speed - 0.005
ElseIf KeyAscii = 32 Then Unload Me
End If



Label1.Caption = "SPEED: " & (speed * 1000)


End Sub

Private Sub Form_Load()
For i = 0 To 199

Xreal(i) = (Rnd * 30 * Form1.Width) - 15 * Form1.Width
Yreal(i) = (Rnd * 30 * Form1.Height) - 15 * Form1.Height
Z(i) = 10
Next i
moveX = 0
moveY = 0
speed = 0.1
Label1.Caption = "SPEED: " & speed * 1000
ShowCursor 0

For i = 0 To 5
Zshoot(i) = 10
shoot(i) = 0
Next i
'Form1.Picture = LoadPicture("f:\private\vb-projects\space game\spade.bmp")
'Xshoot = Rnd * Form1.Width
'Yshoot = Rnd * Form1.Height
shootcolor = RGB(0, 255, 0)
points = 0
Dim nisse As Long
nisse = 0
'r = sndPlaySound(ByVal "f:\private\vb-projects\space game\SpaceLoop.aiff.wav", &H1 + &H8)
MediaPlayer1.AutoStart = True
bildnr = 1
bildcount = 1
stage = 1
comets = 0
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
laser = 1
Form1.ForeColor = RGB(0, 255, 255)
Form1.Line (Form1.Width / 2, Form1.Height - 1500)-(X, Y)
xlaser = X
ylaser = Y

'If X > XshootOld - 200 And X < XshootOld + 200 And Y > YshootOld - 200 And Y < YshootOld + 200 Then
'MediaPlayer2.Play

'r = sndPlaySound(ByVal "f:\private\vb-projects\space game\Big_bang-Public_D-298.wav", &H1)
'r = sndPlaySound(ByVal "f:\private\vb-projects\space game\SpaceLoop.aiff.wav", &H1 + &H8)
'points = points + (100 * speed)
'Label2.Caption = "POINTS: " & points
'shoot = 0
'Form1.FillColor = RGB(0, 0, 0)
'Form1.FillStyle = 0
'Form1.Circle (XshootOld, YshootOld), 140 - (8 * ZshootOld), RGB(0, 0, 0)
'XshootOld = 0
'YshootOld = 0
'Xshoot = Rnd * Form1.Width
'Yshoot = Rnd * Form1.Height
'Zshoot = 10
'Else
MediaPlayer3.Play

'r = sndPlaySound(ByVal "f:\private\vb-projects\space game\Laser-Public_D-33.wav", &H1)
points = points - 1
Label2.Caption = "POINTS: " & points

'End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If NowX < X Then
'moveX = moveX + 10
'End If
'If NowX > X Then
'moveX = moveX - 10
'End If

'If NowY < Y Then
'moveY = moveY + 10
'End If
'If NowY > Y Then
'moveY = moveY - 10
'End If
Form1.ForeColor = RGB(0, 0, 0)
Form1.FillStyle = 1
Form1.Line (0, NowY)-(Form1.Width, NowY)
Form1.Line (NowX, 0)-(NowX, Form1.Height)

Form1.Circle (NowX, NowY), 200, RGB(0, 0, 0)

moveX = (X - (Form1.Width / 2)) / 3
moveY = (Y - (Form1.Height / 2)) / 3



NowX = X
NowY = Y
Form1.ForeColor = RGB(0, 255, 0)
Form1.Line (0, Y)-(Form1.Width, Y)
Form1.Line (X, 0)-(X, Form1.Height)
'If X > XshootOld - 200 And X < XshootOld + 200 And Y > YshootOld - 200 And Y < YshootOld + 200 Then
shootcolor = RGB(0, 255, 0)
'Else
'shootcolor = RGB(0, 255, 0)
'End If

Form1.Circle (X, Y), 200, shootcolor

End Sub

Private Sub Form_Resize()
Picture1.Top = Form1.Height - 1500
Picture1.Left = 0
Picture1.Width = Form1.Width
Picture1.Height = 1500
Label4.Left = Form1.Width / 2 - 1000


End Sub

Private Sub Form_Unload(Cancel As Integer)
ShowCursor 1
End

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
i = Index
laser = 1
Form1.ForeColor = RGB(0, 255, 255)
Form1.Line (Form1.Width / 2, Form1.Height - 1500)-(X + Image1(i).Left, Y + Image1(i).Top)
xlaser = X + Image1(i).Left
ylaser = Y + Image1(i).Top

'If X > XshootOld - 200 And X < XshootOld + 200 And Y > YshootOld - 200 And Y < YshootOld + 200 Then
MediaPlayer2.Play

'r = sndPlaySound(ByVal "f:\private\vb-projects\space game\Big_bang-Public_D-298.wav", &H1)
'r = sndPlaySound(ByVal "f:\private\vb-projects\space game\SpaceLoop.aiff.wav", &H1 + &H8)
points = points + (100 * speed)
Label2.Caption = "POINTS: " & points
shoot(i) = 0
Form1.FillColor = RGB(0, 0, 0)
Form1.FillStyle = 0
Form1.Circle (XshootOld(i), YshootOld(i)), 140 - (8 * ZshootOld(i)), RGB(0, 0, 0)
XshootOld(i) = 0
YshootOld(i) = 0
Xshoot(i) = Rnd * Form1.Width
Yshoot(i) = Rnd * Form1.Height
Zshoot(i) = 10
Image1(i).Visible = False
Image1(i).Top = 0
Image1(i).Left = 0
Unload Image1(i)
'Else
'MediaPlayer3.Play

'r = sndPlaySound(ByVal "f:\private\vb-projects\space game\Laser-Public_D-33.wav", &H1)
'points = points - 1
'Label2.Caption = "POINTS: " & points

'End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim i As Integer
i = Index

Form1.ForeColor = RGB(0, 0, 0)
Form1.FillStyle = 1
Form1.Line (0, NowY)-(Form1.Width, NowY)
Form1.Line (NowX, 0)-(NowX, Form1.Height)

Form1.Circle (NowX, NowY), 200, RGB(0, 0, 0)

moveX = ((X + Image1(i).Left) - (Form1.Width / 2)) / 3
moveY = ((Y + Image1(i).Top) - (Form1.Height / 2)) / 3



NowX = X + Image1(i).Left
NowY = Y + Image1(i).Top
Form1.ForeColor = RGB(0, 255, 0)
Form1.Line (0, NowY)-(Form1.Width, NowY)
Form1.Line (NowX, 0)-(NowX, Form1.Height)
'If X + Image1.Left > XshootOld - 200 And X + Image1.Left < XshootOld + 200 And Y + Image1.Top > YshootOld - 200 And Y + Image1.Top < YshootOld + 200 Then
'shootcolor = RGB(255, 0, 0)
'Else
shootcolor = RGB(255, 0, 0)
'End If

Form1.Circle (X + Image1(i).Left, Y + Image1(i).Top), 200, shootcolor

End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


Form1.ForeColor = RGB(0, 0, 0)
Form1.FillStyle = 1
Form1.Line (0, NowY)-(Form1.Width, NowY)
Form1.Line (NowX, 0)-(NowX, Form1.Height)

Form1.Circle (NowX, NowY), 200, RGB(0, 0, 0)

moveX = ((X + Label4.Left) - (Form1.Width / 2)) / 3
moveY = ((Y + Label4.Top) - (Form1.Height / 2)) / 3



NowX = X + Label4.Left
NowY = Y + Label4.Top
Form1.ForeColor = RGB(0, 255, 0)
Form1.Line (0, NowY)-(Form1.Width, NowY)
Form1.Line (NowX, 0)-(NowX, Form1.Height)
'If X + Image1.Left > XshootOld - 200 And X + Image1.Left < XshootOld + 200 And Y + Image1.Top > YshootOld - 200 And Y + Image1.Top < YshootOld + 200 Then
'shootcolor = RGB(255, 0, 0)
'Else
shootcolor = RGB(255, 0, 0)
'End If

Form1.Circle (X + Label4.Left, Y + Label4.Top), 200, shootcolor

Label4.Width = Label4.Width * 1


End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
If KeyAscii = 56 And speed < 0.3 Then
speed = speed + 0.005
ElseIf KeyAscii = 50 And speed > 0.02 Then
speed = speed - 0.005
ElseIf KeyAscii = 32 Then
Unload Form1
ElseIf KeyAscii = 115 Or KeyAscii = 83 Then
Label4.Caption = "STAGE " & stage
Label5.Caption = "STAGE " & stage
Timer3.Enabled = True


End If



Label1.Caption = "SPEED: " & (speed * 1000)


End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form1.ForeColor = RGB(0, 0, 0)
Form1.FillStyle = 1
Form1.Line (0, Form1.Height - 1500)-(Form1.Width, Form1.Height - 1500)
Form1.Line (NowX, 0)-(NowX, Form1.Height - 1500)
Form1.Circle (NowX, Form1.Height - 1500), 200, RGB(0, 0, 0)

moveX = (X - (Form1.Width / 2)) / 3
'moveY = (Y - (Form1.Height / 2)) / 10



NowX = X
NowY = Form1.Height - 1500
Form1.ForeColor = RGB(0, 255, 0)
Form1.Line (0, Form1.Height - 1500)-(Form1.Width, Form1.Height - 1500)
Form1.Line (X, 0)-(X, Form1.Height - 1500)
Form1.Circle (X, Form1.Height - 1500), 200, RGB(0, 255, 0)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

For i = 0 To 199



Form1.FillColor = RGB(0, 0, 0)
Form1.FillStyle = 0
Form1.Circle (Xspar(i), Yspar(i)), 80 - (8 * Zspar(i)), RGB(0, 0, 0)


Next i



'Cls
For i = 0 To 199
Xreal(i) = Xreal(i) - moveX
Yreal(i) = Yreal(i) + moveY
Next i

For i = 0 To 199
Xrec(i) = (Form1.Width / 2) + ((Xreal(i) - (Form1.Width / 2)) / Z(i))
Yrec(i) = (Form1.Height / 2) - ((Yreal(i) - (Form1.Height / 2)) / Z(i))



If Xrec(i) > Form1.Width + 3000 Or Xrec(i) < -3000 Then
Xreal(i) = (Rnd * 30 * Form1.Width) - 15 * Form1.Width
Yreal(i) = (Rnd * 30 * Form1.Height) - 15 * Form1.Height
Z(i) = 10
ElseIf Yrec(i) > Form1.Height + 3000 Or Yrec(i) < -3000 Then
Xreal(i) = (Rnd * 30 * Form1.Width) - 15 * Form1.Width
Yreal(i) = (Rnd * 30 * Form1.Height) - 15 * Form1.Height
Z(i) = 10
ElseIf Z(i) < 0 Then
Xreal(i) = (Rnd * 30 * Form1.Width) - 15 * Form1.Width
Yreal(i) = (Rnd * 30 * Form1.Height) - 15 * Form1.Height

Z(i) = 10
Else
Xspar(i) = Xrec(i)
Yspar(i) = Yrec(i)
Zspar(i) = Z(i)

color = 255 - (Z(i) * 10)
Form1.FillColor = RGB(color, color, color)
Form1.FillStyle = 0


Form1.Circle (Xrec(i), Yrec(i)), 80 - (8 * Z(i)), RGB(color, color, color)
Z(i) = Z(i) - speed
End If


Next i
For i = 0 To 5
If shoot(i) = 1 Then

'Form1.FillColor = RGB(0, 0, 0)
'Form1.FillStyle = 0
'Form1.Circle (XshootOld, YshootOld), 140 - (8 * ZshootOld), RGB(0, 0, 0)
Xshoot(i) = Xshoot(i) - moveX
Yshoot(i) = Yshoot(i) + moveY

XshootOld(i) = (Form1.Width / 2) + ((Xshoot(i) - (Form1.Width / 2)) / Zshoot(i))
YshootOld(i) = (Form1.Height / 2) - ((Yshoot(i) - (Form1.Height / 2)) / Zshoot(i))


If XshootOld(i) > Form1.Width Or XshootOld(i) < 1 Then
shoot(i) = 0
Xshoot(i) = Rnd * Form1.Width
Yshoot(i) = Rnd * Form1.Height
Zshoot(i) = 10
points = points - 5
Label2.Caption = "POINTS: " & points
Image1(i).Visible = False
Unload Image1(i)
ElseIf YshootOld(i) > Form1.Height Or YshootOld(i) < 1 Then
shoot(i) = 0
Xshoot(i) = Rnd * Form1.Width
Yshoot(i) = Rnd * Form1.Height
Zshoot(i) = 10
points = points - 5
Label2.Caption = "POINTS: " & points
Image1(i).Visible = False
Unload Image1(i)
Else
'Form1.FillColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
'Form1.FillStyle = 0
'Form1.Circle (XshootOld, YshootOld), 140 - (8 * Zshoot), RGB(0, 0, 0)
Image1(i).Picture = LoadPicture("sten" & bildnr & ".bmp")
Image1(i).Visible = True

Image1(i).Height = 1000 - (100 * Zshoot(i))
Image1(i).Width = 1000 - (100 * Zshoot(i))

Image1(i).Left = XshootOld(i)
Image1(i).Top = YshootOld(i)
bildcount = bildcount + 1
If bildcount = 5 Then
bildnr = bildnr + 1
bildcount = 1
End If
If bildnr = 21 Then bildnr = 1

End If



'XshootOld = Xshoot
'YshootOld = Yshoot
ZshootOld(i) = Zshoot(i)
Zshoot(i) = Zshoot(i) - speed

End If

Next i

Form1.ForeColor = RGB(0, 255, 0)
Form1.FillStyle = 1
Form1.Line (0, NowY)-(Form1.Width, NowY)
Form1.Line (NowX, 0)-(NowX, Form1.Height)
Form1.Circle (NowX, NowY), 200, shootcolor

If laser = 10 Then
Form1.ForeColor = RGB(0, 0, 0)
Form1.Line (Form1.Width / 2, Form1.Height - 1500)-(xlaser, ylaser)
laser = 0
ElseIf laser > 0 Then

laser = laser + 1

End If
End Sub

Private Sub Timer2_Timer()
'Dim Image1(5) As Image
'Load Image1(5)
If comets = 30 Then
    stage = stage + 1
    comets = 0
    Timer1.Enabled = False
    Timer2.Enabled = False
    Cls
    Label4.Enabled = True
    Label4.Visible = True
    
    If stage < 21 Then
    Label4.Caption = "STAGE " & stage
    Label5.Caption = "STAGE " & stage
    Timer3.Enabled = True
    Else
    Label4.Caption = "MISSION COMPLETE"
    End If
Else

Dim i As Integer



For i = 1 To stage
If shoot(i) = 0 Then
Load Image1(i)
'set image(i) =





Image1(i).Picture = LoadPicture("sten" & bildnr & ".bmp")
Xshoot(i) = Rnd * Form1.Width
Yshoot(i) = Rnd * Form1.Height
shoot(i) = 1
End If
Next i


End If

comets = comets + 1
End Sub

Private Sub Timer3_Timer()
Timer1.Enabled = True
Timer2.Enabled = True
Label4.Enabled = False
Label4.Visible = False
Timer3.Enabled = False

End Sub
