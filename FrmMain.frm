VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00E17E35&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Worm"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5820
   ForeColor       =   &H00000000&
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Scr 
      BackColor       =   &H00E0E0E0&
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   3900
      Left            =   0
      Picture         =   "FrmMain.frx":08CA
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   0
      Top             =   0
      Width           =   5835
      Begin VB.Timer TimHideDamagedCar 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1860
         Top             =   0
      End
      Begin VB.Timer TimMoveMan 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   2880
         Top             =   0
      End
      Begin VB.Timer TimShowMan 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   2460
         Top             =   0
      End
      Begin VB.Timer TimMoveCar 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1440
         Top             =   0
      End
      Begin VB.Timer TimShowCar 
         Enabled         =   0   'False
         Interval        =   6000
         Left            =   1020
         Top             =   0
      End
      Begin VB.Timer TimMoveWorm 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.Image Apple 
         Height          =   240
         Left            =   5280
         Picture         =   "FrmMain.frx":16020C
         Stretch         =   -1  'True
         Top             =   2880
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image DamagedCar 
         Height          =   480
         Index           =   1
         Left            =   5280
         Picture         =   "FrmMain.frx":160AD6
         Top             =   1620
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image DamagedCar 
         Height          =   480
         Index           =   0
         Left            =   4740
         Picture         =   "FrmMain.frx":1613A0
         Top             =   1620
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Wall 
         Height          =   480
         Index           =   0
         Left            =   5280
         Picture         =   "FrmMain.frx":161C6A
         Top             =   2340
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgCar 
         Height          =   480
         Index           =   1
         Left            =   5280
         Picture         =   "FrmMain.frx":161F74
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgCar 
         Height          =   480
         Index           =   0
         Left            =   4740
         Picture         =   "FrmMain.frx":162286
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Car 
         Height          =   480
         Index           =   0
         Left            =   4200
         Picture         =   "FrmMain.frx":162B50
         Tag             =   "0"
         Top             =   1200
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Man 
         Height          =   480
         Index           =   0
         Left            =   4140
         Picture         =   "FrmMain.frx":16341A
         Tag             =   "0"
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   1
         Left            =   4440
         Picture         =   "FrmMain.frx":163724
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   2
         Left            =   4800
         Picture         =   "FrmMain.frx":163A2E
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   3
         Left            =   5100
         Picture         =   "FrmMain.frx":163D38
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   4
         Left            =   5400
         Picture         =   "FrmMain.frx":164042
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   5
         Left            =   4500
         Picture         =   "FrmMain.frx":16434C
         Top             =   540
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   6
         Left            =   4800
         Picture         =   "FrmMain.frx":164656
         Top             =   540
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   7
         Left            =   5100
         Picture         =   "FrmMain.frx":164960
         Top             =   540
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image ImgMan 
         Height          =   480
         Index           =   8
         Left            =   5400
         Picture         =   "FrmMain.frx":164C6A
         Top             =   540
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label BtnOptions 
      BackStyle       =   0  'Transparent
      Caption         =   "Click here for options"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4380
      Width           =   1575
   End
   Begin VB.Label CapLevel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   3900
      TabIndex        =   4
      Top             =   4065
      Width           =   120
   End
   Begin VB.Label InfoLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Index           =   1
      Left            =   3060
      TabIndex        =   3
      Top             =   4020
      Width           =   780
   End
   Begin VB.Label CapScore 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1020
      TabIndex        =   2
      Top             =   4065
      Width           =   1935
   End
   Begin VB.Label InfoLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4020
      Width           =   855
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Main"
      Begin VB.Menu MnuStartGame 
         Caption         =   "Start Game"
      End
      Begin VB.Menu MnuOpenLevel 
         Caption         =   "Open"
         Begin VB.Menu MnuOpenUserLevel 
            Caption         =   "User Level"
         End
         Begin VB.Menu MnuOpenStandardLevel 
            Caption         =   "Standard Level"
         End
      End
      Begin VB.Menu MnuGameMode 
         Caption         =   "GameMode"
         Begin VB.Menu MnuUserLevel 
            Caption         =   "User Levels"
         End
         Begin VB.Menu MnuStandardLevel 
            Caption         =   "Standard Levels"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnOptions_Click()
GamePause
Me.PopupMenu MnuMain
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim PreJahat As Byte
PreJahat = Jahat 'Hold previous direction
Select Case KeyCode 'change worm's direction according to keys
 Case 37
   If Jahat <> 2 And Not Paused Then Jahat = 4
 Case 39
   If Jahat <> 4 And Not Paused Then Jahat = 2
 Case 38
   If Jahat <> 3 And Not Paused Then Jahat = 1
 Case 40
   If Jahat <> 1 And Not Paused Then Jahat = 3
 Case 80
   GamePause
   ShowMsg "Game paused" + vbNewLine + "press ok button to resume"
   DoEvents
   DrawWorm 'redraw the worm
   GameResume
 Case 83
   'Start game from first level if user has lost . else restart current level
   LoadLevel Info.CurLevel
End Select
If PreJahat <> Jahat Then TimMoveWorm_Timer
End Sub

Sub ShowApple()
Dim Can As Boolean
Apple.Visible = False
'Find a free place and move the apple there
Do
  Apple.Top = Ran((Scr.Height - Apple.Height) \ WormWidth) * WormWidth
  Apple.Left = Ran((Scr.Width - Apple.Width) \ WormWidth) * WormWidth
  Can = True
  For I = 1 To Info.WallCount
      'Is the apple on a wall
      If Apple.Top + Apple.Height > Wall(I).Top And Apple.Top < Wall(I).Top + Wall(I).Height Then
        If Apple.Left + Apple.Width > Wall(I).Left And Apple.Left < Wall(I).Left + Wall(I).Width Then
          Can = False
        End If
      End If
  Next
Loop Until Can
Apple.Visible = True
DrawWorm
End Sub



Private Sub Form_Load()
MnuMain.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' this line is required to stop sound of this level
StopSound Info.LevelSound
End Sub



Private Sub MnuExit_Click()
' this line is required to stop sound of this level
StopSound Info.LevelSound
'End program
End
End Sub

Private Sub MnuHelp_Click()
FrmHelp.Show 1, Me
End Sub

Private Sub MnuOpenStandardLevel_Click()
Dim J
J = InputBox("Enter The Level Number", "")
If Val(J) Then
  Info.CurLevel = Val(J)
  Info.LevelsDir = CPath + "Levels\"
  LoadLevel Info.CurLevel
End If
End Sub

Private Sub MnuOpenUserLevel_Click()
Dim J
J = InputBox("Enter The Level Number", "")
If Val(J) Then
  Info.CurLevel = Val(J)
  Info.LevelsDir = CPath + "Levels\User\"
  LoadLevel Info.CurLevel
End If
End Sub

Private Sub MnuStandardLevel_Click()
MnuStandardLevel.Checked = True
MnuUserLevel.Checked = False
Info.LevelsDir = CPath + "Levels\"
Info.CurLevel = 1
End Sub

Private Sub MnuStartGame_Click()
LoadLevel Info.CurLevel
End Sub

Private Sub MnuUserLevel_Click()
MnuUserLevel.Checked = True
MnuStandardLevel.Checked = False
Info.LevelsDir = CPath + "Levels\User\"
Info.CurLevel = 1
End Sub


Private Sub TimHideDamagedCar_Timer()
DamagedCar(0).Visible = False
DamagedCar(1).Visible = False
TimHideDamagedCar = False
DrawWorm
End Sub

Private Sub TimMoveCar_Timer()
Dim I, N As Integer
For I = 1 To Car.Count - 1
  If Car(I).Visible Then
    If Car(I).Tag Then
      Car(I).Left = Car(I).Left + 1
    Else
      Car(I).Left = Car(I).Left - 1
    End If
    'Draw The Dots Which Are Cleared
    For N = 1 To Info.WormLen
      If WormDot(N).Top + WormWidth >= Car(I).Top - 10 And WormDot(N).Top <= Car(I).Top + Car(I).Height + 10 Then
        If WormDot(N).Left + WormWidth >= Car(I).Left - 10 And WormDot(N).Left <= Car(I).Left + Car(I).Width + 10 Then
          Scr.Circle (WormDot(N).Left, WormDot(N).Top), HalfWormWidth
        End If
     End If
    Next
    'Has The Car Faced The Wall
    For J = 1 To Wall.Count - 1
      If Car(I).Top + Car(I).Height > Wall(J).Top And Car(I).Top < Wall(J).Top + Wall(J).Height Then
        If Car(I).Left + Car(I).Width > Wall(J).Left And Car(I).Left < Wall(J).Left + Wall(J).Width Then
           Car(I).Picture = ImgCar(Car(I).Tag + 2)
           PlaySound CPath + "Sounds\Destroy.WAV"
           DamagedCar(Car(I).Tag).Top = Car(I).Top
           DamagedCar(Car(I).Tag).Left = Car(I).Left
           Car(I).Visible = False
           DamagedCar(Car(I).Tag).Visible = True
           TimHideDamagedCar = True
           DrawWorm
        End If
      End If
    Next
  
  End If
Next
End Sub

Private Sub TimMoveMan_Timer()
Dim I As Integer
For I = 1 To UBound(RecMan)
  If RecMan(I).Visible Then
     'Change Man's pic
     If RecMan(I).D < 20 / (TimMoveMan.Interval / 6) Then
       RecMan(I).D = RecMan(I).D + 1
     Else
       RecMan(I).D = 0
       If Man(I).Tag Then Man(I).Tag = 0 Else Man(I).Tag = 1
       If RecMan(I).Jahat Then
         RecMan(I).D = 0
         Man(I).Picture = ImgMan(RecMan(I).Jahat * 2 - Man(I).Tag)
       Else
         Man(I).Picture = ImgMan(5 + Man(I).Tag)
       End If
     End If
   'Change Man's Direction
    If WormDot(1).Left < Man(I).Left + Man(I).Width + 30 And WormDot(1).Left > Man(I).Left - 30 And WormDot(1).Top > Man(I).Top - 30 And WormDot(1).Top < Man(I).Top + Man(I).Height + 30 Then
      If RecMan(I).NearWorm = False Then
        RecMan(I).NearWorm = True
        RecMan(I).Jahat = Ran(4)
        PlaySound CPath + "Sounds\hey.WAV"
      End If
    Else
      RecMan(I).NearWorm = False
    End If
    'Move Man
    Select Case RecMan(I).Jahat
     Case 1
       Man(I).Top = Man(I).Top - 1
     Case 2
       Man(I).Left = Man(I).Left + 1
     Case 3
       Man(I).Top = Man(I).Top + 1
     Case 4
       Man(I).Left = Man(I).Left - 1
    End Select
    'Draw The Dots Which Are Cleared
    For N = 1 To Info.WormLen
      If WormDot(N).Top + WormWidth >= Man(I).Top - 10 And WormDot(N).Top <= Man(I).Top + Man(I).Height + 10 Then
        If WormDot(N).Left + WormWidth >= Man(I).Left - 10 And WormDot(N).Left <= Man(I).Left + Man(I).Width + 10 Then
          Scr.Circle (WormDot(N).Left, WormDot(N).Top), HalfWormWidth
        End If
     End If
    Next
    'Check Man's Position
    If Man(I).Top < -Man(I).Height Or Man(I).Left < -Man(I).Width Or Man(I).Top > Scr.Height Or Man(I).Left > Scr.Width Then RecMan(I).Visible = False
  End If
Next
End Sub

Private Sub TimMoveWorm_Timer()
Rec.Top = WormDot(Info.WormLen).Top - HalfWormWidth
Rec.Left = WormDot(Info.WormLen).Left - HalfWormWidth
Rec.Bottom = WormDot(Info.WormLen).Top + HalfWormWidth + 1
Rec.Right = WormDot(Info.WormLen).Left + HalfWormWidth + 1
RedrawWindow Scr.hWnd, Rec, 0, 1

'Remove Last Dot
For I = Info.WormLen To 2 Step -1
  WormDot(I).Left = WormDot(I - 1).Left
  WormDot(I).Top = WormDot(I - 1).Top
Next
Select Case Jahat 'move worm accordding to the direction
 Case 1
   WormDot(1).Top = WormDot(1).Top - WormWidth
 Case 2
   WormDot(1).Left = WormDot(1).Left + WormWidth
 Case 3
   WormDot(1).Top = WormDot(1).Top + WormWidth
 Case 4
   WormDot(1).Left = WormDot(1).Left - WormWidth
End Select
ControlWorm
'Draw worm's head
Scr.Circle (WormDot(1).Left, WormDot(1).Top), HalfWormWidth
Scr.Circle (WormDot(1).Left, WormDot(1).Top), HalfWormWidth - 2, QBColor(4)
'Clear Previous Head
Scr.Circle (WormDot(2).Left, WormDot(2).Top), HalfWormWidth, 0
End Sub

Sub ControlWorm()
Dim I As Integer
Dim Lost As Boolean
Lost = False
'Has The Worm Got Into Itself?
For I = 2 To Info.WormLen
  If WormDot(1).Top = WormDot(I).Top And WormDot(1).Left = WormDot(I).Left Then
    Lost = True
    I = Info.WormLen
  End If
Next
'Has The Worm Got Out Of The Form?
If WormDot(1).Top < 0 Or WormDot(1).Top + WormWidth > Scr.Height Or WormDot(1).Left < 0 Or WormDot(1).Left + WormWidth > Scr.Width Then
  Lost = True
End If
'Has Worm Eaten A Man?
For I = 1 To UBound(RecMan)
  If RecMan(I).Visible Then
    If WormDot(1).Top + WormWidth > Man(I).Top And WormDot(1).Top < Man(I).Top + Man(I).Height Then
      If WormDot(1).Left + WormWidth > Man(I).Left And WormDot(1).Left < Man(I).Left + Man(I).Width Then
        RecMan(I).Visible = False
        Man(I).Visible = False
        PlaySound CPath + "Sounds\ManDie.WAV"
        AddToWormLen Info.ManAddToWormLen
        Info.Score = Info.Score + Info.ManScore
        ScoreChanged
        DrawWorm
      End If
    End If
  End If
Next
'Has Worm Eaten A Car?
For I = 1 To Car.Count - 1
  If Car(I).Visible Then
    If WormDot(1).Top + WormWidth > Car(I).Top And WormDot(1).Top < Car(I).Top + Car(I).Height Then
      If WormDot(1).Left + WormWidth > Car(I).Left And WormDot(1).Left < Car(I).Left + Car(I).Width Then
        Car(I).Visible = False
        PlaySound CPath + "Sounds\CarDie.WAV"
        AddToWormLen Info.CarAddToWormLen
        Info.Score = Info.Score + Info.CarScore
        ScoreChanged
        DrawWorm
      End If
    End If
  End If
Next
For I = 1 To Wall.Count - 1
  If WormDot(1).Top + HalfWormWidth > Wall(I).Top And WormDot(1).Top < Wall(I).Top + Wall(I).Height Then
    If WormDot(1).Left + HalfWormWidth > Wall(I).Left And WormDot(1).Left < Wall(I).Left + Wall(I).Width Then
      Lost = True
    End If
  End If
Next
'Has The Worm Eaten The Apple
If WormDot(1).Top + WormWidth > Apple.Top And WormDot(1).Top < Apple.Top + Apple.Height Then
  If WormDot(1).Left + WormWidth > Apple.Left And WormDot(1).Left < Apple.Left + Apple.Width Then
    PlaySound CPath + "Sounds\AppleEat.WAV"
    AddToWormLen Info.AppleAddToWormLen
    ShowApple
    Info.Score = Info.Score + Info.AppleScore
    ScoreChanged
    DrawWorm
  End If
End If
If Lost Then
  WormDot(1).Top = -1000: WormDot(2).Top = -1000
  GamePause
  Apple.Visible = False
  Scr.Cls
  ShowMsg "      Game Over      ", CPath + "Sounds\GameOver.WAV", 26
  Info.CurLevel = 1
  Me.CapScore.Caption = 0
  Me.CapLevel.Caption = 1
End If
End Sub

Private Sub TimShowCar_Timer()
Dim I As Integer
I = Car.Count
If I > Info.CarCount Then
  TimShowCar = False
  Exit Sub
End If
Load Car(I)
If Ran(2) = 1 Then
  Car(I).Tag = 0
  Car(I).Picture = ImgCar(0)
  Car(I).Left = Scr.Width + Car(I).Width
  Car(I).Top = Ran((Scr.Height - Car(I).Height) \ WormWidth) * WormWidth
Else
  Car(I).Tag = 1
  Car(I).Picture = ImgCar(1)
  Car(I).Left = -Car(I).Width
  Car(I).Top = Ran((Scr.Height - Car(I).Height) \ WormWidth) * WormWidth
End If
Car(I).Visible = True
DrawWorm
PlaySound CPath + "Sounds\CarEnter.WAV"
End Sub

Private Sub TimShowMan_Timer()
Dim I As Integer
I = UBound(RecMan) + 1
If I > Info.ManCount Then
  TimShowMan = False
  Exit Sub
End If
ReDim Preserve RecMan(I)
Load Man(I)
'Find A Random Place
Man(I).Top = Ran((Scr.Height - Man(I).Height) \ WormWidth) * WormWidth
Man(I).Left = Ran((Scr.Width - Man(I).Width) \ WormWidth) * WormWidth

RecMan(I).Visible = True
RecMan(I).Jahat = Ran(4)
RecMan(I).D = 0
Man(I).Visible = True
DrawWorm
End Sub
Sub AddToWormLen(L As Integer)
ReDim Preserve WormDot(UBound(WormDot) + L)
For I = Info.WormLen To Info.WormLen + L
  WormDot(I).Left = WormDot(Info.WormLen).Left
  WormDot(I).Top = WormDot(Info.WormLen).Top
Next
Info.WormLen = Info.WormLen + L
End Sub
Sub DrawWorm()
For I = 1 To Info.WormLen
  Scr.Circle (WormDot(I).Left, WormDot(I).Top), HalfWormWidth
Next
End Sub
