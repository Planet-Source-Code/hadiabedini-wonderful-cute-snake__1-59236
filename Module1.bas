Attribute VB_Name = "ModCommon"
Type TypMan
  Jahat As Byte 'jahat means direction
  D As Byte
  NearWorm As Boolean
  Visible As Boolean
End Type
Type TypDot
  Left As Integer
  Top As Integer
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type TypInfo
  CurLevel As Integer
  ManCount As Integer
  CarCount As Integer
  CarScore As Integer
  ManScore As Integer
  AppleScore As Integer
  WallCount As Integer
  AppleAddToWormLen As Integer
  ManAddToWormLen As Integer
  CarAddToWormLen As Integer
  WormLen As Integer
  Score As Long
  RequiredScore As Long
  LevelsDir As String
  LevelName As String
  LevelSound As String
End Type
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function mciSendString Lib "WINMM.DLL" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "WINMM.DLL" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Public RecMan() As TypMan
Public Rec As RECT, Info As TypInfo
Public WormDot() As TypDot, Block() As RECT, WormWidth As Integer, HalfWormWidth As Single
Public Jahat As Byte
Public TwipX, TwipY
Public Paused As Boolean, CPath As String

Sub Main()
Randomize
CPath = App.Path + "\"
TwipX = Screen.TwipsPerPixelX
TwipY = Screen.TwipsPerPixelY
WormWidth = 8
HalfWormWidth = WormWidth / 2
Jahat = 1
Info.CurLevel = 1
Info.LevelsDir = CPath + "Levels\"
FrmMain.Show
End Sub


Sub GamePause()
Paused = True
FrmMain.TimMoveWorm = False
FrmMain.TimMoveCar = False
FrmMain.TimMoveMan = False
FrmMain.TimShowCar = False
FrmMain.TimShowMan = False
End Sub
Sub GameResume()
Paused = False
FrmMain.TimMoveWorm = True
FrmMain.TimMoveCar = True
FrmMain.TimMoveMan = True
FrmMain.TimShowCar = True
FrmMain.TimShowMan = True
End Sub

Sub ShowMsg(Str As String, Optional Snd As String = "", Optional FontSize As Integer = 18)
Unload FrmMsg
StopSound Info.LevelSound
FrmMsg.MsgLabel.FontSize = FontSize
FrmMsg.MsgLabel.Caption = Str
FrmMsg.MsgLabel.Top = 200: FrmMsg.MsgLabel.Left = 200
DoEvents
FrmMsg.Height = FrmMsg.MsgLabel.Height + FrmMsg.OKButton.Height + 800
FrmMsg.Width = FrmMsg.MsgLabel.Width + 500
FrmMsg.OKButton.Top = FrmMsg.Height - FrmMsg.OKButton.Height - 200
FrmMsg.OKButton.Left = FrmMsg.Width / 2 - FrmMsg.OKButton.Width / 2
PlaySound Snd
FrmMsg.Show 1, FrmMain
End Sub

Sub PlaySound(StrPath As String)
On Error Resume Next
Call mciSendString("play " & StrPath, 0&, 0, 0)
End Sub
Sub StopSound(StrPath As String)
On Error Resume Next
Call mciSendString("stop " & StrPath, 0&, 0, 0)
End Sub
Function Ran(I As Variant) As Variant
Ran = Fix(Rnd * I) + 1
End Function
Sub ScoreChanged()
'Show current score
FrmMain.CapScore = CStr(Info.Score) + " OF " + CStr(Info.RequiredScore)
If Info.Score >= Info.RequiredScore Then
  GamePause
  If Len(Dir(Info.LevelsDir + "Level_" + CStr(Info.CurLevel + 1) + ".Txt")) = 0 Then
    ShowMsg "Congradulations" + vbNewLine + "You have passed all levels" + vbNewLine + "Thank you for playing this" + vbNewLine + "game ." + vbNewLine + "Author : Hadi Abedini 18,M", CPath + "Sounds\WinGame.WAV"
    FrmMain.Apple.Visible = False
  Else
    ShowMsg "You Won" + vbNewLine + "Click OK Button To Start Level " + CStr(Info.CurLevel + 1), CPath + "Sounds\WinGame.WAV"
    LoadLevel Info.CurLevel + 1
  End If
End If
End Sub
