Attribute VB_Name = "ModLevel"
Sub LoadLevel(Level As Integer)
Dim S(16) As Variant, I As Integer, T As String
'Check wether the file exists or not
If Dir(Info.LevelsDir + "Level_" + CStr(Level) + ".Txt") = "" Then
  ShowMsg "Level " + CStr(Level) + "  Doesn't exist"
  Exit Sub
End If
ResetAll
StopSound Info.LevelSound
Info.CurLevel = Level
Open Info.LevelsDir + "Level_" + CStr(Level) + ".Txt" For Input As #1
For I = 0 To 16
  Input #1, S(I)
Next

'Use loaded values
With FrmMain
.TimShowCar.Interval = S(1)
.TimMoveCar.Interval = S(0)
Info.CarCount = S(2)
Info.CarScore = S(3)
Info.CarAddToWormLen = S(4)

.TimShowMan.Interval = S(5)
.TimMoveMan.Interval = S(6)
Info.ManCount = S(7)
Info.ManScore = S(8)
Info.ManAddToWormLen = S(9)

Info.AppleScore = S(10)
Info.AppleAddToWormLen = S(11)
.TimMoveWorm.Interval = S(12)
Info.WormLen = S(13)
ReDim WormDot(Info.WormLen)
WormDot(1).Top = .Scr.Height - WormWidth
WormDot(1).Left = S(16) * WormWidth + WormWidth / 2
For I = 2 To Info.WormLen
  WormDot(I).Top = WormDot(1).Top '+ (I * WormWidth)
  WormDot(I).Left = WormDot(1).Left '.Scr.Width / 2 + (I * WormWidth)
Next
Info.LevelName = S(14)
Info.RequiredScore = S(15)

GetLevelMap 'Load the map

Input #1, T
If Len(T) Then 'if this level has a background image
  .Scr.Picture = LoadPicture(CPath + "Images\" + T)
Else
  .Scr.Picture = LoadPicture()
End If
Input #1, T
Info.LevelSound = CPath + "Sounds\" + T
PlaySound Info.LevelSound
.CapLevel.Caption = CStr(Info.CurLevel) + "  " + Info.LevelName
End With
Close #1
FrmMain.ShowApple
End Sub

Sub ResetAll()
Info.WallCount = 0
Info.Score = 0
FrmMain.CapScore.Caption = 0
ReDim RecMan(0)
ReDim WormDot(0)
For I = 1 To FrmMain.Wall.Count - 1
  Unload FrmMain.Wall(I)
Next
For I = 1 To FrmMain.Car.Count - 1
  Unload FrmMain.Car(I)
Next
For I = 1 To FrmMain.Man.Count - 1
  Unload FrmMain.Man(I)
Next
Jahat = 1
GameResume
End Sub
Sub GetLevelMap()
Dim S As String, Cx, Cy As Integer, I, T As Integer
Cx = 0: Cy = 0
Input #1, S
Do While LCase(S) <> "end"
  For I = 1 To Len(S)
    Select Case Mid(S, I, 1)
      Case "W"
        T = FrmMain.Wall.Count
        Load FrmMain.Wall(T)
        Info.WallCount = Info.WallCount + 1
        FrmMain.Wall(T).Top = Cy
        FrmMain.Wall(T).Left = Cx
        FrmMain.Wall(T).Visible = True
        Cx = Cx + FrmMain.Wall(T).Width
      Case " "
        Cx = Cx + WormWidth
    End Select
  Next
  Input #1, S
  Cy = Cy + FrmMain.Wall(0).Height
  Cx = 0
Loop
End Sub
