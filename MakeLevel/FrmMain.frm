VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Level Maker"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnOpenLevel 
      Caption         =   "Open a level"
      Height          =   555
      Left            =   7500
      TabIndex        =   48
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton BtnOpenUserLevel 
      Caption         =   "Open a user level"
      Height          =   555
      Left            =   5820
      TabIndex        =   47
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton BtnNewLevel 
      Caption         =   "New level"
      Height          =   555
      Left            =   120
      TabIndex        =   46
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton BtnSaveAsUserLevel 
      Caption         =   "Save as user level"
      Height          =   555
      Left            =   1800
      TabIndex        =   11
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton BtnSaveAsLevel 
      Caption         =   "Save as level"
      Height          =   555
      Left            =   3480
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Level Info"
      ForeColor       =   &H00808080&
      Height          =   5295
      Left            =   6900
      TabIndex        =   9
      Top             =   120
      Width           =   2895
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   16
         Left            =   1980
         TabIndex        =   49
         Text            =   "30"
         Top             =   1140
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   15
         Left            =   1980
         TabIndex        =   44
         Text            =   "10"
         Top             =   1980
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   14
         Left            =   1980
         TabIndex        =   42
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   13
         Left            =   1980
         TabIndex        =   40
         Text            =   "10"
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   12
         Left            =   1980
         TabIndex        =   38
         Text            =   "120"
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton BtnBrowseSound 
         Caption         =   "..."
         Height          =   315
         Left            =   2460
         TabIndex        =   14
         Top             =   4800
         Width           =   315
      End
      Begin VB.CommandButton BtnBrowsePic 
         Caption         =   "..."
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   4800
         Width           =   315
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Worm left at start"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   50
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label CapLabel 
         Caption         =   "Required Score To Win Level"
         Height          =   375
         Index           =   15
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Name of the level"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   43
         Top             =   1620
         Width           =   1245
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Worm length at start"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   41
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Move worm every"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Level Sound"
         Height          =   195
         Left            =   1500
         TabIndex        =   15
         Top             =   4860
         Width           =   900
      End
      Begin VB.Shape ImgBorder 
         Height          =   1815
         Left            =   180
         Top             =   2820
         Width           =   2595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Level Image"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   4860
         Width           =   870
      End
      Begin VB.Image ImgBox 
         Height          =   1785
         Left            =   195
         Stretch         =   -1  'True
         Top             =   2835
         Width           =   2565
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map"
      ForeColor       =   &H00808080&
      Height          =   3915
      Left            =   3540
      TabIndex        =   5
      Top             =   1500
      Width           =   3255
      Begin VB.TextBox TxtMap 
         Height          =   2115
         Left            =   420
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   660
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type the level map"
         Height          =   195
         Left            =   420
         TabIndex        =   8
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   $"FrmMain.frx":000C
         Height          =   885
         Left            =   420
         TabIndex        =   7
         Top             =   2940
         Width           =   2610
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Car Info"
      ForeColor       =   &H00808080&
      Height          =   2655
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3315
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   4
         Left            =   1980
         TabIndex        =   22
         Text            =   "2"
         Top             =   2100
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   3
         Left            =   1980
         TabIndex        =   20
         Text            =   "1"
         Top             =   1680
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   2
         Left            =   1980
         TabIndex        =   18
         Text            =   "10"
         Top             =   1260
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   0
         Left            =   1980
         TabIndex        =   16
         Text            =   "15"
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   3
         Text            =   "9000"
         Top             =   360
         Width           =   555
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Increase worm length "
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1560
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Increase score "
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   21
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Car count"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Move all cars every"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   900
         Width           =   1380
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Show a new car every"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   1590
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Man Info"
      ForeColor       =   &H00808080&
      Height          =   2595
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2820
      Width           =   3315
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   9
         Left            =   1980
         TabIndex        =   32
         Text            =   "2"
         Top             =   2040
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   8
         Left            =   1980
         TabIndex        =   30
         Text            =   "1"
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   7
         Left            =   1980
         TabIndex        =   28
         Text            =   "10"
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   6
         Left            =   1980
         TabIndex        =   26
         Text            =   "50"
         Top             =   780
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   5
         Left            =   1980
         TabIndex        =   24
         Text            =   "8000"
         Top             =   300
         Width           =   555
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Increase worm length "
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   33
         Top             =   2100
         Width           =   1560
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Increase score "
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   31
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Man count"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   29
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Move all men every"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Show a new man every"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Apple Info"
      ForeColor       =   &H00808080&
      Height          =   1335
      Index           =   0
      Left            =   3540
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   11
         Left            =   2040
         TabIndex        =   36
         Text            =   "1"
         Top             =   780
         Width           =   555
      End
      Begin VB.TextBox Box 
         Height          =   315
         Index           =   10
         Left            =   2040
         TabIndex        =   34
         Text            =   "1"
         Top             =   360
         Width           =   555
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Increase worm length "
         Height          =   195
         Index           =   11
         Left            =   300
         TabIndex        =   37
         Top             =   840
         Width           =   1560
      End
      Begin VB.Label CapLabel 
         AutoSize        =   -1  'True
         Caption         =   "Increase score "
         Height          =   195
         Index           =   10
         Left            =   300
         TabIndex        =   35
         Top             =   420
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cpath As String
Public BgSound As String, BgImage As String


Private Sub BtnBrowsePic_Click()
FrmOpenFile.Show 1, Me
End Sub

Private Sub BtnBrowseSound_Click()
'Prepare FrmOpenFile to browse a sound file
FrmOpenFile.PlayFrame.Visible = True
FrmOpenFile.ImgBoxFrame.Visible = False

FrmOpenFile.Show 1, Me
End Sub

Private Sub BtnNewLevel_Click()
'Reset all values
For I = 0 To Box.Count - 1
  Box(I).Text = Box(I).Tag
Next
BgImage = ""
BgSound = ""
TxtMap.Text = ""
ImgBox.Picture = LoadPicture()
End Sub

Private Sub BtnOpenLevel_Click()
LoadLevel Cpath + "Levels\Level_"
End Sub

Private Sub BtnOpenUserLevel_Click()
LoadLevel Cpath + "Levels\User\Level_"
End Sub

Private Sub BtnSaveAsLevel_Click()
SaveLevel Cpath + "Levels\Level_"
End Sub

Private Sub BtnSaveAsUserLevel_Click()
SaveLevel Cpath + "Levels\User\Level_"
End Sub



Private Sub Form_Load()
Cpath = App.Path + "\"
'Store default values
For I = 0 To Box.Count - 1
  Box(I).Tag = Box(I).Text
Next
End Sub


Sub SaveLevel(P As String)
Dim S As String, I As Integer, J, MapLine
'get level number
J = InputBox("Enter The Level Number", "")
If Len(J) = 0 Then Exit Sub
'split level map so that it can be written in a text file
TxtMap = UCase(TxtMap)
MapLine = Split(TxtMap, vbNewLine)
Open P + CStr(J) + ".Txt" For Output As #1
'Print all texts in textboxes
For I = 0 To Box.Count - 1
  Print #1, Trim(Box(I).Text)
Next
'Write level map in the file
For I = 0 To UBound(MapLine)
  Write #1, MapLine(I)
Next
Print #1, "End"
'Save file names of level sound and image
Write #1, BgImage
Write #1, BgSound
Close #1
End Sub
Sub LoadLevel(P As String)
Dim S As String, I As Integer, J
'get level number
J = InputBox("Enter The Level Number", "")
If Len(J) = 0 Then Exit Sub
If Dir(P + CStr(J) + ".Txt") = "" Then
  MsgBox "This level does not exist !"
  Exit Sub
End If
Open P + CStr(J) + ".Txt" For Input As #1
For I = 0 To Box.Count - 1
  Input #1, S
  Box(I).Text = S
Next

'Load level map
TxtMap = ""
Input #1, S
Do While LCase(S) <> "end"
  TxtMap = TxtMap + S + vbNewLine
  Input #1, S
Loop

'Load file names of level sound and image
Input #1, BgImage
If Len(BgImage) Then ImgBox.Picture = LoadPicture(Cpath + "Images\" + BgImage)
Input #1, BgSound
Close #1
End Sub

Private Sub Label_Click(Index As Integer)

End Sub
