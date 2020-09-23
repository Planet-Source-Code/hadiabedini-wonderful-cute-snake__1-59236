VERSION 5.00
Begin VB.Form FrmOpenFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open File"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "FrmOpenFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame PlayFrame 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   4680
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
      Begin VB.CommandButton BtnStop 
         Caption         =   "Stop"
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton BtnPlay 
         Caption         =   "Play"
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame ImgBoxFrame 
      BorderStyle     =   0  'None
      Height          =   1155
      Left            =   4680
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
      Begin VB.Shape ImgBorder 
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image ImgBox 
         Height          =   1070
         Left            =   10
         Stretch         =   -1  'True
         Top             =   10
         Width           =   1190
      End
   End
   Begin VB.FileListBox FileBox 
      Height          =   2820
      Left            =   2460
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.DirListBox DirBox 
      Height          =   2565
      Left            =   180
      TabIndex        =   3
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox DriveBox 
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmOpenFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String, FileName As String, IsPlaying As Boolean
Option Explicit

Private Sub BtnPlay_Click()
PlaySound Path
IsPlaying = True
End Sub

Private Sub BtnStop_Click()
StopSound Path
IsPlaying = False
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub DirBox_Change()
FileBox.Path = DirBox.Path
End Sub

Private Sub DriveBox_Change()
DirBox.Path = DriveBox.Drive
End Sub

Private Sub FileBox_Click()
On Error Resume Next
'Stop any sounds
If IsPlaying Then
  StopSound Path
  IsPlaying = False
End If
'Hold current file path and file name
If Len(FileBox.Path) = 3 Then Path = FileBox.Path + FileBox.FileName Else Path = FileBox.Path + "\" + FileBox.FileName
FileName = FileBox.FileName
'If the selected file is an image file show it
ImgBox.Picture = LoadPicture(Path)
End Sub

Private Sub OKButton_Click()
'Stop any sounds
If IsPlaying Then
  StopSound Path
  IsPlaying = False
End If
If Len(Path) Then 'If user has selected a file
  If PlayFrame.Visible Then 'If user is browsing a sound file
      'If the selected file does not exist in sounds folder copy it
      If Dir(FrmMain.Cpath + "Sounds\" + FileName) = "" Then FileCopy Path, FrmMain.Cpath + "Sounds\" + FileName
      FrmMain.BgSound = FileName
      Unload Me
  Else
      FrmMain.ImgBox.Picture = ImgBox.Picture
      'If the selected file does not exist in images folder copy it
      If Dir(FrmMain.Cpath + "Images\" + FileName) = "" Then FileCopy Path, FrmMain.Cpath + "Images\" + FileName
      FrmMain.BgImage = FileName
      Unload Me
  End If
End If
End Sub
