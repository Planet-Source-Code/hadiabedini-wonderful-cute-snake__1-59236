VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   3675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5490
   Icon            =   "FrmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   765
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "FrmHelp.frx":000C
      Top             =   2100
      Width           =   5115
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4140
      TabIndex        =   0
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Try my site and tell me your idea about it"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2940
      Width           =   2820
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "If you need help or you can help me in programming send me a message"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1500
      Width           =   5100
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "I am Hadi Abedini (Male , From Iran) , I love programming and chating"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1260
      Width           =   4875
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "About me"
      ForeColor       =   &H00404080&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Press ""S"" to restart the game"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Eat every thing you see except walls"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "How to play this game ?"
      ForeColor       =   &H00404080&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
End Sub
