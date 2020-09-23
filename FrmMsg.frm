VERSION 5.00
Begin VB.Form FrmMsg 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   2760
   ClientTop       =   3360
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0080C0FF&
      Caption         =   "OK"
      Height          =   375
      Left            =   2505
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1860
      Width           =   675
   End
   Begin VB.Label MsgLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15
      TabIndex        =   1
      Top             =   0
      Width           =   225
   End
End
Attribute VB_Name = "FrmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
End Sub
