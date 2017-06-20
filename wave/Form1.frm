VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "deinit"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "play"
      Height          =   1095
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "red"
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "init"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
WaveInInit CHANNEL, SAMPLES, BITS
WaveOutInit CHANNEL, SAMPLES, BITS
End Sub

Private Sub Command2_Click()
WaveInRecord
End Sub

Private Sub Command3_Click()
WaveOutPlayback
End Sub

Private Sub Command4_Click()
toend = True
WaveInDeinit
WaveOutDeinit
End Sub

Private Sub Form_Load()

End Sub
