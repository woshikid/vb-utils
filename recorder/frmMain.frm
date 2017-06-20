VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recorder"
   ClientHeight    =   2115
   ClientLeft      =   5670
   ClientTop       =   3405
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3090
   Begin VB.CommandButton Command2 
      Caption         =   "Deinit"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Init"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton PlayBack 
      Caption         =   "PlayBack"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Record 
      Caption         =   "Record"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "MainForm"
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
WaveInDeinit
WaveOutDeinit

End Sub

Private Sub Form_Load()
RegisterWinProc Me.hWnd

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
UnRegisterWinProc Me.hWnd

End Sub

Private Sub PlayBack_Click()
WaveOutPlayback

End Sub

Private Sub Record_Click()
WaveInRecord

End Sub

