VERSION 5.00
Begin VB.Form Server 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1695
      Left            =   0
      TabIndex        =   13
      Top             =   2160
      Width           =   5055
      Begin VB.TextBox txtLog 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox txtMessage 
         Height          =   285
         Left            =   120
         MaxLength       =   32
         TabIndex        =   7
         Top             =   0
         Width           =   3975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   255
         Left            =   4200
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.ListBox lstServer 
      Height          =   1230
      ItemData        =   "Server.frx":0000
      Left            =   120
      List            =   "Server.frx":0002
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   600
      MaxLength       =   12
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   2400
      TabIndex        =   9
      Top             =   720
      Width           =   2535
      Begin VB.Frame frmCommands 
         Height          =   1215
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
         Begin VB.CommandButton cmdListen 
            Caption         =   "Listen"
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   960
            MaxLength       =   5
            TabIndex        =   4
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   315
            Left            =   1320
            TabIndex        =   6
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Port:"
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
      Height          =   255
      Left            =   30
      TabIndex        =   12
      Top             =   285
      Width           =   495
   End
End
Attribute VB_Name = "Server"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is an example for CSocketPlus
'by Emiliano Scavuzzo
'**************************************************************************************

Option Explicit
Private WithEvents spServer As CSocketPlus
Attribute spServer.VB_VarHelpID = -1

Private Sub Form_Load()
Set spServer = New CSocketPlus
Client.Show
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Error_Handler
spServer.ArrayAdd txtServer.Text
lstServer.AddItem txtServer.Text
txtServer.Text = ""

Exit Sub
Error_Handler:
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdRemove_Click()
spServer.ArrayRemove lstServer.List(lstServer.ListIndex)
lstServer.RemoveItem lstServer.ListIndex
cmdRemove.Enabled = False
frmCommands.Visible = False
cmdSend.Enabled = False
txtLog.Text = ""
End Sub

Private Sub cmdListen_Click()
On Error GoTo Error_Handler

Dim strIndex As String
strIndex = lstServer.List(lstServer.ListIndex)
spServer.Bind strIndex, txtPort, spServer.LocalIP(strIndex)
spServer.Listen strIndex
ShowInfo
Exit Sub

Error_Handler:
    MsgBox Err.Description, vbCritical, "Error on server " & strIndex
End Sub

Private Sub cmdClose_Click()
On Error GoTo Error_Handler

Dim strIndex As String
strIndex = lstServer.List(lstServer.ListIndex)
spServer.CloseSck strIndex
spServer.Tag(strIndex) = ""
cmdSend.Enabled = False
ShowInfo
Exit Sub

Error_Handler:
    MsgBox Err.Description, vbCritical, "Error of " & strIndex
End Sub

Private Sub cmdSend_Click()
Dim strIndex As String
strIndex = lstServer.List(lstServer.ListIndex)

spServer.SendData strIndex, strIndex & ":" & vbCrLf & txtMessage.Text
txtMessage.Text = ""
End Sub

Private Sub lstServer_Click()
If lstServer.ListIndex >= 0 Then ShowInfo
End Sub

Private Sub lstServer_GotFocus()
If lstServer.ListIndex >= 0 Then ShowInfo
End Sub

Private Sub txtServer_Change()
If Len(txtServer.Text) > 0 Then
    cmdAdd.Enabled = True
Else
    cmdAdd.Enabled = False
End If
End Sub

Private Sub ShowInfo()
Dim strIndex As String
strIndex = lstServer.List(lstServer.ListIndex)

cmdRemove.Enabled = True
frmCommands.Visible = True

frmCommands.Caption = strIndex
txtPort.Text = spServer.LocalPort(strIndex)
txtLog.Text = spServer.Tag(strIndex)

If spServer.State(strIndex) = sckClosed Then
    txtPort.Locked = False
    txtPort.BackColor = &H80000005
    cmdListen.Enabled = True
    cmdClose.Enabled = False
    cmdSend.Enabled = False
Else
    txtPort.Locked = True
    txtPort.BackColor = &H8000000F
    cmdListen.Enabled = False
    cmdClose.Enabled = True
    If spServer.State(strIndex) = sckConnected Then cmdSend.Enabled = True
End If
End Sub

Private Sub spServer_ConnectionRequest(ByVal Index As Variant, ByVal requestID As Long)
spServer.CloseSck Index
spServer.Accept Index, requestID
cmdSend.Enabled = True
End Sub

Private Sub spServer_DataArrival(ByVal Index As Variant, ByVal bytesTotal As Long)
Dim strData As String
spServer.GetData Index, strData
spServer.Tag(Index) = spServer.Tag(Index) & strData & vbCrLf
ShowInfo
End Sub

Private Sub spServer_CloseSck(ByVal Index As Variant)
MsgBox "Connection closed", vbInformation, "Message on server " & Index
spServer.CloseSck Index
spServer.Tag(Index) = ""
cmdSend.Enabled = False
ShowInfo
End Sub

Private Sub spServer_Error(ByVal Index As Variant, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Description, vbCritical, "Error on server " & Index
spServer.CloseSck Index
spServer.Tag(Index) = ""
cmdSend.Enabled = False
ShowInfo
End Sub

Private Sub txtLog_Change()
txtLog.SelStart = Len(txtLog)
txtLog.SelLength = 0
End Sub

