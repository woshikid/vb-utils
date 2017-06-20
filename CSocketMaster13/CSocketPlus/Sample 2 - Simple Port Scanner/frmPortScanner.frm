VERSION 5.00
Begin VB.Form frmPortScanner 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Simple Port Scanner"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Index           =   0
      Left            =   3840
      Top             =   600
   End
   Begin VB.PictureBox picProgress 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   4155
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3480
      Width           =   4215
      Begin VB.Label lblProgress 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   285
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "2000"
      Top             =   1035
      Width           =   615
   End
   Begin VB.TextBox txtConnections 
      Height          =   285
      Left            =   1200
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "35"
      Top             =   1035
      Width           =   495
   End
   Begin VB.ListBox lstOpenPorts 
      Height          =   1620
      ItemData        =   "frmPortScanner.frx":0000
      Left            =   240
      List            =   "frmPortScanner.frx":0002
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtEndPort 
      Height          =   285
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "65535"
      Top             =   550
      Width           =   615
   End
   Begin VB.TextBox txtStartPort 
      Height          =   285
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "1"
      Top             =   550
      Width           =   615
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   70
      Width           =   3615
   End
   Begin VB.Label Label6 
      Caption         =   "ms."
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Time-out:"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Connections:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "to:"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ports from:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Host/IP:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPortScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************
'This code is an example for CSocketPlus
'by Emiliano Scavuzzo
'
'NOTE: I suggest running this demo as a compiled EXE
'for speed's sake, or you could remove all the comments
'from the class.
'**************************************************************************************

Option Explicit

Private WithEvents spSockets As CSocketPlus
Attribute spSockets.VB_VarHelpID = -1

Private lngStartPort As Long
Private lngEndPort As Long
Private lngCurrentPort As Long
Private intConnections As Integer
Private intCurrentConnections As Integer
Private strRemoteHost As String
Private lngTimeOut As Long


Private Sub cmdScan_Click()

DisableControls

Set spSockets = New CSocketPlus
lstOpenPorts.Clear

If TakeInfo = False Then Exit Sub

Dim intCount As Integer

For intCount = 1 To intConnections
    spSockets.ArrayAdd intCount
    If TestNextPort(intCount) = False Then Exit Sub
Next

End Sub


Private Sub cmdStop_Click()
CleanSystem
End Sub


Private Sub DisableControls()
cmdScan.Enabled = False
txtHost.Enabled = False
txtStartPort.Enabled = False
txtEndPort.Enabled = False
txtConnections.Enabled = False
txtTimeOut.Enabled = False
cmdStop.Enabled = True
End Sub


Private Sub EnableControls()
cmdStop.Enabled = False
txtHost.Enabled = True
txtStartPort.Enabled = True
txtEndPort.Enabled = True
txtConnections.Enabled = True
txtTimeOut.Enabled = True
cmdScan.Enabled = True
End Sub


Private Function TakeInfo() As Boolean
On Error GoTo Error_Handler

If Trim(txtHost.Text) = "" Then Err.Raise 1, , "You must enter the host."
strRemoteHost = txtHost.Text

If IsNumeric(txtStartPort.Text) Then
    lngStartPort = Val(txtStartPort.Text)
    lngCurrentPort = lngStartPort
    If lngStartPort < 1 Or lngStartPort > 65535 Then Err.Raise 1, , "Wrong initial port."
Else
    Err.Raise 1, , "Wrong initial port."
End If

If IsNumeric(txtEndPort.Text) Then
    lngEndPort = Val(txtEndPort.Text)
    If lngEndPort < 1 Or lngEndPort > 65535 Then Err.Raise 1, , "Wrong end port."
Else
    Err.Raise 1, , "Wrong end port."
End If

If lngStartPort > lngEndPort Then Err.Raise 1, , "Initial port is bigger than end port."

If IsNumeric(txtConnections.Text) Then
    intConnections = Val(txtConnections.Text)
    If intConnections < 1 Then Err.Raise 1, , "Wrong connection number."
    If intConnections > 100 Then Err.Raise 1, , "The number you entered for connections must be less than 100."
Else
    Err.Raise 1, , "Wrong connection number."
End If

If IsNumeric(txtTimeOut.Text) Then
    lngTimeOut = Val(txtTimeOut.Text)
    If lngTimeOut < 1 Or lngTimeOut > 65535 Then Err.Raise 1, , "Wrong time-out."
Else
    Err.Raise 1, , "Wrong time-out."
End If

TakeInfo = True

Exit Function
Error_Handler:
    Dim strDescription As String
    strDescription = Err.Description
    CleanSystem
    MsgBox strDescription, vbExclamation, "Error"
End Function


Private Sub CleanSystem()
Set spSockets = Nothing

lblProgress.Width = 0

Dim intCount As Integer

On Error Resume Next
For intCount = 1 To intConnections
   Unload Timer(intCount)
Next
On Error GoTo 0

intCurrentConnections = 0

EnableControls
End Sub

'Scan next port
Private Function TestNextPort(ByVal Index As Integer) As Boolean
On Error GoTo Error_Handler

If lngCurrentPort <= lngEndPort Then

    intCurrentConnections = intCurrentConnections + 1
    spSockets.Connect Index, strRemoteHost, lngCurrentPort
    lblProgress.Width = picProgress.Width * (lngCurrentPort - lngStartPort) / (lngEndPort - lngStartPort)
    
    lngCurrentPort = lngCurrentPort + 1
    
    Load Timer(Index)
    Timer(Index).Interval = lngTimeOut
    Timer(Index).Enabled = True
Else
    If intCurrentConnections = 0 Then
        CleanSystem
        MsgBox "Scan completed", vbOKOnly, "Done"
    End If
End If

TestNextPort = True

Exit Function
Error_Handler:
CleanSystem
MsgBox "Error scanning " + strRemoteHost, vbCritical, "Error"

End Function


Private Sub spSockets_Connect(ByVal Index As Variant)
Unload Timer(Index)
lstOpenPorts.AddItem "[OPEN PORT] " & spSockets.RemotePort(Index)
intCurrentConnections = intCurrentConnections - 1
spSockets.CloseSck Index
TestNextPort Index
End Sub


Private Sub spSockets_Error(ByVal Index As Variant, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Unload Timer(Index)
spSockets.CloseSck Index
intCurrentConnections = intCurrentConnections - 1
TestNextPort Index
End Sub


Private Sub Timer_Timer(Index As Integer)
Debug.Print "Time-out"

Unload Timer(Index)
spSockets.CloseSck Index
intCurrentConnections = intCurrentConnections - 1
TestNextPort Index

End Sub
