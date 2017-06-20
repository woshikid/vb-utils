VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socket5 proxy"
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "2222"
      Top             =   120
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   0
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim step1 As Boolean
Dim step2 As Boolean

Private Sub Command1_Click()
    If Command1.Caption = "start" Then
        Winsock1.LocalPort = Val(Text1.Text)
        Winsock1.Listen
        Command1.Caption = "stop"
    Else
        Winsock1.Close
        Command1.Caption = "start"
    End If
End Sub

Private Sub Text1_Change()
    Dim i As Long
    i = Val(Text1.Text)
    If i > 65535 Then
        i = 65535
    ElseIf i < 0 Then
        i = 0
    End If
    Text1.Text = i
End Sub

Private Sub Winsock1_Close()
    Winsock2.Close
    Winsock1.Close
    Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
    step1 = True
    step2 = True
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim bytes() As Byte
    Winsock1.GetData bytes, vbString + vbArray
    If step1 = True Then
        If bytesTotal <> 3 Or bytes(0) <> 5 Then
            Winsock1_Close
            Exit Sub
        End If
        ReDim bytes(1)
        bytes(0) = 5
        bytes(1) = 0
        Winsock1.SendData bytes
        step1 = False
    ElseIf step2 = True Then
        If bytes(0) <> 5 Or bytes(1) <> 1 Then
            Winsock1_Close
            Exit Sub
        End If
        Dim i As Integer
        Dim addr As String
        If bytes(3) = 1 Then
            i = bytes(4) * 256 + bytes(5)
            addr = Str(bytes(4)) + "." + Str(bytes(5)) + "." + Str(bytes(6)) + "." + Str(bytes(7))
            i = bytes(8) * 256 + bytes(9)
        ElseIf bytes(3) = 3 Then
            For i = 4 To bytesTotal - 3
                addr = addr + Chr(bytes(i))
            Next i
            i = bytes(bytesTotal - 2) * 256 + bytes(bytesTotal - 1)
        Else
            Winsock1_Close
            Exit Sub
        End If
        Winsock2.Close
        Winsock2.RemoteHost = addr
        Winsock2.RemotePort = i
        Winsock2.Connect
        step2 = False
    Else
        Winsock2.SendData bytes
    End If
End Sub

Private Sub Winsock2_Close()
    Winsock2.Close
    Winsock1.Close
End Sub

Private Sub Winsock2_Connect()
    Dim bytes() As Byte
    ReDim bytes(9)
    bytes(0) = 5
    bytes(1) = 0
    bytes(2) = 0
    bytes(3) = 1
    Winsock1.SendData bytes
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
    Dim bytes() As Byte
    Winsock2.GetData bytes, vbString + vbArray
    Winsock1.SendData bytes
End Sub

