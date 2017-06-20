VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   3000
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents TCP As CSocketPlus
Attribute TCP.VB_VarHelpID = -1

Private Sub Command1_Click()
    TCP.ArrayAdd "aa"
    TCP.LocalPort("aa") = 1002
    TCP.Listen "aa"
    
    TCP.ArrayAdd "bb"
    TCP.RemoteHost("bb") = "127.0.0.1"
    TCP.RemotePort("bb") = 1002
    TCP.Connect "bb"
    
    
    
End Sub

Private Sub Command2_Click()
    Dim a() As Byte
    ReDim a(255)
    Dim i As Long
    For i = 0 To 255
        a(i) = i
    Next i
    TCP.SendData "bb", a
End Sub

Private Sub Form_Load()
    Set TCP = New CSocketPlus
End Sub


Private Sub TCP_ConnectionRequest(ByVal Index As Variant, ByVal requestID As Long)
TCP.CloseSck Index
TCP.Accept Index, requestID
Text1.Text = ""
End Sub


Private Sub TCP_DataArrival(ByVal Index As Variant, ByVal bytesTotal As Long)
Dim strData() As Byte
TCP.GetData "aa", strData

MsgBox bytesTotal
MsgBox UBound(strData) + 1

Dim i As Long
For i = 0 To UBound(strData)
   Text1.Text = Text1.Text + CStr(strData(i)) + vbNewLine
Next i
'Text1.Text = CStr(strData)
End Sub
