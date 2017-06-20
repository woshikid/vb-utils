VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "server"
   ClientHeight    =   915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   120
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ready"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "d:\test.rar"
      Top             =   120
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label1 
      Caption         =   "client IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim receivedLen As Long
Dim ifReady As Boolean
Dim cache() As Byte
Dim cacheStart As Long
Dim cacheEnd As Long

Private Sub Command1_Click()
    Winsock1.LocalPort = 9998
    Winsock1.Bind
    Winsock1.RemoteHost = Text2.Text
    Winsock1.RemotePort = 9999
    receivedLen = 0
    Open Text1.Text For Binary As #1
    ifReady = True
    cacheStart = 1
    cacheEnd = cacheStart + UBound(cache)
    Text1.Enabled = False
    Text2.Enabled = False
    Command1.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    ifReady = False
    Dim cacheLen As Long
    cacheLen = 1024 '1K
    cacheLen = cacheLen * 1024 '1M
    cacheLen = cacheLen * 2 '2M of cache
    ReDim cache(cacheLen - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'write the data still in buffer
    ReDim Preserve cache(receivedLen - cacheStart)
    Put #1, cacheStart, cache 'write file
    Close #1
End Sub

Private Sub Timer1_Timer()
    Form1.Caption = Int(receivedLen / 1024) & "KB received"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim dataBytes() As Byte
    Dim pos As Long
    Dim packageLen As Integer
    Dim i As Integer
    Dim tempPos As Long
    
    'read data
    Winsock1.GetData dataBytes, vbArray + vbByte
    
    'read header part II
    packageLen = dataBytes(4) * 256 + dataBytes(5)
    
    'check the length of package,throw bad ones
    If packageLen <> UBound(dataBytes) + 1 Then Exit Sub
    
    'read header part I
    pos = 0
    For i = 0 To 3
        pos = pos * 256 + dataBytes(i)
    Next i
    
    'not expected yet,throw away
    If pos > receivedLen + 1 Then
        Exit Sub
    End If
    
    'send replay
    Winsock1.SendData pos
    
    'the package already received.
    'it seems that the sender didn't got the replay ack correctly
    'so just send another ack but ignore the data
    If pos < receivedLen + 1 Then
        Exit Sub
    End If
    
    'the header must cut
    packageLen = packageLen - 6
    
    'cache work start
    If receivedLen + packageLen > cacheEnd Then 'cache full
        Put #1, cacheStart, cache 'write file
        'change cache pointer
        cacheStart = cacheEnd + 1
        cacheEnd = cacheStart + UBound(cache)
    End If
    
    'write data to cache
    tempPos = pos - cacheStart
    For i = 0 To packageLen - 1
        cache(tempPos + i) = dataBytes(i + 6)
    Next i 'cache work end
    
    receivedLen = receivedLen + packageLen 'modify the length
End Sub

