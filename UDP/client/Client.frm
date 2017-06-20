VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Client 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "4K"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1K"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   1080
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "127.0.0.1"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "c:\test.rar"
      Top             =   720
      Width           =   4455
   End
   Begin VB.Timer Speed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   120
   End
   Begin VB.CommandButton SendButton 
      Caption         =   "Send"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Sender 
      Left            =   840
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer TimeOut 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label3 
      Caption         =   "server IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fileLen As Long 'length of the file
Dim sendedLen As Long 'length of data sended
Dim sendingLen As Long 'length of data sending but not checked
'packagelen no more than 8K-6,it should be integer,
'but somewhere in the program it will overflow temply
'so it was dimed as long
Dim packageLen As Long
Dim sendWindow As Integer 'packages send at on time
Dim windowLimit As Integer 'the increase limit of sendwindow
Dim sendLimit As Long 'the edge of the send windows
Dim ifSending As Boolean
Dim ackCount As Integer 'count the arrival of ack
Dim lastPos As Long 'to calculate the transfer speed
Dim cache() As Byte 'disk cache
Dim cacheStart As Long 'start of the cache,the position in the file
Dim cacheEnd As Long 'end of cache

Private Sub Form_Load()
    Sender.LocalPort = 9999
    Sender.Bind
    packageLen = 1024
    Dim cacheLen As Long
    'can't set the cachelen at one time it will overflow
    cacheLen = 1024 '1K
    cacheLen = cacheLen * 1024 '1M
    cacheLen = 2 * cacheLen '2M of cache
    ReDim cache(cacheLen - 1)
End Sub

Private Sub Option1_Click()
    packageLen = 1024
End Sub

Private Sub Option2_Click()
    packageLen = 4096
End Sub

Private Sub Sender_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim message As Long
    Sender.GetData message, vbLong
    
    If message <> sendedLen + 1 Then 'throw any other ack
        Exit Sub
    End If
    
    sendedLen = sendedLen + packageLen 'mark the sended part
    'Label1.Caption = sendedLen / 1024 & "KB/" & Int(fileLen / 1024) & "KB"
    If sendedLen >= fileLen Then
        TimeOut.Enabled = False 'send finish
        Speed.Enabled = False
        sendLimit = 0 'close the window
        Close #1 'close the file
        MsgBox "finished"
        Text1.Enabled = True
        SendButton.Enabled = True
        Exit Sub
    End If
        
    'change the window
    ackCount = ackCount + 1
    If ackCount >= sendWindow Then 'if the window should change
        ackCount = 0 'restart the count
        If sendWindow >= windowLimit Then
            sendWindow = sendWindow + 1
        Else 'did't reach the increase limit
            sendWindow = sendWindow * 2
        End If
        
        'in the local area network it's not easily timeout so the window might
        'reach 65536 or more(more than an integer can store,overflow),so it should be limited
        'and it makes no sence to make such a large window while the sending process
        'will be traped by this event at any time
        If sendWindow > 2048 Then sendWindow = 2048
        
        'reduce the timeout
        TimeOut.Interval = Int(TimeOut.Interval / 2)
        If TimeOut.Interval < 10 Then TimeOut.Interval = 10
    End If
        
    'change the sendlimit
    sendLimit = sendedLen + packageLen * sendWindow 'if the packagelen is integer,it overflow here
    If sendLimit > fileLen Then sendLimit = fileLen 'meet the end of file
        
    'timeout restart
    TimeOut.Enabled = False
    TimeOut.Enabled = True
        
    'if it is still sending when dataarrive
    If ifSending = False Then
        Call sendData 'send rest data
    End If
End Sub

Private Sub SendButton_Click()
    Text1.Enabled = False
    Text2.Enabled = False
    Sender.RemoteHost = Text2.Text
    Sender.RemotePort = 9998
    SendButton.Enabled = False
    sendFile Text1.Text
End Sub

Private Sub sendFile(ByVal fileName As String)
    'check if the file is existed
    If Dir(fileName) = "" Then Exit Sub
    
    Open fileName For Binary As #1
    fileLen = LOF(1)
    'no data
    If fileLen = 0 Then Exit Sub
    'notice that file pointer started with 1
    'array pointer started with 0
    sendedLen = 0
    sendingLen = 0
    sendWindow = 1
    windowLimit = 1024
    sendLimit = packageLen * sendWindow
    'check if the file is very small
    If sendLimit > fileLen Then sendLimit = fileLen
    ackCount = 0
    lastPos = 0
    cacheEnd = 0 'cache's empty
    TimeOut.Enabled = True
    Speed.Enabled = True
    Call sendData
End Sub

Private Sub sendData() 'send some of the data that left
    On Error Resume Next
    Dim dataBytes() As Byte 'array that to be send
    Dim tempLen As Long
    Dim i As Integer
    Dim tempPos As Long
    
    ifSending = True 'sending start
    Do While sendingLen < sendLimit 'send the data in the window
        'calculate the packagelen,templen=the length of real data in package
        tempLen = IIf((sendLimit - sendingLen) >= packageLen, packageLen, (sendLimit - sendingLen))
        
        'cache work start
        If sendingLen + tempLen > cacheEnd Then 'out of cache
            Get #1, sendingLen + 1, cache 'rebuffer
            'cachepointer change
            cacheStart = sendingLen + 1
            cacheEnd = cacheStart + UBound(cache) 'no +1
        End If
        
        'before the cache,this may happen when timeout
        If sendingLen + 1 < cacheStart Then
            ReDim dataBytes(tempLen - 1)
            Get #1, sendingLen + 1, dataBytes 'read file,no cache to use
            
            ReDim Preserve dataBytes(tempLen + 5) 'make the package bigger
            'make the room for the header
            For i = tempLen - 1 To 0 Step -1
                dataBytes(i + 6) = dataBytes(i)
            Next i
        Else 'use cache
            'make room for data and header
            ReDim dataBytes(tempLen + 5)
            tempPos = sendingLen + 1 - cacheStart
            'copy data form cache
            For i = 0 To tempLen - 1
                dataBytes(i + 6) = cache(tempPos + i)
            Next i
        End If 'cache work end
        
        'make the header
        tempLen = sendingLen + 1
        'the first part shows the position in the file
        For i = 3 To 0 Step -1
            'divide the long data into bytes
            dataBytes(i) = tempLen Mod 256
            tempLen = Int(tempLen / 256)
        Next i
        
        tempLen = UBound(dataBytes) + 1 'total len of the package
        'record the length of the package in the header
        dataBytes(5) = tempLen Mod 256
        dataBytes(4) = Int(tempLen / 256)
                
        'send data
        Sender.sendData dataBytes
        sendingLen = sendingLen + tempLen - 6 'reset sendinglen

        'share the cpu time with other threads
        DoEvents
    Loop 'send another package in the window
    
    ifSending = False
End Sub

Private Sub TimeOut_Timer()
    windowLimit = Int(sendWindow / 2)
    If windowLimit = 0 Then windowLimit = 1
    sendWindow = 1
    ackCount = 0
    sendingLen = sendedLen 'reset the sign
    
    sendLimit = sendedLen + packageLen * sendWindow
    If sendLimit > fileLen Then sendLimit = fileLen 'meet the end of file
    
    TimeOut.Interval = TimeOut.Interval * 2 'double the timeout
    'don't wast time in waiting,just resend,but 3 sec is needed
    If TimeOut.Interval > 3200 Then TimeOut.Interval = 3200
    
    If ifSending = False Then
        Call sendData 'resend the data for timeout
    End If
End Sub

Private Sub Speed_Timer()
    Label2.Caption = (sendedLen - lastPos) / 1024 & "KB/s"
    Label1.Caption = sendedLen / 1024 & "KB/" & Int(fileLen / 1024) & "KB"
    lastPos = sendedLen
End Sub
