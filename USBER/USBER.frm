VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USBer"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2985
      ItemData        =   "USBER.frx":0000
      Left            =   3000
      List            =   "USBER.frx":0002
      TabIndex        =   6
      Top             =   0
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Run"
      Height          =   300
      Left            =   1920
      TabIndex        =   4
      Top             =   2700
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2700
      Width           =   1815
   End
   Begin MSWinsockLib.Winsock UDP2 
      Left            =   1560
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   840
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Default         =   -1  'True
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   1815
      ItemData        =   "USBER.frx":0004
      Left            =   0
      List            =   "USBER.frx":0006
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin MSWinsockLib.Winsock UDP 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim lastIndex As Integer

Private Sub Command1_Click()
    On Error Resume Next
    Form1.Caption = "Scaning..."
    Timer1.Enabled = True
    Text2.Locked = True
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    List1.Clear
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    If Text1.Text = "" Then
        Exit Sub
    End If
    
    UDP2.RemoteHost = Text1.Text
    UDP2.SendData "done"
    
    Dim i As Integer
    For i = 0 To List2.ListCount - 1
        If ipValue(List2.List(i)) > ipValue(Text1.Text) Then
            found = True
            Exit For
        End If
    Next i
    
    List2.AddItem Text1.Text, i
    List1.RemoveItem lastIndex
    List2.Selected(i) = True
    Text1.Text = ""
    Form1.Caption = "USBer  " & List1.ListCount & " | " & List2.ListCount
End Sub

Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then End
    App.TaskVisible = False

    UDP.LocalPort = 9885
    UDP.RemotePort = 9884
    UDP.Bind
    UDP2.RemotePort = 9883
    
    Dim AddHeader As String
    AddHeader = UDP.LocalIP
    Do While Right(AddHeader, 1) <> "." And Len(AddHeader) > 0
        AddHeader = Left(AddHeader, Len(AddHeader) - 1)
    Loop
    
    AddHeader = Left(AddHeader, Len(AddHeader) - 1)
    Do While Right(AddHeader, 1) <> "." And Len(AddHeader) > 0
        AddHeader = Left(AddHeader, Len(AddHeader) - 1)
    Loop
    Text2.Text = Left(AddHeader, Len(AddHeader) - 1)
    
    Dim ip As String
    Open "tftp.txt" For Input As #1
    If Err Then
    Else
        Do While Not EOF(1)
            Line Input #1, ip
            List2.AddItem ip
        Loop
    End If
    Close #1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Open "tftp.txt" For Output As #1
        Dim i As Integer
        For i = 0 To List2.ListCount - 1
            Print #1, List2.List(i)
        Next i
    Close #1
End Sub

Private Sub List1_DblClick()
    On Error Resume Next
    UDP2.RemoteHost = List1.Text
    UDP2.SendData UDP2.LocalIP
    Text1.Text = List1.Text
    lastIndex = List1.ListIndex
End Sub

Private Sub List2_DblClick()
    On Error Resume Next
    List1.AddItem List2.Text
    List2.RemoveItem List2.ListIndex
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Timer1.Enabled = False
    Dim i As Integer
    Dim j As Integer
    For i = 0 To 255
    For j = 1 To 254
        UDP.RemoteHost = Text2.Text & "." & i & "." & j
        UDP.SendData UDP.LocalIP
        DoEvents
    Next j
    Sleep 100
    Next i
    Form1.Caption = "USBer"
    Text2.Locked = False
End Sub

Private Sub UDP_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim remoteIP As String
    Dim found As Boolean
    Dim i As Integer
    
    UDP.GetData remoteIP
    
    If remoteIP = "" Then Exit Sub
    
    found = False
    
    For i = 0 To List2.ListCount - 1
        If List2.List(i) = remoteIP Then
            found = True
            Exit For
        End If
    Next i

    If found = False Then
        For i = 0 To List1.ListCount - 1
            If List1.List(i) = remoteIP Then
                found = True
                Exit For
            End If
        Next i
    End If
    
    If found = False And remoteIP <> "" Then List1.AddItem remoteIP
End Sub

Private Function ipValue(ByVal ip As String) As Double
    On Error Resume Next
    Dim i As Integer
    Dim tempValue As Double
    tempValue = 0
    Dim j As Integer
    
    For j = 0 To 2
        i = InStr(1, ip, ".")
        tempValue = tempValue * 256
        tempValue = tempValue + Val(Left(ip, i - 1))
        ip = Right(ip, Len(ip) - i)
    Next j
    
    tempValue = tempValue * 256
    tempValue = tempValue + Val(ip)
    ipValue = tempValue
End Function
