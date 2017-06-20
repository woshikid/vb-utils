VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocketPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'
'Name.......... CSocketPlus
'File.......... CSocketPlus.cls
'Version....... 1.2
'Dependencies.. Requires modSocketPlus.bas code module
'Description... Winsock api implementation class
'Author........ Emiliano Scavuzzo <anshoku@yahoo.com>
'Date.......... February, 4th 2005

'Copyright (c) 2004 by Emiliano Scavuzzo
'Rosario, Argentina
'
'Based on CSocket by Oleg Gdalevich
'Subclassing based on WinSubHook2 by Paul Caton <Paul_Caton@hotmail.com>
'
'********************************************************************************

Option Explicit

'==============================================================================
'API FUNCTIONS
'==============================================================================

Private Declare Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function api_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare Function api_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Private Declare Function api_htons Lib "ws2_32.dll" Alias "htons" (ByVal hostshort As Integer) As Integer
Private Declare Function api_ntohs Lib "ws2_32.dll" Alias "ntohs" (ByVal netshort As Integer) As Integer
Private Declare Function api_connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare Function api_gethostname Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function api_gethostbyname Lib "ws2_32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_getsockname Lib "ws2_32.dll" Alias "getsockname" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_getpeername Lib "ws2_32.dll" Alias "getpeername" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Private Declare Function api_inet_addr Lib "ws2_32.dll" Alias "inet_addr" (ByVal cp As String) As Long
Private Declare Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_sendto Lib "ws2_32.dll" Alias "sendto" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef toaddr As sockaddr_in, ByVal tolen As Long) As Long
Private Declare Function api_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
Private Declare Function api_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function api_recvfrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef from As sockaddr_in, ByRef fromlen As Long) As Long
Private Declare Function api_WSACancelAsyncRequest Lib "ws2_32.dll" Alias "WSACancelAsyncRequest" (ByVal hAsyncTaskHandle As Long) As Long
Private Declare Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Private Declare Function api_inet_ntoa Lib "ws2_32.dll" Alias "inet_ntoa" (ByVal inn As Long) As Long
Private Declare Function api_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As Long, ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare Function api_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
'Private Declare Function api_gethostbyaddr Lib "ws2_32.dll" Alias "gethostbyaddr" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long

'==============================================================================
'CONSTANTS
'==============================================================================
Public Enum SockState
    sckClosed = 0
    sckOpen
    sckListening
    sckConnectionPending
    sckResolvingHost
    sckHostResolved
    sckConnecting
    sckConnected
    sckClosing
    sckError
End Enum

Private Const SOMAXCONN As Long = 5

Public Enum ProtocolConstants
    sckTCPProtocol = 0
End Enum

Private Const MSG_PEEK  As Long = &H2

'==============================================================================
'EVENTS
'==============================================================================

Public Event CloseSck(ByVal Index As Variant)
Public Event Connect(ByVal Index As Variant)
Public Event ConnectionRequest(ByVal Index As Variant, ByVal requestID As Long)
Public Event DataArrival(ByVal Index As Variant, ByVal bytesTotal As Long)
Public Event Error(ByVal Index As Variant, ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendComplete(ByVal Index As Variant)
Public Event SendProgress(ByVal Index As Variant, ByVal bytesSent As Long, ByVal bytesRemaining As Long)

'==============================================================================
'MEMBER VARIABLES
'==============================================================================

Private Type typSocket
    lngSocketHandle           As Long                 'socket handle
    enmState                  As SockState            'socket state
    strTag                    As String               'tag
    strRemoteHost             As String               'remote host
    lngRemotePort             As Long                 'remote port
    strRemoteHostIP           As String               'remote host ip
    lngLocalPort              As Long                 'local port
    lngLocalPortBind          As Long                 'temporary local port
    strLocalIP                As String               'local IP
    enmProtocol               As ProtocolConstants    'protocol used (TCP / UDP)
    lngMemoryPointer          As Long                 'memory pointer used as buffer when resolving host
    lngMemoryHandle           As Long                 'buffer memory handle
    strSendBuffer             As String               'local incoming buffer
    strRecvBuffer             As String               'local outgoing buffer
    lngSendBufferLen          As Long                 'winsock buffer size for sends
    lngRecvBufferLen          As Long                 'winsock buffer size for receives
    colWaitingResolutions     As Collection           'hosts waiting to be resolved by the system ("R" & async handle - destination)
    varUserIndex              As Variant              'user index, set by user (long or string)
End Type

Private m_colString_Index           As Collection   '(index: "I" & string index, item: socket index)
Private m_colUserIndex_Index        As Collection   '(index: "I" & user index, item: socket index)
Private m_colSockHandle_Index       As Collection   '(index: "S" & socket handle, item: socket index)
Private m_colWaitingResolutionsS    As Collection   '(index: "R" & async handle, item: socket index)

Private m_Socket() As typSocket 'socket array
Private m_EmptyIndex() As Long  'socket array empty indexes


Private m_blnAcceptClass As Boolean 'if True then this is an Accept socket class


'  ****  WARNING WARNING WARNING WARNING ******
'This sub MUST be the first on the class. DO NOT attempt
'to change it's location or the code will CRASH.
'This sub receives system messages from our WndProc.
Public Sub WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
On Error Resume Next
Select Case uMsg

Case RESOLVE_MESSAGE
    
    PostResolution wParam, modSocketPlus.HiWord(lParam)
    
Case SOCKET_MESSAGE
    
    PostSocket wParam, LoWord(lParam), modSocketPlus.HiWord(lParam)
    
End Select
End Sub

Private Sub Class_Initialize()
On Error Resume Next

Set m_colSockHandle_Index = New Collection
Set m_colUserIndex_Index = New Collection
Set m_colString_Index = New Collection
Set m_colWaitingResolutionsS = New Collection

ReDim m_Socket(0) As typSocket
ReDim m_EmptyIndex(0) As Long

'initiate processes and winsock service
modSocketPlus.InitiateProcesses
End Sub

Private Sub Class_Terminate()
On Error Resume Next
Dim lngIndex As Long

'clean hostname resolution system
For lngIndex = 1 To UBound(m_Socket)
    CleanResolutionSystem lngIndex
Next

'destroy socket if it exists
If Not m_blnAcceptClass Then
    For lngIndex = 1 To UBound(m_Socket)
        DestroySocket lngIndex
    Next
End If

Erase m_Socket
Erase m_EmptyIndex

'clean processes and finish winsock service
modSocketPlus.FinalizeProcesses

Set m_colSockHandle_Index = Nothing
Set m_colUserIndex_Index = Nothing
Set m_colString_Index = Nothing
Set m_colWaitingResolutionsS = Nothing

End Sub

'==============================================================================
'PROPERTIES
'==============================================================================

Public Property Get RemotePort(Index As Variant) As Long
On Error Resume Next
RemotePort = m_Socket(GetSocketIndex(Index)).lngRemotePort
End Property

Public Property Let RemotePort(Index As Variant, ByVal lngPort As Long)
On Error Resume Next
With m_Socket(GetSocketIndex(Index))
    If .enmState <> sckClosed Then
        ErrRaise sckInvalidOp, "CSocketPlus.RemotePort", "Invalid operation at current state"
        Exit Property
    End If

    If lngPort < 0 Or lngPort > 65535 Then
        ErrRaise sckInvalidArg, "CSocketPlus.RemotePort", "The argument passed to a function was not in the correct format or in the specified range."
        Exit Property
    Else
    .lngRemotePort = lngPort
    End If
End With
End Property

Public Property Get RemoteHost(Index As Variant) As String
On Error Resume Next
RemoteHost = m_Socket(GetSocketIndex(Index)).strRemoteHost
End Property

Public Property Let RemoteHost(Index As Variant, ByVal strHost As String)
On Error Resume Next
With m_Socket(GetSocketIndex(Index))
    If .enmState <> sckClosed Then
        ErrRaise sckInvalidOp, "CSocketPlus.RemoteHost", "Invalid operation at current state"
        Exit Property
    End If

    .strRemoteHost = strHost
End With
End Property

Public Property Get RemoteHostIP(Index As Variant) As String
On Error Resume Next
RemoteHostIP = m_Socket(GetSocketIndex(Index)).strRemoteHostIP
End Property

Public Property Get LocalPort(Index As Variant) As Long
On Error Resume Next
With m_Socket(GetSocketIndex(Index))
    If .lngLocalPortBind = 0 Then
        LocalPort = .lngLocalPort
    Else
        LocalPort = .lngLocalPortBind
    End If
End With
End Property

Public Property Let LocalPort(Index As Variant, ByVal lngPort As Long)
On Error Resume Next
With m_Socket(GetSocketIndex(Index))
    If .enmState <> sckClosed Then
        ErrRaise sckInvalidOp, "CSocketPlus.LocalPort", "Invalid operation at current state"
        Exit Property
    End If
    If lngPort < 0 Or lngPort > 65535 Then
        ErrRaise sckInvalidArg, "CSocketPlus.LocalPort", "The argument passed to a function was not in the correct format or in the specified range."
        Exit Property
    Else
        .lngLocalPort = lngPort
    End If
End With
End Property

Public Property Get State(Index As Variant) As SockState
On Error Resume Next
State = m_Socket(GetSocketIndex(Index)).enmState
End Property

Public Property Get LocalHostName() As String
On Error Resume Next
LocalHostName = GetLocalHostName
End Property

Public Property Get LocalIP(Index As Variant) As String
On Error Resume Next
With m_Socket(GetSocketIndex(Index))
    If .enmState = sckConnected Then
        LocalIP = .strLocalIP
    Else
        LocalIP = GetLocalIP
    End If
End With
End Property

Public Property Get BytesReceived(Index As Variant) As Long
On Error Resume Next
Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)
BytesReceived = LenB(m_Socket(lngIndex).strRecvBuffer)
End Property

Public Property Get SocketHandle(Index As Variant) As Long
On Error Resume Next
SocketHandle = m_Socket(GetSocketIndex(Index)).lngSocketHandle
End Property

Public Property Get Tag(Index As Variant) As String
On Error Resume Next
Tag = m_Socket(GetSocketIndex(Index)).strTag
End Property

Public Property Let Tag(Index As Variant, ByVal strTag As String)
On Error Resume Next
m_Socket(GetSocketIndex(Index)).strTag = strTag
End Property

Public Property Get Protocol(Index As Variant) As ProtocolConstants
On Error Resume Next
Protocol = m_Socket(GetSocketIndex(Index)).enmProtocol
End Property

Public Property Let Protocol(Index As Variant, ByVal enmProtocol As ProtocolConstants)
On Error Resume Next
With m_Socket(GetSocketIndex(Index))
    If .enmState <> sckClosed Then
        ErrRaise sckInvalidOp, "CSocketPlus.Protocol", "Invalid operation at current state"
        Exit Property
    Else
        .enmProtocol = enmProtocol
    End If
End With
End Property


Public Sub Accept(Index As Variant, requestID As Long)
On Error Resume Next

Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)


If m_Socket(lngIndex).enmState <> sckClosed Then
    ErrRaise sckInvalidOp, "CSocketPlus.Accept", "Invalid operation at current state"
    Exit Sub
End If

m_Socket(lngIndex).lngSocketHandle = requestID
m_Socket(lngIndex).enmProtocol = sckTCPProtocol
ProcessOptions lngIndex

If Not modSocketPlus.IsAcceptRegistered(requestID) Then
    If IsSocketRegistered(requestID) Then
        m_Socket(lngIndex).lngSocketHandle = INVALID_SOCKET
        m_Socket(lngIndex).lngRecvBufferLen = 0
        m_Socket(lngIndex).lngSendBufferLen = 0
        ErrRaise sckBadState, "CSocketPlus.Accept", "Wrong protocol or connection state for the requested transaction or request"
        Exit Sub
    Else
        m_blnAcceptClass = True
        m_Socket(lngIndex).enmState = sckConnected: Debug.Print "STATE: sckConnected"
        GetLocalInfo m_Socket(lngIndex).lngSocketHandle, m_Socket(lngIndex).lngLocalPortBind, m_Socket(lngIndex).strLocalIP
        modSocketPlus.RegisterSocket m_Socket(lngIndex).lngSocketHandle, ObjPtr(Me), False
        Exit Sub
    End If
End If

Dim clsSocket As CSocketPlus
Set clsSocket = modSocketPlus.GetAcceptClass(requestID)
modSocketPlus.UnregisterAccept requestID

GetLocalInfo m_Socket(lngIndex).lngSocketHandle, m_Socket(lngIndex).lngLocalPortBind, m_Socket(lngIndex).strLocalIP
GetRemoteInfo m_Socket(lngIndex).lngSocketHandle, m_Socket(lngIndex).lngRemotePort, m_Socket(lngIndex).strRemoteHostIP, m_Socket(lngIndex).strRemoteHost
    
m_Socket(lngIndex).enmState = sckConnected: Debug.Print "STATE: sckConnected"

If clsSocket.BytesReceived("Temporal_Socket") > 0 Then
    clsSocket.GetData "Temporal_Socket", m_Socket(lngIndex).strRecvBuffer
End If
    
m_colSockHandle_Index.Add lngIndex, "S" & requestID
modSocketPlus.Subclass_ChangeOwner requestID, ObjPtr(Me)

If LenB(m_Socket(lngIndex).strRecvBuffer) > 0 Then RaiseEvent DataArrival(Index, LenB(m_Socket(lngIndex).strRecvBuffer))

If clsSocket.State("Temporal_Socket") = sckClosing Then
    m_Socket(lngIndex).enmState = sckClosing: Debug.Print "STATE: sckClosing"
    RaiseEvent CloseSck(Index)
End If

Set clsSocket = Nothing

End Sub


'Create a new element in the socket array.
'Returns de index used. This can be used in case the
'user didn't pass the Index argument.
Public Function ArrayAdd(Optional Index As Variant) As Variant

On Error GoTo Error_Handler

Dim lngIndex As Long
lngIndex = GetFreeSocketIndex

m_Socket(lngIndex).lngSocketHandle = INVALID_SOCKET
Set m_Socket(lngIndex).colWaitingResolutions = New Collection
    
If IsMissing(Index) Then
    Dim Cont As Long
    
    Do
        Cont = Cont + 1
    Loop Until Not IsUserIndexRegistered(Cont)
        
    m_colUserIndex_Index.Add lngIndex, "I" & Cont
    
    m_Socket(lngIndex).varUserIndex = Cont
    ArrayAdd = Cont
    Debug.Print "OK Added '" & Cont & "' to the socket array"
Else
    m_Socket(lngIndex).varUserIndex = Index
    ArrayAdd = Index
    Select Case VarType(Index)
    Case vbString
        m_colString_Index.Add lngIndex, "I" & Index
    Case vbInteger To vbLong
        m_colUserIndex_Index.Add lngIndex, "I" & Index
    Case Else
        ErrRaise sckWrongIndex, "CSocketPlus.ArrayAdd", "Unsupported index type."
        Exit Function
    End Select
        
    Debug.Print "OK Added '" & Index & "' to the socket array"
End If


Exit Function

Error_Handler:
    Dim lngErrorCode As Long
    Dim strErrorDesc As String
    
    lngErrorCode = Err.Number
    strErrorDesc = Err.Description
    
    If lngErrorCode = 457 Then
        lngErrorCode = sckUsedIndex
        strErrorDesc = "This index is already associated with an element of this array."
    End If
    
    CleanSocketArray lngIndex
    ErrRaise lngErrorCode, "CSocketPlus.ArrayAdd", strErrorDesc
End Function

'Remove an element in the socket array
Public Function ArrayRemove(Index As Variant) As Variant
On Error Resume Next
Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)

CloseSck Index
    
Select Case VarType(Index)
Case vbString
    m_colString_Index.Remove "I" & Index
Case vbInteger To vbLong
    m_colUserIndex_Index.Remove "I" & Index
Case Else
    ErrRaise sckWrongIndex, "CSocketPlus.ArrayRemove", "Unsupported index type."
    Exit Function
End Select

CleanSocketArray lngIndex

ArrayRemove = Index

Debug.Print "OK Removed '" & Index & "' from the socket array"
End Function

'Returns TRUE if the index passed is in use by the array
Public Function ArrayIndexInUse(Index As Variant) As Boolean
On Error GoTo Error_Handler

GetSocketIndex Index
ArrayIndexInUse = True

Exit Function
Error_Handler:

    Dim lngErrorCode As Long
    Dim strErrorDesc As String
    
    lngErrorCode = Err.Number
    strErrorDesc = Err.Description
    
    Const strSource As String = "CSocketPlus.ArrayIndexInUse"
    
    If Err.Number = sckWrongIndex Then
        ErrRaise sckWrongIndex, strSource, "Unsupported index type."
    ElseIf Err.Number = sckMissingIndex Then
        ArrayIndexInUse = False
    Else
        ErrRaise lngErrorCode, strSource, strErrorDesc
    End If

End Function


'Returns amount of 'sockets' created by user.
Public Function ArrayCount() As Long
On Error Resume Next
ArrayCount = m_colString_Index.Count + m_colUserIndex_Index.Count
End Function

Public Sub Bind(Index As Variant, Optional LocalPort As Variant, Optional LocalIP As Variant)
On Error Resume Next
Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)

If m_Socket(lngIndex).enmState <> sckClosed Then
    ErrRaise sckInvalidOp, "CSocketPlus.Bind", "Invalid operation at current state"
    Exit Sub
End If

If BindInternal(lngIndex, LocalPort, LocalIP) Then
    m_Socket(lngIndex).enmState = sckOpen: Debug.Print "STATE: sckOpen"
End If
End Sub

Public Sub CloseSck(Index As Variant)
On Error Resume Next
Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)

With m_Socket(lngIndex)
    If .lngMemoryHandle = INVALID_SOCKET Then Exit Sub

    .enmState = sckClosing: Debug.Print "STATE: sckClosing"
    CleanResolutionSystem lngIndex
    DestroySocket GetSocketIndex(Index)

    .lngLocalPortBind = 0
    .strRemoteHostIP = ""
    .strRecvBuffer = ""
    .strSendBuffer = ""
    .lngSendBufferLen = 0
    .lngRecvBufferLen = 0
    .enmState = sckClosed: Debug.Print "STATE: sckClosed"
End With

End Sub

'Tries to connect to RemoteHost if it was passed, or uses
'm_strRemoteHost instead. If it is a hostname tries to
'resolve it first.
Public Sub Connect(Index As Variant, Optional RemoteHost As Variant, Optional RemotePort As Variant)
On Error Resume Next
Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)
If m_Socket(lngIndex).enmState <> sckClosed Then
    ErrRaise sckInvalidOp, "CSocketPlus.Connect", "Invalid operation at current state"
    Exit Sub
End If

If Not IsMissing(RemoteHost) Then
        m_Socket(lngIndex).strRemoteHost = CStr(RemoteHost)
End If

'for some reason we get a GPF if we try to
'resolve a null string, so we replace it with
'an empty string
If m_Socket(lngIndex).strRemoteHost = vbNullString Then
    m_Socket(lngIndex).strRemoteHost = ""
End If

'check if RemotePort is a number between 1 and 65535
If Not IsMissing(RemotePort) Then
    If IsNumeric(RemotePort) Then
        If CLng(RemotePort) > 65535 Or CLng(RemotePort) < 1 Then
            ErrRaise sckInvalidArg, "CSocketPlus.Connect", "The argument passed to a function was not in the correct format or in the specified range."
            Exit Sub
        Else
            m_Socket(lngIndex).lngRemotePort = CLng(RemotePort)
        End If
    Else
        ErrRaise sckUnsupported, "CSocketPlus.Connect", "Unsupported variant type."
        Exit Sub
    End If
End If

'create a socket if there isn't one yet
If Not SocketExists(lngIndex) Then Exit Sub

'Here we bind the socket
If Not BindInternal(lngIndex) Then Exit Sub

'try to get a 32 bits long that is used to identify a host
Dim lngAddress As Long
lngAddress = ResolveIfHostname(lngIndex, m_Socket(lngIndex).strRemoteHost)

'We've got two options here:
'1) m_strRemoteHost was an IP, so a resolution wasn't
'   necessary, and now lngAddress is a 32 bits long and
'   we proceed to connect.
'2) m_strRemoteHost was a hostname, so a resolution was
'   necessary and it's taking place right now. We leave
'   silently.

If lngAddress <> vbNull Then
    ConnectToIP lngIndex, lngAddress, 0
End If

End Sub

Public Sub GetData(Index As Variant, ByRef data As Variant)
On Error Resume Next

Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)

If m_Socket(lngIndex).enmState <> sckConnected And Not m_blnAcceptClass Then
    ErrRaise sckBadState, "CSocketPlus.GetData", "Wrong protocol or connection state for the requested transaction or request"
    Exit Sub
End If

If VarType(data) = vbArray + vbByte Then
    Dim tempArray() As Byte
    tempArray = m_Socket(lngIndex).strRecvBuffer
    data = tempArray
ElseIf VarType(data) = vbString Then
    data = m_Socket(lngIndex).strRecvBuffer
End If

m_Socket(lngIndex).strRecvBuffer = vbNullString

End Sub

Public Sub Listen(Index As Variant)
On Error Resume Next
Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)

If m_Socket(lngIndex).enmState <> sckClosed And m_Socket(lngIndex).enmState <> sckOpen Then
    ErrRaise sckInvalidOp, "CSocketPlus.Listen", "Invalid operation at current state"
    Exit Sub
End If

If Not SocketExists(lngIndex) Then Exit Sub
If Not BindInternal(lngIndex) Then Exit Sub

Dim lngResult As Long

lngResult = api_listen(m_Socket(lngIndex).lngSocketHandle, SOMAXCONN)

If lngResult = SOCKET_ERROR Then
    Dim lngErrorCode As Long
    lngErrorCode = Err.LastDllError
    ErrRaise lngErrorCode, "CSocketPlus.Listen", GetErrorDescription(lngErrorCode)
    Exit Sub
Else
    m_Socket(lngIndex).enmState = sckListening: Debug.Print "STATE: sckListening"
End If

End Sub

Public Sub SendData(Index As Variant, data As Variant)
On Error Resume Next
Dim lngIndex As Long
lngIndex = GetSocketIndex(Index)

If m_Socket(lngIndex).enmState <> sckConnected Then
    ErrRaise sckBadState, "CSocketPlus.SendData", "Wrong protocol or connection state for the requested transaction or request"
    Exit Sub
End If

'if there's already something in the buffer that means we are
'already sending data, so we put the new data in the buffer
'and exit silently
If LenB(m_Socket(lngIndex).strSendBuffer) > 0 Then
    m_Socket(lngIndex).strSendBuffer = m_Socket(lngIndex).strSendBuffer + CStr(data)
    Exit Sub
Else
    m_Socket(lngIndex).strSendBuffer = CStr(data)
End If

'send the data
SendBufferedDataTCP lngIndex

End Sub

'Destroys the socket if it exists and unregisters it
'from control list.
Private Sub DestroySocket(ByVal Index As Long)
On Error Resume Next
If Not m_Socket(Index).lngSocketHandle = INVALID_SOCKET Then

    Dim lngResult As Long
    
    lngResult = api_closesocket(m_Socket(Index).lngSocketHandle) ' m_lngSocketHandle)
    
    If lngResult = SOCKET_ERROR Then
        
        m_Socket(Index).enmState = sckError: Debug.Print "STATE: sckError"
        Dim lngErrorCode As Long
        lngErrorCode = Err.LastDllError
        ErrRaise lngErrorCode, "CSocketPlus.DestroySocket", GetErrorDescription(lngErrorCode)
        Exit Sub
    
    Else
        
        Debug.Print "OK Destroyed socket " & m_Socket(Index).lngSocketHandle
        modSocketPlus.UnregisterSocket m_Socket(Index).lngSocketHandle
        m_colSockHandle_Index.Remove "S" & m_Socket(Index).lngSocketHandle
        m_Socket(Index).lngSocketHandle = INVALID_SOCKET
    End If

End If
End Sub

'Tries to create a socket if there isn't one yet and registers
'it to the control list.
'Returns TRUE if it has success
Private Function SocketExists(ByVal lngIndex As Long) As Boolean
On Error Resume Next
SocketExists = True
Dim lngResult As Long
Dim lngErrorCode As Long

'check if there is a socket already
If m_Socket(lngIndex).lngSocketHandle = INVALID_SOCKET Then
    
    lngResult = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    
    If lngResult = INVALID_SOCKET Then
        
        m_Socket(lngIndex).enmState = sckError: Debug.Print "STATE: sckError"
        Debug.Print "ERROR trying to create socket"
        SocketExists = False
        lngErrorCode = Err.LastDllError
        Dim blnCancelDisplay As Boolean
        blnCancelDisplay = True
        RaiseEvent Error(m_Socket(lngIndex).varUserIndex, lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocketPlus.SocketExists", "", 0, blnCancelDisplay)
        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocketPlus.SocketExists"
    Else
        
        Debug.Print "OK Created socket: " & lngResult
        m_colSockHandle_Index.Add lngIndex, "S" & lngResult
        m_Socket(lngIndex).lngSocketHandle = lngResult
        'set and get some socket options
        ProcessOptions lngIndex
        SocketExists = modSocketPlus.RegisterSocket(m_Socket(lngIndex).lngSocketHandle, ObjPtr(Me), True)
    
    End If
End If
End Function

'When the system resolves a hostname in asynchronous way we
'call this function to decide what to do with the result.
Private Sub PostResolution(ByVal lngAsynHandle As Long, ByVal lngErrorCode As Long)
On Error Resume Next
Dim lngIndex As Long

'find out which socket requested the host resolution
lngIndex = m_colWaitingResolutionsS.Item("R" & lngAsynHandle)
m_colWaitingResolutionsS.Remove ("R" & lngAsynHandle)

'erase that record from the collection since we won't need it any longer
m_Socket(lngIndex).colWaitingResolutions.Remove "R" & lngAsynHandle
UnregisterResolution lngAsynHandle

If m_Socket(lngIndex).enmState <> sckResolvingHost Then Exit Sub

If lngErrorCode = 0 Then 'if there weren't errors trying to resolve the hostname
    
     m_Socket(lngIndex).enmState = sckHostResolved: Debug.Print "STATE: sckHostResolved"
    
    Dim udtHostent As HOSTENT
    Dim lngPtrToIP As Long
    Dim arrIpAddress(1 To 4) As Byte
    Dim lngRemoteHostAddress As Long
    Dim Count As Integer
    Dim strIpAddress As String
    
    api_CopyMemory udtHostent, ByVal m_Socket(lngIndex).lngMemoryPointer, LenB(udtHostent)
    api_CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
    api_CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
    api_CopyMemory lngRemoteHostAddress, ByVal lngPtrToIP, 4
    
    'free memory, won't need it any longer
    FreeMemory lngIndex
    
    'We turn the 32 bits long into a readable string.
    'Note: we don't need this string. I put this here just
    'in case you need it.
    For Count = 1 To 4
        strIpAddress = strIpAddress & arrIpAddress(Count) & "."
    Next
        
    strIpAddress = Left(strIpAddress, Len(strIpAddress) - 1)
    
    ConnectToIP lngIndex, lngRemoteHostAddress, 0

Else 'there were errors trying to resolve the hostname

    'free buffer memory
    FreeMemory lngIndex
 
    ConnectToIP lngIndex, vbNull, lngErrorCode
        
End If
End Sub

'This procedure is called by the WindowProc callback function.
'The lngEventID argument is an ID of the network event
'occurred for the socket. The lngErrorCode argument contains
'an error code only if an error was occurred during an
'asynchronous execution.
Private Sub PostSocket(ByVal lngSocketHandle As Long, ByVal lngEventID As Long, ByVal lngErrorCode As Long)
On Error Resume Next

Dim lngIndex As Long
'find out which socket the message belongs to
lngIndex = m_colSockHandle_Index.Item("S" & lngSocketHandle)

Dim varIndex As Variant
varIndex = m_Socket(lngIndex).varUserIndex

Dim blnCancelDisplay As Boolean

'handle any possible error
If lngErrorCode <> 0 Then
    m_Socket(lngIndex).enmState = sckError: Debug.Print "STATE: sckError"
    blnCancelDisplay = True
    RaiseEvent Error(varIndex, lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocketPlus.PostSocket", "", 0, blnCancelDisplay)
    If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocketPlus.PostSocket"
    Exit Sub
End If

Dim udtSockAddr As sockaddr_in
Dim lngResult As Long
Dim lngBytesReceived As Long

Select Case lngEventID

'======================================================================

Case FD_CONNECT

    'Arrival of this message means that the connection initiated by the call
    'of the connect Winsock API function was successfully established.

    Debug.Print "FD_CONNECT " & m_Socket(lngIndex).lngSocketHandle
    
    If m_Socket(lngIndex).enmState <> sckConnecting Then
        Debug.Print "WARNING: Omitting FD_CONNECT"
        Exit Sub
    End If
    
    'Get the local parameters
    GetLocalInfo m_Socket(lngIndex).lngSocketHandle, m_Socket(lngIndex).lngLocalPortBind, m_Socket(lngIndex).strLocalIP
    
    'Get the connection local end-point parameters
    GetRemoteInfo m_Socket(lngIndex).lngSocketHandle, m_Socket(lngIndex).lngRemotePort, m_Socket(lngIndex).strRemoteHostIP, m_Socket(lngIndex).strRemoteHost
    
    m_Socket(lngIndex).enmState = sckConnected: Debug.Print "STATE: sckConnected"
    RaiseEvent Connect(varIndex)

'======================================================================

Case FD_WRITE

    'This message means that the socket in a write-able
    'state, that is, buffer for outgoing data of the transport
    'service is empty and ready to receive data to send through
    'the network.
    
    Debug.Print "FD_WRITE " & m_Socket(lngIndex).lngSocketHandle
    
    If m_Socket(lngIndex).enmState <> sckConnected Then
        Debug.Print "WARNING: Omitting FD_WRITE"
        Exit Sub
    End If
    
    If LenB(m_Socket(lngIndex).strSendBuffer) > 0 Then
        SendBufferedDataTCP lngIndex
    End If
    
'======================================================================

Case FD_READ

    'Some data has arrived for this socket.

    Debug.Print "FD_READ " & m_Socket(lngIndex).lngSocketHandle
    
        
    If m_Socket(lngIndex).enmState <> sckConnected Then
        Debug.Print "WARNING: Omitting FD_READ"
        Exit Sub
    End If
    
    'Call the RecvDataToBuffer function that move arrived data
    'from the Winsock buffer to the local one and returns number
    'of bytes received.

    lngBytesReceived = RecvDataToBuffer(lngIndex)

    If lngBytesReceived > 0 Then
        RaiseEvent DataArrival(varIndex, LenB(m_Socket(lngIndex).strRecvBuffer))
    End If


    
    
'======================================================================

Case FD_ACCEPT

    'When the socket is in a listening state, arrival of this message
    'means that a connection request was received. Call the accept
    'Winsock API function in order to create a new socket for the
    'requested connection.
  
    Debug.Print "FD_ACCEPT " & m_Socket(lngIndex).lngSocketHandle
    If m_Socket(lngIndex).enmState <> sckListening Then
        Debug.Print "WARNING: Omitting FD_ACCEPT"
        Exit Sub
    End If
    
    lngResult = api_accept(m_Socket(lngIndex).lngSocketHandle, udtSockAddr, LenB(udtSockAddr))
    
    If lngResult = INVALID_SOCKET Then
        lngErrorCode = Err.LastDllError
        m_Socket(lngIndex).enmState = sckError: Debug.Print "STATE: sckError"
        blnCancelDisplay = True
        RaiseEvent Error(varIndex, lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocketPlus.PostSocket", "", 0, blnCancelDisplay)
        If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocketPlus.PostSocket"
    Else
        'We assign a temporal instance of CSocketPlus to
        'handle this new socket until user accepts (or not)
        'the new connection
        modSocketPlus.RegisterAccept lngResult
        
        'We change remote info before firing ConnectionRequest
        'event so the user can see which host is trying to
        'connect.

        Dim lngTempRP As Long
        Dim strTempRHIP As String
        Dim strTempRH As String
        lngTempRP = m_Socket(lngIndex).lngRemotePort
        strTempRHIP = m_Socket(lngIndex).strRemoteHostIP
        strTempRH = m_Socket(lngIndex).strRemoteHost
        
        GetRemoteInfo lngResult, m_Socket(lngIndex).lngRemotePort, m_Socket(lngIndex).strRemoteHostIP, m_Socket(lngIndex).strRemoteHost
        
        Debug.Print "OK Accepted socket: " & lngResult
        RaiseEvent ConnectionRequest(varIndex, lngResult)
        
        'we return original info
        If m_Socket(lngIndex).enmState = sckListening Then
            m_Socket(lngIndex).lngRemotePort = lngTempRP
            m_Socket(lngIndex).strRemoteHostIP = strTempRHIP
            m_Socket(lngIndex).strRemoteHost = strTempRH
        End If

        'This is very important. If the connection wasn't accepted
        'we must close the socket.
        If IsAcceptRegistered(lngResult) Then
            api_closesocket lngResult
            modSocketPlus.UnregisterSocket lngResult
            modSocketPlus.UnregisterAccept lngResult
            Debug.Print "OK Closed accepted socket: " & lngResult
        End If
    End If
    
'======================================================================
    
Case FD_CLOSE
    
    'This message means that the remote host is closing the conection
    
    Debug.Print "FD_CLOSE " & m_Socket(lngIndex).lngSocketHandle
    
    If m_Socket(lngIndex).enmState <> sckConnected Then
        Debug.Print "WARNING: Omitting FD_CLOSE"
        Exit Sub
    End If
    
    m_Socket(lngIndex).enmState = sckClosing: Debug.Print "STATE: sckClosing"
    RaiseEvent CloseSck(varIndex)
    
End Select
End Sub

'Connect to a given 32 bits long ip
Private Sub ConnectToIP(ByVal lngIndex As Long, ByVal lngRemoteHostAddress As Long, ByVal lngErrorCode As Long)
On Error Resume Next

Dim blnCancelDisplay As Boolean
Dim varIndex As Variant
varIndex = m_Socket(lngIndex).varUserIndex

'Check and handle errors
If lngErrorCode <> 0 Then
    m_Socket(lngIndex).enmState = sckError: Debug.Print "STATE: sckError"
    blnCancelDisplay = True
    RaiseEvent Error(varIndex, lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocketPlus.ConnectToIP", "", 0, blnCancelDisplay)
    If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocketPlus.ConnectToIP"
    Exit Sub
End If

Debug.Print "OK Connecting to: " + m_Socket(lngIndex).strRemoteHost + " " + m_Socket(lngIndex).strRemoteHostIP
m_Socket(lngIndex).enmState = sckConnecting: Debug.Print "STATE: sckConnecting"

Dim udtSockAddr As sockaddr_in
Dim lngResult As Long

'Build the sockaddr_in structure to pass it to the connect
'Winsock API function as an address of the remote host.
With udtSockAddr
    .sin_addr = lngRemoteHostAddress
    .sin_family = AF_INET
    .sin_port = api_htons(modSocketPlus.UnsignedToInteger(m_Socket(lngIndex).lngRemotePort))
End With

'Call the connect Winsock API function in order to establish connection.
lngResult = api_connect(m_Socket(lngIndex).lngSocketHandle, udtSockAddr, LenB(udtSockAddr))

'Check and handle errors
If lngResult = SOCKET_ERROR Then
    lngErrorCode = Err.LastDllError
    If lngErrorCode <> WSAEWOULDBLOCK Then
        If lngErrorCode = WSAEADDRNOTAVAIL Then
            ErrRaise WSAEADDRNOTAVAIL, "CSocketPlus.ConnectToIP", GetErrorDescription(WSAEADDRNOTAVAIL)
            Exit Sub
        Else
            m_Socket(lngIndex).enmState = sckError: Debug.Print "STATE: sckError"
            blnCancelDisplay = True
            RaiseEvent Error(varIndex, lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocketPlus.ConnectToIP", "", 0, blnCancelDisplay)
            If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocketPlus.ConnectToIP"
        End If
    End If
End If

End Sub

'This function binds a socket to a local port and IP.
'Retunrs TRUE if it has success.
Private Function BindInternal(ByVal lngIndex As Long, Optional ByVal varLocalPort As Variant, Optional ByVal varLocalIP As Variant) As Boolean
On Error Resume Next
If m_Socket(lngIndex).enmState = sckOpen Then
    BindInternal = True
    Exit Function
End If

Dim lngLocalPortInternal As Long
Dim strLocalHostInternal As String
Dim strIP As String
Dim lngAddressInternal As Long
Dim lngResult As Long
Dim lngErrorCode As Long

BindInternal = False

'Check if varLocalPort is a number between 0 and 65535
If Not IsMissing(varLocalPort) Then
    
    If IsNumeric(varLocalPort) Then
        If varLocalPort < 0 Or varLocalPort > 65535 Then
            BindInternal = False
            ErrRaise sckInvalidArg, "CSocketPlus.BindInternal", "The argument passed to a function was not in the correct format or in the specified range."
            Exit Function
        Else
            lngLocalPortInternal = CLng(varLocalPort)
        End If
    Else
        BindInternal = False
        ErrRaise sckUnsupported, "CSocketPlus.BindInternal", "Unsupported variant type."
        Exit Function
    End If
    
Else
    
    lngLocalPortInternal = m_Socket(lngIndex).lngLocalPort
    
End If

If Not IsMissing(varLocalIP) Then
    If varLocalIP <> vbNullString Then
        strLocalHostInternal = CStr(varLocalIP)
    Else
        strLocalHostInternal = ""
    End If
Else
    strLocalHostInternal = ""
End If

'get a 32 bits long IP
lngAddressInternal = ResolveIfHostnameSync(strLocalHostInternal, strIP, lngResult)

If lngResult <> 0 Then
    ErrRaise sckInvalidArg, "CSocketPlus.BindInternal", "Invalid argument"
    Exit Function
End If

'create a socket if there isn't one yet
If Not SocketExists(lngIndex) Then Exit Function

Dim udtSockAddr As sockaddr_in

With udtSockAddr
    .sin_addr = lngAddressInternal
    .sin_family = AF_INET
    .sin_port = api_htons(modSocketPlus.UnsignedToInteger(lngLocalPortInternal))
End With

'bind the socket
lngResult = api_bind(m_Socket(lngIndex).lngSocketHandle, udtSockAddr, LenB(udtSockAddr))

If lngResult = SOCKET_ERROR Then

    lngErrorCode = Err.LastDllError
    ErrRaise lngErrorCode, "CSocketPlus.BindInternal", GetErrorDescription(lngErrorCode)
    Exit Function
    
Else

    If lngLocalPortInternal <> 0 Then
    
        Debug.Print "OK Bind HOST: " & strLocalHostInternal & " PORT: " & lngLocalPortInternal
        m_Socket(lngIndex).lngLocalPort = lngLocalPortInternal
        
    Else
        lngResult = GetLocalPort(m_Socket(lngIndex).lngSocketHandle)
        
        If lngResult = SOCKET_ERROR Then
            lngErrorCode = Err.LastDllError
            ErrRaise lngErrorCode, "CSocketPlus.BindInternal", GetErrorDescription(lngErrorCode)
            Exit Function
        Else
            Debug.Print "OK Bind HOST: " & strLocalHostInternal & " PORT: " & lngResult
            m_Socket(lngIndex).lngLocalPortBind = lngResult
        End If
        
    End If
    
    BindInternal = True
End If
End Function

'Allocate some memory for HOSTEN structure and returns
'a pointer to this buffer if no error occurs.
'Returns 0 if it fails.
Private Function AllocateMemory(ByVal lngIndex As Long) As Long
On Error Resume Next
With m_Socket(lngIndex)
    .lngMemoryHandle = api_GlobalAlloc(GMEM_FIXED, MAXGETHOSTSTRUCT)

    If .lngMemoryHandle <> 0 Then
        .lngMemoryPointer = api_GlobalLock(.lngMemoryHandle)
    
        If .lngMemoryPointer <> 0 Then
            api_GlobalUnlock (.lngMemoryHandle)
            AllocateMemory = .lngMemoryPointer
        Else
            api_GlobalFree (.lngMemoryHandle)
            AllocateMemory = .lngMemoryPointer '0
        End If

    Else
        AllocateMemory = m_Socket(lngIndex).lngMemoryHandle '0
    End If
End With
End Function

'Free memory allocated by AllocateMemory
Private Sub FreeMemory(ByVal lngIndex As Long)
On Error Resume Next
With m_Socket(lngIndex)
    If .lngMemoryHandle <> 0 Then
        .lngMemoryPointer = 0
        api_GlobalFree .lngMemoryHandle
        .lngMemoryHandle = 0
        Debug.Print "OK Freed resolution memory"
    End If
End With
End Sub

Private Function GetLocalHostName() As String
On Error Resume Next
Dim strHostNameBuf As String * LOCAL_HOST_BUFF
Dim lngResult As Long

lngResult = api_gethostname(strHostNameBuf, LOCAL_HOST_BUFF)

If lngResult = SOCKET_ERROR Then
    GetLocalHostName = vbNullString
    Dim lngErrorCode As Long
    lngErrorCode = Err.LastDllError
    ErrRaise lngErrorCode, "CSocketPlus.GetLocalHostName", GetErrorDescription(lngErrorCode)
    Exit Function
Else
    GetLocalHostName = Left(strHostNameBuf, InStr(1, strHostNameBuf, vbNullChar) - 1)
End If
End Function

'Get local IP when the socket isn't connected yet
Private Function GetLocalIP() As String
On Error Resume Next
Dim lngResult As Long
Dim lngPtrToIP As Long
Dim strLocalHost As String
Dim arrIpAddress(1 To 4) As Byte
Dim Count As Integer
Dim udtHostent As HOSTENT
Dim strIpAddress As String

strLocalHost = GetLocalHostName

lngResult = api_gethostbyname(strLocalHost)

If lngResult = 0 Then
    GetLocalIP = vbNullString
    Dim lngErrorCode As Long
    lngErrorCode = Err.LastDllError
    ErrRaise lngErrorCode, "CSocketPlus.GetLocalIP", GetErrorDescription(lngErrorCode)
    Exit Function
Else
    api_CopyMemory udtHostent, ByVal lngResult, LenB(udtHostent)
    api_CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
    api_CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
   
    For Count = 1 To 4
        strIpAddress = strIpAddress & arrIpAddress(Count) & "."
    Next
   
    strIpAddress = Left(strIpAddress, Len(strIpAddress) - 1)
    GetLocalIP = strIpAddress
End If
End Function

'If Host is an IP doesn't resolve anything and returns a
'a 32 bits long IP.
'If Host isn't an IP then returns vbNull, tries to resolve it
'in asynchronous way.
Private Function ResolveIfHostname(ByVal lngIndex As Long, ByVal Host As String) As Long
On Error Resume Next
Dim lngAddress As Long
lngAddress = api_inet_addr(Host)

If lngAddress = INADDR_NONE Then 'if Host isn't an IP
    
    ResolveIfHostname = vbNull
    m_Socket(lngIndex).enmState = sckResolvingHost: Debug.Print "STATE: sckResolvingHost"
    
    If AllocateMemory(lngIndex) Then
        
        Dim lngAsynHandle As Long
        lngAsynHandle = modSocketPlus.ResolveHost(Host, m_Socket(lngIndex).lngMemoryPointer, ObjPtr(Me))
        
        If lngAsynHandle = 0 Then
            FreeMemory lngIndex
            m_Socket(lngIndex).enmState = sckError: Debug.Print "STATE: sckError"
            Dim lngErrorCode As Long
            lngErrorCode = Err.LastDllError
            Dim blnCancelDisplay As Boolean
            blnCancelDisplay = True
            RaiseEvent Error(m_Socket(lngIndex).varUserIndex, lngErrorCode, GetErrorDescription(lngErrorCode), 0, "CSocketPlus.ResolveIfHostname", "", 0, blnCancelDisplay)
            If blnCancelDisplay = False Then MsgBox GetErrorDescription(lngErrorCode), vbOKOnly, "CSocketPlus.ResolveIfHostname"
        Else
            m_Socket(lngIndex).colWaitingResolutions.Add lngAsynHandle, "R" & lngAsynHandle
            m_colWaitingResolutionsS.Add lngIndex, "R" & lngAsynHandle
            Debug.Print "Resolving host " & Host; " with handle " & lngAsynHandle
        End If
        
    Else
        
        m_Socket(lngIndex).enmState = sckError: Debug.Print "STATE: sckError"
        Debug.Print "Error trying to allocate memory"
        ErrRaise sckOutOfMemory, "CSocketPlus.ResolveIfHostname", "Out of memory"
        Exit Function
        
    End If
    
Else 'if Host is an IP doen't need to resolve anything
    ResolveIfHostname = lngAddress
End If
End Function

'Resolves a host (if necessary) in synchronous way
'If succeeds returns a 32 bits long IP,
'strHostIP = readable IP string and lngErrorCode = 0
'If fails returns vbNull,
'strHostIP = vbNullString and lngErrorCode <> 0
Private Function ResolveIfHostnameSync(ByVal Host As String, ByRef strHostIP As String, ByRef lngErrorCode As Long) As Long
On Error Resume Next
Dim lngPtrToHOSTENT As Long
Dim udtHostent As HOSTENT
Dim lngAddress As Long
Dim lngPtrToIP As Long
Dim arrIpAddress(1 To 4) As Byte
Dim Count As Integer

lngAddress = api_inet_addr(Host)

If lngAddress = INADDR_NONE Then 'if Host isn't an IP
    
    lngPtrToHOSTENT = api_gethostbyname(Host)
    
    If lngPtrToHOSTENT = 0 Then
        lngErrorCode = Err.LastDllError
        strHostIP = vbNullString
        ResolveIfHostnameSync = vbNull
    Else
        api_CopyMemory udtHostent, ByVal lngPtrToHOSTENT, LenB(udtHostent)
        api_CopyMemory lngPtrToIP, ByVal udtHostent.hAddrList, 4
        api_CopyMemory arrIpAddress(1), ByVal lngPtrToIP, 4
        api_CopyMemory lngAddress, ByVal lngPtrToIP, 4
        
        For Count = 1 To 4
            strHostIP = strHostIP & arrIpAddress(Count) & "."
        Next
        
        strHostIP = Left(strHostIP, Len(strHostIP) - 1)
        
        lngErrorCode = 0
        ResolveIfHostnameSync = lngAddress
    End If
    
Else 'if Host is an IP doen't need to resolve anything
    
    lngErrorCode = 0
    strHostIP = Host
    ResolveIfHostnameSync = lngAddress
    
End If
End Function

'Returns local port from a connected or bound socket.
'Returns SOCKET_ERROR if fails.
Private Function GetLocalPort(ByVal lngSocket As Long) As Long
On Error Resume Next
Dim udtSockAddr As sockaddr_in
Dim lngResult As Long

lngResult = api_getsockname(lngSocket, udtSockAddr, LenB(udtSockAddr))

If lngResult = SOCKET_ERROR Then
    GetLocalPort = SOCKET_ERROR
Else
    GetLocalPort = modSocketPlus.IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
End If
End Function

'Send buffered data if we are using TCP protocol.
Private Sub SendBufferedDataTCP(ByVal lngIndex As Long)
On Error Resume Next

Dim arrData()       As Byte
Dim lngBufferLength As Long
Dim lngResult    As Long

lngBufferLength = LenB(m_Socket(lngIndex).strSendBuffer)

Do Until lngResult = SOCKET_ERROR Or lngBufferLength = 0

    If lngBufferLength > m_Socket(lngIndex).lngSendBufferLen Then lngBufferLength = m_Socket(lngIndex).lngSendBufferLen
    ReDim arrData(lngBufferLength - 1)
    api_CopyMemory arrData(0), ByVal StrPtr(m_Socket(lngIndex).strSendBuffer), lngBufferLength

    lngResult = api_send(m_Socket(lngIndex).lngSocketHandle, arrData(0), lngBufferLength, 0&)

    If lngResult > 0 Then
        Dim lngLeftLength As Long
        lngLeftLength = LenB(m_Socket(lngIndex).strSendBuffer) - lngResult
        
        If lngLeftLength > 0 Then
            ReDim arrData(lngLeftLength - 1)
            api_CopyMemory arrData(0), ByVal StrPtr(m_Socket(lngIndex).strSendBuffer) + lngResult, lngLeftLength
            m_Socket(lngIndex).strSendBuffer = CStr(arrData)
        Else
            m_Socket(lngIndex).strSendBuffer = vbNullString
        End If
    End If

    lngBufferLength = LenB(m_Socket(lngIndex).strSendBuffer)
    
Loop
    
End Sub

'This function retrieves data from the Winsock buffer
'into the class local buffer. The function returns number
'of bytes retrieved (received).
Private Function RecvDataToBuffer(ByVal lngIndex As Long) As Long
On Error Resume Next
Dim arrBuffer() As Byte
Dim lngBytesReceived As Long

With m_Socket(lngIndex)

    ReDim arrBuffer(.lngRecvBufferLen - 1)

    lngBytesReceived = api_recv(.lngSocketHandle, arrBuffer(0), .lngRecvBufferLen, 0&)

    If lngBytesReceived > 0 Then
    
        ReDim Preserve arrBuffer(lngBytesReceived - 1)
        .strRecvBuffer = CStr(arrBuffer)
        RecvDataToBuffer = lngBytesReceived
    
    End If

End With
End Function

'Retrieves some socket options.
Private Sub ProcessOptions(ByVal lngIndex As Long)
On Error Resume Next
Dim lngResult As Long
Dim lngBuffer As Long
Dim lngErrorCode As Long

With m_Socket(lngIndex)

    lngResult = api_getsockopt(.lngSocketHandle, SOL_SOCKET, SO_RCVBUF, lngBuffer, LenB(lngBuffer))

    If lngResult = SOCKET_ERROR Then
        lngErrorCode = Err.LastDllError
        ErrRaise lngErrorCode, "CSocketPlus.ProcessOptions", GetErrorDescription(lngErrorCode)
        Exit Sub
    Else
        .lngRecvBufferLen = lngBuffer
    End If

    lngResult = api_getsockopt(.lngSocketHandle, SOL_SOCKET, SO_SNDBUF, lngBuffer, LenB(lngBuffer))

    If lngResult = SOCKET_ERROR Then
        lngErrorCode = Err.LastDllError
        ErrRaise lngErrorCode, "CSocketPlus.ProcessOptions", GetErrorDescription(lngErrorCode)
        Exit Sub
    Else
        .lngSendBufferLen = lngBuffer
    End If


    Debug.Print "Winsock buffer size for sends: " & .lngRecvBufferLen
    Debug.Print "Winsock buffer size for receives: " & .lngSendBufferLen
    
End With
End Sub

'Clean resolution system that is in charge of
'asynchronous hostname resolutions.
Private Sub CleanResolutionSystem(ByVal lngIndex As Long)
On Error Resume Next
Dim varAsynHandle As Variant
Dim lngResult As Long

'cancel async resolutions if they're still running
For Each varAsynHandle In m_Socket(lngIndex).colWaitingResolutions
    lngResult = api_WSACancelAsyncRequest(varAsynHandle)
    If lngResult = 0 Then
        modSocketPlus.UnregisterResolution varAsynHandle
        
        Set m_Socket(lngIndex).colWaitingResolutions = Nothing
        Set m_Socket(lngIndex).colWaitingResolutions = New Collection

        'free memory buffer where resolution results are stored
        FreeMemory lngIndex
        
    End If
Next

End Sub

'Retrieves local info from a connected socket.
'If succeeds returns TRUE and loads the arguments.
'If fails returns FALSE and arguments are not loaded.
Private Function GetLocalInfo(ByVal lngSocket As Long, ByRef lngLocalPort As Long, ByRef strLocalIP As String) As Boolean
On Error Resume Next
GetLocalInfo = False
Dim lngResult As Long
Dim udtSockAddr As sockaddr_in

lngResult = api_getsockname(lngSocket, udtSockAddr, LenB(udtSockAddr))

If lngResult = SOCKET_ERROR Then
    lngLocalPort = 0
    strLocalIP = ""
Else
    GetLocalInfo = True
    lngLocalPort = IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
    strLocalIP = StringFromPointer(api_inet_ntoa(udtSockAddr.sin_addr))
End If

End Function

'Retrieves remote info from a connected socket.
'If succeeds returns TRUE and loads the arguments.
'If fails returns FALSE and arguments are not loaded.
Private Function GetRemoteInfo(ByVal lngSocket As Long, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
On Error Resume Next
GetRemoteInfo = False
Dim lngResult As Long
Dim udtSockAddr As sockaddr_in

lngResult = api_getpeername(lngSocket, udtSockAddr, LenB(udtSockAddr))

If lngResult = 0 Then
    GetRemoteInfo = True
    GetRemoteInfoFromSI udtSockAddr, lngRemotePort, strRemoteHostIP, strRemoteHost
Else
   lngRemotePort = 0
   strRemoteHostIP = ""
   strRemoteHost = ""
End If
End Function

'Gets remote info from a sockaddr_in structure.
Private Sub GetRemoteInfoFromSI(ByRef udtSockAddr As sockaddr_in, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String)
On Error Resume Next
'Dim lngResult As Long
'Dim udtHostent As HOSTENT

lngRemotePort = IntegerToUnsigned(api_ntohs(udtSockAddr.sin_port))
strRemoteHostIP = StringFromPointer(api_inet_ntoa(udtSockAddr.sin_addr))
'lngResult = api_gethostbyaddr(udtSockAddr.sin_addr, 4&, AF_INET)

'If lngResult <> 0 Then
'    api_CopyMemory udtHostent, ByVal lngResult, LenB(udtHostent)
'    strRemoteHost = StringFromPointer(udtHostent.hName)
'Else
    strRemoteHost = ""
'End If

End Sub

'Removes the socket Index from the socket array
'If it is the last one cleans the array
Private Sub CleanSocketArray(ByVal lngIndex As Long)

On Error GoTo Error_Handler

Dim lngNextEmptyIndex As Long
lngNextEmptyIndex = UBound(m_EmptyIndex) + 1

If UBound(m_Socket) = lngNextEmptyIndex Then
    ReDim m_Socket(0) As typSocket
    ReDim m_EmptyIndex(0) As Long
Else 'if it is the last in the array we can erase it from the array
    If UBound(m_Socket) = lngIndex Then
        ReDim Preserve m_Socket(0 To lngIndex - 1) As typSocket
    Else 'if it isn't the last in the array we can't erase it
        With m_Socket(lngIndex)
        Set .colWaitingResolutions = Nothing
        Set .colWaitingResolutions = New Collection
        .enmProtocol = sckTCPProtocol
        .enmState = sckClosed
        .lngLocalPort = -1
        .lngLocalPortBind = -1
        .lngMemoryHandle = 0
        .lngMemoryPointer = 0
        .lngRecvBufferLen = 0
        .lngRemotePort = -1
        .lngSendBufferLen = 0
        .lngSocketHandle = -1
        .strLocalIP = ""
        .strRecvBuffer = ""
        .strRemoteHost = ""
        .strRemoteHostIP = ""
        .strSendBuffer = ""
        .strTag = ""
        .varUserIndex = 0
        End With
        ReDim Preserve m_EmptyIndex(lngNextEmptyIndex)
        m_EmptyIndex(lngNextEmptyIndex) = lngIndex
    End If
End If

Exit Sub

Error_Handler:
    ErrRaise sckOutOfMemory, "CSocketPlus.CleanSocketArray", "Out of memory"
End Sub


'Returns a free index to be used on the socket array.
'Redims the array if necessary.
Private Function GetFreeSocketIndex() As Long
On Error GoTo Error_Handler:

Dim lngFreeBound As Long
Dim lngIndex As Long

lngFreeBound = UBound(m_EmptyIndex)

If lngFreeBound = 0 Then    'if there isn't any empty index

    lngIndex = UBound(m_Socket) + 1
    ReDim Preserve m_Socket(lngIndex)
    
Else 'if there are some empty indexes we take the last one
    lngIndex = m_EmptyIndex(UBound(m_EmptyIndex))
    ReDim Preserve m_EmptyIndex(UBound(m_EmptyIndex) - 1)

End If

GetFreeSocketIndex = lngIndex

Exit Function

Error_Handler:
    ErrRaise sckOutOfMemory, "CSocketPlus.GetFreeSocketIndex", "Out of memory"
End Function

'Get socket array index from user index or string index
Private Function GetSocketIndex(Index As Variant) As Long
On Error GoTo Error_Handler

Select Case VarType(Index)
    Case vbString
        GetSocketIndex = m_colString_Index("I" & Index)
    Case vbInteger To vbLong
        GetSocketIndex = m_colUserIndex_Index("I" & Index)
    Case Else
        ErrRaise sckWrongIndex
        Exit Function
End Select

Exit Function

Error_Handler:
    Dim lngErrorCode As Long
    Dim strErrorDesc As String
    
    lngErrorCode = Err.Number
    strErrorDesc = Err.Description
    
    Const strSource As String = "CSocketPlus.GetSocketIndex"
    
    If Err.Number = sckWrongIndex Then
        ErrRaise sckWrongIndex, strSource, "Unsupported index type."
    ElseIf Err.Number = 5 Then
        ErrRaise sckMissingIndex, strSource, "Unknown index."
    Else
        ErrRaise lngErrorCode, strSource, strErrorDesc
    End If
End Function

'returns true if user index is registered on
'm_colUserIndex_Index
Private Function IsUserIndexRegistered(ByVal lngIndex As Long) As Boolean
On Error GoTo Error_Handler

m_colUserIndex_Index.Item ("I" & lngIndex)
IsUserIndexRegistered = True

Exit Function

Error_Handler:
    IsUserIndexRegistered = False
End Function
