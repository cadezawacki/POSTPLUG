
Attribute VB_Name = "socket_filter"
'
' reference site https://stackoverflow.com/questions/49028281/vba-with-winsock2-send-sends-wrong-data
' edited by robotmanya (2018.10.28) (https://blog.naver.com/monkey5255/221386590654)

' Constants ----------------------------------------------------------
'Const ip = "nykpwm465753"
'Const port = "5001"

Const INVALID_SOCKET = -1
Const WSADESCRIPTION_LEN = 256
Const SOCKET_ERROR = -1
Const SD_SEND = 1

' Typ definitions ----------------------------------------------------
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADESCRIPTION_LEN) As Byte
    szSystemStatus(0 To WSADESCRIPTION_LEN) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type ADDRINFO
    ai_flags As Long
    ai_family As Long
    ai_socktype As Long
    ai_protocol As Long
    ai_addrlen As Long
    ai_canonName As LongPtr 'strptr
    ai_addr As LongPtr 'p sockaddr
    ai_next As LongPtr 'p addrinfo
End Type


' Enums ---------------------------------------------------------------
Enum AF
    AF_UNSPEC = 0
    AF_INET = 2
    AF_IPX = 6
    AF_APPLETALK = 16
    AF_NETBIOS = 17
    AF_INET6 = 23
    AF_IRDA = 26
    AF_BTH = 32
End Enum

Enum sock_type
    SOCK_STREAM = 1
    SOCK_DGRAM = 2
    SOCK_RAW = 3
    SOCK_RDM = 4
    SOCK_SEQPACKET = 5
End Enum
' External functions --------------------------------------------------

Public Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef data As WSADATA) As Long
Public Declare PtrSafe Function connect Lib "ws2_32.dll" (ByVal socket As Long, ByVal SOCKADDR As LongPtr, ByVal namelen As Long) As Long
Public Declare PtrSafe Sub WSACleanup Lib "ws2_32.dll" ()
Private Declare PtrSafe Function GetAddrInfo Lib "ws2_32.dll" Alias "getaddrinfo" (ByVal NodeName As String, ByVal ServName As String, ByVal lpHints As LongPtr, lpResult As LongPtr) As Long
Public Declare PtrSafe Function ws_socket Lib "ws2_32.dll" Alias "socket" (ByVal AF As Long, ByVal stype As Long, ByVal Protocol As Long) As Long
Public Declare PtrSafe Function closesocket Lib "ws2_32.dll" (ByVal socket As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare PtrSafe Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByVal buf As String, ByVal buflen As Long, ByVal flags As Long) As Long
' Public Declare ptrsafe Function recv Lib "wsock32.dll" (ByVal socket As Long, ByVal buffer As String, ByVal bufferLength As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function SendWithPtr Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByVal bufPtr As Long, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function ShutDown Lib "ws2_32.dll" Alias "shutdown" (ByVal s As Long, ByVal how As Long) As Long
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function VarPtrArray Lib "VBE7" Alias "VarPtr" (var() As Any) As Long
Public Declare PtrSafe Function Recv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Byte, ByVal buflen As Long, ByVal flags As Long) As Long

Dim m_wsaData As WSADATA
Dim m_RetVal As Integer
Dim m_ConnSocket As Long
Dim retVal As Long
Dim lastError As Long
Dim iRC As Long
Dim MAX_BUF_SIZE As Integer
Dim CONNECTED As Boolean

Public Function am_i_connected() As Boolean
    am_i_connected = CONNECTED
End Function

Public Sub test2()
    
    a = Login("me", "test")
    
End Sub

Public Sub tester()
    
    Call ConnectToWs
    a = 1
    sendMsg ("Test <1>")
    sendMsg ("Test <2>")
    b = 2
    Call DisconnectFromWs
    c = 3

End Sub

'Login Button Click Event
Function Login(id As String, pw As String) As String
    
    Login = 0
    ip = [SDR.Host]
    port = [Webhook.Port]
    
    Dim m_Hints As ADDRINFO
    Dim pAddrInfo As LongPtr
    
    m_ConnSocket = INVALID_SOCKET
    MAX_BUF_SIZE = 512
    
    'Socket Settings
    retVal = WSAStartup(MAKEWORD(2, 2), m_wsaData)
    If (retVal <> 0) Then
        LogError "WSAStartup failed with error " & retVal, WSAGetLastError()
        Call WSACleanup
        Exit Function
    End If

    m_Hints.ai_family = AF.AF_UNSPEC
    m_Hints.ai_socktype = sock_type.SOCK_STREAM

    retVal = GetAddrInfo(ip, port, VarPtr(m_Hints), pAddrInfo)
    If (retVal <> 0) Then
        LogError "Cannot resolve address " & ip & " and port " & port & ", error " & retVal, WSAGetLastError()
        Call WSACleanup
        Exit Function
    End If

    m_Hints.ai_next = pAddrInfo
    Dim CONNECTED As Boolean: CONNECTED = False
    Do While m_Hints.ai_next > 0
        CopyMemory m_Hints, ByVal m_Hints.ai_next, LenB(m_Hints)

        m_ConnSocket = ws_socket(m_Hints.ai_family, m_Hints.ai_socktype, m_Hints.ai_protocol)

        If (m_ConnSocket = INVALID_SOCKET) Then
            LogError "Error opening socket, error " & retVal
        Else
            Dim connectionResult As Long

            connectionResult = connect(m_ConnSocket, m_Hints.ai_addr, m_Hints.ai_addrlen)

            If connectionResult <> SOCKET_ERROR Then
                CONNECTED = True
                Exit Do
            End If

            LogError "connect() to socket failed"
            closesocket (m_ConnSocket)
        End If
    Loop

    If Not CONNECTED Then
        LogError "Fatal error: unable to connect to the server", WSAGetLastError()
        retVal = closesocket(m_ConnSocket)
        Call WSACleanup
        Exit Function
    End If

    'After Socket Connected
    Dim SendBuf As String
    SendBuf = id + "|" + pw

    'Send Login Data
    retVal = Send(m_ConnSocket, SendBuf, Len(SendBuf), 0)

    If retVal = SOCKET_ERROR Then
        LogError "send() failed", WSAGetLastError()
        retVal = closesocket(m_ConnSocket)
        Call WSACleanup
        Exit Function
    Else
        ' Debug.Print "sent " & retVal & " bytes"
    End If

    ' shutdown the connection since no more data will be sent
    retVal = ShutDown(m_ConnSocket, SD_SEND)
    If retVal <> 0 Then
        LogError "send socket close failed", WSAGetLastError()
        retVal = closesocket(m_ConnSocket)
        Call WSACleanup
    Else
        Debug.Print "send socket closed"
    End If
    
    Login = "yippee"

End Function

Public Function ConnectToWs() As Boolean

    Dim m_Hints As ADDRINFO
    Dim pAddrInfo As LongPtr
    ip = [SDR.Host]
    port = [Webhook.Port]
    
    m_ConnSocket = INVALID_SOCKET
    CONNECTED = False
    MAX_BUF_SIZE = 512
    ConnectToWs = False
    
    'Socket Settings
    retVal = WSAStartup(MAKEWORD(2, 2), m_wsaData)
    If (retVal <> 0) Then
        LogError "WSAStartup failed with error " & retVal, WSAGetLastError()
        Call WSACleanup
        Exit Function
    End If

    m_Hints.ai_family = AF.AF_UNSPEC
    m_Hints.ai_socktype = sock_type.SOCK_STREAM

    retVal = GetAddrInfo(ip, port, 0, pAddrInfo)
    If (retVal <> 0) Then
        LogError "Cannot resolve address " & ip & " and port " & port & ", error " & retVal, WSAGetLastError()
        Call WSACleanup
        Exit Function
    End If

    m_Hints.ai_next = pAddrInfo
    Dim CONNECTED_ As Boolean: CONNECTED_ = False
    Do While m_Hints.ai_next > 0
        CopyMemory m_Hints, ByVal m_Hints.ai_next, LenB(m_Hints)

        m_ConnSocket = ws_socket(m_Hints.ai_family, m_Hints.ai_socktype, m_Hints.ai_protocol)

        If (m_ConnSocket = INVALID_SOCKET) Then
            LogError "Error opening socket, error " & retVal
        Else
            Dim connectionResult As Long

            connectionResult = connect(m_ConnSocket, m_Hints.ai_addr, m_Hints.ai_addrlen)

            If connectionResult <> SOCKET_ERROR Then
                CONNECTED_ = True
                Exit Do
            End If

            LogError "connect() to socket failed"
            closesocket (m_ConnSocket)
        End If
    Loop

    If Not CONNECTED_ Then
        LogError "Fatal error: unable to connect to the server", WSAGetLastError()
        retVal = closesocket(m_ConnSocket)
        Call WSACleanup
        Exit Function
        ConnectToWs = False
        CONNECTED = False
    Else
        ConnectToWs = True
        CONNECTED = True
        Debug.Print "Successfully connected to WS"
        ThisWorkbook.Worksheets("Main").Range("HubStatus").value = "ON"
    End If
    
End Function

Public Function sendMsg(msg As String) As Boolean
    
    ip = [SDR.Host]
    port = [Webhook.Port]

    'After Socket Connected
    Dim SendBuf As String
    SendBuf = msg
    sendMsg = False

    'Send Login Data
    retVal = Send(m_ConnSocket, SendBuf, Len(SendBuf), 0)

    If retVal = SOCKET_ERROR Then
        LogError "send() failed", WSAGetLastError()
        retVal = closesocket(m_ConnSocket)
        Call WSACleanup
        Exit Function
    Else
        ' Debug.Print "sent " & retVal & " bytes"
        sendMsg = True
    End If
    
End Function

Sub DisconnectFromWs()
    
    ip = [SDR.Host]
    port = [Webhook.Port]

    ' shutdown the connection since no more data will be sent
    retVal = ShutDown(m_ConnSocket, SD_SEND)
    If retVal <> 0 Then
        LogError "send socket close failed", WSAGetLastError()
        retVal = closesocket(m_ConnSocket)
        Call WSACleanup
    Else
        Debug.Print "send socket closed"
    End If
    CONNECTED = False
End Sub

Public Sub i_am_not_connected()
    CONNECTED = False
End Sub

'Function GetMsg() As String
'    Dim m_wsaData As WSADATA
'    Dim m_RetVal As Integer
'    Dim m_Hints As ADDRINFO
'    Dim m_ConnSocket As Long: m_ConnSocket = INVALID_SOCKET
'    Dim pAddrInfo As LongPtr
'    Dim retVal As Long
'    Dim lastError As Long
'    Dim iRC As Long
'    Dim MAX_BUF_SIZE As Integer: MAX_BUF_SIZE = 1024
'
'    'Socket Settings
'    retVal = WSAStartup(MAKEWORD(2, 2), m_wsaData)
'    If (retVal <> 0) Then
'        LogError "WSAStartup failed with error " & retVal, WSAGetLastError()
'        Call WSACleanup
'        Exit Function
'    End If
'
'    m_Hints.ai_family = AF.AF_UNSPEC
'    m_Hints.ai_socktype = sock_type.SOCK_STREAM
'
'    retVal = GetAddrInfo(ip, port, VarPtr(m_Hints), pAddrInfo)
'    If (retVal <> 0) Then
'        LogError "Cannot resolve address " & ip & " and port " & port & ", error " & retVal, WSAGetLastError()
'        Call WSACleanup
'        Exit Function
'    End If
'
'    m_Hints.ai_next = pAddrInfo
'    Dim connected As Boolean: connected = False
'    Do While m_Hints.ai_next > 0
'        CopyMemory m_Hints, ByVal m_Hints.ai_next, LenB(m_Hints)
'
'        m_ConnSocket = ws_socket(m_Hints.ai_family, m_Hints.ai_socktype, m_Hints.ai_protocol)
'
'        If (m_ConnSocket = INVALID_SOCKET) Then
'            LogError "Error opening socket, error " & retVal
'        Else
'            Dim connectionResult As Long
'
'            connectionResult = connect(m_ConnSocket, m_Hints.ai_addr, m_Hints.ai_addrlen)
'
'            If connectionResult <> SOCKET_ERROR Then
'                connected = True
'                Exit Do
'            End If
'
'            LogError "connect() to socket failed"
'            closesocket (m_ConnSocket)
'        End If
'    Loop
'
'    If Not connected Then
'        LogError "Fatal error: unable to connect to the server", WSAGetLastError()
'        retVal = closesocket(m_ConnSocket)
'        Call WSACleanup
'        Exit Function
'    End If
'
'    'Recieve From Server (Login Success : 1, Fail : 0)
'    ' Dim recvBuf As Byte
'    Dim recvBuf As Byte
'    retVal = Recv(m_ConnSocket, recvBuf, MAX_BUF_SIZE, 0)
'    test = StrPtr(recvBuf)
'    If retVal = SOCKET_ERROR Then
'        LogError "recv() failed", WSAGetLastError()
'        retVal = closesocket(m_ConnSocket)
'        Call WSACleanup
'        Exit Function
'    Else
'        ' Debug.Print "recieved " & RetVal & " bytes"
'    End If
'
'    'Login Check (s : success(id,pw correspond, f : fail)
'    ' response = Chr(recvBuf)
'    GetMsg = "test"
'
'
'
'    retVal = closesocket(m_ConnSocket)
'    If retVal <> 0 Then
'    LogError "closesocket() failed", WSAGetLastError()
'    Call WSACleanup
'    Else
'        Debug.Print "closed socket"
'    End If
'End Function

Public Function MAKEWORD(lo As Byte, hi As Byte) As Integer
    MAKEWORD = lo + hi * 256& Or 32768 * (hi > 127)
End Function

Private Sub LogError(msg As String, Optional ErrorCode As Long = -1)
    If ErrorCode > -1 Then
        msg = msg & " (error code " & ErrorCode & ")"
    End If

    Debug.Print msg
End Sub

Public Sub test_login()
    Dim username As String
    username = Environ$("username")
    a = Login("zawackic", "KHC")
    bb = 3
End Sub






