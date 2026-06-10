
Attribute VB_Name = "modBridge"

Option Explicit

' ============================================================
'  lifecycle + helpers for the WebSocket bridge
' ============================================================

Public gHost As cBridgeHost
Public gSuspendEvents As Boolean
Private gBatchRouter As Object   ' Scripting.Dictionary

' Heartbeat / keep-alive timer
Private gHeartbeatCallback As String    ' fully-qualified "Module.Sub" to call
Private gHeartbeatSeconds As Long       ' interval in seconds (0 = disabled)
Private gHeartbeatNextRun As Date       ' scheduled time for next fire
Private gHeartbeatActive As Boolean     ' is the timer running?

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If


Public Sub StartBridge()
    ' Initialize message router
    modSocketRouter.InitRouter
    'modSocketRouter.RouteDefault "modBridge.HandleUnknown"

    ' Configure heartbeat (fires SendPing every X seconds while connected)
    SetHeartbeat "modBridge.SendPing", 30

    If gHost Is Nothing Then
        Set gHost = New cBridgeHost
    End If
    gHost.StartUp "ws://cds-sn-api-dev.sik.intranet.barcapint.com/ws/cds?username=" & User_GetUserName()
    
End Sub

Public Sub StopBridge()
    If Not gHost Is Nothing Then gHost.ShutDown
    Set gHost = Nothing
End Sub

Public Sub EnsureBridge()
    ' Restart bridge only if truly dead (not just mid-reconnect)
    If gHost Is Nothing Then
        StartBridge
    ElseIf gHost.b Is Nothing Then
        Set gHost = Nothing
        StartBridge
    ElseIf Not gHost.b.IsRunning Then
        ' Loop has exited entirely � restart
        Set gHost = Nothing
        StartBridge
    End If
    ' If IsRunning=True but IsConnected=False, it's reconnecting � leave it alone
End Sub

' ---- Example extensibility: pure-VBA commands (no DLL rebuild) ----

Public Sub SendPing()
    modBridge.EnsureBridge
    If gHost Is Nothing Then Exit Sub
    gHost.b.Send "ping", ""
End Sub

Public Sub SendClear()
    If gHost Is Nothing Then Exit Sub
    gHost.b.Send "clear", ""
End Sub

Public Sub RegisterBatch(requestId As String, batch As cHttpBatch)
    If gBatchRouter Is Nothing Then Set gBatchRouter = CreateObject("Scripting.Dictionary")
    If gBatchRouter.exists(requestId) Then
        Set gBatchRouter(requestId) = batch
    Else
        gBatchRouter.Add requestId, batch
    End If
End Sub

Public Sub RouteHttpResult(requestId As String, r As cHttpResponse)
    If gBatchRouter Is Nothing Then Exit Sub
    If Not gBatchRouter.exists(requestId) Then Exit Sub
    Dim batch As cHttpBatch
    Set batch = gBatchRouter(requestId)
    gBatchRouter.Remove requestId
    If Not batch Is Nothing Then batch.OnComplete requestId, r
End Sub

' ============================================================
'  HEARTBEAT / KEEP-ALIVE TIMER (async via Application.OnTime)
' ============================================================

Public Sub SetHeartbeat(callbackId As String, Optional intervalSeconds As Long = 10)
    ' Configure a periodic callback while connected.
    ' Example: SetHeartbeat "modBridge.SendPing", 10
    gHeartbeatCallback = callbackId
    gHeartbeatSeconds = intervalSeconds
    If intervalSeconds > 0 Then
        StartHeartbeat
    Else
        StopHeartbeat
    End If
End Sub

Public Sub StartHeartbeat()
    If gHeartbeatSeconds <= 0 Or Len(gHeartbeatCallback) = 0 Then Exit Sub
    gHeartbeatActive = True
    gHeartbeatNextRun = Now + TimeSerial(0, 0, gHeartbeatSeconds)
    Application.OnTime gHeartbeatNextRun, "modBridge.HeartbeatTick"
End Sub

Public Sub StopHeartbeat()
    If Not gHeartbeatActive Then Exit Sub
    On Error Resume Next
    Application.OnTime gHeartbeatNextRun, "modBridge.HeartbeatTick", , False
    On Error GoTo 0
    gHeartbeatActive = False
End Sub

Public Sub HeartbeatTick()
    ' Fired by Application.OnTime � runs async on Excel's idle loop
    If Not gHeartbeatActive Then Exit Sub
    If gHeartbeatSeconds <= 0 Then Exit Sub

    ' Only fire if connected
    If Not gHost Is Nothing Then
        If Not gHost.b Is Nothing Then
            If gHost.b.IsConnected Then
                On Error Resume Next
                Application.Run gHeartbeatCallback
                On Error GoTo 0
            End If
        End If
    End If

    ' Schedule next tick
    If gHeartbeatActive Then
        gHeartbeatNextRun = Now + TimeSerial(0, 0, gHeartbeatSeconds)
        Application.OnTime gHeartbeatNextRun, "modBridge.HeartbeatTick"
    End If
End Sub

' ============================================================
'  Message router callbacks
' ============================================================

Public Sub HandlePong(jsonPayload As String)
    Debug.Print "Application Ping: " & jsonPayload
End Sub

Public Sub HandlePing(jsonPayload As String)
    Debug.Print "Server Ping: " & jsonPayload
End Sub

Public Sub HandleStatus(jsonPayload As String)
    [lastFetchTime].Value2 = ExtractTimeEastern(modSocketRouter.ExtractJsonValue(jsonPayload, "timestamp"))
End Sub

Public Sub HandleUnknown(msgType As String, jsonPayload As String)
    Debug.Print "unhandled [" & msgType & "]: " & Left$(jsonPayload, 200)
End Sub




