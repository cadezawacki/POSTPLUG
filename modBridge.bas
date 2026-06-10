
Attribute VB_Name = "modBridge"

Option Explicit

' ============================================================
'  Lifecycle + helpers for the ExcelBridge XLL add-in.
'
'  HTTP and WS are independent:
'    EnsureHttp   - loads/attaches the XLL only (async REST ready,
'                   no socket is ever opened by this path)
'    StartBridge  - additionally starts the WebSocket feed
' ============================================================

Public gSuspendEvents As Boolean
Public gHttpVerbose As Boolean   ' opt-in: Debug.Print + HttpLog sheet per HTTP completion
Private gBatchRouter As Object   ' Scripting.Dictionary
Private gAttached As Boolean     ' EB_Attach done for this workbook

Private Const HTTP_DEFAULT_TIMEOUT_MS As Long = 30000
Private Const WS_MAX_RECONNECT_DELAY_MS As Long = 15000
Private Const WS_BATCH_WINDOW_MS As Long = 16
Private Const WS_BATCH_MAX_COUNT As Long = 200

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

' ============================================================
'  ADD-IN LIFECYCLE
' ============================================================

' True if the ExcelBridge XLL is loaded in this Excel session.
Public Function AddinLoaded() As Boolean
    On Error Resume Next
    Dim v As Variant
    v = Application.Run("EB_Version")
    AddinLoaded = (Err.Number = 0)
    If AddinLoaded Then AddinLoaded = Not IsError(v)
    On Error GoTo 0
End Function

' Load the XLL if it isn't already. Looks for an explicit path in the
' 'BridgeXllPath' named range first, then next to the workbook.
Public Sub EnsureAddin()
    If AddinLoaded() Then Exit Sub

    Dim candidates(0 To 2) As String
    Dim n As Long: n = 0

    Dim configured As String
    configured = ConfiguredXllPath()
    If Len(configured) > 0 Then candidates(n) = configured: n = n + 1

    Dim base As String
    base = ThisWorkbook.Path & Application.PathSeparator
    ' Packed first: it is self-contained. The unpacked flavor only works with
    ' ExcelBridge.dll AND Newtonsoft.Json.dll sitting beside it.
    #If Win64 Then
        candidates(n) = base & "ExcelBridge-AddIn64-packed.xll": n = n + 1
        candidates(n) = base & "ExcelBridge-AddIn64.xll": n = n + 1
    #Else
        candidates(n) = base & "ExcelBridge-AddIn-packed.xll": n = n + 1
        candidates(n) = base & "ExcelBridge-AddIn.xll": n = n + 1
    #End If

    Dim i As Long
    For i = 0 To n - 1
        If Len(candidates(i)) > 0 Then
            If Len(Dir$(candidates(i))) > 0 Then
                If Application.RegisterXLL(candidates(i)) Then Exit For
            End If
        End If
    Next i

    If Not AddinLoaded() Then
        Err.Raise vbObjectError + 513, "modBridge.EnsureAddin", _
            "ExcelBridge add-in is not loaded. Place the XLL next to the workbook " & _
            "or set its full path in the 'BridgeXllPath' named range."
    End If
End Sub

Private Function ConfiguredXllPath() As String
    On Error Resume Next
    ConfiguredXllPath = CStr(ThisWorkbook.Names("BridgeXllPath").RefersToRange.Value2)
    On Error GoTo 0
End Function

' HTTP-only entry point: loads the XLL and attaches this workbook as the
' callback target. Never opens (or restarts) a WebSocket.
Public Sub EnsureHttp()
    EnsureAddin
    If Not gAttached Then
        Application.Run "EB_Attach", ThisWorkbook.name
        Application.Run "EB_HttpSetDefaultTimeout", HTTP_DEFAULT_TIMEOUT_MS
        gAttached = True
    End If
End Sub

' Stop callbacks into this workbook (call before close).
Public Sub DetachBridge()
    gAttached = False
    If Not AddinLoaded() Then Exit Sub
    On Error Resume Next
    Application.Run "EB_Detach"
    On Error GoTo 0
End Sub

' ============================================================
'  WEBSOCKET LIFECYCLE
' ============================================================

Public Sub StartBridge()
    ' Idempotent router init: a restart must not wipe registered routes.
    modSocketRouter.EnsureRouter

    ' Configure heartbeat (fires SendPing every X seconds while connected)
    SetHeartbeat "modBridge.SendPing", 30

    EnsureHttp
    Application.Run "EB_WsConfig", True, WS_MAX_RECONNECT_DELAY_MS, False, _
                    WS_BATCH_WINDOW_MS, WS_BATCH_MAX_COUNT
    Application.Run "EB_WsStart", BridgeWsUrl()
End Sub

Private Function BridgeWsUrl() As String
    BridgeWsUrl = "ws://cds-sn-api-dev.sik.intranet.barcapint.com/ws/cds?username=" & User_GetUserName()
End Function

Public Sub StopBridge()
    If Not AddinLoaded() Then Exit Sub
    On Error Resume Next
    Application.Run "EB_WsStop"
    On Error GoTo 0
End Sub

Public Sub EnsureBridge()
    ' Restart the WS only if the loop has exited entirely; while it is
    ' running (connected OR mid-reconnect) leave it alone.
    EnsureHttp
    If Not CBool(Application.Run("EB_WsIsRunning")) Then
        StartBridge
    End If
End Sub

' ---- Example extensibility: pure-VBA commands (no DLL rebuild) ----

Public Sub SendPing()
    EnsureBridge
    Application.Run "EB_WsSend", "ping", ""
End Sub

Public Sub SendClear()
    If Not AddinLoaded() Then Exit Sub
    Application.Run "EB_WsSend", "clear", ""
End Sub

' ============================================================
'  HTTP BATCH ROUTING
' ============================================================

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
    ' Cancel any pending tick first: re-arming without cancelling spawns a
    ' second OnTime chain (each restart would add another, multiplying pings).
    StopHeartbeat
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
    ' Fired by Application.OnTime - runs async on Excel's idle loop
    If Not gHeartbeatActive Then Exit Sub
    If gHeartbeatSeconds <= 0 Then Exit Sub

    ' Only fire if connected
    On Error Resume Next
    If AddinLoaded() Then
        If CBool(Application.Run("EB_WsIsConnected")) Then
            Application.Run gHeartbeatCallback
        End If
    End If
    On Error GoTo 0

    ' Schedule next tick
    If gHeartbeatActive Then
        gHeartbeatNextRun = Now + TimeSerial(0, 0, gHeartbeatSeconds)
        Application.OnTime gHeartbeatNextRun, "modBridge.HeartbeatTick"
    End If
End Sub

' ============================================================
'  SELF-TEST
'  Run from the Immediate window:
'      modBridge.BridgeSelfTest
'      modBridge.BridgeSelfTest "https://your-endpoint/health"
' ============================================================

Public Sub BridgeSelfTest(Optional testUrl As String = "")
    On Error Resume Next

    ' [1] Is the XLL loaded in THIS Excel instance?
    Err.Clear
    Dim ver As Variant
    ver = Application.Run("EB_Version")
    If Err.Number <> 0 Then
        Debug.Print "[1] EB_Version not callable (err " & Err.Number & "): XLL not loaded here. Trying EnsureAddin..."
        Err.Clear
        EnsureAddin
        If Err.Number <> 0 Then
            Debug.Print "    EnsureAddin failed: " & Err.Description
            Exit Sub
        End If
        ver = Application.Run("EB_Version")
        If Err.Number <> 0 Then
            Debug.Print "    Still not callable after RegisterXLL. Check: 64/32-bit XLL flavor, " & _
                        "and that the file isn't blocked (right-click -> Properties -> Unblock)."
            Exit Sub
        End If
    End If
    Debug.Print "[1] XLL loaded. EB_Version = " & CStr(ver)
    If CStr(ver) <> "2.0.0" Then
        Debug.Print "    WARNING: expected 2.0.0 - a stale build is loaded. Close Excel and replace the XLL."
    End If
    Err.Clear
    Dim diag As Variant
    diag = Application.Run("EB_Diag")
    If Err.Number = 0 Then
        Debug.Print "    " & CStr(diag)
        If InStr(1, CStr(diag), "newtonsoft=MISSING", vbTextCompare) > 0 Then
            Debug.Print "    -> Dependencies missing: deploy the single-file " & _
                        "ExcelBridge-AddIn64-packed.xll (the unpacked XLL needs " & _
                        "ExcelBridge.dll AND Newtonsoft.Json.dll beside it). Restart Excel after swapping."
            Exit Sub
        End If
    End If

    ' [2] Attach this workbook as the callback target
    Err.Clear
    EnsureHttp
    If Err.Number <> 0 Then
        Debug.Print "[2] EnsureHttp failed: " & Err.Description
        Exit Sub
    End If
    Debug.Print "[2] Attached '" & ThisWorkbook.name & "' (EB_Attach)"

    If Len(testUrl) = 0 Then
        Debug.Print "[3] Skipped HTTP round trip - pass a URL: BridgeSelfTest ""https://..."""
        Exit Sub
    End If

    ' [3] Blocking round trip (exercises Run marshaling + chunked results)
    Err.Clear
    Dim r As cHttpResponse
    Set r = modHttp.HttpGetSync(testUrl, "", 15000)
    Debug.Print "[3] sync GET -> status " & r.Status & ", " & Len(r.body) & " chars, " & _
                r.ElapsedMs & "ms" & IIf(Len(r.ErrorMsg) > 0, ", error: " & r.ErrorMsg, "")

    ' [4] Async round trip (exercises QueueAsMacro -> EB_OnHttpBatch delivery)
    Err.Clear
    Dim b As cHttpBatch
    Set b = New cHttpBatch
    b.Init
    b.AddGet testUrl, "", 15000
    If b.WaitAll(20000) Then
        Dim ids As Variant: ids = b.ids()
        Dim resp As cHttpResponse
        Set resp = b.GetResult(CStr(ids(LBound(ids))))
        Debug.Print "[4] async GET delivered -> status " & resp.Status & _
                    IIf(Len(resp.ErrorMsg) > 0, ", error: " & resp.ErrorMsg, "")
    Else
        Debug.Print "[4] async GET TIMED OUT. EB_LastDispatchError = " & _
                    CStr(Application.Run("EB_LastDispatchError"))
    End If

    On Error GoTo 0
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
