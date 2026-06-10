
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cBridgeHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
'  cBridgeHost: hosts the ExcelBridge.Bridge COM object and
'  receives its events via WithEvents.
' ============================================================

Public WithEvents b As ExcelBridge.Bridge
Attribute b.VB_VarHelpID = -1

' Create the COM object only. HTTP works from this point on;
' no WebSocket connection is opened until StartWs is called.
Public Sub EnsureObject()
    If b Is Nothing Then
        Set b = New ExcelBridge.Bridge
        b.AutoReconnect = True
        b.MaxReconnectDelayMs = 15000
        b.RaiseGenericOnMessage = False
        b.HttpDefaultTimeoutMs = 30000
    End If
End Sub

' Start (or restart) the WebSocket loop on the SAME Bridge object so that
' in-flight HTTP requests keep their event sink.
Public Sub StartWs(URL As String)
    EnsureObject
    b.Start URL
End Sub

Public Sub StartUp(URL As String)
    StartWs URL
End Sub

Public Sub ShutDown()
    On Error Resume Next
    If Not b Is Nothing Then b.Stop
    Set b = Nothing
    On Error GoTo 0
End Sub

Public Sub Push(r As Long, c As Long, v As String)
    If b Is Nothing Then Exit Sub
    b.PushCell r - 1, c - 1, v
End Sub

Public Sub PushBatch(arr As Variant)
    If b Is Nothing Then Exit Sub
    b.PushCells arr
End Sub

' ============================================================
'  Typed WebSocket events
' ============================================================

Private Sub b_OnCellUpdate(ByVal r As Long, ByVal c As Long, ByVal v As String)
    On Error GoTo cleanup
    gSuspendEvents = True
    ThisWorkbook.Sheets("Grid").cells(r + 1, c + 1).value = v
cleanup:
    gSuspendEvents = False
End Sub

Private Sub b_OnCellBatch(ByVal cells As Variant)
    On Error GoTo cleanup
    gSuspendEvents = True
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Grid")
    Dim i As Long
    For i = 0 To UBound(cells, 1)
        ws.cells(CLng(cells(i, 0)) + 1, CLng(cells(i, 1)) + 1).value = cells(i, 2)
    Next i
cleanup:
    gSuspendEvents = False
End Sub

Private Sub b_OnFullGrid(ByVal jsonGrid As String)
    On Error GoTo cleanup
    gSuspendEvents = True

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Grid")
    Dim rowsArr() As String
    Dim s As String: s = jsonGrid

    If Left$(s, 2) = "[[" Then s = mid$(s, 3)
    If Right$(s, 2) = "]]" Then s = Left$(s, Len(s) - 2)
    rowsArr = Split(s, "],[")

    Dim nRows As Long: nRows = UBound(rowsArr) - LBound(rowsArr) + 1
    If nRows = 0 Then GoTo cleanup

    Dim firstCols() As String
    firstCols = SplitJsonStrings(rowsArr(0))
    Dim nCols As Long: nCols = UBound(firstCols) - LBound(firstCols) + 1
    If nCols = 0 Then GoTo cleanup

    Dim vals() As String
    ReDim vals(1 To nRows, 1 To nCols)
    Dim r As Long, c As Long, cols() As String
    For r = 1 To nRows
        cols = SplitJsonStrings(rowsArr(r - 1))
        For c = 1 To nCols
            If (c - 1) <= UBound(cols) Then vals(r, c) = cols(c - 1)
        Next c
    Next r

    ws.Range(ws.cells(1, 1), ws.cells(nRows, nCols)).value = vals
cleanup:
    gSuspendEvents = False
End Sub

' ============================================================
'  Status / diagnostics
' ============================================================

Private Sub b_OnConnected()
    SetStatus "connected"
End Sub

Private Sub b_OnDisconnected(ByVal reason As String)
    SetStatus "disconnected: " & reason
End Sub

Private Sub b_OnReconnecting(ByVal attempt As Long, ByVal delayMs As Long)
    SetStatus "reconnecting (#" & attempt & ", " & delayMs & "ms)"
End Sub

Private Sub b_OnError(ByVal Message As String)
    Debug.Print "Bridge error: " & Message
End Sub

Private Sub b_OnLog(ByVal level As String, ByVal Message As String)
    Debug.Print "[" & level & "] " & Message
End Sub

Private Sub b_OnMessage(ByVal msgType As String, ByVal jsonPayload As String)
    modSocketRouter.RouteMessage msgType, jsonPayload
End Sub

Private Sub b_OnMessageBatch(ByVal jsonArray As String)
    modSocketRouter.RouteBatch jsonArray
End Sub

' ============================================================
'  HTTP events: fire-and-forget completions land here
' ============================================================

Private Sub b_OnHttpResponse(ByVal requestId As String, ByVal statusCode As Long, _
                             ByVal body As String, ByVal headersJson As String, _
                             ByVal ElapsedMs As Long)
    Dim r As cHttpResponse
    Set r = New cHttpResponse
    r.Status = statusCode
    r.body = body
    r.headersJson = headersJson
    r.ElapsedMs = ElapsedMs
    r.ErrorMsg = ""

    modBridge.RouteHttpResult requestId, r

    ' Diagnostics are opt-in: with many parallel requests the per-response
    ' sheet write + Debug.Print dominate the completion path.
    If modBridge.gHttpVerbose Then
        Debug.Print "HTTP " & statusCode & " [" & requestId & "] " & ElapsedMs & "ms: " & Left$(body, 200)
        LogHttpToSheet requestId, statusCode, ElapsedMs, body
    End If
End Sub

Private Sub LogHttpToSheet(requestId As String, ByVal statusCode As Long, _
                           ByVal ElapsedMs As Long, body As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HttpLog")
    If ws Is Nothing Then Exit Sub

    Dim prevEvents As Boolean: prevEvents = Application.enableEvents
    Application.enableEvents = False
    Dim nextRow As Long
    nextRow = ws.cells(ws.rows.count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    ws.cells(nextRow, 1).value = Now
    ws.cells(nextRow, 2).value = requestId
    ws.cells(nextRow, 3).value = statusCode
    ws.cells(nextRow, 4).value = ElapsedMs
    ws.cells(nextRow, 5).value = Left$(body, 32767)
    Application.enableEvents = prevEvents
    On Error GoTo 0
End Sub

Private Sub b_OnHttpError(ByVal requestId As String, ByVal Message As String, ByVal ElapsedMs As Long)
    Dim r As cHttpResponse
    Set r = New cHttpResponse
    r.Status = 0
    r.body = ""
    r.headersJson = "{}"
    r.ElapsedMs = ElapsedMs
    r.ErrorMsg = Message

    modBridge.RouteHttpResult requestId, r

    If modBridge.gHttpVerbose Then
        Debug.Print "HTTP ERROR [" & requestId & "] " & ElapsedMs & "ms: " & Message
    End If
End Sub


' ============================================================
'  Helpers
' ============================================================

Private Sub SetStatus(s As String)
    On Error Resume Next
    Application.enableEvents = False
    ThisWorkbook.Sheets("Grid").Range("G1").value = s
    Application.enableEvents = True
    On Error GoTo 0
End Sub

' Parse a JSON fragment like  "a","b","c with \"quote\""  into a string array.
' Handles escaped quotes, backslashes, and common escape sequences.
Private Function SplitJsonStrings(ByVal s As String) As String()
    Dim out() As String: ReDim out(0 To 63)
    Dim n As Long: n = 0
    Dim sLen As Long: sLen = Len(s)
    Dim i As Long, ch As String, nxt As String
    Dim inStr_ As Boolean
    ' Fixed-size scratch buffer written via Mid$ assignment: appending with
    ' "buf = buf & ch" reallocates the string per character (O(len^2)).
    Dim buf As String: buf = Space$(sLen)
    Dim bufN As Long: bufN = 0
    i = 1
    Do While i <= sLen
        ch = mid$(s, i, 1)
        If inStr_ Then
            If ch = "\" And i < sLen Then
                nxt = mid$(s, i + 1, 1)
                Select Case nxt
                    Case """": bufN = bufN + 1: mid$(buf, bufN, 1) = """": i = i + 1
                    Case "\":  bufN = bufN + 1: mid$(buf, bufN, 1) = "\":  i = i + 1
                    Case "/":  bufN = bufN + 1: mid$(buf, bufN, 1) = "/":  i = i + 1
                    Case "n":  bufN = bufN + 1: mid$(buf, bufN, 1) = vbLf: i = i + 1
                    Case "r":  bufN = bufN + 1: mid$(buf, bufN, 1) = vbCr: i = i + 1
                    Case "t":  bufN = bufN + 1: mid$(buf, bufN, 1) = vbTab: i = i + 1
                    Case Else: bufN = bufN + 1: mid$(buf, bufN, 1) = ch
                End Select
            ElseIf ch = """" Then
                If n > UBound(out) Then ReDim Preserve out(0 To (n + 1) * 2)
                out(n) = Left$(buf, bufN)
                n = n + 1
                bufN = 0
                inStr_ = False
            Else
                bufN = bufN + 1
                mid$(buf, bufN, 1) = ch
            End If
        ElseIf ch = """" Then
            inStr_ = True
        End If
        i = i + 1
    Loop
    If n = 0 Then
        ReDim out(0 To 0)
    Else
        ReDim Preserve out(0 To n - 1)
    End If
    SplitJsonStrings = out
End Function






