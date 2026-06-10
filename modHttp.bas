
Attribute VB_Name = "modHttp"

Option Explicit

' ============================================================
'  Async HTTP via the ExcelBridge XLL. No WebSocket required.
'
'  Fire-and-forget  (returns requestId immediately, response
'  arrives via modBridgeEvents.EB_OnHttpBatch on the main thread):
'
'      Dim id As String
'      id = HttpGet("https://api.example.com/users/42")
'      id = HttpPost("https://api.example.com/users", _
'                    "{""name"":""Alice""}")
'
'  Blocking (returns cHttpResponse record; blocks VBA only):
'
'      Dim r As cHttpResponse
'      Set r = HttpGetSync("https://api.example.com/users/42")
'      If r.Status = 200 Then Debug.Print r.Body
'
'  Custom headers (one per line):
'
'      Set r = HttpRequestSync("POST", url, jsonBody, _
'          "Authorization: Bearer abc" & vbLf & _
'          "X-Trace-Id: 42")
'
'  Cancel an in-flight fire-and-forget request:
'
'      HttpCancel id
' ============================================================

' Excel's evaluator caps VBA -> XLL string arguments at 32,767 chars;
' bodies above this are staged into the add-in in chunks.
Private Const CHUNK_MAX As Long = 30000

' ---------- Fire-and-forget ----------

Public Function HttpGet(URL As String, _
                        Optional headers As String = "", _
                        Optional timeoutMs As Long = 0) As String
    HttpGet = HttpSendFF("GET", URL, "", headers, timeoutMs)
End Function

Public Function HttpPost(URL As String, body As String, _
                         Optional headers As String = "", _
                         Optional timeoutMs As Long = 0) As String
    HttpPost = HttpSendFF("POST", URL, body, headers, timeoutMs)
End Function

Public Function HttpPut(URL As String, body As String, _
                        Optional headers As String = "", _
                        Optional timeoutMs As Long = 0) As String
    HttpPut = HttpSendFF("PUT", URL, body, headers, timeoutMs)
End Function

Public Function HttpDelete(URL As String, _
                           Optional headers As String = "", _
                           Optional timeoutMs As Long = 0) As String
    HttpDelete = HttpSendFF("DELETE", URL, "", headers, timeoutMs)
End Function

Public Function HttpPatch(URL As String, body As String, _
                          Optional headers As String = "", _
                          Optional timeoutMs As Long = 0) As String
    HttpPatch = HttpSendFF("PATCH", URL, body, headers, timeoutMs)
End Function

Public Function HttpSendFF(method As String, URL As String, body As String, _
                           headers As String, timeoutMs As Long) As String
    ' EnsureHttp loads/attaches the XLL only - it never opens a WebSocket.
    modBridge.EnsureHttp
    If Len(body) > CHUNK_MAX Then
        HttpSendFF = Application.Run("EB_HttpSendAsyncBody", method, URL, _
                                     StageBody(body), headers, timeoutMs)
    Else
        HttpSendFF = Application.Run("EB_HttpSendAsync", method, URL, _
                                     body, headers, timeoutMs)
    End If
End Function

Public Sub HttpCancel(requestId As String)
    If Not modBridge.AddinLoaded() Then Exit Sub
    Application.Run "EB_HttpCancel", requestId
End Sub

' ---------- Blocking ----------

Public Function HttpGetSync(URL As String, _
                            Optional headers As String = "", _
                            Optional timeoutMs As Long = 0) As cHttpResponse
    Set HttpGetSync = HttpRequestSync("GET", URL, "", headers, timeoutMs)
End Function

Public Function HttpPostSync(URL As String, body As String, _
                             Optional headers As String = "", _
                             Optional timeoutMs As Long = 0) As cHttpResponse
    Set HttpPostSync = HttpRequestSync("POST", URL, body, headers, timeoutMs)
End Function

Public Function HttpPutSync(URL As String, body As String, _
                            Optional headers As String = "", _
                            Optional timeoutMs As Long = 0) As cHttpResponse
    Set HttpPutSync = HttpRequestSync("PUT", URL, body, headers, timeoutMs)
End Function

Public Function HttpDeleteSync(URL As String, _
                               Optional headers As String = "", _
                               Optional timeoutMs As Long = 0) As cHttpResponse
    Set HttpDeleteSync = HttpRequestSync("DELETE", URL, "", headers, timeoutMs)
End Function

Public Function HttpPatchSync(URL As String, body As String, _
                              Optional headers As String = "", _
                              Optional timeoutMs As Long = 0) As cHttpResponse
    Set HttpPatchSync = HttpRequestSync("PATCH", URL, body, headers, timeoutMs)
End Function

Public Function HttpRequestSync(method As String, URL As String, body As String, _
                                headers As String, timeoutMs As Long) As cHttpResponse
    Dim r As cHttpResponse
    Set r = New cHttpResponse

    On Error GoTo fail
    modBridge.EnsureHttp

    ' Result row: { status, inlineBody, headersJson, elapsedMs, error, bodyToken, bodyLength }
    Dim result As Variant
    If Len(body) > CHUNK_MAX Then
        result = Application.Run("EB_HttpSendSyncBody", method, URL, _
                                 StageBody(body), headers, timeoutMs)
    Else
        result = Application.Run("EB_HttpSendSync", method, URL, _
                                 body, headers, timeoutMs)
    End If

    Dim v() As Variant
    v = NormalizeRow(result)

    r.Status = CLng(v(0))
    r.body = CStr(v(1))
    r.headersJson = CStr(v(2))
    r.ElapsedMs = CLng(v(3))
    r.ErrorMsg = CStr(v(4))

    ' Bodies too large to return inline are fetched in chunks
    Dim bodyToken As String: bodyToken = CStr(v(5))
    If Len(bodyToken) > 0 Then
        r.body = FetchResult(bodyToken, CLng(v(6)))
    End If

    Set HttpRequestSync = r
    Exit Function
fail:
    r.ErrorMsg = "VBA error: " & Err.Description
    Set HttpRequestSync = r
End Function

' ---------- Internal helpers ----------

' Stage a large request body into the add-in, CHUNK_MAX chars at a time.
Private Function StageBody(body As String) As String
    Dim token As String
    token = Application.Run("EB_BodyBegin")

    Dim p As Long: p = 1
    Do While p <= Len(body)
        Application.Run "EB_BodyAppend", token, mid$(body, p, CHUNK_MAX)
        p = p + CHUNK_MAX
    Loop
    StageBody = token
End Function

' Fetch a large sync-response body from the add-in chunk store.
Private Function FetchResult(token As String, totalLen As Long) As String
    If totalLen <= 0 Then
        Application.Run "EB_ResultRelease", token
        Exit Function
    End If

    Dim buf As String: buf = Space$(totalLen)
    Dim pos As Long: pos = 1
    Dim idx As Long: idx = 0
    Dim chunk As String

    Do While pos <= totalLen
        chunk = Application.Run("EB_ResultChunk", token, idx)
        If Len(chunk) = 0 Then Exit Do
        mid$(buf, pos, Len(chunk)) = chunk
        pos = pos + Len(chunk)
        idx = idx + 1
    Loop

    Application.Run "EB_ResultRelease", token
    FetchResult = Left$(buf, pos - 1)
End Function

' Application.Run may hand back the XLL result row as a 1-D array (any base)
' or a 2-D single-row array depending on marshaling. Normalize to 0-based 1-D.
Private Function NormalizeRow(ByVal result As Variant) As Variant()
    Dim out() As Variant
    Dim n As Long, i As Long

    Dim ub2 As Long: ub2 = -2147483647
    On Error Resume Next
    ub2 = UBound(result, 2)
    On Error GoTo 0

    If ub2 <> -2147483647 Then
        ' 2-D single row
        n = ub2 - LBound(result, 2) + 1
        ReDim out(0 To n - 1)
        For i = 0 To n - 1
            out(i) = result(LBound(result, 1), LBound(result, 2) + i)
        Next i
    Else
        n = UBound(result) - LBound(result) + 1
        ReDim out(0 To n - 1)
        For i = 0 To n - 1
            out(i) = result(LBound(result) + i)
        Next i
    End If
    NormalizeRow = out
End Function
