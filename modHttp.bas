
Attribute VB_Name = "modHttp"

Option Explicit

' ============================================================
'  Fire-and-forget  (returns requestId immediately, response
'  comes via cBridgeHost.b_OncHttpResponse / b_OnHttpError):
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
'      r = HttpRequestSync("POST", url, jsonBody, _
'          "Authorization: Bearer abc" & vbLf & _
'          "X-Trace-Id: 42")
'
'  Cancel an in-flight fire-and-forget request:
'
'      HttpCancel id
' ============================================================


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
    ' EnsureHttp, not EnsureBridge: REST calls must work with no WebSocket
    ' connected and must never trigger a socket (re)connect.
    modBridge.EnsureHttp
    HttpSendFF = modBridge.gHost.b.HttpSendAsync(method, URL, body, headers, timeoutMs)
End Function

Public Sub HttpCancel(requestId As String)
    If modBridge.gHost Is Nothing Then Exit Sub
    modBridge.gHost.b.HttpCancel requestId
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

    modBridge.EnsureHttp
    If modBridge.gHost Is Nothing Then
        r.ErrorMsg = "Bridge not started"
        Set HttpRequestSync = r
        Exit Function
    End If

    Dim result As Variant
    On Error GoTo fail
    result = modBridge.gHost.b.HttpSendSync(method, URL, body, headers, timeoutMs)

    r.Status = CLng(result(0))
    r.body = CStr(result(1))
    r.headersJson = CStr(result(2))
    r.ElapsedMs = CLng(result(3))
    r.ErrorMsg = CStr(result(4))
    Set HttpRequestSync = r
    Exit Function
fail:
    r.ErrorMsg = "VBA error: " & Err.Description
    Set HttpRequestSync = r
End Function






