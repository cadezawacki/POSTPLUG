
Attribute VB_Name = "modBridgeEvents"

Option Explicit

' ============================================================
'  Sinks invoked BY the ExcelBridge XLL via Application.Run,
'  always on Excel's main thread (QueueAsMacro delivery).
'
'  The sub names and signatures below are a fixed contract with
'  the add-in (see VbaDispatcher in ExcelInterface.cs) - do not
'  rename them. This module replaces the old cBridgeHost class
'  (COM WithEvents sink).
' ============================================================

' ============================================================
'  HTTP: batched completions
'  rows = { requestId, status, body, headersJson, elapsedMs, errorMsg }
' ============================================================

Public Sub EB_OnHttpBatch(ByVal results As Variant)
    Dim rLo As Long: rLo = LBound(results, 1)
    Dim rHi As Long: rHi = UBound(results, 1)
    Dim c As Long: c = LBound(results, 2)

    Dim i As Long
    Dim requestId As String
    Dim resp As cHttpResponse

    For i = rLo To rHi
        Set resp = New cHttpResponse
        requestId = CStr(results(i, c))
        resp.Status = CLng(results(i, c + 1))
        resp.body = CStr(results(i, c + 2))
        resp.headersJson = CStr(results(i, c + 3))
        resp.ElapsedMs = CLng(results(i, c + 4))
        resp.ErrorMsg = CStr(results(i, c + 5))

        modBridge.RouteHttpResult requestId, resp

        ' Diagnostics are opt-in: with many parallel requests the per-response
        ' sheet write + Debug.Print dominate the completion path.
        If modBridge.gHttpVerbose Then
            If Len(resp.ErrorMsg) > 0 Then
                Debug.Print "HTTP ERROR [" & requestId & "] " & resp.ElapsedMs & "ms: " & resp.ErrorMsg
            Else
                Debug.Print "HTTP " & resp.Status & " [" & requestId & "] " & _
                            resp.ElapsedMs & "ms: " & Left$(resp.body, 200)
                LogHttpToSheet requestId, resp.Status, resp.ElapsedMs, resp.body
            End If
        End If
    Next i
End Sub

' ============================================================
'  WebSocket: routed messages
' ============================================================

Public Sub EB_OnMessage(ByVal msgType As String, ByVal jsonPayload As String)
    modSocketRouter.RouteMessage msgType, jsonPayload
End Sub

Public Sub EB_OnMessageBatch(ByVal jsonArray As String)
    modSocketRouter.RouteBatch jsonArray
End Sub

' ============================================================
'  WebSocket: typed grid events
' ============================================================

Public Sub EB_OnCellUpdate(ByVal r As Long, ByVal c As Long, ByVal v As String)
    On Error GoTo cleanup
    gSuspendEvents = True
    ThisWorkbook.Sheets("Grid").cells(r + 1, c + 1).value = v
cleanup:
    gSuspendEvents = False
End Sub

Public Sub EB_OnCellBatch(ByVal cells As Variant)
    On Error GoTo cleanup
    gSuspendEvents = True
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Grid")
    Dim rLo As Long: rLo = LBound(cells, 1)
    Dim cLo As Long: cLo = LBound(cells, 2)
    Dim i As Long
    For i = rLo To UBound(cells, 1)
        ws.cells(CLng(cells(i, cLo)) + 1, CLng(cells(i, cLo + 1)) + 1).value = cells(i, cLo + 2)
    Next i
cleanup:
    gSuspendEvents = False
End Sub

Public Sub EB_OnFullGrid(ByVal jsonGrid As String)
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

Public Sub EB_OnWsStatus(ByVal state As String, ByVal detail As String)
    If Len(detail) > 0 Then
        SetStatus state & ": " & detail
    Else
        SetStatus state
    End If
End Sub

Public Sub EB_OnLog(ByVal level As String, ByVal Message As String)
    Debug.Print "[" & level & "] " & Message
End Sub

' ============================================================
'  Helpers
' ============================================================

Private Sub SetStatus(s As String)
    On Error Resume Next
    Dim prevEvents As Boolean: prevEvents = Application.enableEvents
    Application.enableEvents = False
    ThisWorkbook.Sheets("Grid").Range("G1").value = s
    Application.enableEvents = prevEvents
    On Error GoTo 0
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
