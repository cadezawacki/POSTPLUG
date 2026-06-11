
Attribute VB_Name = "modSocketRouter"

Option Explicit

' ============================================================
'  Dictionary-based WebSocket message router.
'  Routes by field value or field presence with O(1) lookup.
' ============================================================

' VALUE routes: {fieldName: {fieldValue: callbackId}}
Private gValueRoutes As Object

' PRESENCE routes: {fieldName: callbackId}
Private gPresenceRoutes As Object

' Fallback callback (receives msgType, jsonPayload)
Private gFallbackCallback As String

' Pre-computed list of field names to inspect
Private gWatchKeys() As String
Private gWatchKeyCount As Long

' ============================================================
'  INITIALIZATION
' ============================================================

' Idempotent: initializes the router only on first call, preserving any
' routes already registered. Use InitRouter to force a full reset.
Public Sub EnsureRouter()
    If gValueRoutes Is Nothing Then InitRouter
End Sub

Public Sub InitRouter()
    Set gValueRoutes = CreateObject("Scripting.Dictionary")
    gValueRoutes.CompareMode = vbTextCompare
    Set gPresenceRoutes = CreateObject("Scripting.Dictionary")
    gPresenceRoutes.CompareMode = vbTextCompare
    gFallbackCallback = ""
    gWatchKeyCount = 0
    ReDim gWatchKeys(0 To 0)
End Sub

' ============================================================
'  REGISTRATION API
' ============================================================

Public Sub RouteOnValue(fieldName As String, fieldValue As String, callbackId As String)
    If Not gValueRoutes.exists(fieldName) Then
        Dim inner As Object
        Set inner = CreateObject("Scripting.Dictionary")
        inner.CompareMode = vbTextCompare
        gValueRoutes.Add fieldName, inner
    End If
    gValueRoutes(fieldName)(fieldValue) = callbackId
    RebuildWatchKeys
End Sub

Public Sub RouteOnPresence(fieldName As String, callbackId As String)
    If gPresenceRoutes.exists(fieldName) Then
        gPresenceRoutes(fieldName) = callbackId
    Else
        gPresenceRoutes.Add fieldName, callbackId
    End If
    RebuildWatchKeys
End Sub

Public Sub RouteDefault(callbackId As String)
    gFallbackCallback = callbackId
End Sub

Public Sub UnrouteValue(fieldName As String, fieldValue As String)
    If gValueRoutes.exists(fieldName) Then
        Dim inner As Object: Set inner = gValueRoutes(fieldName)
        If inner.exists(fieldValue) Then inner.Remove fieldValue
        If inner.count = 0 Then gValueRoutes.Remove fieldName
    End If
    RebuildWatchKeys
End Sub

Public Sub UnroutePresence(fieldName As String)
    If gPresenceRoutes.exists(fieldName) Then gPresenceRoutes.Remove fieldName
    RebuildWatchKeys
End Sub

' ============================================================
'  DISPATCH � SINGLE MESSAGE
' ============================================================

Public Sub RouteMessage(msgType As String, jsonPayload As String)
    On Error GoTo RouteErr

    ' Fast path: check "type" value routes first
    If Len(msgType) > 0 And gValueRoutes.exists("type") Then
        Dim typeRoutes As Object: Set typeRoutes = gValueRoutes("type")
        If typeRoutes.exists(msgType) Then
            Application.Run typeRoutes(msgType), jsonPayload
            Exit Sub
        End If
    End If

    ' Check other watched keys
    Dim i As Long
    For i = 0 To gWatchKeyCount - 1
        Dim k As String: k = gWatchKeys(i)
        If k = "type" Then GoTo NextKey

        Dim v As String
        v = ExtractJsonValue(jsonPayload, k)
        If Len(v) > 0 Then
            ' Value routes for this key
            If gValueRoutes.exists(k) Then
                Dim valDict As Object: Set valDict = gValueRoutes(k)
                If valDict.exists(v) Then
                    Application.Run valDict(v), jsonPayload
                    Exit Sub
                End If
            End If
            ' Presence route
            If gPresenceRoutes.exists(k) Then
                Application.Run gPresenceRoutes(k), jsonPayload
                Exit Sub
            End If
        End If
NextKey:
    Next i

    ' Fallback
    If Len(gFallbackCallback) > 0 Then
        Application.Run gFallbackCallback, msgType, jsonPayload
    End If
    Exit Sub

RouteErr:
    Debug.Print "RouteMessage error: " & Err.Description & " | key=" & msgType
End Sub

' ============================================================
'  DISPATCH � BATCH (splits JSON array, routes each)
' ============================================================

Public Sub RouteBatch(jsonArray As String)
    Dim pos As Long, msgStart As Long, depth As Long
    Dim msgJson As String
    Dim arrLen As Long: arrLen = Len(jsonArray)

    ' Find first '{'
    pos = InStr(1, jsonArray, "{")
    If pos = 0 Then Exit Sub

    msgStart = pos
    depth = 0

    Do While pos <= arrLen
        Select Case mid$(jsonArray, pos, 1)
            Case "{"
                depth = depth + 1
            Case "}"
                depth = depth - 1
                If depth = 0 Then
                    msgJson = mid$(jsonArray, msgStart, pos - msgStart + 1)
                    RouteSingleFromBatch msgJson
                    ' Advance to next '{'
                    pos = InStr(pos + 1, jsonArray, "{")
                    If pos = 0 Then Exit Do
                    msgStart = pos
                    depth = 0
                End If
            Case """"
                ' Skip string contents (avoid counting braces inside strings)
                pos = pos + 1
                Do While pos <= arrLen
                    Select Case mid$(jsonArray, pos, 1)
                        Case "\"
                            pos = pos + 1 ' skip escaped char
                        Case """"
                            Exit Do
                    End Select
                    pos = pos + 1
                Loop
        End Select
        pos = pos + 1
    Loop
End Sub

Private Sub RouteSingleFromBatch(json As String)
    On Error Resume Next
    Dim t As String: t = ExtractJsonValue(json, "type")
    RouteMessage t, json
    On Error GoTo 0
End Sub

' ============================================================
'  LIGHTWEIGHT JSON KEY EXTRACTOR
'  Finds top-level "key":"value" � O(1) for short messages.
' ============================================================

Public Function ExtractJsonValue(json As String, key As String) As String
    Dim needle As String: needle = """" & key & """:"
    Dim p As Long: p = InStr(1, json, needle, vbBinaryCompare)
    If p = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    p = p + Len(needle)
    ' Skip whitespace
    Do While p <= Len(json)
        Select Case mid$(json, p, 1)
            Case " ", vbTab: p = p + 1
            Case Else: Exit Do
        End Select
    Loop

    If p > Len(json) Then ExtractJsonValue = "": Exit Function

    Dim ch As String: ch = mid$(json, p, 1)

    If ch = """" Then
        ' String value: find closing unescaped quote
        Dim endP As Long: endP = p + 1
        Do While endP <= Len(json)
            Select Case mid$(json, endP, 1)
                Case "\"
                    endP = endP + 2
                Case """"
                    Exit Do
                Case Else
                    endP = endP + 1
            End Select
        Loop
        ExtractJsonValue = mid$(json, p + 1, endP - p - 1)
    ElseIf ch = "{" Or ch = "[" Then
        ' Nested object/array � not a simple routable value
        ExtractJsonValue = ""
    Else
        ' Number, bool, null: read until delimiter
        Dim endN As Long: endN = p
        Do While endN <= Len(json)
            Select Case mid$(json, endN, 1)
                Case ",", "}", "]", " ", vbCr, vbLf, vbTab
                    Exit Do
                Case Else
                    endN = endN + 1
            End Select
        Loop
        ExtractJsonValue = mid$(json, p, endN - p)
    End If
End Function

' ============================================================
'  JSON UTILITY FUNCTIONS
' ============================================================

Public Function JsonHasKey(json As String, key As String) As Boolean
    ' O(1) check for key existence
    Dim needle As String: needle = """" & key & """:"
    JsonHasKey = (InStr(1, json, needle, vbBinaryCompare) > 0)
End Function

Public Function JsonGet(json As String, key As String, Optional fallback As String = "") As String
    ' Extract value with fallback default if key missing
    Dim v As String: v = ExtractJsonValue(json, key)
    If Len(v) = 0 Then
        JsonGet = fallback
    Else
        JsonGet = v
    End If
End Function

Public Function JsonGetOrError(json As String, key As String, Optional errorCallbackId As String = "") As String
    ' Extract value; fire error callback if missing
    Dim v As String: v = ExtractJsonValue(json, key)
    If Len(v) = 0 Then
        If Len(errorCallbackId) > 0 Then
            On Error Resume Next
            Application.Run errorCallbackId, key, json
            On Error GoTo 0
        End If
        JsonGetOrError = ""
    Else
        JsonGetOrError = v
    End If
End Function

Public Function JsonGetLong(json As String, key As String, Optional fallback As Long = 0) As Long
    Dim v As String: v = ExtractJsonValue(json, key)
    If Len(v) = 0 Then
        JsonGetLong = fallback
    Else
        JsonGetLong = CLng(v)
    End If
End Function

Public Function JsonGetDouble(json As String, key As String, Optional fallback As Double = 0) As Double
    Dim v As String: v = ExtractJsonValue(json, key)
    If Len(v) = 0 Then
        JsonGetDouble = fallback
    Else
        JsonGetDouble = CDbl(v)
    End If
End Function

Public Function JsonGetBool(json As String, key As String, Optional fallback As Boolean = False) As Boolean
    Dim v As String: v = ExtractJsonValue(json, key)
    If Len(v) = 0 Then
        JsonGetBool = fallback
    Else
        JsonGetBool = (v = "true" Or v = "True" Or v = "1")
    End If
End Function

Public Function JsonKeys(json As String) As String()
    ' Returns array of all top-level key names in the JSON object.
    ' Scans for "key": patterns at depth 0.
    Dim result() As String
    Dim n As Long: n = 0
    ReDim result(0 To 31)

    Dim arrLen As Long: arrLen = Len(json)
    Dim pos As Long: pos = 1
    Dim depth As Long: depth = 0
    Dim inStr_ As Boolean: inStr_ = False
    Dim atRoot As Boolean

    Do While pos <= arrLen
        Dim ch As String: ch = mid$(json, pos, 1)

        If inStr_ Then
            If ch = "\" Then
                pos = pos + 1
            ElseIf ch = """" Then
                inStr_ = False
            End If
            pos = pos + 1
            GoTo ContinueKeysLoop
        End If

        Select Case ch
            Case """"
                ' Check if this is a root-level key (depth = 1, preceded by { or ,)
                If depth = 1 Then
                    ' Find the closing quote
                    Dim keyStart As Long: keyStart = pos + 1
                    Dim keyEnd As Long: keyEnd = keyStart
                    Do While keyEnd <= arrLen
                        Select Case mid$(json, keyEnd, 1)
                            Case "\"
                                keyEnd = keyEnd + 1
                            Case """"
                                Exit Do
                            Case Else
                        End Select
                        keyEnd = keyEnd + 1
                    Loop
                    ' Check if followed by ":"
                    Dim afterQuote As Long: afterQuote = keyEnd + 1
                    Do While afterQuote <= arrLen And mid$(json, afterQuote, 1) = " "
                        afterQuote = afterQuote + 1
                    Loop
                    If afterQuote <= arrLen And mid$(json, afterQuote, 1) = ":" Then
                        ' It's a key
                        If n > UBound(result) Then ReDim Preserve result(0 To n * 2)
                        result(n) = mid$(json, keyStart, keyEnd - keyStart)
                        n = n + 1
                    End If
                    pos = keyEnd
                End If
                inStr_ = True
            Case "{", "["
                depth = depth + 1
            Case "}", "]"
                depth = depth - 1
        End Select
        pos = pos + 1
ContinueKeysLoop:
    Loop

    If n = 0 Then
        ReDim result(0 To 0)
        result(0) = ""
    Else
        ReDim Preserve result(0 To n - 1)
    End If
    JsonKeys = result
End Function

Public Function JsonValues(json As String, keys() As String) As String()
    ' Bulk extract: returns parallel array of values for given keys.
    ' Faster than calling ExtractJsonValue N times for many keys.
    Dim n As Long: n = UBound(keys) - LBound(keys) + 1
    Dim result() As String
    ReDim result(LBound(keys) To UBound(keys))

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        result(i) = ExtractJsonValue(json, keys(i))
    Next i
    JsonValues = result
End Function

Public Function JsonPick(json As String, ParamArray keys() As Variant) As String()
    ' Extract multiple values by key names. Returns array in same order.
    ' Usage: vals = JsonPick(json, "ticker", "action", "notional")
    Dim n As Long: n = UBound(keys) - LBound(keys) + 1
    Dim result() As String
    ReDim result(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        result(i) = ExtractJsonValue(json, CStr(keys(i)))
    Next i
    JsonPick = result
End Function

Public Function JsonToDict(json As String) As Object
    ' Parse all top-level keys into a Scripting.Dictionary.
    ' Useful when you need to access many fields from one message.
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbBinaryCompare

    Dim k() As String: k = JsonKeys(json)
    If Len(k(0)) = 0 And UBound(k) = 0 Then
        Set JsonToDict = d
        Exit Function
    End If

    Dim i As Long
    For i = LBound(k) To UBound(k)
        If Not d.exists(k(i)) Then
            d.Add k(i), ExtractJsonValue(json, k(i))
        End If
    Next i
    Set JsonToDict = d
End Function

Public Function JsonCount(json As String) As Long
    ' Count top-level keys in a JSON object
    Dim k() As String: k = JsonKeys(json)
    If Len(k(0)) = 0 And UBound(k) = 0 Then
        JsonCount = 0
    Else
        JsonCount = UBound(k) - LBound(k) + 1
    End If
End Function

' ============================================================
'  INTERNAL
' ============================================================

Private Sub RebuildWatchKeys()
    Dim allKeys As Object: Set allKeys = CreateObject("Scripting.Dictionary")
    allKeys.CompareMode = vbTextCompare

    Dim k As Variant
    For Each k In gValueRoutes.keys
        If Not allKeys.exists(CStr(k)) Then allKeys.Add CStr(k), True
    Next k
    For Each k In gPresenceRoutes.keys
        If Not allKeys.exists(CStr(k)) Then allKeys.Add CStr(k), True
    Next k

    gWatchKeyCount = allKeys.count
    If gWatchKeyCount = 0 Then
        ReDim gWatchKeys(0 To 0)
        Exit Sub
    End If

    ReDim gWatchKeys(0 To gWatchKeyCount - 1)
    Dim i As Long: i = 0
    For Each k In allKeys.keys
        gWatchKeys(i) = CStr(k)
        i = i + 1
    Next k
End Sub



