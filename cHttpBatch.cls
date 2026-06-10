
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHttpBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mResults As Object          ' Scripting.Dictionary: id -> cHttpResponse
Private mPending As Long
Private mCallback As String         ' optional: name of a Public Sub to call on each completion
Private mCallbackTarget As Object   ' optional: object that has the callback method
Private mSimpleCallback As String
Private mAllCompleteCallback As String  ' optional: Public Sub fired once when the last request completes

Public Sub Init()
    Set mResults = CreateObject("Scripting.Dictionary")
    mPending = 0
    mCallback = ""
    Set mCallbackTarget = Nothing
    mSimpleCallback = ""
    mAllCompleteCallback = ""
End Sub

' Per-request callback: invoked each time a request completes.
' The target object must implement:  Public Sub <callbackName>(requestId As String, r As cHttpResponse)
Public Sub SetCallback(Target As Object, methodName As String)
    Set mCallbackTarget = Target
    mCallback = methodName
End Sub

Public Sub SetSimpleCallback(methodName As String)
    mSimpleCallback = methodName
End Sub

' Fired once, asynchronously, when the final pending request completes.
' The target must be:  Public Sub <methodName>(batch As cHttpBatch)
Public Sub SetAllCompleteCallback(methodName As String)
    mAllCompleteCallback = methodName
End Sub

' Add a fire-and-forget request. Returns the request id.
Public Function Add(method As String, URL As String, _
                    Optional body As String = "", _
                    Optional headers As String = "", _
                    Optional timeoutMs As Long = 0) As String
    Dim id As String
    id = modHttp.HttpSendFF(method, URL, body, headers, timeoutMs)

    ' Store Nothing as a placeholder; replaced when the result arrives
    mResults.Add id, Nothing
    mPending = mPending + 1

    modBridge.RegisterBatch id, Me
    Add = id
End Function

' Convenience wrappers
Public Function AddGet(URL As String, Optional headers As String = "", Optional timeoutMs As Long = 0) As String
    AddGet = Add("GET", URL, "", headers, timeoutMs)
End Function

Public Function AddPost(URL As String, body As String, Optional headers As String = "", Optional timeoutMs As Long = 0) As String
    AddPost = Add("POST", URL, body, headers, timeoutMs)
End Function

Public Function AddPut(URL As String, body As String, Optional headers As String = "", Optional timeoutMs As Long = 0) As String
    AddPut = Add("PUT", URL, body, headers, timeoutMs)
End Function

Public Function AddDelete(URL As String, Optional headers As String = "", Optional timeoutMs As Long = 0) As String
    AddDelete = Add("DELETE", URL, "", headers, timeoutMs)
End Function

' Block until every request in the batch has completed or errored,
' or until overallTimeoutMs has elapsed (0 = no overall cap).
' Returns True if all completed, False if the overall timeout fired.
Public Function WaitAll(Optional overallTimeoutMs As Long = 0) As Boolean
    Dim startTick As Double
    startTick = Timer

    Do While mPending > 0
        DoEvents   ' let COM events deliver OnHttpResponse / OnHttpError into cBridgeHost
        Sleep 5    ' tiny yield so we're not burning 100% CPU

        If overallTimeoutMs > 0 Then
            Dim elapsedSec As Double
            elapsedSec = Timer - startTick
            If elapsedSec < 0 Then elapsedSec = elapsedSec + 86400  ' Timer resets at midnight
            If elapsedSec * 1000 > overallTimeoutMs Then
                WaitAll = False
                Exit Function
            End If
        End If
    Loop
    WaitAll = True
End Function

' Called by modBridge.RouteHttpResult when a matching request completes.
' Do NOT call directly.
Friend Sub OnComplete(requestId As String, r As cHttpResponse)
    If Not mResults.exists(requestId) Then Exit Sub
    Set mResults(requestId) = r
    mPending = mPending - 1

    If Len(mCallback) > 0 And Not mCallbackTarget Is Nothing Then
        On Error Resume Next
        CallByName mCallbackTarget, mCallback, VbMethod, requestId, r
        On Error GoTo 0
    ElseIf Len(mSimpleCallback) > 0 Then
        On Error Resume Next
        Application.Run mSimpleCallback, requestId, r
        On Error GoTo 0
    End If

    If mPending = 0 And Len(mAllCompleteCallback) > 0 Then
        On Error Resume Next
        Application.Run mAllCompleteCallback, Me
        On Error GoTo 0
    End If
End Sub

' Retrieve a single result by id
Public Function GetResult(requestId As String) As cHttpResponse
    If mResults.exists(requestId) Then
        Set GetResult = mResults(requestId)
    End If
End Function

' Enumerate all ids
Public Function ids() As Variant
    ids = mResults.keys
End Function

Public Property Get count() As Long
    count = mResults.count
End Property

Public Property Get Pending() As Long
    Pending = mPending
End Property




