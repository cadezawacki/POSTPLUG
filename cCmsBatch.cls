
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCmsBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
'  cCmsBatch - asynchronous GET/SET of any number of CMS curves.
'
'  Unlike the legacy CMS_GetBulkCurveQuoteData (one blocking SOAP round trip
'  for the whole batch) this launches the requests through the ExcelBridge
'  HTTP engine and surfaces results as they land:
'
'    * per-curve callback   - fired as each curve's response is parsed
'                             Public Sub MyOnCurve(curve As cCmsCurve, batch As cCmsBatch)
'    * all-complete callback- fired exactly once when the last request lands
'                             Public Sub MyOnAllDone(batch As cCmsBatch)
'
'  Both callbacks are invoked on Excel's main thread (the bridge delivers
'  completions via QueueAsMacro), so they may touch sheets safely.
'
'  Construct via modCms.CMS_GetCurvesAsync / CMS_SetCurvesAsync /
'  CMS_SetCurves5yAsync rather than directly.
' =============================================================================

Private mHttp As cHttpBatch
Private mAction As String           ' "GET" / "SET"
Private mCurves As Object           ' Scripting.Dictionary key -> cCmsCurve (insertion-ordered)
Private mReqKeys As Object          ' Scripting.Dictionary requestId -> Variant() of curve keys
Private mRowKeys As Variant         ' per input row: curve key ("" = blank/invalid row)
Private mOnCurve As String          ' "Module.Sub" run per curve
Private mOnAllDone As String        ' "Module.Sub" run once at the end
Private mPending As Long            ' outstanding HTTP requests
Private mDone As Boolean
Private mLaunched As Boolean
Private mFailedCount As Long
Private mLaunchTick As Double       ' Timer at FinishLaunch
Private mWallMs As Long             ' launch -> last delivery (wall clock)

Public Sub Init(ByVal Action As String, Optional ByVal OnCurve As String = "", _
                Optional ByVal OnAllDone As String = "")
    mAction = UCase$(Action)
    mOnCurve = OnCurve
    mOnAllDone = OnAllDone
    Set mCurves = CreateObject("Scripting.Dictionary")
    mCurves.CompareMode = vbTextCompare
    Set mReqKeys = CreateObject("Scripting.Dictionary")
    Set mHttp = New cHttpBatch
    mHttp.Init
    mHttp.SetCallback Me, "OnHttpDone"
    mPending = 0
    mDone = False
    mLaunched = False
    mFailedCount = 0
End Sub

' ---------------------------------------------------------------------------
'  Introspection
' ---------------------------------------------------------------------------

Public Property Get Action() As String
    Action = mAction
End Property

Public Property Get Count() As Long
    Count = mCurves.Count
End Property

Public Property Get PendingRequests() As Long
    PendingRequests = mPending
End Property

Public Property Get IsDone() As Boolean
    IsDone = mDone
End Property

Public Property Get FailedCount() As Long
    FailedCount = mFailedCount
End Property

' Curves whose outcome is known (OK or FAILED).
Public Property Get CompletedCurves() As Long
    Dim k As Variant
    Dim n As Long
    For Each k In mCurves.Keys
        If Len(mCurves(k).Status) > 0 Then n = n + 1
    Next k
    CompletedCurves = n
End Property

' Wall-clock ms from FinishLaunch to the last delivery (0 until done).
Public Property Get WallMs() As Long
    WallMs = mWallMs
End Property

' Slowest server round trip in the batch (engine-measured, excludes any
' delivery lag). WallMs >> MaxRequestMs means deliveries sat waiting in the
' queue; WallMs ~ MaxRequestMs means delivery was immediate.
Public Property Get MaxRequestMs() As Long
    Dim k As Variant
    For Each k In mCurves.Keys
        If mCurves(k).ElapsedMs > MaxRequestMs Then MaxRequestMs = mCurves(k).ElapsedMs
    Next k
End Property

Public Function Keys() As Variant
    Keys = mCurves.Keys
End Function

Public Function Curve(ByVal Key As String) As cCmsCurve
    If mCurves.Exists(Key) Then Set Curve = mCurves(Key)
End Function

Public Function HasCurve(ByVal Key As String) As Boolean
    HasCurve = mCurves.Exists(Key)
End Function

' ---------------------------------------------------------------------------
'  Launching (Friend-style: called by modCms only)
' ---------------------------------------------------------------------------

' Track a curve; returns False if the key was already present (duplicate row).
Public Function AddCurve(ByVal Curve As cCmsCurve) As Boolean
    If mCurves.Exists(Curve.Key) Then Exit Function
    mCurves.Add Curve.Key, Curve
    AddCurve = True
End Function

' Remember the per-input-row key layout so ToArray can mirror the caller's
' range row-for-row (duplicates and blank rows included).
Public Sub SetRowKeys(ByVal RowKeys As Variant)
    mRowKeys = RowKeys
End Sub

' Send one HTTP request covering the given curve keys.
Public Sub SendRequest(ByVal Url As String, ByVal Body As String, ByVal Headers As String, _
                       ByVal TimeoutMs As Long, ByVal CurveKeys As Variant)
    Dim id As String
    Dim k As Variant
    id = mHttp.AddPost(Url, Body, Headers, TimeoutMs)
    mReqKeys.Add id, CurveKeys
    For Each k In CurveKeys
        If mCurves.Exists(CStr(k)) Then mCurves(CStr(k)).RequestId = id
    Next k
    mPending = mPending + 1
End Sub

' Mark a curve as failed without sending anything (bad input, missing store
' entry, ...). The per-curve callback still fires so UIs see it.
Public Sub FailCurve(ByVal Curve As cCmsCurve, ByVal Reason As String)
    Curve.Status = modCms.CMS_STATUS_FAILED
    Curve.ErrorMsg = Reason
    mFailedCount = mFailedCount + 1
    modCms.NotifySubscribers Curve
    FireOnCurve Curve
End Sub

' Call after all SendRequest/FailCurve calls. If nothing was actually sent,
' completes the batch immediately (all-done callback included).
Public Sub FinishLaunch()
    mLaunched = True
    mLaunchTick = Timer
    If mPending = 0 Then CompleteBatch
End Sub

' ---------------------------------------------------------------------------
'  Completion path (invoked by cHttpBatch via CallByName - do not call)
' ---------------------------------------------------------------------------

Public Sub OnHttpDone(RequestId As String, r As cHttpResponse)
    Dim curveKeys As Variant
    Dim k As Variant
    Dim cv As cCmsCurve
    Dim rosetta As String
    Dim transportError As String

    If Not mReqKeys.Exists(RequestId) Then Exit Sub
    modCms.NoteDelivery
    curveKeys = mReqKeys(RequestId)
    mReqKeys.Remove RequestId

    ' Transport-level failure?
    If Len(r.ErrorMsg) > 0 Then
        transportError = r.ErrorMsg
    ElseIf r.Status < 200 Or r.Status >= 300 Then
        transportError = "HTTP " & r.Status & ": " & Left$(r.Body, 300)
    Else
        On Error Resume Next
        rosetta = modCms.ExtractRosetta(r.Body)
        If Err.Number <> 0 Then transportError = Err.Description
        On Error GoTo 0
    End If

    If Len(transportError) = 0 Then
        ' Parse the Rosetta payload into the curves carried by this request.
        On Error Resume Next
        modCms.ParseRosettaResponse rosetta, (mAction = "SET"), mCurves, curveKeys
        If Err.Number <> 0 Then transportError = "Parse error: " & Err.Description
        On Error GoTo 0
    End If

    For Each k In curveKeys
        If mCurves.Exists(CStr(k)) Then
            Set cv = mCurves(CStr(k))
            cv.HttpStatus = r.Status
            cv.ElapsedMs = r.ElapsedMs
            If Len(transportError) > 0 Then
                cv.Status = modCms.CMS_STATUS_FAILED
                cv.ErrorMsg = transportError
            End If
            If cv.Status = modCms.CMS_STATUS_FAILED Then mFailedCount = mFailedCount + 1
            ' Persist to the in-memory curve store.
            If cv.IsOk Then
                cv.FetchedAt = Now()
                If mAction = "SET" Then
                    modCms.StoreAfterSet cv
                Else
                    modCms.StoreUpsert cv
                End If
            End If
            ' Recalc any sheet cells subscribed to this curve (status updates
            ' included, so FAILED surfaces in CURVE(t,"Status") cells too).
            modCms.NotifySubscribers cv
            FireOnCurve cv
        End If
    Next k

    mPending = mPending - 1
    If mLaunched And mPending = 0 Then CompleteBatch
End Sub

Private Sub CompleteBatch()
    Dim elapsedSec As Double
    If mDone Then Exit Sub
    mDone = True
    If mLaunchTick > 0 Then
        elapsedSec = Timer - mLaunchTick
        If elapsedSec < 0 Then elapsedSec = elapsedSec + 86400  ' Timer resets at midnight
        mWallMs = CLng(elapsedSec * 1000)
    End If
    ' Break the cCmsBatch <-> cHttpBatch reference cycle so both can be freed,
    ' and drop out of the active-batch registry (CMS_LastBatch keeps the most
    ' recent one reachable for inspection).
    mHttp.SetCallback Nothing, ""
    modCms.UnregisterActiveBatch Me
    If Len(mOnAllDone) > 0 Then
        On Error Resume Next
        Application.Run mOnAllDone, Me
        If Err.Number <> 0 Then Debug.Print "cCmsBatch: all-done callback '" & mOnAllDone & "' failed: " & Err.Description
        On Error GoTo 0
    End If
End Sub

Private Sub FireOnCurve(ByVal Curve As cCmsCurve)
    If Len(mOnCurve) = 0 Then Exit Sub
    On Error Resume Next
    Application.Run mOnCurve, Curve, Me
    If Err.Number <> 0 Then Debug.Print "cCmsBatch: per-curve callback '" & mOnCurve & "' failed: " & Err.Description
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
'  Blocking wait (for the sync compatibility wrappers / Immediate window).
'  Pumps messages so the bridge can deliver completions. Returns False on
'  overall timeout (0 = wait forever).
' ---------------------------------------------------------------------------

Public Function WaitAll(Optional ByVal OverallTimeoutMs As Long = 0) As Boolean
    Dim startTick As Double
    Dim elapsedSec As Double
    startTick = Timer
    Do While Not mDone
        DoEvents
        modBridge.Sleep 5
        If OverallTimeoutMs > 0 Then
            elapsedSec = Timer - startTick
            If elapsedSec < 0 Then elapsedSec = elapsedSec + 86400  ' Timer resets at midnight
            If elapsedSec * 1000 > OverallTimeoutMs Then
                WaitAll = False
                Exit Function
            End If
        End If
    Loop
    WaitAll = True
End Function

' ---------------------------------------------------------------------------
'  Output: 2D array, one row per INPUT row (duplicates/blanks preserved),
'  17 tenor columns + 8 scalar columns. Optionally prepend the header row.
' ---------------------------------------------------------------------------

Public Function ToArray(Optional ByVal IncludeHeader As Boolean = False) As Variant
    Dim rowKeys As Variant
    Dim nRows As Long, nCols As Long
    Dim out() As Variant
    Dim rowData As Variant
    Dim header As Variant
    Dim i As Long, j As Long, base As Long
    Dim k As String

    If IsEmpty(mRowKeys) Then
        rowKeys = mCurves.Keys
    Else
        rowKeys = mRowKeys
    End If

    nRows = UBound(rowKeys) - LBound(rowKeys) + 1
    nCols = 25
    base = IIf(IncludeHeader, 1, 0)
    ReDim out(0 To nRows - 1 + base, 0 To nCols - 1)

    If IncludeHeader Then
        header = modCms.CMS_GetCurveQuoteDataHeader17()
        For j = 0 To nCols - 1
            out(0, j) = header(j)
        Next j
    End If

    For i = 0 To nRows - 1
        k = CStr(rowKeys(LBound(rowKeys) + i))
        If Len(k) > 0 And mCurves.Exists(k) Then
            rowData = mCurves(k).ToRow()
            For j = 0 To nCols - 1
                out(i + base, j) = rowData(j)
            Next j
        Else
            out(i + base, 23) = modCms.CMS_STATUS_FAILED
            out(i + base, 24) = "Curve FourTuple is empty"
        End If
    Next i

    ToArray = out
End Function
