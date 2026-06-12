
Attribute VB_Name = "modCms"
Option Explicit

' =============================================================================
'  modCms - asynchronous CMS curve marking over the ExcelBridge HTTP engine.
'
'  Replaces the legacy CMS Data Access add-in modules (modCmsGetCurveSpreads,
'  modCmsSetCurveSpreads, CMarketData, CWebServiceHelper, modConst, ...) and
'  the xmlBats/winhttpjs SET hack, keeping byte-compatible wire formats:
'
'    GET  - Rosetta <action>GET</action>, one <marketSet> per curve, tenor grid
'           TENORS_17_STANDARD, genericKeys tag (LIVE / NYOISCLOSE) +
'           returnCreditCurveType.
'    SET  - Rosetta <action>SET</action>, one <point> per SET tenor (the
'           12-standard: 0M 3M 6M 9M 1Y 2Y 3Y 4Y 5Y 6Y 7Y 10Y - we do NOT set
'           8Y 9Y 15Y 20Y 30Y), contractualSpread -777 on standard-contract
'           products, exactly as xmlBats built them.
'
'  Transports (CMS_Configure / 'CmsEndpointUrl' & 'CmsTransport' named ranges):
'    SOAP (default) - rpc/encoded envelope POSTed to .../CreditMarkingService/
'                     (the proven legacy path; SOAPAction is "" per the WSDL).
'    REST           - the modernized JSON wrapper POSTed to
'                     <base>/api/v1/getMarketData | setMarketData; carries the
'                     SAME Rosetta XML inside <cms-request><operation ...>.
'
'  Units (mirroring the legacy add-in / CMarketData):
'    GET responses: creditSpread x10000 -> bps, creditUpfront x100 -> %.
'    SET requests : values sent verbatim (same units the GET handed you).
'    SET responses: echoed verbatim - no scaling.
'
'  Everything async runs through cCmsBatch: requests are launched in parallel
'  (or grouped via CurvesPerRequest), a per-curve callback fires as each
'  response lands, and an all-complete callback fires once at the end.
' =============================================================================

' --------------------------------------------------------------------------
'  Public constants
' --------------------------------------------------------------------------
Public Const CMS_STATUS_OK As String = "OK"
Public Const CMS_STATUS_FAILED As String = "FAILED"

Public Const CMS_TAG_LIVE As String = "LIVE"
Public Const CMS_TAG_NYOISCLOSE As String = "NYOISCLOSE"

' 5Y position on the 17-standard grid
Public Const CMS_IDX_5Y As Long = 8

' --------------------------------------------------------------------------
'  Private constants (formerly modConst)
' --------------------------------------------------------------------------
Private Const ROSETTA_VERSION As String = "5.0.15"
Private Const CMS_APP_ID As String = "CMS - Excel Addin"
Private Const CMS_APP_VERSION As String = "5.0.1"
Private Const CMS_NAMESPACE As String = "http://www.lehman.com/fta/CreditMarkingService"
Private Const DEFAULT_SOAP_ENDPOINT As String = "http://cms-lxp.lehman.com/CreditMarkingService/"
Private Const DEFAULT_LOCATION As String = "NYC"
Private Const DEFAULT_LOGICAL_TIME As String = "LIVE"
Private Const CONTRACTUAL_SPREAD_DUMMY As Long = -777
Private Const DEFAULT_TIMEOUT_MS As Long = 120000
Private Const STANDARD_JAVA_ERROR As String = "Exception during processing: javax.ejb.EJBException: Unexpected exception: javax.ejb.EJBException:"
Private Const TOO_MANY_ERRORS_CMS_MESSAGE As String = "Too many errors while processing request. Please check the request."

' --------------------------------------------------------------------------
'  Module state
' --------------------------------------------------------------------------
Private gConfigured As Boolean
Private gTransport As String        ' "SOAP" | "REST"
Private gEndpointUrl As String      ' SOAP endpoint, or REST base url (no /api/v1)
Private gImpersonatedUser As String

Private gUserName As String         ' cached identity
Private gUserDomain As String
Private gMachineName As String

Private gTenorIndex As Object       ' "5Y" -> 8 etc.
Private gStore As Object            ' fourTuple -> cCmsCurve (latest known state)
Private gTickerIndex As Object      ' TICKER -> fourTuple key (one curve per ticker)
Private gActiveBatches As Collection ' in-flight cCmsBatch objects (kept alive + inspectable)
Private gLastBatch As cCmsBatch     ' most recently launched batch (kept after completion)

Private gWatchdogOn As Boolean      ' OnTime delivery watchdog (armed while batches in flight)
Private gWatchdogNextRun As Date
Private gDeliveryCount As Long      ' total HTTP completions delivered into cCmsBatch
Private gWatchdogFlushCount As Long ' deliveries that were STUCK until a watchdog tick flushed them

Private gSubs As Object             ' bucket (KEY or TICKER) -> Dictionary(cell address -> True)
Private gWanted As Object           ' fourTuple keys queued for auto-fetch
Private gFetchStamp As Object       ' key -> time of last auto-fetch launch (cooldown)
Private gPrevWanted As Object       ' keys queued for a T-1 close fetch
Private gPrevStamp As Object        ' key -> time of last T-1 fetch launch (cooldown)
Private gAutoPrevClose As Variant   ' Empty = default (True)

Private gNotifyMode As Long         ' see CMS_NOTIFY_* (default DIRTY)
Private gNotifyModeSet As Boolean
Private gWatchdogSeconds As Long    ' 0 = default (1s)

Private Const AUTOFETCH_COOLDOWN_SEC As Long = 30
Private Const PREVFETCH_COOLDOWN_SEC As Long = 300

' What a store change does to subscribed cells:
'   OFF   - nothing (subscriptions kept; read on your own schedule)
'   DIRTY - mark dirty: automatic mode recalcs after the delivery macro;
'           manual mode picks them up at the user's next recalc  (DEFAULT)
'   CALC  - recalculate subscribed cells immediately (heaviest; can feel
'           laggy when large batches land)
Public Const CMS_NOTIFY_OFF As Long = 0
Public Const CMS_NOTIFY_DIRTY As Long = 1
Public Const CMS_NOTIFY_CALC As Long = 2

' Text shown by CURVE() cells while a queued fetch is outstanding.
Public Const CMS_PENDING_TEXT As String = "#Pending"

' =============================================================================
'  CONFIGURATION
' =============================================================================

' Optional explicit configuration. Without it the module reads the
' 'CmsEndpointUrl' / 'CmsTransport' named ranges, falling back to the
' legacy PROD SOAP endpoint.
Public Sub CMS_Configure(Optional ByVal EndpointUrl As String = "", _
                         Optional ByVal Transport As String = "", _
                         Optional ByVal ImpersonateUser As String = "")
    EnsureConfig
    If Len(EndpointUrl) > 0 Then gEndpointUrl = EndpointUrl
    If Len(Transport) > 0 Then gTransport = UCase$(Transport)
    If Len(ImpersonateUser) > 0 Then gImpersonatedUser = ImpersonateUser
End Sub

' How store changes reach subscribed cells (CMS_NOTIFY_OFF/DIRTY/CALC).
Public Sub CMS_SetNotifyMode(ByVal Mode As Long)
    gNotifyMode = Mode
    gNotifyModeSet = True
End Sub

Public Function CMS_GetNotifyMode() As Long
    If gNotifyModeSet Then
        CMS_GetNotifyMode = gNotifyMode
    Else
        CMS_GetNotifyMode = CMS_NOTIFY_DIRTY
    End If
End Function

' Watchdog tick interval in seconds (min/default 1). The watchdog only runs
' while requests are in flight or fetches are queued.
Public Sub CMS_SetWatchdogInterval(ByVal Seconds As Long)
    If Seconds < 1 Then Seconds = 1
    gWatchdogSeconds = Seconds
End Sub

' Automatic once-a-day T-1 close fetch after each curve's first live GET.
Public Sub CMS_SetAutoPrevClose(ByVal Enabled As Boolean)
    gAutoPrevClose = Enabled
End Sub

Public Function CMS_AutoPrevCloseEnabled() As Boolean
    If IsEmpty(gAutoPrevClose) Then
        CMS_AutoPrevCloseEnabled = True   ' default ON
    Else
        CMS_AutoPrevCloseEnabled = CBool(gAutoPrevClose)
    End If
End Function

Public Function CMS_EndpointUrl() As String
    EnsureConfig
    CMS_EndpointUrl = gEndpointUrl
End Function

Public Function CMS_Transport() As String
    EnsureConfig
    CMS_Transport = gTransport
End Function

Private Sub EnsureConfig()
    If gConfigured Then Exit Sub
    gEndpointUrl = NamedRangeText("CmsEndpointUrl")
    If Len(gEndpointUrl) = 0 Then gEndpointUrl = DEFAULT_SOAP_ENDPOINT
    gTransport = UCase$(NamedRangeText("CmsTransport"))
    If gTransport <> "REST" Then gTransport = "SOAP"
    gConfigured = True
End Sub

Private Function NamedRangeText(ByVal Name As String) As String
    On Error Resume Next
    NamedRangeText = CStr(ThisWorkbook.Names(Name).RefersToRange.Value2)
    On Error GoTo 0
End Function

' =============================================================================
'  BATCH REGISTRY, PUMP & DIAGNOSTICS
'
'  Delivery model (read this before debugging "callbacks never fire"):
'  the XLL does its I/O on .NET worker threads, but the final hop into VBA
'  (EB_OnHttpBatch -> cHttpBatch -> cCmsBatch -> your callbacks) is queued via
'  QueueAsMacro and can only execute when Excel's main thread is able to run
'  a macro: no VBA executing, NOT in VBE break mode, no cell being edited,
'  no modal dialog. When Excel is genuinely idle, delivery is automatic and
'  immediate. While your own VBA is running, DoEvents opens the gate (that is
'  what WaitAll does). Sitting at a breakpoint / after hitting Debug on an
'  error dialog blocks ALL delivery until you resume or reset.
'
'  Also: pressing End/Reset (or editing code in break mode) wipes every
'  module-level variable - the curve store, modBridge's request router, the
'  active-batch registry. In-flight responses arriving after a reset are
'  dropped. If you reset mid-test, relaunch the batch.
' =============================================================================

' Every launched batch is held here until complete, so it stays alive and
' inspectable even after the caller's local variable goes out of scope.
Public Sub RegisterActiveBatch(ByVal Batch As cCmsBatch)
    If gActiveBatches Is Nothing Then Set gActiveBatches = New Collection
    gActiveBatches.Add Batch
    Set gLastBatch = Batch
    EnsureWatchdog
End Sub

Public Sub UnregisterActiveBatch(ByVal Batch As cCmsBatch)
    Dim i As Long
    If gActiveBatches Is Nothing Then Exit Sub
    For i = gActiveBatches.Count To 1 Step -1
        If gActiveBatches(i) Is Batch Then gActiveBatches.Remove i
    Next i
End Sub

' The most recently launched batch (kept after completion for inspection):
'   ?CMS_LastBatch.IsDone, CMS_LastBatch.PendingRequests, CMS_LastBatch.FailedCount
Public Function CMS_LastBatch() As cCmsBatch
    Set CMS_LastBatch = gLastBatch
End Function

Public Function CMS_ActiveBatchCount() As Long
    If Not gActiveBatches Is Nothing Then CMS_ActiveBatchCount = gActiveBatches.Count
End Function

' Manually pump message delivery from the Immediate window or a test sub:
' runs DoEvents until all active batches complete or TimeoutMs elapses.
' Production code does NOT need this - the watchdog below flushes delivery -
' but in the Immediate window nothing pumps for you between statements.
Public Sub CMS_Pump(Optional ByVal TimeoutMs As Long = 2000)
    Dim startTick As Double
    Dim elapsedSec As Double
    startTick = Timer
    Do
        DoEvents
        modBridge.Sleep 5
        elapsedSec = Timer - startTick
        If elapsedSec < 0 Then elapsedSec = elapsedSec + 86400
        If elapsedSec * 1000 >= TimeoutMs Then Exit Do
    Loop While CMS_ActiveBatchCount() > 0
End Sub

' ---------------------------------------------------------------------------
'  Delivery watchdog.
'
'  The bridge's queued completions normally execute the moment Excel goes
'  idle, but a batch with NO pump anywhere (e.g. a SET-only flow launched and
'  forgotten) can sit on a stuck delivery if Excel never returns to a
'  macro-capable state the bridge can reach (focus parked in the VBE, etc.).
'  While any batch is in flight we keep a 1-second Application.OnTime tick
'  armed; entering the tick gives Excel a macro context and its DoEvents
'  flushes any queued deliveries. When delivery is working normally the tick
'  finds nothing to do and disarms itself once the last batch completes.
' ---------------------------------------------------------------------------

Private Sub EnsureWatchdog()
    If gWatchdogOn Then Exit Sub
    ' Armed while requests are in flight OR auto-fetches are queued: the queue
    ' is filled from UDF calc context, which cannot launch HTTP itself and -
    ' in manual calc mode - cannot rely on a SheetCalculate event to drain it.
    If CMS_ActiveBatchCount() = 0 And WantedCount() = 0 Then Exit Sub
    On Error Resume Next
    gWatchdogNextRun = Now + TimeSerial(0, 0, IIf(gWatchdogSeconds < 1, 1, gWatchdogSeconds))
    Application.OnTime gWatchdogNextRun, "modCms.CmsWatchdogTick"
    gWatchdogOn = (Err.Number = 0)
    On Error GoTo 0
End Sub

' Fired by Application.OnTime - do not call directly.
Public Sub CmsWatchdogTick()
    Dim before As Long
    gWatchdogOn = False
    ' We are in a macro context now; DoEvents lets the bridge's queued
    ' EB_OnHttpBatch deliveries run if they were stuck waiting for one.
    ' Deliveries landing inside THIS DoEvents were, by definition, waiting:
    ' count them so CMS_Diag can show whether the watchdog is load-bearing.
    ' (Slight over-attribution is possible if a tick fires inside someone
    ' else's pump - fine for diagnostics.)
    before = gDeliveryCount
    DoEvents
    If gDeliveryCount > before Then
        gWatchdogFlushCount = gWatchdogFlushCount + (gDeliveryCount - before)
        Debug.Print "CMS watchdog " & Format$(Now, "hh:nn:ss") & ": flushed " & _
                    (gDeliveryCount - before) & " stuck deliver(ies)"
    End If
    CMS_FlushAutoFetch
    SweepStaleBatches
    If CMS_ActiveBatchCount() > 0 Or WantedCount() > 0 Then EnsureWatchdog
End Sub

' A batch whose deliveries were lost (state reset mid-flight) would otherwise
' stay "active" forever and keep the watchdog ticking. Abandon anything that
' has produced nothing for far longer than any request timeout.
Private Const BATCH_ABANDON_MS As Long = 600000   ' 10 minutes

Private Sub SweepStaleBatches()
    Dim i As Long
    Dim b As cCmsBatch
    If gActiveBatches Is Nothing Then Exit Sub
    For i = gActiveBatches.Count To 1 Step -1
        Set b = gActiveBatches(i)
        If b.IsDone Then
            gActiveBatches.Remove i
        ElseIf b.AgeMs > BATCH_ABANDON_MS Then
            Debug.Print "CMS: abandoning stuck " & b.Action & " batch after " & _
                        (b.AgeMs \ 1000) & "s (" & b.PendingRequests & " request(s) undelivered)"
            b.Abandon "Abandoned: no delivery within " & (BATCH_ABANDON_MS \ 1000) & _
                      "s - responses likely lost to a state reset; relaunch"
        End If
    Next i
End Sub

Private Function WantedCount() As Long
    If Not gWanted Is Nothing Then WantedCount = gWanted.Count
    If Not gPrevWanted Is Nothing Then WantedCount = WantedCount + gPrevWanted.Count
End Function

' Called by cCmsBatch.OnHttpDone on every delivered completion.
Public Sub NoteDelivery()
    gDeliveryCount = gDeliveryCount + 1
End Sub

Public Function CMS_DeliveryCount() As Long
    CMS_DeliveryCount = gDeliveryCount
End Function

' > 0 means deliveries in this session genuinely sat stuck until the
' watchdog freed them (i.e. the SET-only stall was real, not user error).
Public Function CMS_WatchdogFlushCount() As Long
    CMS_WatchdogFlushCount = gWatchdogFlushCount
End Function

' Cancel a pending watchdog tick. Call from Workbook_BeforeClose alongside
' bridge_stop so a scheduled OnTime cannot reopen the workbook after close.
Public Sub CMS_StopWatchdog()
    If Not gWatchdogOn Then Exit Sub
    On Error Resume Next
    Application.OnTime gWatchdogNextRun, "modCms.CmsWatchdogTick", , False
    On Error GoTo 0
    gWatchdogOn = False
End Sub

Private Function CalcModeName() As String
    Select Case Application.Calculation
        Case xlCalculationAutomatic: CalcModeName = "automatic"
        Case xlCalculationManual: CalcModeName = "manual"
        Case Else: CalcModeName = "semiautomatic"
    End Select
End Function

' End-to-end smoke test - run from the Immediate window: modCms.CMS_SelfTest
' Walks registration state, fires one verbose GET for the first stored
' curve, pumps until it lands, and prints the outcome at every step.
Public Sub CMS_SelfTest()
    Dim k As String
    Dim b As cCmsBatch
    Dim cv As cCmsCurve
    Dim prevVerbose As Boolean

    Debug.Print "=== CMS self-test " & Format$(Now, "hh:nn:ss") & " ==="
    Debug.Print CMS_Diag()

    If CMS_RegisteredCount() = 0 Then
        Debug.Print "STORE IS EMPTY - register_curve_by_range did not run or raised."
        Debug.Print "Check your registration range has 5 columns (Ticker, Ccy, DebtClass, Product, QuoteConvention)."
        Exit Sub
    End If

    k = CStr(CMS_Store().Keys()(0))
    Debug.Print "test GET for: " & k

    prevVerbose = modBridge.gHttpVerbose
    modBridge.gHttpVerbose = True
    On Error GoTo Cleanup

    Set b = CMS_GetCurvesAsync(k)
    If b.WaitAll(30000) Then
        Set cv = b.Curve(CStr(b.Keys()(0)))
        Debug.Print "result: status=" & cv.Status & "  5Y=" & CStr(cv.Quote("5Y")) & _
                    "  http=" & cv.HttpStatus & "  elapsed=" & cv.ElapsedMs & "ms" & _
                    IIf(Len(cv.ErrorMsg) > 0, "  error=" & cv.ErrorMsg, "")
        Debug.Print "FetchedAt=" & Format$(cv.FetchedAt, "yyyy-mm-dd hh:nn:ss")
    Else
        Debug.Print "TIMED OUT after 30s. EB_LastDispatchError=" & modBridge.BridgeLastDispatchError()
    End If

Cleanup:
    If Err.Number <> 0 Then Debug.Print "self-test raised: " & Err.Description
    modBridge.gHttpVerbose = prevVerbose
    Debug.Print "=== self-test done ==="
End Sub

' One-shot health report:  ?modCms.CMS_Diag
Public Function CMS_Diag() As String
    Dim s As String
    Dim i As Long
    Dim b As cCmsBatch
    EnsureConfig
    s = "transport=" & gTransport & "; endpoint=" & gEndpointUrl & vbNewLine
    s = s & "bridge loaded=" & modBridge.AddinLoaded() & _
        "; last dispatch error=" & modBridge.BridgeLastDispatchError() & vbNewLine
    s = s & "store curves=" & CMS_Store().Count & _
        "; pending 5Y=" & CMS_PendingCount() & _
        "; deliveries=" & gDeliveryCount & _
        "; watchdog flushes=" & gWatchdogFlushCount & vbNewLine
    s = s & "cell subscriptions=" & CMS_SubscriptionCount() & _
        "; autofetch queued=" & IIf(gWanted Is Nothing, 0, gWanted.Count) & _
        "; T-1 queued=" & IIf(gPrevWanted Is Nothing, 0, gPrevWanted.Count) & vbNewLine
    s = s & "notify mode=" & CMS_GetNotifyMode() & " (0=off 1=dirty 2=calc)" & _
        "; auto prev close=" & CMS_AutoPrevCloseEnabled() & _
        "; prev biz day=" & Format$(CMS_PrevBizDay(), "yyyy-mm-dd") & vbNewLine
    ' Calc-state visibility: ForceFullCalculation=True turns EVERY calc
    ' trigger (incl. each formula commit) into a full-workbook rebuild.
    s = s & "calc mode=" & CalcModeName() & _
        "; ForceFullCalculation=" & ThisWorkbook.ForceFullCalculation & _
        "; watchdog armed=" & gWatchdogOn & vbNewLine
    s = s & "active batches=" & CMS_ActiveBatchCount()
    If Not gActiveBatches Is Nothing Then
        For i = 1 To gActiveBatches.Count
            Set b = gActiveBatches(i)
            s = s & vbNewLine & "  [" & i & "] " & b.Action & ": " & _
                b.CompletedCurves & "/" & b.Count & " curves done, " & _
                b.PendingRequests & " request(s) pending, " & b.FailedCount & " failed"
        Next i
    End If
    If Not gLastBatch Is Nothing Then
        s = s & vbNewLine & "last batch: " & gLastBatch.Action & _
            " done=" & gLastBatch.IsDone & " curves=" & gLastBatch.Count & _
            " failed=" & gLastBatch.FailedCount & _
            " wall=" & gLastBatch.WallMs & "ms slowest request=" & _
            gLastBatch.MaxRequestMs & "ms"
    End If
    CMS_Diag = s
End Function

' =============================================================================
'  TENOR MODEL
' =============================================================================

' GET grid: TENORS_17_STANDARD
Public Function CMS_GetTenors17() As Variant
    CMS_GetTenors17 = Array("0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y", _
                            "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "30Y")
End Function

' SET grid: the 12 tenors we mark (no 8Y/9Y/15Y/20Y/30Y)
Public Function CMS_SetTenors12() As Variant
    CMS_SetTenors12 = Array("0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y", _
                            "5Y", "6Y", "7Y", "10Y")
End Function

' Indices of the 12 SET tenors within the 17-grid
Public Function CMS_SetTenorIndices() As Variant
    CMS_SetTenorIndices = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 13)
End Function

' "5Y" -> 8; -1 if unknown
Public Function CMS_TenorIndex(ByVal Tenor As String) As Long
    Dim t As Variant
    Dim i As Long
    If gTenorIndex Is Nothing Then
        Set gTenorIndex = CreateObject("Scripting.Dictionary")
        gTenorIndex.CompareMode = vbTextCompare
        i = 0
        For Each t In CMS_GetTenors17()
            gTenorIndex.Add CStr(t), i
            i = i + 1
        Next t
    End If
    Tenor = UCase$(Trim$(Tenor))
    If gTenorIndex.Exists(Tenor) Then
        CMS_TenorIndex = gTenorIndex(Tenor)
    Else
        CMS_TenorIndex = -1
    End If
End Function

' Result-row header: 17 tenors + the 8 scalar columns
Public Function CMS_GetCurveQuoteDataHeader17() As Variant
    CMS_GetCurveQuoteDataHeader17 = Array( _
        "0M", "3M", "6M", "9M", "1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y", "8Y", "9Y", "10Y", "15Y", "20Y", "30Y", _
        "Recovery", "Last Marked On", "Last Marked By", "Owner", "Is Parent", "Curve Type", "Status", "Error Message")
End Function

' =============================================================================
'  CURVE MATH - the 5Y perturbation
'
'  Front end (0M..4Y) : keep the RATIO to the old 5Y   -> ratio * new5y
'  5Y                 : the new mark
'  Long end (6Y..30Y) : keep the DIFFERENCE to the old 5Y -> new5y + diff
' =============================================================================

Public Function CMS_ApplyNew5y(ByVal Quotes17 As Variant, ByVal New5y As Double) As Variant
    Dim out(0 To 16) As Variant
    Dim old5y As Double
    Dim i As Long

    If IsEmpty(Quotes17(CMS_IDX_5Y)) Or Not IsNumeric(Quotes17(CMS_IDX_5Y)) Then
        Err.Raise vbObjectError + 600, "modCms.CMS_ApplyNew5y", "Existing curve has no 5Y mark"
    End If
    old5y = CDbl(Quotes17(CMS_IDX_5Y))
    If old5y = 0 Then
        Err.Raise vbObjectError + 601, "modCms.CMS_ApplyNew5y", "Existing 5Y mark is zero - cannot scale the front end"
    End If

    For i = 0 To 7   ' 0M 3M 6M 9M 1Y 2Y 3Y 4Y
        If IsNumeric(Quotes17(i)) And Not IsEmpty(Quotes17(i)) Then
            out(i) = (CDbl(Quotes17(i)) / old5y) * New5y
        End If
    Next i
    out(CMS_IDX_5Y) = New5y
    For i = 9 To 16  ' 6Y 7Y 8Y 9Y 10Y 15Y 20Y 30Y
        If IsNumeric(Quotes17(i)) And Not IsEmpty(Quotes17(i)) Then
            out(i) = New5y + (CDbl(Quotes17(i)) - old5y)
        End If
    Next i
    CMS_ApplyNew5y = out
End Function

' =============================================================================
'  PUBLIC ASYNC API
' =============================================================================

' ---------------------------------------------------------------------------
' Bulk async GET. FiveTuples accepts any curve spec:
'   - Range/2D array, 5 cols: Ticker, Ccy, DebtClass, Product, QuoteConvention
'     (e.g. D9:H450; identities are registered for later ticker-only use)
'   - Range/2D array, 1 col : registered tickers or fourTuple keys
'   - a single ticker/key string
' Blank-ticker rows are skipped; duplicates collapse onto one request but
' still appear row-for-row in batch.ToArray(). Unregistered tickers raise
' before anything is sent.
'
'   OnCurve    : "Module.Sub" with signature (curve As cCmsCurve, batch As cCmsBatch)
'   OnAllDone  : "Module.Sub" with signature (batch As cCmsBatch)
'   CurvesPerRequest : 1 = one HTTP request per curve (max parallelism,
'                      default); N>1 groups N marketSets per request.
' ---------------------------------------------------------------------------
Public Function CMS_GetCurvesAsync(ByVal FiveTuples As Variant, _
                                   Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                   Optional ByVal QuoteDate As Date, _
                                   Optional ByVal OnCurve As String = "", _
                                   Optional ByVal OnAllDone As String = "", _
                                   Optional ByVal CurvesPerRequest As Long = 1, _
                                   Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim batch As cCmsBatch
    Dim curves As Collection
    Dim rowKeys As Variant
    Dim groupCurves As Collection
    Dim i As Long
    Dim dateStr As String

    If CurvesPerRequest < 1 Then CurvesPerRequest = 1
    dateStr = QuoteDateString(QuoteDate)
    Tag = UCase$(Tag)

    Set batch = New cCmsBatch
    batch.Init "GET", OnCurve, OnAllDone
    RegisterActiveBatch batch

    Set curves = NormalizeFiveTuples(FiveTuples, batch, rowKeys)
    batch.SetRowKeys rowKeys

    ' Chunk into requests of CurvesPerRequest marketSets each
    Set groupCurves = New Collection
    For i = 1 To curves.Count
        groupCurves.Add curves(i)
        If groupCurves.Count = CurvesPerRequest Or i = curves.Count Then
            SendGetRequest batch, groupCurves, Tag, dateStr, TimeoutMs
            Set groupCurves = New Collection
        End If
    Next i

    batch.FinishLaunch
    TraceLaunch batch, "GET"
    Set CMS_GetCurvesAsync = batch
End Function

' ---------------------------------------------------------------------------
' Bulk async SET ("CMS_SetBulkCurveQuoteData" - the call that never existed).
' One HTTP request per curve, all in flight at once.
'
'   FiveTuples : any curve spec (5-col tuples, 1-col registered tickers/keys,
'                or a single ticker string) - N rows
'   Quotes     : 2D array/Range, N rows x 12 columns (SET tenor order
'                0M 3M 6M 9M 1Y 2Y 3Y 4Y 5Y 6Y 7Y 10Y) or N x 17 columns
'                (17-standard order; the 8Y/9Y/15Y/20Y/30Y columns are ignored).
'   Recoveries : optional N x 1 recovery rates (only sent when > 0)
' ---------------------------------------------------------------------------
Public Function CMS_SetCurvesAsync(ByVal FiveTuples As Variant, _
                                   ByVal Quotes As Variant, _
                                   Optional ByVal Recoveries As Variant, _
                                   Optional ByVal OnCurve As String = "", _
                                   Optional ByVal OnAllDone As String = "", _
                                   Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim batch As cCmsBatch
    Dim curves As Collection
    Dim rowKeys As Variant
    Dim cv As cCmsCurve
    Dim q As Variant
    Dim rec As Variant
    Dim r As Long, nRows As Long
    Dim Key As String
    Dim quotes12 As Variant
    Dim recovery As Double

    Set batch = New cCmsBatch
    batch.Init "SET", OnCurve, OnAllDone
    RegisterActiveBatch batch

    Set curves = NormalizeFiveTuples(FiveTuples, batch, rowKeys)
    batch.SetRowKeys rowKeys

    q = ToArray2D(Quotes)
    If Not IsMissing(Recoveries) And Not IsEmpty(Recoveries) Then rec = ToArray2D(Recoveries)

    nRows = UBound(rowKeys) + 1
    If UBound(q, 1) - LBound(q, 1) + 1 < nRows Then
        Err.Raise vbObjectError + 602, "modCms.CMS_SetCurvesAsync", _
            "Quotes array has fewer rows than FiveTuples"
    End If

    ' Walk INPUT rows so quote rows stay aligned with their five-tuples even
    ' when blank or duplicate rows are skipped.
    For r = 0 To nRows - 1
        Key = CStr(rowKeys(r))
        If Len(Key) > 0 Then
            Set cv = batch.Curve(Key)
            If Len(cv.RequestId) = 0 Then   ' first occurrence of this curve
                quotes12 = ExtractQuotes12(q, LBound(q, 1) + r)
                recovery = 0
                If Not IsEmpty(rec) Then
                    If LBound(rec, 1) + r <= UBound(rec, 1) Then
                        If IsNumeric(rec(LBound(rec, 1) + r, LBound(rec, 2))) Then _
                            recovery = CDbl(rec(LBound(rec, 1) + r, LBound(rec, 2)))
                    End If
                End If
                cv.PendingQuotes = Expand12To17(quotes12)
                SendSetRequest batch, cv, quotes12, recovery, TimeoutMs
            End If
        End If
    Next r

    batch.FinishLaunch
    TraceLaunch batch, "SET"
    Set CMS_SetCurvesAsync = batch
End Function

' ---------------------------------------------------------------------------
' The perturbation workflow: re-mark curves off a new 5Y level, computing the
' rest of the curve in memory from the last GET (no sheet round trip).
'
'   KeysAndNew5y : 2D array/Range, N rows x 2: (ticker OR fourTuple key,
'                  new 5Y level). Tickers must be registered (raises before
'                  anything is sent); curves must have been GETted so the
'                  cached quotes can drive the transform.
' ---------------------------------------------------------------------------
Public Function CMS_SetCurves5yAsync(ByVal KeysAndNew5y As Variant, _
                                     Optional ByVal OnCurve As String = "", _
                                     Optional ByVal OnAllDone As String = "", _
                                     Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim batch As cCmsBatch
    Dim arr As Variant
    Dim i As Long
    Dim Key As String
    Dim new5y As Double
    Dim stored As cCmsCurve
    Dim cv As cCmsCurve
    Dim rowKeys() As Variant
    Dim n As Long

    Set batch = New cCmsBatch
    batch.Init "SET", OnCurve, OnAllDone
    RegisterActiveBatch batch

    arr = ToArray2D(KeysAndNew5y)
    n = UBound(arr, 1) - LBound(arr, 1) + 1
    ReDim rowKeys(0 To n - 1)

    ' Resolution pre-pass: tickers -> keys, raising for unregistered tickers
    ' BEFORE any request is launched.
    For i = 0 To n - 1
        Key = Trim$(NzStr(arr(LBound(arr, 1) + i, LBound(arr, 2))))
        If Len(Key) = 0 Then
            rowKeys(i) = ""
        Else
            rowKeys(i) = CMS_ResolveKey(Key)
        End If
    Next i

    For i = 0 To n - 1
        Key = CStr(rowKeys(i))
        If Len(Key) = 0 Then GoTo NextRow

        Set cv = New cCmsCurve
        Set stored = StoreGet(Key)
        If stored Is Nothing Then
            ParseKeyInto Key, cv
            If batch.AddCurve(cv) Then batch.FailCurve cv, "Curve not in store - GET it before setting off a new 5Y"
            GoTo NextRow
        End If
        If Not IsNumeric(arr(LBound(arr, 1) + i, LBound(arr, 2) + 1)) Then
            ParseKeyInto Key, cv
            If batch.AddCurve(cv) Then batch.FailCurve cv, "New 5Y level is not numeric"
            GoTo NextRow
        End If
        new5y = CDbl(arr(LBound(arr, 1) + i, LBound(arr, 2) + 1))

        cv.Init stored.Ticker, stored.Ccy, stored.DebtClass, stored.Product, stored.QuoteConvention
        cv.LabelOverride = stored.LabelOverride

        On Error Resume Next
        cv.PendingQuotes = CMS_ApplyNew5y(stored.Quotes, new5y)
        If Err.Number <> 0 Then
            Dim transformErr As String
            transformErr = Err.Description
            Err.Clear
            On Error GoTo 0
            If batch.AddCurve(cv) Then batch.FailCurve cv, transformErr
            GoTo NextRow
        End If
        On Error GoTo 0

        If batch.AddCurve(cv) Then
            SendSetRequest batch, cv, Project17To12(cv.PendingQuotes), 0, TimeoutMs
        End If
NextRow:
    Next i

    batch.SetRowKeys rowKeys
    batch.FinishLaunch
    TraceLaunch batch, "SET-5Y"
    Set CMS_SetCurves5yAsync = batch
End Function

' One line per launched batch in the Immediate window - the first thing to
' check when "nothing happens": did anything actually go out?
Private Sub TraceLaunch(ByVal Batch As cCmsBatch, ByVal Label As String)
    Debug.Print "CMS " & Label & " " & Format$(Now, "hh:nn:ss") & ": " & _
                Batch.Count & " curve(s), " & Batch.PendingRequests & " request(s) sent" & _
                IIf(Batch.FailedCount > 0, ", " & Batch.FailedCount & " failed at launch", "")
End Sub

' ---------------------------------------------------------------------------
' Async refresh of a single curve (returns the batch; result via callbacks
' or batch.Curve(key) after batch.IsDone). Identity parts beyond Ticker are
' optional: omit them to use the registered identity for that ticker.
' ---------------------------------------------------------------------------
Public Function CMS_GetCurveAsync(ByVal Ticker As String, _
                                  Optional ByVal Ccy As String = "", _
                                  Optional ByVal DebtClass As String = "", _
                                  Optional ByVal Product As String = "", _
                                  Optional ByVal QuoteConvention As String = "", _
                                  Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                  Optional ByVal QuoteDate As Date, _
                                  Optional ByVal OnCurve As String = "", _
                                  Optional ByVal OnAllDone As String = "") As cCmsBatch
    Dim idc As cCmsCurve
    Dim tuples(0 To 0, 0 To 4) As Variant
    Set idc = BuildIdentity(Ticker, Ccy, DebtClass, Product, QuoteConvention)
    tuples(0, 0) = idc.Ticker: tuples(0, 1) = idc.Ccy: tuples(0, 2) = idc.DebtClass
    tuples(0, 3) = idc.Product: tuples(0, 4) = idc.QuoteConvention
    Set CMS_GetCurveAsync = CMS_GetCurvesAsync(tuples, Tag, QuoteDate, OnCurve, OnAllDone)
End Function

' ---------------------------------------------------------------------------
' Async SET of a single curve. TermQuotes: 12 values in SET-tenor order or a
' 17-grid row (Range / 1D / 2D; columns or transposed both fine). Identity
' parts optional - ticker-only works for registered curves.
' ---------------------------------------------------------------------------
Public Function CMS_SetCurveAsync(ByVal Ticker As String, ByVal TermQuotes As Variant, _
                                  Optional ByVal Ccy As String = "", _
                                  Optional ByVal DebtClass As String = "", _
                                  Optional ByVal Product As String = "", _
                                  Optional ByVal RecoveryRate As Double = 0, _
                                  Optional ByVal QuoteConvention As String = "", _
                                  Optional ByVal OnCurve As String = "", _
                                  Optional ByVal OnAllDone As String = "", _
                                  Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim idc As cCmsCurve
    Dim tuples(0 To 0, 0 To 4) As Variant
    Dim recs(0 To 0, 0 To 0) As Variant
    Set idc = BuildIdentity(Ticker, Ccy, DebtClass, Product, QuoteConvention)
    tuples(0, 0) = idc.Ticker: tuples(0, 1) = idc.Ccy: tuples(0, 2) = idc.DebtClass
    tuples(0, 3) = idc.Product: tuples(0, 4) = idc.QuoteConvention
    recs(0, 0) = RecoveryRate
    Set CMS_SetCurveAsync = CMS_SetCurvesAsync(tuples, RowToArray2D(TermQuotes), recs, _
                                               OnCurve, OnAllDone, TimeoutMs)
End Function

' ---------------------------------------------------------------------------
' Async 5Y re-mark of a single registered curve (store-cached quotes drive
' the rest of the curve).
' ---------------------------------------------------------------------------
Public Function CMS_SetCurve5yAsync(ByVal TickerOrKey As String, ByVal New5y As Double, _
                                    Optional ByVal OnCurve As String = "", _
                                    Optional ByVal OnAllDone As String = "", _
                                    Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim arr(0 To 0, 0 To 1) As Variant
    arr(0, 0) = TickerOrKey: arr(0, 1) = New5y
    Set CMS_SetCurve5yAsync = CMS_SetCurves5yAsync(arr, OnCurve, OnAllDone, TimeoutMs)
End Function

' ---------------------------------------------------------------------------
' Ticker-list GET that tolerates any shape (1D array, horizontal row,
' vertical column) - normalizes to the 1-column ticker spec.
' ---------------------------------------------------------------------------
Public Function CMS_GetTickersAsync(ByVal Tickers As Variant, _
                                    Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                    Optional ByVal QuoteDate As Date, _
                                    Optional ByVal OnCurve As String = "", _
                                    Optional ByVal OnAllDone As String = "", _
                                    Optional ByVal CurvesPerRequest As Long = 1, _
                                    Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Set CMS_GetTickersAsync = CMS_GetCurvesAsync(ToTickerColumn(Tickers), Tag, QuoteDate, _
                                                 OnCurve, OnAllDone, CurvesPerRequest, TimeoutMs)
End Function

' =============================================================================
'  SYNC COMPATIBILITY WRAPPERS (legacy signatures, 17-standard layout)
'  These block the calling VBA only - the bridge still does the I/O async.
' =============================================================================

' One curve -> 1D Variant(0..24): 17 tenor quotes + 8 scalars.
' Ccy/DebtClass are optional: CMS_GetCurveQuoteData("AEP") works once the
' ticker is registered (legacy positional calls are unchanged).
Public Function CMS_GetCurveQuoteData(ByVal Ticker As String, _
                                      Optional ByVal Ccy As String = "", _
                                      Optional ByVal DebtClass As String = "", _
                                      Optional ByVal Product As String = "DERIV", _
                                      Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                      Optional ByVal QuoteDate As Date, _
                                      Optional ByVal QuoteConvention As String = "", _
                                      Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As Variant
    Dim batch As cCmsBatch
    Set batch = CMS_GetCurveAsync(Ticker, Ccy, DebtClass, Product, QuoteConvention, Tag, QuoteDate)
    If Not batch.WaitAll(TimeoutMs) Then
        Err.Raise vbObjectError + 603, "modCms.CMS_GetCurveQuoteData", "Timed out waiting for CMS"
    End If
    CMS_GetCurveQuoteData = SingleBatchRow(batch, "modCms.CMS_GetCurveQuoteData")
End Function

' Bulk GET -> 2D Variant (one row per input row, 25 columns).
Public Function CMS_GetBulkCurveQuoteData(ByVal FiveTuples As Variant, _
                                          Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                          Optional ByVal QuoteDate As Date, _
                                          Optional ByVal CurvesPerRequest As Long = 1, _
                                          Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As Variant
    Dim batch As cCmsBatch
    Set batch = CMS_GetCurvesAsync(FiveTuples, Tag, QuoteDate, , , CurvesPerRequest)
    If Not batch.WaitAll(TimeoutMs) Then
        Err.Raise vbObjectError + 603, "modCms.CMS_GetBulkCurveQuoteData", "Timed out waiting for CMS"
    End If
    CMS_GetBulkCurveQuoteData = batch.ToArray(False)
End Function

' One curve SET -> echoed row (legacy CMS_SetCurveQuoteData shape/order).
Public Function CMS_SetCurveQuoteData(ByVal Ticker As String, ByVal Ccy As String, _
                                      ByVal DebtClass As String, ByVal Product As String, _
                                      ByVal TermQuotes As Variant, _
                                      Optional ByVal RecoveryRate As Double = 0, _
                                      Optional ByVal QuoteConvention As String = "SPREADS", _
                                      Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As Variant
    Dim batch As cCmsBatch
    Set batch = CMS_SetCurveAsync(Ticker, TermQuotes, Ccy, DebtClass, Product, _
                                  RecoveryRate, QuoteConvention)
    If Not batch.WaitAll(TimeoutMs) Then
        Err.Raise vbObjectError + 603, "modCms.CMS_SetCurveQuoteData", "Timed out waiting for CMS"
    End If
    CMS_SetCurveQuoteData = SingleBatchRow(batch, "modCms.CMS_SetCurveQuoteData")
End Function

' Ticker-only blocking SET (the legacy signature can't take optional middle
' args). TermQuotes: 12 SET-tenor values or a 17-grid row.
Public Function CMS_SetTickerQuoteData(ByVal TickerOrKey As String, ByVal TermQuotes As Variant, _
                                       Optional ByVal RecoveryRate As Double = 0, _
                                       Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As Variant
    Dim batch As cCmsBatch
    Set batch = CMS_SetCurveAsync(TickerOrKey, TermQuotes, , , , RecoveryRate)
    If Not batch.WaitAll(TimeoutMs) Then
        Err.Raise vbObjectError + 603, "modCms.CMS_SetTickerQuoteData", "Timed out waiting for CMS"
    End If
    CMS_SetTickerQuoteData = SingleBatchRow(batch, "modCms.CMS_SetTickerQuoteData")
End Function

' Result row of a single-curve batch (the key may have come from the registry).
Private Function SingleBatchRow(ByVal Batch As cCmsBatch, ByVal Caller As String) As Variant
    Dim ks As Variant
    ks = Batch.Keys()
    If UBound(ks) < LBound(ks) Then
        Err.Raise vbObjectError + 604, Caller, "Invalid curve identifiers"
    End If
    SingleBatchRow = Batch.Curve(CStr(ks(LBound(ks)))).ToRow()
End Function

' Bulk SET, blocking variant -> 2D result array. For the non-blocking version
' use CMS_SetCurvesAsync / CMS_SetCurves5yAsync.
Public Function CMS_SetBulkCurveQuoteData(ByVal FiveTuples As Variant, _
                                          ByVal Quotes As Variant, _
                                          Optional ByVal Recoveries As Variant, _
                                          Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As Variant
    Dim batch As cCmsBatch
    Set batch = CMS_SetCurvesAsync(FiveTuples, Quotes, Recoveries)
    If Not batch.WaitAll(TimeoutMs) Then
        Err.Raise vbObjectError + 603, "modCms.CMS_SetBulkCurveQuoteData", "Timed out waiting for CMS"
    End If
    CMS_SetBulkCurveQuoteData = batch.ToArray(False)
End Function

' =============================================================================
'  IN-MEMORY CURVE STORE
'  Latest known state per fourTuple, refreshed by every successful GET and
'  merged on every successful SET. Lets the 5Y workflow run without keeping
'  full curves on any sheet (write them out only if you want to see them).
' =============================================================================

Public Function CMS_Store() As Object
    If gStore Is Nothing Then
        Set gStore = CreateObject("Scripting.Dictionary")
        gStore.CompareMode = vbTextCompare
    End If
    Set CMS_Store = gStore
End Function

Public Function StoreGet(ByVal Key As String) As cCmsCurve
    If CMS_Store().Exists(Key) Then Set StoreGet = gStore(Key)
End Function

Public Sub StoreUpsert(ByVal Curve As cCmsCurve)
    Dim old As cCmsCurve
    ' A refresh GET delivers a NEW object; carry over session state that
    ' lives on the store entry so it survives the replacement.
    If CMS_Store().Exists(Curve.Key) Then
        Set old = gStore(Curve.Key)
        If Not old Is Curve Then
            If old.PrevDate <> 0 And Curve.PrevDate = 0 Then
                Curve.PrevQuotes = old.PrevQuotes
                Curve.PrevDate = old.PrevDate
                Curve.PrevFetchedAt = old.PrevFetchedAt
            End If
            If Not IsEmpty(old.Pending5y) And IsEmpty(Curve.Pending5y) Then
                Curve.Pending5y = old.Pending5y
            End If
        End If
    End If
    Set gStore(Curve.Key) = Curve
    IndexTicker Curve
End Sub

' After a successful SET: the response echoes the 12 sent tenors verbatim.
' Merge echo over the full computed curve, then upsert.
Public Sub StoreAfterSet(ByVal Curve As cCmsCurve)
    Dim merged(0 To 16) As Variant
    Dim i As Long
    Dim stored As cCmsCurve

    If IsEmpty(Curve.PendingQuotes) Then
        ' Raw SET without a computed full curve: merge echo over stored quotes.
        Set stored = StoreGet(Curve.Key)
        If Not stored Is Nothing Then Curve.PendingQuotes = stored.Quotes
    End If

    For i = 0 To 16
        If Not IsEmpty(Curve.Quotes(i)) Then
            merged(i) = Curve.Quotes(i)
        ElseIf IsArray(Curve.PendingQuotes) Then
            merged(i) = Curve.PendingQuotes(i)
        End If
    Next i
    Curve.Quotes = merged
    StoreUpsert Curve
End Sub

Public Sub CMS_StoreClear()
    If Not gStore Is Nothing Then gStore.RemoveAll
End Sub

' Dump store (or a subset of keys) to a 2D array for an optional sheet view.
Public Function CMS_StoreToArray(Optional ByVal Keys As Variant, _
                                 Optional ByVal IncludeHeader As Boolean = True) As Variant
    Dim useKeys As Variant
    Dim out() As Variant
    Dim header As Variant
    Dim rowData As Variant
    Dim i As Long, j As Long, base As Long
    Dim k As String

    If IsMissing(Keys) Or IsEmpty(Keys) Then
        useKeys = CMS_Store().Keys
    Else
        useKeys = Keys
    End If

    base = IIf(IncludeHeader, 1, 0)
    ReDim out(0 To UBound(useKeys) - LBound(useKeys) + base, 0 To 26)

    If IncludeHeader Then
        out(0, 0) = "Curve"
        header = CMS_GetCurveQuoteDataHeader17()
        For j = 0 To 24
            out(0, j + 1) = header(j)
        Next j
        out(0, 26) = "Pending 5Y"
    End If

    For i = 0 To UBound(useKeys) - LBound(useKeys)
        k = CStr(useKeys(LBound(useKeys) + i))
        out(i + base, 0) = k
        If CMS_Store().Exists(k) Then
            rowData = gStore(k).ToRow()
            For j = 0 To 24
                out(i + base, j + 1) = rowData(j)
            Next j
            out(i + base, 26) = gStore(k).Pending5y
        End If
    Next i
    CMS_StoreToArray = out
End Function

' Convenience: write a finished batch (or the store) below/right of a cell.
Public Sub CMS_WriteBatchToRange(ByVal Batch As cCmsBatch, ByVal TopLeft As Range, _
                                 Optional ByVal IncludeHeader As Boolean = False)
    Dim arr As Variant
    arr = Batch.ToArray(IncludeHeader)
    TopLeft.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).Value2 = arr
End Sub

' =============================================================================
'  REGISTRATION & TICKER-ONLY ACCESS
'  One curve per unique ticker: registering (explicitly, or implicitly by
'  GETting/SETting with the full tuple) maps TICKER -> fourTuple key, after
'  which every get/set/store call accepts just the ticker.
' =============================================================================

' Register one curve's identity so it can be referenced by ticker alone.
' Does NOT overwrite an already-stored curve's quotes.
Public Sub register_curve(ByVal Ticker As String, ByVal Ccy As String, _
                          ByVal DebtClass As String, ByVal Product As String, _
                          Optional ByVal QuoteConvention As String = "SPREADS")
    Dim cv As New cCmsCurve
    cv.Init Ticker, Ccy, DebtClass, Product, QuoteConvention
    RegisterIdentity cv
End Sub

' Register a block of curves from a Range/2D array with the usual 5 columns
' (Ticker, Ccy, DebtClass, Product, QuoteConvention). Blank rows are skipped.
Public Sub register_curve_by_range(ByVal FiveTuples As Variant)
    Dim arr As Variant
    Dim r As Long, c0 As Long
    Dim Ticker As String

    arr = ToArray2D(FiveTuples)
    If UBound(arr, 2) - LBound(arr, 2) + 1 < 5 Then
        Err.Raise vbObjectError + 640, "modCms.register_curve_by_range", _
            "Range must have 5 columns: Ticker, Ccy, DebtClass, Product, QuoteConvention"
    End If
    c0 = LBound(arr, 2)
    For r = LBound(arr, 1) To UBound(arr, 1)
        Ticker = Trim$(NzStr(arr(r, c0)))
        If Len(Ticker) > 0 And Ticker <> "0" Then
            register_curve Ticker, NzStr(arr(r, c0 + 1)), NzStr(arr(r, c0 + 2)), _
                           NzStr(arr(r, c0 + 3)), NzStr(arr(r, c0 + 4))
        End If
    Next r
End Sub

' "AEP" -> "AEP.USD.SENIOR_NORE_14.SNAC100" (anything containing "." is
' treated as a full fourTuple key and passed through). Raises if the ticker
' has never been registered.
Public Function CMS_ResolveKey(ByVal TickerOrKey As String) As String
    Dim s As String
    s = UCase$(Trim$(TickerOrKey))
    If Len(s) = 0 Then
        Err.Raise vbObjectError + 641, "modCms.CMS_ResolveKey", "Empty ticker/key"
    End If
    If InStr(1, s, ".") > 0 Then
        CMS_ResolveKey = s
        Exit Function
    End If
    If Not gTickerIndex Is Nothing Then
        If gTickerIndex.Exists(s) Then
            CMS_ResolveKey = gTickerIndex(s)
            Exit Function
        End If
    End If
    Err.Raise vbObjectError + 642, "modCms.CMS_ResolveKey", _
        "Ticker '" & s & "' is not registered. Call register_curve / " & _
        "register_curve_by_range, or GET it once with the full tuple."
End Function

' Stored curve by ticker or key (Nothing if registered key has gone missing).
Public Function CMS_GetStored(ByVal TickerOrKey As String) As cCmsCurve
    Set CMS_GetStored = StoreGet(CMS_ResolveKey(TickerOrKey))
End Function

' Like CMS_GetStored but raises instead of returning Nothing.
Private Function RequireStored(ByVal TickerOrKey As String) As cCmsCurve
    Dim Key As String
    Key = CMS_ResolveKey(TickerOrKey)
    Set RequireStored = StoreGet(Key)
    If RequireStored Is Nothing Then
        Err.Raise vbObjectError + 643, "modCms.RequireStored", _
            "Curve '" & Key & "' is not in the store - GET it or register it first."
    End If
End Function

' Identity from explicit parts, or from the registry when only Ticker is
' given (Ccy AND DebtClass empty). Explicit parts also register the identity.
Private Function BuildIdentity(ByVal Ticker As String, ByVal Ccy As String, _
                               ByVal DebtClass As String, ByVal Product As String, _
                               ByVal QuoteConvention As String) As cCmsCurve
    Dim cv As New cCmsCurve
    Dim stored As cCmsCurve
    Dim Key As String

    If Len(Trim$(Ccy)) = 0 And Len(Trim$(DebtClass)) = 0 Then
        ' Ticker-only (or raw fourTuple key)
        Key = CMS_ResolveKey(Ticker)
        Set stored = StoreGet(Key)
        If stored Is Nothing Then
            ParseKeyInto Key, cv
            If Len(QuoteConvention) > 0 Then cv.QuoteConvention = UCase$(QuoteConvention)
        Else
            cv.Init stored.Ticker, stored.Ccy, stored.DebtClass, stored.Product, _
                    IIf(Len(QuoteConvention) > 0, QuoteConvention, stored.QuoteConvention)
            cv.LabelOverride = stored.LabelOverride
        End If
    ElseIf Len(Trim$(Ccy)) = 0 Or Len(Trim$(DebtClass)) = 0 Or Len(Trim$(Product)) = 0 Then
        Err.Raise vbObjectError + 644, "modCms.BuildIdentity", _
            "Provide either the Ticker alone (registered curve) or the full Ticker/Ccy/DebtClass/Product"
    Else
        If Len(QuoteConvention) = 0 Then QuoteConvention = "SPREADS"
        cv.Init Ticker, Ccy, DebtClass, Product, QuoteConvention
        RegisterIdentity cv
    End If
    Set BuildIdentity = cv
End Function

' Identity-only registration: keep an existing stored curve (and its quotes),
' just make sure the ticker index points at it.
Private Sub RegisterIdentity(ByVal Curve As cCmsCurve)
    If CMS_Store().Exists(Curve.Key) Then
        IndexTicker gStore(Curve.Key)
    Else
        StoreUpsert Curve
    End If
End Sub

Private Sub IndexTicker(ByVal Curve As cCmsCurve)
    If gTickerIndex Is Nothing Then
        Set gTickerIndex = CreateObject("Scripting.Dictionary")
        gTickerIndex.CompareMode = vbTextCompare
    End If
    If gTickerIndex.Exists(Curve.Ticker) Then
        If StrComp(CStr(gTickerIndex(Curve.Ticker)), Curve.Key, vbTextCompare) <> 0 Then
            Debug.Print "modCms: ticker '" & Curve.Ticker & "' remapped " & _
                        gTickerIndex(Curve.Ticker) & " -> " & Curve.Key
        End If
        gTickerIndex(Curve.Ticker) = Curve.Key
    Else
        gTickerIndex.Add Curve.Ticker, Curve.Key
    End If
End Sub

' =============================================================================
'  CURVE ANALYTICS (off the store - no I/O)
'  Unregistered tickers raise; missing quotes return Empty.
' =============================================================================

Public Function CMS_Quote(ByVal TickerOrKey As String, _
                          Optional ByVal Tenor As String = "5Y") As Variant
    CMS_Quote = RequireStored(TickerOrKey).Quote(Tenor)
End Function

' Slope between two tenors of one curve: Quote(TenorB) - Quote(TenorA).
' Default is the 5s10s: CMS_CurveDiff("AEP") = 10Y - 5Y.
Public Function CMS_CurveDiff(ByVal TickerOrKey As String, _
                              Optional ByVal TenorA As String = "5Y", _
                              Optional ByVal TenorB As String = "10Y") As Variant
    Dim cv As cCmsCurve
    Dim a As Variant, b As Variant
    Set cv = RequireStored(TickerOrKey)
    a = cv.Quote(TenorA): b = cv.Quote(TenorB)
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    CMS_CurveDiff = CDbl(b) - CDbl(a)
End Function

' Ratio between two tenors of one curve: Quote(TenorB) / Quote(TenorA).
Public Function CMS_CurveRatio(ByVal TickerOrKey As String, _
                               Optional ByVal TenorA As String = "5Y", _
                               Optional ByVal TenorB As String = "10Y") As Variant
    Dim cv As cCmsCurve
    Dim a As Variant, b As Variant
    Set cv = RequireStored(TickerOrKey)
    a = cv.Quote(TenorA): b = cv.Quote(TenorB)
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    If CDbl(a) = 0 Then Exit Function
    CMS_CurveRatio = CDbl(b) / CDbl(a)
End Function

' Spread between two tickers at one tenor: QuoteA - QuoteB.
Public Function CMS_TickerDiff(ByVal TickerA As String, ByVal TickerB As String, _
                               Optional ByVal Tenor As String = "5Y") As Variant
    Dim a As Variant, b As Variant
    a = RequireStored(TickerA).Quote(Tenor)
    b = RequireStored(TickerB).Quote(Tenor)
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    CMS_TickerDiff = CDbl(a) - CDbl(b)
End Function

' Ratio between two tickers at one tenor: QuoteA / QuoteB.
Public Function CMS_TickerRatio(ByVal TickerA As String, ByVal TickerB As String, _
                                Optional ByVal Tenor As String = "5Y") As Variant
    Dim a As Variant, b As Variant
    a = RequireStored(TickerA).Quote(Tenor)
    b = RequireStored(TickerB).Quote(Tenor)
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    If CDbl(b) = 0 Then Exit Function
    CMS_TickerRatio = CDbl(a) / CDbl(b)
End Function

' =============================================================================
'  PENDING 5Y LEVELS
'  Stage amended 5Y marks on the cached curves (e.g. from a sheet-change
'  handler), then drain them all into one async SET when the user presses
'  the mark button. No cell re-reads needed at mark time.
' =============================================================================

' Stage (or restage) a pending 5Y level on a registered curve.
Public Sub CMS_StagePending5y(ByVal TickerOrKey As String, ByVal Level As Double)
    Dim cv As cCmsCurve
    Set cv = RequireStored(TickerOrKey)
    cv.Pending5y = Level
    NotifySubscribers cv
End Sub

' Un-stage a pending level (user deleted their amend). Returns the cached
' ORIGINAL 5Y so the caller can restore it into the cell (Empty if the curve
' was never fetched).
Public Function CMS_ClearPending5y(ByVal TickerOrKey As String) As Variant
    Dim cv As cCmsCurve
    Set cv = RequireStored(TickerOrKey)
    cv.Pending5y = Empty
    NotifySubscribers cv
    CMS_ClearPending5y = cv.Quotes(CMS_IDX_5Y)
End Function

Public Function CMS_Pending5y(ByVal TickerOrKey As String) As Variant
    CMS_Pending5y = RequireStored(TickerOrKey).Pending5y
End Function

Public Function CMS_HasPending(ByVal TickerOrKey As String) As Boolean
    CMS_HasPending = Not IsEmpty(RequireStored(TickerOrKey).Pending5y)
End Function

' Keys of every curve with a staged level (0-based array; empty array if none).
Public Function CMS_PendingKeys() As Variant
    Dim k As Variant
    Dim out() As Variant
    Dim n As Long
    ReDim out(0 To CMS_Store().Count)   ' upper bound; trimmed below
    For Each k In gStore.Keys
        If Not IsEmpty(gStore(k).Pending5y) Then
            out(n) = k
            n = n + 1
        End If
    Next k
    If n = 0 Then
        CMS_PendingKeys = Array()
    Else
        ReDim Preserve out(0 To n - 1)
        CMS_PendingKeys = out
    End If
End Function

Public Function CMS_PendingCount() As Long
    Dim k As Variant
    For Each k In CMS_Store().Keys
        If Not IsEmpty(gStore(k).Pending5y) Then CMS_PendingCount = CMS_PendingCount + 1
    Next k
End Function

Public Sub CMS_ClearAllPending()
    Dim k As Variant
    For Each k In CMS_Store().Keys
        If Not IsEmpty(gStore(k).Pending5y) Then
            gStore(k).Pending5y = Empty
            NotifySubscribers gStore(k)
        End If
    Next k
End Sub

' Drain: SET every staged curve off its pending 5Y (full curve computed from
' the cached quotes), clearing the staged levels as they are consumed.
' Returns Nothing when nothing is staged.
Public Function CMS_MarkPendingAsync(Optional ByVal OnCurve As String = "", _
                                     Optional ByVal OnAllDone As String = "", _
                                     Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim Keys As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim cv As cCmsCurve

    Keys = CMS_PendingKeys()
    If UBound(Keys) - LBound(Keys) + 1 <= 0 Then Exit Function   ' nothing staged

    ReDim arr(0 To UBound(Keys) - LBound(Keys), 0 To 1)
    For i = 0 To UBound(Keys) - LBound(Keys)
        Set cv = gStore(CStr(Keys(LBound(Keys) + i)))
        arr(i, 0) = cv.Key
        arr(i, 1) = cv.Pending5y
        cv.Pending5y = Empty
    Next i

    Set CMS_MarkPendingAsync = CMS_SetCurves5yAsync(arr, OnCurve, OnAllDone, TimeoutMs)
End Function

' =============================================================================
'  WHOLE-STORE CONVENIENCE WRAPPERS
'  Refresh / fetch everything the workbook knows about in one call. All of
'  these return Nothing when there is nothing applicable to do.
' =============================================================================

' Every fourTuple key in the store (registered and/or fetched), 0-based.
Public Function CMS_RegisteredKeys() As Variant
    CMS_RegisteredKeys = CMS_Store().Keys
End Function

' Every registered ticker, 0-based.
Public Function CMS_RegisteredTickers() As Variant
    If gTickerIndex Is Nothing Then
        CMS_RegisteredTickers = Array()
    Else
        CMS_RegisteredTickers = gTickerIndex.Keys
    End If
End Function

Public Function CMS_RegisteredCount() As Long
    CMS_RegisteredCount = CMS_Store().Count
End Function

' Force-refresh EVERY curve in the store - one parallel async GET across the
' lot. This INCLUDES curves that were only registered and never fetched
' (registration puts the identity in the store), so register + GetAll/
' RefreshAll is the standard sheet-load sequence. Raises if the store is
' empty - that means registration didn't happen.
Public Function CMS_RefreshAllAsync(Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                    Optional ByVal QuoteDate As Date, _
                                    Optional ByVal OnCurve As String = "", _
                                    Optional ByVal OnAllDone As String = "", _
                                    Optional ByVal CurvesPerRequest As Long = 1, _
                                    Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    If CMS_Store().Count = 0 Then
        Err.Raise vbObjectError + 650, "modCms.CMS_RefreshAllAsync", _
            "The curve store is empty - nothing to fetch. Call register_curve / " & _
            "register_curve_by_range (or GET with full tuples) first."
    End If
    Set CMS_RefreshAllAsync = LaunchGetForKeys(CMS_Store().Keys, Tag, QuoteDate, _
                                               OnCurve, OnAllDone, CurvesPerRequest, TimeoutMs)
End Function

' Alias for discoverability: "get all" = refresh all registered curves.
Public Function CMS_GetAllAsync(Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                Optional ByVal QuoteDate As Date, _
                                Optional ByVal OnCurve As String = "", _
                                Optional ByVal OnAllDone As String = "", _
                                Optional ByVal CurvesPerRequest As Long = 1, _
                                Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Set CMS_GetAllAsync = CMS_RefreshAllAsync(Tag, QuoteDate, OnCurve, OnAllDone, _
                                              CurvesPerRequest, TimeoutMs)
End Function

' Re-GET only the curves whose last attempt FAILED. Also clears their
' auto-fetch cooldown stamps so CURVE() cells resume updating.
Public Function CMS_RefreshFailedAsync(Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                       Optional ByVal QuoteDate As Date, _
                                       Optional ByVal OnCurve As String = "", _
                                       Optional ByVal OnAllDone As String = "", _
                                       Optional ByVal CurvesPerRequest As Long = 1, _
                                       Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim k As Variant
    Dim picked As Collection
    Set picked = New Collection
    For Each k In CMS_Store().Keys
        If gStore(k).Status = CMS_STATUS_FAILED Then
            picked.Add CStr(k)
            If Not gFetchStamp Is Nothing Then
                If gFetchStamp.Exists(CStr(k)) Then gFetchStamp.Remove CStr(k)
            End If
        End If
    Next k
    Set CMS_RefreshFailedAsync = LaunchGetForCollection(picked, Tag, QuoteDate, _
                                                        OnCurve, OnAllDone, CurvesPerRequest, TimeoutMs)
End Function

' Re-GET curves never fetched, or fetched more than OlderThanMinutes ago.
Public Function CMS_RefreshStaleAsync(Optional ByVal OlderThanMinutes As Long = 15, _
                                      Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                      Optional ByVal QuoteDate As Date, _
                                      Optional ByVal OnCurve As String = "", _
                                      Optional ByVal OnAllDone As String = "", _
                                      Optional ByVal CurvesPerRequest As Long = 1, _
                                      Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim k As Variant
    Dim picked As Collection
    Set picked = New Collection
    For Each k In CMS_Store().Keys
        If gStore(k).FetchedAt = 0 Then
            picked.Add CStr(k)
        ElseIf DateDiff("n", gStore(k).FetchedAt, Now()) >= OlderThanMinutes Then
            picked.Add CStr(k)
        End If
    Next k
    Set CMS_RefreshStaleAsync = LaunchGetForCollection(picked, Tag, QuoteDate, _
                                                       OnCurve, OnAllDone, CurvesPerRequest, TimeoutMs)
End Function

' Dump the whole store (key + 17 tenors + scalars + pending) below a cell.
Public Sub CMS_WriteStoreToRange(ByVal TopLeft As Range, _
                                 Optional ByVal IncludeHeader As Boolean = True)
    Dim arr As Variant
    arr = CMS_StoreToArray(, IncludeHeader)
    TopLeft.Resize(UBound(arr, 1) + 1, UBound(arr, 2) + 1).Value2 = arr
End Sub

Private Function LaunchGetForKeys(ByVal Keys As Variant, ByVal Tag As String, _
                                  ByVal QuoteDate As Date, ByVal OnCurve As String, _
                                  ByVal OnAllDone As String, ByVal CurvesPerRequest As Long, _
                                  ByVal TimeoutMs As Long) As cCmsBatch
    Dim arr() As Variant
    Dim i As Long
    If UBound(Keys) < LBound(Keys) Then Exit Function   ' empty store -> Nothing
    ReDim arr(0 To UBound(Keys) - LBound(Keys), 0 To 0)
    For i = 0 To UBound(Keys) - LBound(Keys)
        arr(i, 0) = Keys(LBound(Keys) + i)
    Next i
    Set LaunchGetForKeys = CMS_GetCurvesAsync(arr, Tag, QuoteDate, OnCurve, OnAllDone, _
                                              CurvesPerRequest, TimeoutMs)
End Function

Private Function LaunchGetForCollection(ByVal Picked As Collection, ByVal Tag As String, _
                                        ByVal QuoteDate As Date, ByVal OnCurve As String, _
                                        ByVal OnAllDone As String, ByVal CurvesPerRequest As Long, _
                                        ByVal TimeoutMs As Long) As cCmsBatch
    Dim arr() As Variant
    Dim i As Long
    If Picked.Count = 0 Then Exit Function              ' nothing applicable -> Nothing
    ReDim arr(0 To Picked.Count - 1, 0 To 0)
    For i = 1 To Picked.Count
        arr(i - 1, 0) = Picked(i)
    Next i
    Set LaunchGetForCollection = CMS_GetCurvesAsync(arr, Tag, QuoteDate, OnCurve, OnAllDone, _
                                                    CurvesPerRequest, TimeoutMs)
End Function

' =============================================================================
'  CELL SUBSCRIPTIONS - reactive recalc of sheet references
'
'  Cells (or any ranges) subscribe to a curve; whenever that curve's store
'  entry changes (GET arrives, SET confirms or fails, pending level staged or
'  cleared) the subscribed cells are marked dirty and recalculated. The CURVE
'  sheet UDFs self-subscribe via Application.Caller, so =CURVE("AEP","5Y")
'  updates by itself as data lands; CMS_SubscribeRange is the manual hook for
'  anything else (e.g. the row you want refreshed when a pending mark lands).
' =============================================================================

' Subscribe a range to a curve (by ticker or key; unregistered tickers are
' bucketed by name and start firing once the ticker is registered+fetched).
Public Sub CMS_SubscribeRange(ByVal TickerOrKey As String, ByVal Target As Range)
    SubscribeAddress BucketFor(TickerOrKey), AddrOf(Target)
End Sub

Public Sub CMS_UnsubscribeRange(ByVal TickerOrKey As String, ByVal Target As Range)
    Dim bucket As String
    bucket = BucketFor(TickerOrKey)
    If gSubs Is Nothing Then Exit Sub
    If Not gSubs.Exists(bucket) Then Exit Sub
    If gSubs(bucket).Exists(AddrOf(Target)) Then gSubs(bucket).Remove AddrOf(Target)
End Sub

' Drop all subscriptions, or just one curve's.
Public Sub CMS_UnsubscribeAll(Optional ByVal TickerOrKey As String = "")
    If gSubs Is Nothing Then Exit Sub
    If Len(TickerOrKey) = 0 Then
        gSubs.RemoveAll
    ElseIf gSubs.Exists(BucketFor(TickerOrKey)) Then
        gSubs.Remove BucketFor(TickerOrKey)
    End If
End Sub

Public Function CMS_SubscriptionCount() As Long
    Dim b As Variant
    If gSubs Is Nothing Then Exit Function
    For Each b In gSubs.Keys
        CMS_SubscriptionCount = CMS_SubscriptionCount + gSubs(b).Count
    Next b
End Function

' Recalc every cell subscribed to this curve (called from the delivery path
' and the pending-staging functions; runs in a macro context).
Public Sub NotifySubscribers(ByVal Curve As cCmsCurve)
    NotifyBucket Curve.Key
    NotifyBucket Curve.Ticker
End Sub

Private Sub NotifyBucket(ByVal Bucket As String)
    Dim addrs As Variant
    Dim a As Variant
    Dim rng As Range
    Dim bang As Long

    If gSubs Is Nothing Then Exit Sub
    Bucket = UCase$(Trim$(Bucket))
    If Len(Bucket) = 0 Then Exit Sub
    If Not gSubs.Exists(Bucket) Then Exit Sub

    addrs = gSubs(Bucket).Keys   ' snapshot: recalc may re-subscribe
    For Each a In addrs
        Set rng = Nothing
        On Error Resume Next
        bang = InStrRev(CStr(a), "!")
        Set rng = ThisWorkbook.Worksheets(Left$(CStr(a), bang - 1)).Range(Mid$(CStr(a), bang + 1))
        On Error GoTo 0
        If rng Is Nothing Then
            gSubs(Bucket).Remove a   ' sheet/cell gone - prune
        Else
            On Error Resume Next
            Select Case CMS_GetNotifyMode()
                Case CMS_NOTIFY_CALC:
                    ' Heaviest: force the subscribed cells to recalc NOW.
                    If Application.Calculation = xlCalculationManual Then
                        rng.Calculate
                    Else
                        rng.Dirty
                    End If
                Case CMS_NOTIFY_DIRTY:
                    ' Default: cheap marking. Automatic mode recalcs after
                    ' the delivery macro; manual mode picks the cells up at
                    ' the user's next recalc (Shift+F9 / F9).
                    rng.Dirty
                Case Else  ' CMS_NOTIFY_OFF
                    ' leave the cells alone
            End Select
            If Err.Number <> 0 Then
                Debug.Print "CMS notify: recalc failed for " & CStr(a) & _
                            " (" & Err.Description & ")"
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next a
End Sub

Private Function AddrOf(ByVal Target As Range) As String
    AddrOf = Target.Worksheet.Name & "!" & Target.Address
End Function

' Resolved key when possible, otherwise the raw upper-cased ticker (so cells
' can subscribe before the ticker is registered).
Private Function BucketFor(ByVal TickerOrKey As String) As String
    On Error Resume Next
    BucketFor = CMS_ResolveKey(TickerOrKey)
    On Error GoTo 0
    If Len(BucketFor) = 0 Then BucketFor = UCase$(Trim$(TickerOrKey))
End Function

Private Sub SubscribeAddress(ByVal Bucket As String, ByVal Addr As String)
    If Len(Bucket) = 0 Or Len(Addr) = 0 Then Exit Sub
    If gSubs Is Nothing Then
        Set gSubs = CreateObject("Scripting.Dictionary")
        gSubs.CompareMode = vbTextCompare
    End If
    If Not gSubs.Exists(Bucket) Then
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.CompareMode = vbTextCompare
        gSubs.Add Bucket, d
    End If
    gSubs(Bucket)(Addr) = True
End Sub

' =============================================================================
'  AUTO-FETCH QUEUE
'  Sheet UDFs cannot launch HTTP (or much of anything) from calc context, so
'  they queue the key here; the queue is drained right after calculation ends
'  (Workbook_SheetCalculate) or by the delivery watchdog - both proper macro
'  contexts. When the GET lands, subscriptions recalc the waiting cells.
' =============================================================================

' Queue a REGISTERED curve for fetching. Cooldown prevents refetch storms
' from repeated recalcs. Curves whose last fetch FAILED are not auto-retried
' (the cell shows #N/A and CURVE(t,"Error") shows why) - refresh manually.
Public Sub QueueAutoFetch(ByVal Key As String)
    If gWanted Is Nothing Then
        Set gWanted = CreateObject("Scripting.Dictionary")
        gWanted.CompareMode = vbTextCompare
        Set gFetchStamp = CreateObject("Scripting.Dictionary")
        gFetchStamp.CompareMode = vbTextCompare
    End If
    If gFetchStamp.Exists(Key) Then
        If DateDiff("s", gFetchStamp(Key), Now()) < AUTOFETCH_COOLDOWN_SEC Then Exit Sub
    End If
    gWanted(Key) = True
    ' Self-driving in manual calc mode: arm the OnTime watchdog so the queue
    ' drains ~1s from now even if no SheetCalculate event ever fires and no
    ' batch is currently in flight. (Application.OnTime is one of the few
    ' state changes a UDF may make; guarded On Error inside EnsureWatchdog.)
    EnsureWatchdog
End Sub

' Launch one async GET for everything queued (live and T-1 close queues).
' Safe to call often (no-op when empty). Wired into Workbook_SheetCalculate
' and the watchdog tick.
Public Sub CMS_FlushAutoFetch()
    Dim ks As Variant
    Dim arr() As Variant
    Dim i As Long
    Dim pb As cCmsBatch
    Dim pd As Date

    ' --- live quotes ---
    If Not gWanted Is Nothing Then
        If gWanted.Count > 0 Then
            ks = gWanted.Keys
            gWanted.RemoveAll
            ReDim arr(0 To UBound(ks), 0 To 0)
            For i = 0 To UBound(ks)
                arr(i, 0) = ks(i)
                gFetchStamp(ks(i)) = Now()
            Next i
            On Error Resume Next
            CMS_GetCurvesAsync arr
            If Err.Number <> 0 Then Debug.Print "CMS auto-fetch failed to launch: " & Err.Description
            On Error GoTo 0
        End If
    End If

    ' --- T-1 business-day close (NYOISCLOSE) ---
    If Not gPrevWanted Is Nothing Then
        If gPrevWanted.Count > 0 Then
            ks = gPrevWanted.Keys
            gPrevWanted.RemoveAll
            ReDim arr(0 To UBound(ks), 0 To 0)
            For i = 0 To UBound(ks)
                arr(i, 0) = ks(i)
                gPrevStamp(ks(i)) = Now()
            Next i
            pd = CMS_PrevBizDay()
            On Error Resume Next
            Set pb = CMS_GetCurvesAsync(arr, CMS_TAG_NYOISCLOSE, pd)
            If Err.Number <> 0 Then
                Debug.Print "CMS T-1 fetch failed to launch: " & Err.Description
            ElseIf Not pb Is Nothing Then
                pb.PrevCloseMode = True
                pb.PrevAsOf = pd
            End If
            On Error GoTo 0
        End If
    End If
End Sub

' =============================================================================
'  T-1 BUSINESS-DAY CLOSE (NYOISCLOSE)
'  After a curve's first live GET of the day, its previous-business-day close
'  is fetched once (async) and cached on the same store entry - historical
'  closes don't change intraday, so once per day is enough. Default ON;
'  toggle with CMS_SetAutoPrevClose False (then use CMS_GetPrevCloseAsync
'  manually). Access via curve.PrevQuote("5Y"), CMS_PrevQuote, CMS_DayChange,
'  or =CURVE("AEP","Prev5Y") / "PrevDate" / "PrevFetchedAt" in cells.
' =============================================================================

' Previous BUSINESS day (weekends skipped; holidays too if a 'CmsHolidays'
' named range of dates exists in the workbook).
Public Function CMS_PrevBizDay(Optional ByVal FromDate As Date) As Date
    Dim hol As Variant
    If FromDate = 0 Then FromDate = Date
    On Error Resume Next
    hol = ThisWorkbook.Names("CmsHolidays").RefersToRange.Value2
    On Error GoTo 0
    If IsEmpty(hol) Then
        CMS_PrevBizDay = CDate(Application.WorksheetFunction.WorkDay(FromDate, -1))
    Else
        CMS_PrevBizDay = CDate(Application.WorksheetFunction.WorkDay(FromDate, -1, hol))
    End If
End Function

' Queue a curve for a T-1 close fetch if it doesn't already have today's
' (i.e. the current T-1 date's) close cached. Cooldown-guarded; respects the
' CMS_SetAutoPrevClose toggle.
Public Sub QueuePrevCloseFetch(ByVal Key As String)
    Dim stored As cCmsCurve

    If Not CMS_AutoPrevCloseEnabled() Then Exit Sub
    Set stored = StoreGet(Key)
    If stored Is Nothing Then Exit Sub
    If stored.PrevDate = CMS_PrevBizDay() Then Exit Sub   ' already current

    If gPrevWanted Is Nothing Then
        Set gPrevWanted = CreateObject("Scripting.Dictionary")
        gPrevWanted.CompareMode = vbTextCompare
        Set gPrevStamp = CreateObject("Scripting.Dictionary")
        gPrevStamp.CompareMode = vbTextCompare
    End If
    If gPrevStamp.Exists(Key) Then
        If DateDiff("s", gPrevStamp(Key), Now()) < PREVFETCH_COOLDOWN_SEC Then Exit Sub
    End If
    gPrevWanted(Key) = True
    EnsureWatchdog
End Sub

' Route a prev-close GET result into the stored curve's PrevQuotes (the live
' quotes are untouched). Called by cCmsBatch for PrevCloseMode batches.
Public Sub StorePrevClose(ByVal Curve As cCmsCurve, ByVal AsOf As Date)
    Dim stored As cCmsCurve
    Set stored = StoreGet(Curve.Key)
    If stored Is Nothing Then Exit Sub
    stored.PrevQuotes = Curve.Quotes
    stored.PrevDate = AsOf
    stored.PrevFetchedAt = Now()
    NotifySubscribers stored
End Sub

' Explicit T-1 close fetch: any curve spec (tuples / tickers / single
' ticker), or everything in the store when omitted. Returns Nothing when
' there is nothing to fetch.
Public Function CMS_GetPrevCloseAsync(Optional ByVal CurveSpec As Variant, _
                                      Optional ByVal OnCurve As String = "", _
                                      Optional ByVal OnAllDone As String = "", _
                                      Optional ByVal CurvesPerRequest As Long = 1, _
                                      Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As cCmsBatch
    Dim b As cCmsBatch
    Dim pd As Date
    pd = CMS_PrevBizDay()
    If IsMissing(CurveSpec) Or IsEmpty(CurveSpec) Then
        Set b = LaunchGetForKeys(CMS_Store().Keys, CMS_TAG_NYOISCLOSE, pd, _
                                 OnCurve, OnAllDone, CurvesPerRequest, TimeoutMs)
    Else
        Set b = CMS_GetCurvesAsync(CurveSpec, CMS_TAG_NYOISCLOSE, pd, _
                                   OnCurve, OnAllDone, CurvesPerRequest, TimeoutMs)
    End If
    If Not b Is Nothing Then
        b.PrevCloseMode = True
        b.PrevAsOf = pd
    End If
    Set CMS_GetPrevCloseAsync = b
End Function

' Cached T-1 close quote (Empty if not cached yet).
Public Function CMS_PrevQuote(ByVal TickerOrKey As String, _
                              Optional ByVal Tenor As String = "5Y") As Variant
    CMS_PrevQuote = RequireStored(TickerOrKey).PrevQuote(Tenor)
End Function

' Day change: live quote minus T-1 close at a tenor (Empty if either missing).
Public Function CMS_DayChange(ByVal TickerOrKey As String, _
                              Optional ByVal Tenor As String = "5Y") As Variant
    Dim cv As cCmsCurve
    Dim a As Variant, b As Variant
    Set cv = RequireStored(TickerOrKey)
    a = cv.Quote(Tenor): b = cv.PrevQuote(Tenor)
    If IsEmpty(a) Or IsEmpty(b) Then Exit Function
    CMS_DayChange = CDbl(a) - CDbl(b)
End Function

' =============================================================================
'  REQUEST CONSTRUCTION
' =============================================================================

Public Function CMS_CompileFourTuple(ByVal Ticker As String, ByVal Ccy As String, _
                                     ByVal DebtClass As String, ByVal Product As String) As String
    CMS_CompileFourTuple = UCase$(Ticker & "." & Ccy & "." & DebtClass & "." & Product)
End Function

' DebtClass from Tier/RefId + DocClause ("~" for SPECREF products, "_" otherwise)
Public Function CMS_CompileDebtClass(ByVal TierOrRefId As String, ByVal DocClause As String, _
                                     ByVal Product As String) As String
    Dim concatChar As String
    If UCase$(Product) = "SPECREF" Then concatChar = "~" Else concatChar = "_"
    CMS_CompileDebtClass = UCase$(TierOrRefId)
    If Len(DocClause) > 0 Then CMS_CompileDebtClass = CMS_CompileDebtClass & concatChar & UCase$(DocClause)
End Function

Public Function IsStandardContractProduct(ByVal Product As String) As Boolean
    Dim prefixes As Variant
    Dim p As Variant
    prefixes = Array("BLCDS", "SNAC", "STAJ", "STAS", "STEA", "STEC", "STEJ", "STEM", "SUKU")
    Product = UCase$(Product)
    For Each p In prefixes
        If Left$(Product, Len(p)) = p Then
            IsStandardContractProduct = True
            Exit Function
        End If
    Next p
End Function

' --- Rosetta builders --------------------------------------------------------

' GET: one <marketSet> per curve inside a single <marketData>
Private Function BuildRosettaGet(ByVal Curves As Collection, ByVal Tag As String, _
                                 ByVal DateStr As String) As String
    Dim sb As String
    Dim cv As cCmsCurve
    Dim convTag As String
    Dim i As Long

    For i = 1 To Curves.Count
        Set cv = Curves(i)
        convTag = cv.QuoteConvention
        ' Multi-row conventions are not allowed on the wire (legacy bulk rule)
        If convTag = "RUNNING" Or convTag = "EXTENDED" Then convTag = "UPFRONT"
        If Len(convTag) = 0 Then convTag = "SPREADS"
        sb = sb & "<marketSet>" _
                & "<location>" & DEFAULT_LOCATION & "</location>" _
                & "<currency>" & XmlEsc(cv.Ccy) & "</currency>" _
                & "<dataSource>" & XmlEsc(Tag) & "</dataSource>" _
                & "<version>CLOSE</version><type>creditSpread</type>" _
                & "<label>" & XmlEsc(cv.Label) & "</label>" _
                & "<genericKeys>" _
                & "<nameValuePair name=""tag"">" & XmlEsc(Tag) & "</nameValuePair>" _
                & "<nameValuePair name=""returnCreditCurveType"">" & XmlEsc(convTag) & "</nameValuePair>" _
                & "</genericKeys>" _
                & "</marketSet>"
    Next i

    BuildRosettaGet = "<Rosetta version=""" & ROSETTA_VERSION & """><market><marketData>" _
                    & "<action>GET</action><date>" & DateStr & "</date>" _
                    & sb & "</marketData></market></Rosetta>"
End Function

' SET: one <point> per SET tenor, contractualSpread on standard-contract
' products (exactly how xmlBats built them; -777 is just a SNAC/STEC marker)
Private Function BuildRosettaSet(ByVal Curve As cCmsCurve, ByVal Quotes12 As Variant, _
                                 ByVal RecoveryRate As Double) As String
    Dim points As String
    Dim tenors As Variant
    Dim t As Variant
    Dim i As Long
    Dim dataType As String
    Dim hasContractual As Boolean
    Dim pm As String, per As String
    Dim lbl As String

    Select Case Curve.QuoteConvention
        Case "UPFRONT", "EXTENDED":
            dataType = "creditUpfront"
        Case Else:  ' SPREADS, QUOTED_SPREADS, QUOTED, RUNNING
            dataType = "creditSpread"
    End Select
    hasContractual = IsStandardContractProduct(Curve.Product)
    lbl = XmlEsc(Curve.Label)

    tenors = CMS_SetTenors12()
    For i = 0 To UBound(tenors)
        If Not IsEmpty(Quotes12(i)) And IsNumeric(Quotes12(i)) Then
            t = CStr(tenors(i))
            pm = Left$(t, Len(t) - 1)
            per = Right$(t, 1)
            points = points & "<point>" _
                   & "<label>" & lbl & "</label>" _
                   & "<" & dataType & ">" _
                   & "<issuerTicker>" & XmlEsc(Curve.Ticker) & "</issuerTicker>" _
                   & "<debtClass>" & XmlEsc(Curve.DebtClass) & "</debtClass>" _
                   & "<currency>" & XmlEsc(Curve.Ccy) & "</currency>" _
                   & "<creditCurveType>" & XmlEsc(Curve.Product) & "</creditCurveType>" _
                   & "<periodMultiplier>" & pm & "</periodMultiplier>" _
                   & "<period>" & per & "</period>"
            If hasContractual Then
                points = points & "<contractualSpread>" & CONTRACTUAL_SPREAD_DUMMY & "</contractualSpread>"
            End If
            points = points & "</" & dataType & ">" _
                   & "<value>" & NumToXml(CDbl(Quotes12(i))) & "</value>" _
                   & "</point>"
        End If
    Next i

    If RecoveryRate > 0 Then
        points = points & "<point>" _
               & "<label>" & lbl & "</label>" _
               & "<issuerRecovery>" _
               & "<issuerTicker>" & XmlEsc(Curve.Ticker) & "</issuerTicker>" _
               & "<debtClass>" & XmlEsc(Curve.DebtClass) & "</debtClass>" _
               & "<currency>" & XmlEsc(Curve.Ccy) & "</currency>" _
               & "<creditCurveType>" & XmlEsc(Curve.Product) & "</creditCurveType>" _
               & "</issuerRecovery>" _
               & "<value>" & NumToXml(RecoveryRate) & "</value>" _
               & "</point>"
    End If

    BuildRosettaSet = "<Rosetta version=""" & ROSETTA_VERSION & """><market><marketData>" _
                    & "<action>SET</action><date>" & QuoteDateString(0) & "</date>" _
                    & "<marketSet>" _
                    & "<location>" & DEFAULT_LOCATION & "</location>" _
                    & "<logicalTime>" & DEFAULT_LOGICAL_TIME & "</logicalTime>" _
                    & points _
                    & "</marketSet></marketData></market></Rosetta>"
End Function

' --- user-metadata (identical to the legacy CompileUserMetaData) -------------

Private Function UserMetaXml() As String
    Dim userId As String
    EnsureIdentity
    If Len(gImpersonatedUser) > 0 Then userId = gImpersonatedUser Else userId = gUserName
    UserMetaXml = "<user-metadata><user-details>" _
                & "<user-id>" & XmlEsc(userId) & "</user-id>" _
                & "<user-domain>" & XmlEsc(gUserDomain) & "</user-domain>" _
                & "<application-id>" & CMS_APP_ID & "</application-id>" _
                & "<application-version>" & CMS_APP_VERSION & "</application-version>" _
                & "<application-batch-id>" & XmlEsc(ThisWorkbook.FullName) & "</application-batch-id>" _
                & "<id>" & XmlEsc(gUserName) & "</id>" _
                & "<host-id>" & XmlEsc(gMachineName) & "</host-id>" _
                & "</user-details></user-metadata>"
End Function

Private Sub EnsureIdentity()
    If Len(gUserName) > 0 Then Exit Sub
    gUserName = LCase$(User_GetUserName())
    gUserDomain = UCase$(User_GetDomain())
    gMachineName = UCase$(User_GetMachineName())
End Sub

' --- transports ---------------------------------------------------------------

' Wrap a Rosetta payload for the configured transport and hand the batch a
' ready-to-send request. Operation is "getMarketData" / "setMarketData".
Private Sub PostCmsRequest(ByVal Batch As cCmsBatch, ByVal Operation As String, _
                           ByVal Rosetta As String, ByVal TimeoutMs As Long, _
                           ByVal CurveKeys As Variant)
    Dim url As String
    Dim body As String
    Dim headers As String
    EnsureConfig

    If gTransport = "REST" Then
        url = gEndpointUrl
        If Right$(url, 1) = "/" Then url = Left$(url, Len(url) - 1)
        url = url & "/api/v1/" & Operation
        body = BuildRestBody(Operation, Rosetta)
        headers = "Content-Type: application/json"
    Else
        url = gEndpointUrl
        body = BuildSoapEnvelope(Operation, Rosetta)
        headers = "Content-Type: text/xml; charset=utf-8" & vbLf & "SOAPAction: """""
    End If

    If modBridge.gHttpVerbose Then
        Debug.Print "CMS " & Operation & " -> " & url & " (" & Len(body) & " chars, " & _
                    (UBound(CurveKeys) - LBound(CurveKeys) + 1) & " curve(s))"
    End If

    Batch.SendRequest url, body, headers, TimeoutMs, CurveKeys
End Sub

' rpc/encoded SOAP 1.1 envelope - byte-compatible with what xmlBats sent.
Private Function BuildSoapEnvelope(ByVal Operation As String, ByVal Rosetta As String) As String
    BuildSoapEnvelope = _
        "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""no""?>" _
        & "<SOAP-ENV:Envelope xmlns:SOAPSDK1=""http://www.w3.org/2001/XMLSchema""" _
        & " xmlns:SOAPSDK2=""http://www.w3.org/2001/XMLSchema-instance""" _
        & " xmlns:SOAPSDK3=""http://schemas.xmlsoap.org/soap/encoding/""" _
        & " xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">" _
        & "<SOAP-ENV:Body SOAP-ENV:encodingStyle=""http://schemas.xmlsoap.org/soap/encoding/"">" _
        & "<SOAPSDK4:" & Operation & " xmlns:SOAPSDK4=""" & CMS_NAMESPACE & """>" _
        & "<string>" & XmlEsc(UserMetaXml()) & "</string>" _
        & "<string0>" & XmlEsc(Rosetta) & "</string0>" _
        & "</SOAPSDK4:" & Operation & ">" _
        & "</SOAP-ENV:Body></SOAP-ENV:Envelope>"
End Function

' Modernized JSON wrapper (/api/v1/*): same Rosetta inside <cms-request>.
Private Function BuildRestBody(ByVal Operation As String, ByVal Rosetta As String) As String
    Dim req As String
    EnsureIdentity
    req = "<cms-request><operation name=""" & Operation & """><arguments>" _
        & UserMetaXml() & Rosetta _
        & "</arguments></operation></cms-request>"
    BuildRestBody = "{""requestHeader"":{" _
        & """appId"":""" & JsonEsc(CMS_APP_ID) & """," _
        & """userId"":""" & JsonEsc(gUserName) & """," _
        & """hostId"":""" & JsonEsc(gMachineName) & """," _
        & """requestId"":""" & JsonEsc(CMS_APP_ID & "-" & Format$(Now(), "yyyymmddhhnnss")) & """," _
        & """timeStamp"":" & CStr(DateDiff("s", DateSerial(1970, 1, 1), Now())) _
        & "},""request"":""" & JsonEsc(req) & """}"
End Function

' --- launch helpers used by the public API -----------------------------------

Private Sub SendGetRequest(ByVal Batch As cCmsBatch, ByVal Curves As Collection, _
                           ByVal Tag As String, ByVal DateStr As String, ByVal TimeoutMs As Long)
    Dim Keys() As Variant
    Dim i As Long
    If Curves.Count = 0 Then Exit Sub
    ReDim Keys(0 To Curves.Count - 1)
    For i = 1 To Curves.Count
        Keys(i - 1) = Curves(i).Key
    Next i
    PostCmsRequest Batch, "getMarketData", BuildRosettaGet(Curves, Tag, DateStr), TimeoutMs, Keys
End Sub

Private Sub SendSetRequest(ByVal Batch As cCmsBatch, ByVal Curve As cCmsCurve, _
                           ByVal Quotes12 As Variant, ByVal RecoveryRate As Double, _
                           ByVal TimeoutMs As Long)
    Dim Keys(0 To 0) As Variant
    Keys(0) = Curve.Key
    PostCmsRequest Batch, "setMarketData", BuildRosettaSet(Curve, Quotes12, RecoveryRate), _
                   TimeoutMs, Keys
End Sub

' =============================================================================
'  RESPONSE HANDLING
' =============================================================================

' Unwrap the HTTP body down to the Rosetta XML, whatever the transport:
'   { "responseHeader": ..., "response": "<Rosetta.../>" }   (REST)
'   <SOAP-ENV:Envelope>...<...Return>escaped rosetta</...>   (SOAP)
'   <Rosetta .../>                                           (already raw)
Public Function ExtractRosetta(ByVal HttpBody As String) As String
    Dim t As String
    Dim doc As Object
    Dim nd As Object
    Dim code As String

    t = Trim$(HttpBody)
    If Len(t) = 0 Then Err.Raise vbObjectError + 610, "modCms.ExtractRosetta", "Empty response body"

    If Left$(t, 1) = "{" Then
        ' REST JSON wrapper
        code = JsonExtractString(t, "code")
        If Len(code) > 0 And UCase$(code) <> "SUCCESS" And UCase$(code) <> "WARNING" Then
            Err.Raise vbObjectError + 611, "modCms.ExtractRosetta", _
                "CMS REST " & code & ": " & JsonExtractString(t, "description")
        End If
        ExtractRosetta = JsonExtractString(t, "response")
        If Len(ExtractRosetta) = 0 Then
            Err.Raise vbObjectError + 612, "modCms.ExtractRosetta", "REST response carries no 'response' payload"
        End If
        Exit Function
    End If

    If InStr(1, t, "Envelope", vbTextCompare) > 0 And InStr(1, t, "<Rosetta", vbTextCompare) = 0 Then
        ' SOAP: the return payload is the (entity-escaped) text under <Body>
        Set doc = LoadXmlDoc(t)
        Set nd = doc.selectSingleNode("//*[local-name()='faultstring']")
        If Not nd Is Nothing Then
            Err.Raise vbObjectError + 613, "modCms.ExtractRosetta", "SOAP fault: " & nd.Text
        End If
        Set nd = doc.selectSingleNode("//*[local-name()='Body']")
        If nd Is Nothing Then
            Err.Raise vbObjectError + 614, "modCms.ExtractRosetta", "No SOAP Body in response"
        End If
        ExtractRosetta = nd.Text   ' MSXML unescapes the entities for us
        Exit Function
    End If

    ExtractRosetta = t
End Function

' Parse a Rosetta GET/SET response into the matching cCmsCurve objects.
'   CurvesByKey : dictionary fourTuple -> cCmsCurve (the batch's curve map)
'   CurveKeys   : keys carried by THIS request, in request order (positional
'                 fallback when a marketSet has no label)
Public Sub ParseRosettaResponse(ByVal RosettaXml As String, ByVal IsSetResponse As Boolean, _
                                ByVal CurvesByKey As Object, ByVal CurveKeys As Variant)
    Dim doc As Object
    Dim ndStatus As Object
    Dim msNodes As Object
    Dim ms As Object
    Dim ndLabel As Object
    Dim errNodes As Object
    Dim errs() As String
    Dim overallOk As Boolean
    Dim i As Long
    Dim Key As String
    Dim cv As cCmsCurve
    Dim k As Variant

    Set doc = LoadXmlDoc(RosettaXml)

    Set ndStatus = doc.selectSingleNode("//systemMessages/statusMessage/statusCode")
    If ndStatus Is Nothing Then
        Err.Raise vbObjectError + 615, "modCms.ParseRosettaResponse", "No statusCode in CMS response"
    End If
    overallOk = (StrComp(ndStatus.Text, "OK") = 0)

    ' Batch-level error messages (curve-level errors arrive here, keyed by
    ' fourTuple inside the text - the legacy ProcessBulkErrorMessages contract)
    Set errNodes = doc.selectNodes("//systemErrors/errorMessage/errorMessageContent")
    If errNodes.Length > 0 Then
        ReDim errs(0 To errNodes.Length - 1)
        For i = 0 To errNodes.Length - 1
            errs(i) = Replace(errNodes(i).Text, STANDARD_JAVA_ERROR, "")
        Next i
    Else
        ReDim errs(0 To 0)  ' single empty slot keeps loops simple
        errs(0) = ""
    End If

    ' Walk the marketSets, matching by label first, position second
    Set msNodes = doc.selectNodes("/Rosetta/market/marketData/marketSet")
    For i = 0 To msNodes.Length - 1
        Set ms = msNodes(i)
        Key = ""
        Set ndLabel = ms.selectSingleNode("label")
        If ndLabel Is Nothing Then Set ndLabel = ms.selectSingleNode("point/label")
        If Not ndLabel Is Nothing Then Key = UCase$(Trim$(ndLabel.Text))
        If Len(Key) = 0 Or Not CurvesByKey.Exists(Key) Then
            If i <= UBound(CurveKeys) - LBound(CurveKeys) Then Key = CStr(CurveKeys(LBound(CurveKeys) + i))
        End If
        If Len(Key) > 0 And CurvesByKey.Exists(Key) Then
            ParseMarketSetInto CurvesByKey(Key), ms, IsSetResponse
        End If
    Next i

    ' Per-curve status resolution
    For Each k In CurveKeys
        If CurvesByKey.Exists(CStr(k)) Then
            Set cv = CurvesByKey(CStr(k))
            ResolveCurveStatus cv, overallOk, errs, IsSetResponse
        End If
    Next k
End Sub

Private Sub ParseMarketSetInto(ByVal Curve As cCmsCurve, ByVal MarketSetNode As Object, _
                               ByVal IsSetResponse As Boolean)
    Dim nvp As Object
    Dim pt As Object
    Dim child As Object
    Dim ndVal As Object
    Dim nd As Object
    Dim nodeName As String
    Dim idx As Long
    Dim v As Double
    Dim spreads(0 To 16) As Variant
    Dim upfronts(0 To 16) As Variant
    Dim i As Long

    ' Scalars from genericKeys (relative to THIS marketSet)
    For Each nvp In MarketSetNode.selectNodes("genericKeys/nameValuePair")
        Select Case LCase$(GetAttr(nvp, "name"))
            Case "owner":              Curve.Owner = nvp.Text
            Case "marked":             Curve.LastMarkedOn = ConvertToDate(nvp.Text)
            Case "lastmarkedby":       Curve.LastMarkedBy = nvp.Text
            Case "isparent":           Curve.IsParent = nvp.Text
            Case "latestcurvemarkedtype":
                Select Case UCase$(nvp.Text)
                    Case "QUOTED": Curve.CurveType = "QUOTED_SPREADS"
                    Case "SPREAD": Curve.CurveType = "SPREADS"
                    Case Else:     Curve.CurveType = nvp.Text
                End Select
        End Select
    Next nvp

    ' Single pass over the points (no per-tenor XPath - O(points) not O(17 XPaths))
    For Each pt In MarketSetNode.selectNodes("point")
        Set ndVal = pt.selectSingleNode("value")
        If ndVal Is Nothing Then Set ndVal = pt.LastChild
        If Not ndVal Is Nothing Then
            v = Val(ndVal.Text)
            For Each child In pt.childNodes
                nodeName = child.nodeName
                If nodeName = "creditSpread" Or nodeName = "creditUpfront" Then
                    idx = -1
                    Set nd = child.selectSingleNode("periodMultiplier")
                    If Not nd Is Nothing Then
                        Dim per As Object
                        Set per = child.selectSingleNode("period")
                        If Not per Is Nothing Then idx = CMS_TenorIndex(nd.Text & per.Text)
                    End If
                    If idx >= 0 Then
                        If nodeName = "creditSpread" Then
                            ' GET responses carry decimal fractions -> bps
                            If IsSetResponse Then spreads(idx) = v Else spreads(idx) = v * 10000#
                        Else
                            ' GET responses carry decimal fractions -> %
                            If IsSetResponse Then upfronts(idx) = v Else upfronts(idx) = v * 100#
                        End If
                    End If
                ElseIf nodeName = "issuerRecovery" Then
                    Curve.Recovery = v
                End If
            Next child
        End If
    Next pt

    Curve.Spreads = spreads
    Curve.Upfronts = upfronts

    ' Select the quote per the curve's convention (legacy GetTermQuote rules,
    ' including the QUOTED -> creditSpread fallback)
    Dim quotes(0 To 16) As Variant
    For i = 0 To 16
        Select Case Curve.QuoteConvention
            Case "SPREADS", "QUOTED_SPREADS":
                quotes(i) = spreads(i)
            Case "UPFRONT", "EXTENDED":
                quotes(i) = upfronts(i)
            Case "RUNNING":
                If Not IsEmpty(spreads(i)) Then quotes(i) = spreads(i) Else quotes(i) = upfronts(i)
            Case Else:  ' QUOTED and anything unrecognized
                If Not IsEmpty(upfronts(i)) Then quotes(i) = upfronts(i) Else quotes(i) = spreads(i)
        End Select
    Next i
    Curve.Quotes = quotes
End Sub

Private Sub ResolveCurveStatus(ByVal Curve As cCmsCurve, ByVal OverallOk As Boolean, _
                               ByRef Errs() As String, ByVal IsSetResponse As Boolean)
    Dim i As Long
    Dim tooManyErrors As Boolean
    Dim hasData As Boolean

    ' Curve-level error: any batch error message mentioning this fourTuple
    For i = LBound(Errs) To UBound(Errs)
        If Len(Errs(i)) > 0 Then
            If InStr(1, UCase$(Errs(i)), UCase$(TOO_MANY_ERRORS_CMS_MESSAGE), vbTextCompare) > 0 Then
                tooManyErrors = True
            ElseIf InStr(1, UCase$(Errs(i)), Curve.FourTuple, vbTextCompare) > 0 Then
                Curve.Status = CMS_STATUS_FAILED
                Curve.ErrorMsg = Errs(i)
                Exit Sub
            End If
        End If
    Next i

    For i = 0 To 16
        If Not IsEmpty(Curve.Quotes(i)) Then
            hasData = True
            Exit For
        End If
    Next i

    ' A failed SET is a failure regardless of any echoed points (one curve per
    ' SET request). A failed bulk GET still yields per-curve data when present.
    If Not OverallOk And (IsSetResponse Or Not hasData) Then
        Curve.Status = CMS_STATUS_FAILED
        If Len(Errs(LBound(Errs))) > 0 Then
            Curve.ErrorMsg = Errs(LBound(Errs))
        Else
            Curve.ErrorMsg = "Unknown error occurred while retrieving spreads from CMS"
        End If
    ElseIf Not IsSetResponse And Not hasData And IsEmpty(Curve.Recovery) Then
        ' Legacy bulk rule: OK status but nothing for this curve -> failed
        Curve.Status = CMS_STATUS_FAILED
        If tooManyErrors Then
            Curve.ErrorMsg = "Too many errors; skipping current Curve."
        Else
            Curve.ErrorMsg = "No data returned for curve"
        End If
    Else
        Curve.Status = CMS_STATUS_OK
        Curve.ErrorMsg = ""
    End If
End Sub

' =============================================================================
'  INPUT NORMALIZATION
' =============================================================================

' CurveSpec -> Collection of NEW cCmsCurve registered on the batch.
' Accepts, per row:
'   - 5+ columns : full tuple (Ticker,Ccy,DebtClass,Product,QuoteConvention);
'                  each identity is registered for later ticker-only use
'   - 1 column   : ticker (must be registered - raises otherwise, before
'                  anything is sent) or full fourTuple key
'   - a bare String / single cell counts as a 1-column ticker spec
' RowKeys mirrors the input rows ("" for blank rows, the shared key for
' duplicates). NOTE: ticker LISTS must be vertical (1 column).
Private Function NormalizeFiveTuples(ByVal CurveSpec As Variant, ByVal Batch As cCmsBatch, _
                                     ByRef RowKeys As Variant) As Collection
    Dim arr As Variant
    Dim out As New Collection
    Dim cv As cCmsCurve
    Dim stored As cCmsCurve
    Dim r As Long, c0 As Long
    Dim n As Long, nCols As Long
    Dim Keys() As Variant
    Dim Ticker As String
    Dim s As String, Key As String

    arr = ToArray2D(CurveSpec)
    n = UBound(arr, 1) - LBound(arr, 1) + 1
    nCols = UBound(arr, 2) - LBound(arr, 2) + 1
    c0 = LBound(arr, 2)
    ReDim Keys(0 To n - 1)

    If nCols >= 5 Then
        ' Full five-tuple rows
        For r = 0 To n - 1
            Ticker = Trim$(NzStr(arr(LBound(arr, 1) + r, c0)))
            If Len(Ticker) = 0 Or Ticker = "0" Then
                Keys(r) = ""
            Else
                Set cv = New cCmsCurve
                cv.Init Ticker, _
                        NzStr(arr(LBound(arr, 1) + r, c0 + 1)), _
                        NzStr(arr(LBound(arr, 1) + r, c0 + 2)), _
                        NzStr(arr(LBound(arr, 1) + r, c0 + 3)), _
                        NzStr(arr(LBound(arr, 1) + r, c0 + 4))
                RegisterIdentity cv
                Keys(r) = cv.Key
                If Batch.AddCurve(cv) Then out.Add cv
                ' duplicates: row keeps the key, request sent once
            End If
        Next r

    ElseIf nCols = 1 Then
        ' Ticker / fourTuple-key rows resolved against the registry
        For r = 0 To n - 1
            s = Trim$(NzStr(arr(LBound(arr, 1) + r, c0)))
            If Len(s) = 0 Or s = "0" Then
                Keys(r) = ""
            Else
                Key = CMS_ResolveKey(s)   ' raises for unregistered tickers
                Set cv = New cCmsCurve
                Set stored = StoreGet(Key)
                If stored Is Nothing Then
                    ParseKeyInto Key, cv  ' raw key that was never stored
                Else
                    cv.Init stored.Ticker, stored.Ccy, stored.DebtClass, _
                            stored.Product, stored.QuoteConvention
                    cv.LabelOverride = stored.LabelOverride
                End If
                Keys(r) = cv.Key
                If Batch.AddCurve(cv) Then out.Add cv
            End If
        Next r

    Else
        Err.Raise vbObjectError + 620, "modCms", _
            "CurveSpec must have 5 columns (Ticker,Ccy,DebtClass,Product,QuoteConvention) " & _
            "or 1 column (registered tickers / fourTuple keys); got " & nCols & " columns"
    End If

    RowKeys = Keys
    Set NormalizeFiveTuples = out
End Function

Private Sub ParseKeyInto(ByVal Key As String, ByVal Curve As cCmsCurve)
    Dim parts() As String
    parts = Split(Key, ".")
    If UBound(parts) >= 3 Then
        Curve.Init parts(0), parts(1), parts(2), parts(3)
    Else
        Curve.Init Key, "", "", ""
    End If
End Sub

' Row r of a quotes array (12 or 17 wide) -> Variant(0..11) in SET tenor order
Private Function ExtractQuotes12(ByVal q As Variant, ByVal r As Long) As Variant
    Dim out(0 To 11) As Variant
    Dim nCols As Long
    Dim c0 As Long
    Dim idx As Variant
    Dim i As Long

    nCols = UBound(q, 2) - LBound(q, 2) + 1
    c0 = LBound(q, 2)

    If nCols >= 17 Then
        idx = CMS_SetTenorIndices()
        For i = 0 To 11
            If IsNumeric(q(r, c0 + idx(i))) And Not IsEmpty(q(r, c0 + idx(i))) Then out(i) = q(r, c0 + idx(i))
        Next i
    ElseIf nCols >= 12 Then
        For i = 0 To 11
            If IsNumeric(q(r, c0 + i)) And Not IsEmpty(q(r, c0 + i)) Then out(i) = q(r, c0 + i)
        Next i
    Else
        Err.Raise vbObjectError + 621, "modCms", _
            "Quotes must have 12 (SET grid) or 17 (full grid) columns, got " & nCols
    End If
    ExtractQuotes12 = out
End Function

' 12 SET-tenor quotes -> sparse 17-grid array
Private Function Expand12To17(ByVal Quotes12 As Variant) As Variant
    Dim out(0 To 16) As Variant
    Dim idx As Variant
    Dim i As Long
    idx = CMS_SetTenorIndices()
    For i = 0 To 11
        out(idx(i)) = Quotes12(i)
    Next i
    Expand12To17 = out
End Function

' 17-grid quotes -> 12 SET-tenor quotes
Private Function Project17To12(ByVal Quotes17 As Variant) As Variant
    Dim out(0 To 11) As Variant
    Dim idx As Variant
    Dim i As Long
    idx = CMS_SetTenorIndices()
    For i = 0 To 11
        out(i) = Quotes17(idx(i))
    Next i
    Project17To12 = out
End Function

' Anything (Range / scalar / 1D / 2D, any base) -> 2D Variant array
Private Function ToArray2D(ByVal v As Variant) As Variant
    Dim out() As Variant
    Dim i As Long

    If IsObject(v) Then v = v.Value2   ' Range -> value array (or scalar)

    If Not IsArray(v) Then
        ReDim out(0 To 0, 0 To 0)
        out(0, 0) = v
        ToArray2D = out
        Exit Function
    End If

    On Error Resume Next
    Dim ub2 As Long: ub2 = -2147483647
    ub2 = UBound(v, 2)
    On Error GoTo 0
    If ub2 <> -2147483647 Then
        ToArray2D = v
    Else
        ' 1D -> single row
        ReDim out(0 To 0, 0 To UBound(v) - LBound(v))
        For i = 0 To UBound(v) - LBound(v)
            out(0, i) = v(LBound(v) + i)
        Next i
        ToArray2D = out
    End If
End Function

' Any shape of ticker list (scalar, 1D array, row, column) -> N x 1 2D array
Private Function ToTickerColumn(ByVal v As Variant) As Variant
    Dim arr As Variant
    Dim out() As Variant
    Dim i As Long, n As Long
    arr = ToArray2D(v)
    If UBound(arr, 2) - LBound(arr, 2) + 1 = 1 Then
        ToTickerColumn = arr
    ElseIf UBound(arr, 1) - LBound(arr, 1) + 1 = 1 Then
        ' single row -> transpose
        n = UBound(arr, 2) - LBound(arr, 2) + 1
        ReDim out(0 To n - 1, 0 To 0)
        For i = 0 To n - 1
            out(i, 0) = arr(LBound(arr, 1), LBound(arr, 2) + i)
        Next i
        ToTickerColumn = out
    Else
        Err.Raise vbObjectError + 622, "modCms.ToTickerColumn", _
            "Ticker list must be a single row or a single column"
    End If
End Function

' A quotes row given as Range/1D/2D -> 1-row 2D array (transposing a column)
Private Function RowToArray2D(ByVal v As Variant) As Variant
    Dim arr As Variant
    Dim out() As Variant
    Dim i As Long
    arr = ToArray2D(v)
    If UBound(arr, 1) - LBound(arr, 1) > UBound(arr, 2) - LBound(arr, 2) Then
        ' column vector -> transpose to a row
        ReDim out(0 To 0, 0 To UBound(arr, 1) - LBound(arr, 1))
        For i = 0 To UBound(arr, 1) - LBound(arr, 1)
            out(0, i) = arr(LBound(arr, 1) + i, LBound(arr, 2))
        Next i
        RowToArray2D = out
    Else
        RowToArray2D = arr
    End If
End Function

' =============================================================================
'  LOW-LEVEL HELPERS
' =============================================================================

Private Function LoadXmlDoc(ByVal Xml As String) As Object
    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    If Not doc.LoadXML(Xml) Then
        Err.Raise vbObjectError + 630, "modCms.LoadXmlDoc", _
            "Failed to parse XML: " & doc.parseError.reason
    End If
    Set LoadXmlDoc = doc
End Function

Private Function GetAttr(ByVal Node As Object, ByVal AttrName As String) As String
    Dim a As Object
    Set a = Node.Attributes.getNamedItem(AttrName)
    If Not a Is Nothing Then GetAttr = a.Text
End Function

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    XmlEsc = s
End Function

Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEsc = s
End Function

' Extract a top-level-ish string value from a JSON blob: "key":"value"
' (handles escaped characters inside the value)
Private Function JsonExtractString(ByVal Json As String, ByVal Key As String) As String
    Dim p As Long, q As Long
    Dim sb As String
    Dim ch As String, nxt As String

    p = InStr(1, Json, """" & Key & """", vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(Key) + 2, Json, ":", vbBinaryCompare)
    If p = 0 Then Exit Function
    ' skip whitespace
    p = p + 1
    Do While p <= Len(Json) And (Mid$(Json, p, 1) = " " Or Mid$(Json, p, 1) = vbTab Or _
                                 Mid$(Json, p, 1) = vbCr Or Mid$(Json, p, 1) = vbLf)
        p = p + 1
    Loop
    If Mid$(Json, p, 1) <> """" Then Exit Function   ' not a string value
    p = p + 1

    Do While p <= Len(Json)
        ch = Mid$(Json, p, 1)
        If ch = "\" And p < Len(Json) Then
            nxt = Mid$(Json, p + 1, 1)
            Select Case nxt
                Case """": sb = sb & """"
                Case "\":  sb = sb & "\"
                Case "/":  sb = sb & "/"
                Case "n":  sb = sb & vbLf
                Case "r":  sb = sb & vbCr
                Case "t":  sb = sb & vbTab
                Case "u":
                    If p + 5 <= Len(Json) Then
                        sb = sb & ChrW$(CLng("&H" & Mid$(Json, p + 2, 4)))
                        p = p + 4
                    End If
                Case Else: sb = sb & nxt
            End Select
            p = p + 2
        ElseIf ch = """" Then
            Exit Do
        Else
            sb = sb & ch
            p = p + 1
        End If
    Loop
    JsonExtractString = sb
End Function

' Locale-safe number -> XML text ("." decimal separator always)
Private Function NumToXml(ByVal v As Double) As String
    Dim s As String
    s = Trim$(Str$(v))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    NumToXml = s
End Function

' QuoteDate -> yyyymmdd; 0 / missing -> today (legacy behavior)
Private Function QuoteDateString(ByVal QuoteDate As Date) As String
    If QuoteDate = 0 Then QuoteDate = Now()
    QuoteDateString = Format$(QuoteDate, "yyyymmdd")
End Function

' CMS hands back two timestamp formats; degrade to raw text if both fail
' ("2008-12-12 17:04:41.137" and "Mon Dec 15 12:22:10 EST 2008")
Private Function ConvertToDate(ByVal DateString As String) As Variant
    Dim tmpDate As String
    Dim tmpTime As String
    On Error GoTo Fallback
    If InStr(1, DateString, "-") > 0 Then
        tmpDate = Replace(Left$(DateString, 10), "-", "/")
        tmpTime = Mid$(DateString, 12, 8)
    Else
        tmpDate = Mid$(DateString, 9, 2) & "-" & Mid$(DateString, 5, 3) & "-" & Mid$(DateString, 25, 4)
        tmpTime = Mid$(DateString, 12, 8)
    End If
    ConvertToDate = CDate(tmpDate & " " & tmpTime)
    Exit Function
Fallback:
    ConvertToDate = DateString
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function
