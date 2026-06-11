
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
' Bulk async GET. FiveTuples: a Range or 2D array with columns
' Ticker, Ccy, DebtClass, Product, QuoteConvention (e.g. D9:H450 or a subset).
' Blank-ticker rows are skipped; duplicates collapse onto one request but
' still appear row-for-row in batch.ToArray().
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
    Set CMS_GetCurvesAsync = batch
End Function

' ---------------------------------------------------------------------------
' Bulk async SET ("CMS_SetBulkCurveQuoteData" - the call that never existed).
' One HTTP request per curve, all in flight at once.
'
'   FiveTuples : as above (N rows)
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
    Set CMS_SetCurvesAsync = batch
End Function

' ---------------------------------------------------------------------------
' The perturbation workflow: re-mark curves off a new 5Y level, computing the
' rest of the curve in memory from the last GET (no sheet round trip).
'
'   KeysAndNew5y : 2D array/Range, N rows x 2: (fourTuple key, new 5Y level)
'                  Keys are TICKER.CCY.DEBTCLASS.PRODUCT - cCmsCurve.Key.
'                  Curves must already be in the store (i.e. GET them first).
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

    arr = ToArray2D(KeysAndNew5y)
    n = UBound(arr, 1) - LBound(arr, 1) + 1
    ReDim rowKeys(0 To n - 1)

    For i = 0 To n - 1
        Key = UCase$(Trim$(CStr(arr(LBound(arr, 1) + i, LBound(arr, 2)))))
        rowKeys(i) = Key
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
    Set CMS_SetCurves5yAsync = batch
End Function

' ---------------------------------------------------------------------------
' Async refresh of a single curve (returns the batch; result via callbacks
' or batch.Curve(key) after batch.IsDone).
' ---------------------------------------------------------------------------
Public Function CMS_GetCurveAsync(ByVal Ticker As String, ByVal Ccy As String, _
                                  ByVal DebtClass As String, ByVal Product As String, _
                                  Optional ByVal QuoteConvention As String = "SPREADS", _
                                  Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                  Optional ByVal QuoteDate As Date, _
                                  Optional ByVal OnCurve As String = "", _
                                  Optional ByVal OnAllDone As String = "") As cCmsBatch
    Dim tuples(0 To 0, 0 To 4) As Variant
    tuples(0, 0) = Ticker: tuples(0, 1) = Ccy: tuples(0, 2) = DebtClass
    tuples(0, 3) = Product: tuples(0, 4) = QuoteConvention
    Set CMS_GetCurveAsync = CMS_GetCurvesAsync(tuples, Tag, QuoteDate, OnCurve, OnAllDone)
End Function

' =============================================================================
'  SYNC COMPATIBILITY WRAPPERS (legacy signatures, 17-standard layout)
'  These block the calling VBA only - the bridge still does the I/O async.
' =============================================================================

' One curve -> 1D Variant(0..24): 17 tenor quotes + 8 scalars.
Public Function CMS_GetCurveQuoteData(ByVal Ticker As String, ByVal Ccy As String, _
                                      ByVal DebtClass As String, _
                                      Optional ByVal Product As String = "DERIV", _
                                      Optional ByVal Tag As String = CMS_TAG_LIVE, _
                                      Optional ByVal QuoteDate As Date, _
                                      Optional ByVal QuoteConvention As String = "SPREADS", _
                                      Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As Variant
    Dim batch As cCmsBatch
    Dim cv As cCmsCurve
    Set batch = CMS_GetCurveAsync(Ticker, Ccy, DebtClass, Product, QuoteConvention, Tag, QuoteDate)
    If Not batch.WaitAll(TimeoutMs) Then
        Err.Raise vbObjectError + 603, "modCms.CMS_GetCurveQuoteData", "Timed out waiting for CMS"
    End If
    Set cv = batch.Curve(CMS_CompileFourTuple(Ticker, Ccy, DebtClass, Product))
    If cv Is Nothing Then
        Err.Raise vbObjectError + 604, "modCms.CMS_GetCurveQuoteData", "Invalid curve identifiers"
    End If
    CMS_GetCurveQuoteData = cv.ToRow()
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

' One curve SET -> echoed row (legacy CMS_SetCurveQuoteData shape).
Public Function CMS_SetCurveQuoteData(ByVal Ticker As String, ByVal Ccy As String, _
                                      ByVal DebtClass As String, ByVal Product As String, _
                                      ByVal TermQuotes As Variant, _
                                      Optional ByVal RecoveryRate As Double = 0, _
                                      Optional ByVal QuoteConvention As String = "SPREADS", _
                                      Optional ByVal TimeoutMs As Long = DEFAULT_TIMEOUT_MS) As Variant
    Dim tuples(0 To 0, 0 To 4) As Variant
    Dim quotes As Variant
    Dim recs(0 To 0, 0 To 0) As Variant
    Dim batch As cCmsBatch

    tuples(0, 0) = Ticker: tuples(0, 1) = Ccy: tuples(0, 2) = DebtClass
    tuples(0, 3) = Product: tuples(0, 4) = QuoteConvention
    quotes = RowToArray2D(TermQuotes)
    recs(0, 0) = RecoveryRate

    Set batch = CMS_SetCurvesAsync(tuples, quotes, recs)
    If Not batch.WaitAll(TimeoutMs) Then
        Err.Raise vbObjectError + 603, "modCms.CMS_SetCurveQuoteData", "Timed out waiting for CMS"
    End If
    CMS_SetCurveQuoteData = batch.Curve(CMS_CompileFourTuple(Ticker, Ccy, DebtClass, Product)).ToRow()
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
    Set CMS_Store()(Curve.Key) = Curve
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
    ReDim out(0 To UBound(useKeys) - LBound(useKeys) + base, 0 To 25)

    If IncludeHeader Then
        out(0, 0) = "Curve"
        header = CMS_GetCurveQuoteDataHeader17()
        For j = 0 To 24
            out(0, j + 1) = header(j)
        Next j
    End If

    For i = 0 To UBound(useKeys) - LBound(useKeys)
        k = CStr(useKeys(LBound(useKeys) + i))
        out(i + base, 0) = k
        If CMS_Store().Exists(k) Then
            rowData = gStore(k).ToRow()
            For j = 0 To 24
                out(i + base, j + 1) = rowData(j)
            Next j
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

' FiveTuples (Range / 2D array, cols Ticker,Ccy,DebtClass,Product,QuoteConvention)
' -> Collection of NEW cCmsCurve registered on the batch. RowKeys mirrors the
' input rows ("" for blank/short rows, the shared key for duplicates).
Private Function NormalizeFiveTuples(ByVal FiveTuples As Variant, ByVal Batch As cCmsBatch, _
                                     ByRef RowKeys As Variant) As Collection
    Dim arr As Variant
    Dim out As New Collection
    Dim cv As cCmsCurve
    Dim r As Long, c0 As Long
    Dim n As Long
    Dim Keys() As Variant
    Dim Ticker As String

    arr = ToArray2D(FiveTuples)
    If UBound(arr, 2) - LBound(arr, 2) + 1 < 5 Then
        Err.Raise vbObjectError + 620, "modCms", _
            "FiveTuples must have 5 columns: Ticker, Ccy, DebtClass, Product, QuoteConvention"
    End If

    n = UBound(arr, 1) - LBound(arr, 1) + 1
    c0 = LBound(arr, 2)
    ReDim Keys(0 To n - 1)

    For r = 0 To n - 1
        Ticker = Trim$(CStr(NzStr(arr(LBound(arr, 1) + r, c0))))
        If Len(Ticker) = 0 Or Ticker = "0" Then
            Keys(r) = ""
        Else
            Set cv = New cCmsCurve
            cv.Init Ticker, _
                    CStr(NzStr(arr(LBound(arr, 1) + r, c0 + 1))), _
                    CStr(NzStr(arr(LBound(arr, 1) + r, c0 + 2))), _
                    CStr(NzStr(arr(LBound(arr, 1) + r, c0 + 3))), _
                    CStr(NzStr(arr(LBound(arr, 1) + r, c0 + 4)))
            Keys(r) = cv.Key
            If Batch.AddCurve(cv) Then out.Add cv
            ' duplicates: row keeps the key, request sent once
        End If
    Next r

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
