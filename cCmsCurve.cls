
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCmsCurve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
'  cCmsCurve - one credit curve: identity (5-tuple), quotes on the 17-standard
'  tenor grid, and the scalar attributes CMS returns alongside them.
'
'  Units follow the legacy add-in exactly:
'    GET responses : creditSpread x10000 (bps), creditUpfront x100 (%)
'    SET responses : echoed verbatim in request units (no scaling)
'  so Quotes() from a GET are bps for SPREADS-convention curves, and values
'  you pass to a SET should be in those same units.
' =============================================================================

' --- Identity (the "FiveTuple") ---
Public Ticker As String
Public Ccy As String
Public DebtClass As String
Public Product As String
Public QuoteConvention As String   ' SPREADS / QUOTED_SPREADS / UPFRONT / QUOTED / RUNNING

' Optional explicit <label> override. When empty the compiled 4-tuple
' (TICKER.CCY.DEBTCLASS.PRODUCT) is used - they should normally be identical.
Public LabelOverride As String

' --- Quotes on the 17-standard grid (0-based 0..16, Empty = no mark) ---
Public Quotes As Variant     ' selected per QuoteConvention
Public Spreads As Variant    ' raw creditSpread points
Public Upfronts As Variant   ' raw creditUpfront points

' --- Historical closes (NYOISCLOSE), date-keyed, cached once per day ---
' One 17-grid per business date so T-1, T-2, ... T-N coexist. The legacy
' Prev* members (below) remain as T-1 conveniences over this dictionary.
Private mHist As Object      ' "yyyymmdd" -> 17-grid Variant array
Private mHistAt As Object    ' "yyyymmdd" -> Date the close was pulled

' Full 17-point curve queued for a SET (computed client-side); merged with the
' server echo into Quotes/store once the SET succeeds.
Public PendingQuotes As Variant

' Staged-but-not-sent 5Y amend (Empty = none). Set via CMS_StagePending5y,
' drained into an async SET by CMS_MarkPendingAsync.
Public Pending5y As Variant

' --- Scalars ---
Public Recovery As Variant
Public LastMarkedOn As Variant
Public LastMarkedBy As String
Public Owner As String
Public IsParent As String
Public CurveType As String

' --- Outcome ---
Public Status As String      ' "OK" / "FAILED" / "" (pending)
Public ErrorMsg As String

' --- Diagnostics ---
Public RequestId As String
Public HttpStatus As Long
Public ElapsedMs As Long
Public FetchedAt As Date     ' last successful server round trip (0 = never)

Private Sub Class_Initialize()
    ReDim emptyQ(0 To 16) As Variant
    Quotes = emptyQ
    Spreads = emptyQ
    Upfronts = emptyQ
    Set mHist = CreateObject("Scripting.Dictionary")
    mHist.CompareMode = vbTextCompare
    Set mHistAt = CreateObject("Scripting.Dictionary")
    mHistAt.CompareMode = vbTextCompare
    Status = ""
    ErrorMsg = ""
End Sub

Private Function DK(ByVal AsOf As Date) As String
    DK = Format$(AsOf, "yyyymmdd")
End Function

Public Sub Init(ByVal pTicker As String, ByVal pCcy As String, ByVal pDebtClass As String, _
                ByVal pProduct As String, Optional ByVal pQuoteConvention As String = "SPREADS")
    Ticker = UCase$(Trim$(pTicker))
    Ccy = UCase$(Trim$(pCcy))
    DebtClass = UCase$(Trim$(pDebtClass))
    Product = UCase$(Trim$(pProduct))
    QuoteConvention = UCase$(Trim$(pQuoteConvention))
    If Len(QuoteConvention) = 0 Then QuoteConvention = "SPREADS"
End Sub

' TICKER.CCY.DEBTCLASS.PRODUCT - also the dictionary key everywhere.
Public Function FourTuple() As String
    FourTuple = modCms.CMS_CompileFourTuple(Ticker, Ccy, DebtClass, Product)
End Function

Public Function Key() As String
    Key = FourTuple()
End Function

' The <label> that goes on the wire.
Public Function Label() As String
    If Len(LabelOverride) > 0 Then
        Label = LabelOverride
    Else
        Label = FourTuple()
    End If
End Function

' Quote by tenor name ("5Y", "0M", ...). Empty if not marked.
Public Function Quote(ByVal Tenor As String) As Variant
    Dim idx As Long
    idx = modCms.CMS_TenorIndex(Tenor)
    If idx >= 0 Then Quote = Quotes(idx)
End Function

' --- Date-keyed historical close access ---

' Store a 17-grid close for a business date (called by modCms.StoreHistClose).
Public Sub HistSet(ByVal AsOf As Date, ByVal Quotes17 As Variant)
    mHist(DK(AsOf)) = Quotes17
    mHistAt(DK(AsOf)) = Now()
End Sub

Public Function HasHist(ByVal AsOf As Date) As Boolean
    HasHist = mHist.Exists(DK(AsOf))
End Function

' The 17-grid as of AsOf, or an empty 17-grid when not cached.
Public Function HistQuotesFor(ByVal AsOf As Date) As Variant
    If mHist.Exists(DK(AsOf)) Then
        HistQuotesFor = mHist(DK(AsOf))
    Else
        ReDim e(0 To 16) As Variant
        HistQuotesFor = e
    End If
End Function

' One tenor of the AsOf close. Empty if the date or tenor isn't cached.
Public Function HistQuote(ByVal Tenor As String, ByVal AsOf As Date) As Variant
    Dim idx As Long
    If Not mHist.Exists(DK(AsOf)) Then Exit Function
    idx = modCms.CMS_TenorIndex(Tenor)
    If idx >= 0 Then HistQuote = mHist(DK(AsOf))(idx)
End Function

Public Function HistFetchedAtFor(ByVal AsOf As Date) As Variant
    If mHistAt.Exists(DK(AsOf)) Then HistFetchedAtFor = mHistAt(DK(AsOf))
End Function

' Cached business dates (as Date), newest first - for carryover and diag.
Public Function HistDates() As Variant
    Dim k As Variant, out() As Variant, n As Long, i As Long, j As Long, tmp As Date
    If mHist.Count = 0 Then
        HistDates = Array()
        Exit Function
    End If
    ReDim out(0 To mHist.Count - 1)
    For Each k In mHist.Keys
        out(n) = DateSerial(CLng(Left$(CStr(k), 4)), CLng(Mid$(CStr(k), 5, 2)), CLng(Mid$(CStr(k), 7, 2)))
        n = n + 1
    Next k
    For i = 0 To UBound(out) - 1            ' newest first (small N)
        For j = i + 1 To UBound(out)
            If out(j) > out(i) Then tmp = out(i): out(i) = out(j): out(j) = tmp
        Next j
    Next i
    HistDates = out
End Function

' Copy every cached close from another curve (StoreUpsert carryover across a
' refresh GET, which delivers a brand-new object).
Public Sub CopyHistoryFrom(ByVal Other As cCmsCurve)
    Dim d As Variant
    Dim dates As Variant
    dates = Other.HistDates()
    For Each d In dates
        mHist(DK(CDate(d))) = Other.HistQuotesFor(CDate(d))
        mHistAt(DK(CDate(d))) = Other.HistFetchedAtFor(CDate(d))
    Next d
End Sub

Public Function HistCount() As Long
    HistCount = mHist.Count
End Function

' --- T-1 conveniences over the history dictionary (back-compat) ---

' Previous business-day (T-1) close quote by tenor. Empty if not cached.
Public Function PrevQuote(ByVal Tenor As String) As Variant
    PrevQuote = HistQuote(Tenor, modCms.CMS_PrevBizDay())
End Function

Public Property Get PrevQuotes() As Variant
    PrevQuotes = HistQuotesFor(modCms.CMS_PrevBizDay())
End Property

' T-1 business date if its close is cached, else 0 (legacy semantics).
Public Property Get PrevDate() As Date
    Dim d As Date
    d = modCms.CMS_PrevBizDay()
    If HasHist(d) Then PrevDate = d
End Property

Public Property Get PrevFetchedAt() As Variant
    PrevFetchedAt = HistFetchedAtFor(modCms.CMS_PrevBizDay())
End Property

Public Function FiveYear() As Variant
    FiveYear = Quotes(modCms.CMS_IDX_5Y)
End Function

' New full curve implied by a new 5Y mark (ratio front end, parallel long end).
Public Function ApplyNew5y(ByVal New5y As Double) As Variant
    ApplyNew5y = modCms.CMS_ApplyNew5y(Quotes, New5y)
End Function

Public Function IsOk() As Boolean
    IsOk = (Status = modCms.CMS_STATUS_OK)
End Function

Public Function HasQuotes() As Boolean
    Dim i As Long
    For i = 0 To 16
        If Not IsEmpty(Quotes(i)) Then
            HasQuotes = True
            Exit Function
        End If
    Next i
End Function

' Generic field access (sheet UDFs / dynamic readers): any tenor name plus
' every scalar and identity field.
'   Historical closes: "Prev<tenor>" = T-1 (e.g. "Prev5Y"); "Prev<n>_<tenor>"
'   = T-n (e.g. "Prev2_5Y" = T-2 5Y). "PrevDate"/"PrevFetchedAt" (and the
'   offset forms "Prev2_Date" etc.) expose that close's metadata.
' Unknown names return #NAME?.
Public Function Field(ByVal FieldName As String) As Variant
    Dim idx As Long
    Dim fn As String, rest As String, lead As String
    Dim off As Long
    Dim usp As Long
    Dim d As Date

    idx = modCms.CMS_TenorIndex(FieldName)
    If idx >= 0 Then
        Field = Quotes(idx)
        Exit Function
    End If

    fn = UCase$(Replace(Trim$(FieldName), " ", ""))

    ' Historical-close fields: PREV<tenor> (T-1) or PREV<n>_<tenor> (T-n),
    ' plus PREV[<n>_]DATE / FETCHEDAT.
    If Left$(fn, 4) = "PREV" Then
        rest = Mid$(fn, 5)
        off = 1
        usp = InStr(rest, "_")
        If usp > 0 Then
            lead = Left$(rest, usp - 1)
            If Len(lead) = 0 Then
                rest = Mid$(rest, usp + 1)            ' "PREV_5Y" -> T-1
            ElseIf IsAllDigits(lead) Then
                off = CLng(lead)                      ' "PREV2_5Y" -> T-2
                rest = Mid$(rest, usp + 1)
            End If
        End If
        d = modCms.CMS_BizDay(-off)
        idx = modCms.CMS_TenorIndex(rest)
        If idx >= 0 Then
            Field = HistQuotesFor(d)(idx)
            Exit Function
        End If
        Select Case rest
            Case "DATE":
                If HasHist(d) Then Field = d Else Field = Empty
                Exit Function
            Case "FETCHEDAT":
                Field = HistFetchedAtFor(d)
                Exit Function
        End Select
    End If

    Select Case fn
        Case "RECOVERY":                      Field = Recovery
        Case "LASTMARKEDON", "MARKED":        Field = LastMarkedOn
        Case "LASTMARKEDBY":                  Field = LastMarkedBy
        Case "OWNER":                         Field = Owner
        Case "ISPARENT":                      Field = IsParent
        Case "CURVETYPE":                     Field = CurveType
        Case "STATUS":                        Field = Status
        Case "ERROR", "ERRORMSG", "ERRORMESSAGE": Field = ErrorMsg
        Case "PENDING", "PENDING5Y":          Field = Pending5y
        Case "TICKER":                        Field = Ticker
        Case "CCY", "CURRENCY":               Field = Ccy
        Case "DEBTCLASS":                     Field = DebtClass
        Case "PRODUCT":                       Field = Product
        Case "QUOTECONVENTION", "CONVENTION": Field = QuoteConvention
        Case "KEY", "FOURTUPLE", "LABEL":     Field = Key()
        Case "ELAPSEDMS":                     Field = ElapsedMs
        Case "FETCHEDAT":
            ' Full datetime (Now() at delivery); format the cell to see the
            ' time, or use FetchedAtText for an unambiguous string.
            If FetchedAt = 0 Then Field = Empty Else Field = FetchedAt
        Case "FETCHEDATTEXT":
            If FetchedAt = 0 Then Field = "" Else Field = Format$(FetchedAt, "yyyy-mm-dd hh:nn:ss")
        Case Else:                            Field = CVErr(2029) ' #NAME?
    End Select
End Function

' One result row: 17 tenor quotes + the 8 scalar columns
' (Recovery, LastMarkedOn, LastMarkedBy, Owner, IsParent, CurveType, Status, Error).
' Same layout as the legacy CMS_GetCurveQuoteData with TENORS_17_STANDARD.
Public Function ToRow() As Variant
    Dim row(0 To 24) As Variant
    Dim i As Long
    For i = 0 To 16
        row(i) = Quotes(i)
    Next i
    row(17) = Recovery
    row(18) = LastMarkedOn
    row(19) = LastMarkedBy
    row(20) = Owner
    row(21) = IsParent
    row(22) = CurveType
    row(23) = Status
    row(24) = ErrorMsg
    ToRow = row
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        Select Case Asc(Mid$(s, i, 1))
            Case 48 To 57
            Case Else: Exit Function
        End Select
    Next i
    IsAllDigits = True
End Function
