
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

Private Sub Class_Initialize()
    ReDim emptyQ(0 To 16) As Variant
    Quotes = emptyQ
    Spreads = emptyQ
    Upfronts = emptyQ
    Status = ""
    ErrorMsg = ""
End Sub

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
