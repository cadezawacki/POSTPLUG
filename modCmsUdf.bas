
Attribute VB_Name = "modCmsUdf"
Option Explicit

' =============================================================================
'  modCmsUdf - reactive worksheet functions over the CMS curve store.
'
'      =CURVE("AEP")                  ' 5Y quote (default field)
'      =CURVE("AEP", "10Y")           ' any tenor
'      =CURVE("AEP", "Recovery")      ' any scalar: Status, Owner, CurveType,
'                                     '   LastMarkedOn/By, IsParent, Error,
'                                     '   Pending5y, Ccy, Product, Key, ...
'      =CURVEDIFF("AEP")              ' 5s10s steepness (10Y - 5Y)
'      =CURVERATIO("AEP","5Y","10Y")  ' 10Y / 5Y
'      =TICKERDIFF("AEP","XYZ","5Y")  ' AEP - XYZ at 5Y
'      =TICKERRATIO("AEP","XYZ")      ' AEP / XYZ at 5Y
'
'  Reactivity: each call subscribes its own cell (Application.Caller) to the
'  curve(s) it reads, so the cell recalcs automatically whenever the store
'  changes - a GET lands, a SET confirms or fails, a pending level is staged
'  or cleared. The functions are NOT volatile; only affected cells recalc.
'
'  Auto-fetch: when a curve is REGISTERED but has no quotes yet, the key is
'  queued and fetched asynchronously right after calculation ends (see
'  modCms.CMS_FlushAutoFetch); the cell shows #N/A until the data lands, then
'  updates itself. Pass FALSE as the last argument to disable. Unregistered
'  tickers show #N/A and are never fetched (no identity to fetch with) -
'  register_curve / register_curve_by_range first. Curves whose last fetch
'  FAILED are not auto-retried; =CURVE(t,"Error") shows why.
' =============================================================================

Public Function CURVE(ByVal TickerOrKey As String, _
                      Optional ByVal Field As String = "5Y", _
                      Optional ByVal AutoFetch As Boolean = True) As Variant
    Dim Key As String
    Dim cv As cCmsCurve
    Dim v As Variant

    On Error GoTo NA
    SubscribeCaller TickerOrKey

    On Error Resume Next
    Key = modCms.CMS_ResolveKey(TickerOrKey)
    On Error GoTo NA
    If Len(Key) = 0 Then GoTo NA            ' unregistered ticker

    Set cv = modCms.StoreGet(Key)
    If cv Is Nothing Then GoTo NA           ' raw key never stored

    v = cv.Field(Field)
    If IsError(v) Then
        CURVE = v                           ' #NAME? for unknown fields
    ElseIf IsEmpty(v) Then
        If AutoFetch And Not cv.HasQuotes() And cv.Status <> modCms.CMS_STATUS_FAILED Then
            modCms.QueueAutoFetch Key
        End If
        CURVE = CVErr(xlErrNA)              ' pending / no mark at this tenor
    Else
        CURVE = v
    End If
    Exit Function
NA:
    CURVE = CVErr(xlErrNA)
End Function

' Slope between two tenors of one curve: TenorB - TenorA (default 5s10s).
Public Function CURVEDIFF(ByVal TickerOrKey As String, _
                          Optional ByVal TenorA As String = "5Y", _
                          Optional ByVal TenorB As String = "10Y", _
                          Optional ByVal AutoFetch As Boolean = True) As Variant
    Dim a As Variant, b As Variant
    a = CURVE(TickerOrKey, TenorA, AutoFetch)
    b = CURVE(TickerOrKey, TenorB, AutoFetch)
    If IsError(a) Then CURVEDIFF = a: Exit Function
    If IsError(b) Then CURVEDIFF = b: Exit Function
    CURVEDIFF = CDbl(b) - CDbl(a)
End Function

' Ratio between two tenors of one curve: TenorB / TenorA.
Public Function CURVERATIO(ByVal TickerOrKey As String, _
                           Optional ByVal TenorA As String = "5Y", _
                           Optional ByVal TenorB As String = "10Y", _
                           Optional ByVal AutoFetch As Boolean = True) As Variant
    Dim a As Variant, b As Variant
    a = CURVE(TickerOrKey, TenorA, AutoFetch)
    b = CURVE(TickerOrKey, TenorB, AutoFetch)
    If IsError(a) Then CURVERATIO = a: Exit Function
    If IsError(b) Then CURVERATIO = b: Exit Function
    If CDbl(a) = 0 Then CURVERATIO = CVErr(xlErrDiv0): Exit Function
    CURVERATIO = CDbl(b) / CDbl(a)
End Function

' Spread between two tickers at one tenor: TickerA - TickerB.
Public Function TICKERDIFF(ByVal TickerA As String, ByVal TickerB As String, _
                           Optional ByVal Tenor As String = "5Y", _
                           Optional ByVal AutoFetch As Boolean = True) As Variant
    Dim a As Variant, b As Variant
    a = CURVE(TickerA, Tenor, AutoFetch)
    b = CURVE(TickerB, Tenor, AutoFetch)
    If IsError(a) Then TICKERDIFF = a: Exit Function
    If IsError(b) Then TICKERDIFF = b: Exit Function
    TICKERDIFF = CDbl(a) - CDbl(b)
End Function

' Ratio between two tickers at one tenor: TickerA / TickerB.
Public Function TICKERRATIO(ByVal TickerA As String, ByVal TickerB As String, _
                            Optional ByVal Tenor As String = "5Y", _
                            Optional ByVal AutoFetch As Boolean = True) As Variant
    Dim a As Variant, b As Variant
    a = CURVE(TickerA, Tenor, AutoFetch)
    b = CURVE(TickerB, Tenor, AutoFetch)
    If IsError(a) Then TICKERRATIO = a: Exit Function
    If IsError(b) Then TICKERRATIO = b: Exit Function
    If CDbl(b) = 0 Then TICKERRATIO = CVErr(xlErrDiv0): Exit Function
    TICKERRATIO = CDbl(a) / CDbl(b)
End Function

' Subscribe the calling cell to a curve. No-op when invoked from VBA rather
' than a worksheet cell.
Private Sub SubscribeCaller(ByVal TickerOrKey As String)
    On Error Resume Next
    If TypeName(Application.Caller) = "Range" Then
        modCms.CMS_SubscribeRange TickerOrKey, Application.Caller
    End If
    On Error GoTo 0
End Sub
