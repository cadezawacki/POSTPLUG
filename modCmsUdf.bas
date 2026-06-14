
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
'  Historical close fields (date-keyed, NYOISCLOSE; T-1 auto on fresh GET,
'  deeper offsets fetched on demand when the cell evaluates):
'      =CURVE("AEP", "Prev5Y")        ' T-1 close, any tenor
'      =CURVE("AEP", "Prev2_5Y")      ' T-2 close ("Prev<n>_<tenor>")
'      =CURVE("AEP", "PrevDate")      ' which date the T-1 close is for
'      =CURVEDIFF("AEP","Prev5Y","5Y")  ' 1-day change at 5Y
'      =CURVEDIFF("AEP","Prev2_5Y","5Y")' 2-day change at 5Y
'
'  Reactivity: each call subscribes its own cell (Application.Caller) to the
'  curve(s) it reads. How a store change reaches the cell is governed by
'  modCms.CMS_SetNotifyMode: DIRTY (default) marks the cell so the user's
'  next recalc picks it up (instant in automatic calc mode); CALC forces an
'  immediate recalculation; OFF leaves cells alone. Not volatile.
'
'  Auto-fetch: a curve that is REGISTERED but has no quotes yet shows
'  "#Pending" (modCms.CMS_PENDING_TEXT), queues the key, and the async GET
'  launches within ~1s (watchdog) or when calculation ends. Pass FALSE as
'  the last argument to disable. Unregistered tickers show #N/A and are
'  never fetched (no identity) - register_curve first. FAILED curves show
'  #N/A and are not auto-retried; =CURVE(t,"Error") shows why.
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

    Dim off As Long
    v = cv.Field(Field)
    If IsError(v) Then
        CURVE = v                           ' #NAME? for unknown fields
    ElseIf IsEmpty(v) Then
        off = modCms.CMS_HistFieldOffset(Field)
        If off >= 1 Then
            ' Historical close (T-off) not cached yet: queue that offset
            ' on demand (toggle-respecting) and show pending until it lands;
            ' #N/A once the close is cached but unmarked at this tenor.
            If Not modCms.CMS_HistCached(Key, off) Then
                If AutoFetch Then modCms.QueueHistFetchOffset Key, off
                CURVE = modCms.CMS_PENDING_TEXT
            Else
                CURVE = CVErr(xlErrNA)
            End If
        ElseIf Not cv.HasQuotes() And cv.Status <> modCms.CMS_STATUS_FAILED Then
            ' Registered but never fetched: queue and show pending. The value
            ' appears at the next recalc after the fetch lands (or instantly
            ' under CMS_NOTIFY_CALC).
            If AutoFetch Then modCms.QueueAutoFetch Key
            CURVE = modCms.CMS_PENDING_TEXT
        Else
            CURVE = CVErr(xlErrNA)          ' fetched but no mark / FAILED
        End If
    Else
        CURVE = v
    End If
    Exit Function
NA:
    CURVE = CVErr(xlErrNA)
End Function

' Propagate errors AND the "#Pending" marker through derived functions.
Private Function PassThru(ByVal v As Variant) As Boolean
    PassThru = IsError(v)
    If Not PassThru Then PassThru = Not IsNumeric(v)
End Function

' Slope between two tenors of one curve: TenorB - TenorA (default 5s10s).
Public Function CURVEDIFF(ByVal TickerOrKey As String, _
                          Optional ByVal TenorA As String = "5Y", _
                          Optional ByVal TenorB As String = "10Y", _
                          Optional ByVal AutoFetch As Boolean = True) As Variant
    Dim a As Variant, b As Variant
    a = CURVE(TickerOrKey, TenorA, AutoFetch)
    b = CURVE(TickerOrKey, TenorB, AutoFetch)
    If PassThru(a) Then CURVEDIFF = a: Exit Function
    If PassThru(b) Then CURVEDIFF = b: Exit Function
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
    If PassThru(a) Then CURVERATIO = a: Exit Function
    If PassThru(b) Then CURVERATIO = b: Exit Function
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
    If PassThru(a) Then TICKERDIFF = a: Exit Function
    If PassThru(b) Then TICKERDIFF = b: Exit Function
    TICKERDIFF = CDbl(a) - CDbl(b)
End Function

' Ratio between two tickers at one tenor: TickerA / TickerB.
Public Function TICKERRATIO(ByVal TickerA As String, ByVal TickerB As String, _
                            Optional ByVal Tenor As String = "5Y", _
                            Optional ByVal AutoFetch As Boolean = True) As Variant
    Dim a As Variant, b As Variant
    a = CURVE(TickerA, Tenor, AutoFetch)
    b = CURVE(TickerB, Tenor, AutoFetch)
    If PassThru(a) Then TICKERRATIO = a: Exit Function
    If PassThru(b) Then TICKERRATIO = b: Exit Function
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
