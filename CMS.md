# CMS Async Curve Engine — Deep Reference

Fully asynchronous GET/SET of CMS credit curves from Excel/VBA, built on the
ExcelBridge XLL. Replaces the legacy CMS Data Access add-in and the
xmlBats/winhttpjs SET hack with byte-compatible wire formats and a reactive,
ticker-first API. **Zero dependencies on the deprecated add-in modules and no
VBA project references required** (MSXML and Scripting.Dictionary are
late-bound).

---

## Contents

1. [Files & import checklist](#1-files--import-checklist)
2. [Quick start](#2-quick-start)
3. [Configuration & transports](#3-configuration--transports)
4. [Core concepts](#4-core-concepts)
5. [Registration & ticker-first access](#5-registration--ticker-first-access)
6. [GET — every flavor](#6-get--every-flavor)
7. [SET — every flavor](#7-set--every-flavor)
8. [The 5Y perturbation workflow](#8-the-5y-perturbation-workflow)
9. [Pending levels: stage on edit, mark on button](#9-pending-levels-stage-on-edit-mark-on-button)
10. [Curve analytics](#10-curve-analytics)
11. [Reactive sheet functions (UDFs)](#11-reactive-sheet-functions-udfs)
12. [Cell subscriptions](#12-cell-subscriptions)
13. [Whole-store refresh wrappers](#13-whole-store-refresh-wrappers)
14. [The store & the curve object](#14-the-store--the-curve-object)
15. [Generic shared store](#15-generic-shared-store)
16. [Callbacks contract](#16-callbacks-contract)
17. [Delivery model — when callbacks fire](#17-delivery-model--when-callbacks-fire)
18. [Units & conventions](#18-units--conventions)
19. [Wire formats](#19-wire-formats)
20. [Failure semantics](#20-failure-semantics)
21. [Diagnostics & debugging](#21-diagnostics--debugging)
22. [Gotchas](#22-gotchas)
23. [Migration from the legacy add-in](#23-migration-from-the-legacy-add-in)

---

## 1. Files & import checklist

| File | Type | Role |
|---|---|---|
| `modCms.bas` | module | engine: wire formats, transports, parsing, store, registry, subscriptions, public API |
| `cCmsCurve.cls` | class | one curve: identity, 17-grid quotes, scalars, pending level, status |
| `cCmsBatch.cls` | class | one bulk GET/SET in flight: parallel requests + callbacks |
| `modCmsUdf.bas` | module | reactive sheet functions: `CURVE`, `CURVEDIFF`, `TICKERDIFF`, … |
| `modSharedStore.bas` | module | generic namespaced key-value session store |
| `cmsMarker.bas` | module | the workbook's Ctrl+Shift+M publish flow, rebuilt on the engine |

Prerequisites already in the workbook: the ExcelBridge XLL + `modBridge`,
`modBridgeEvents`, `modHttp`, `cHttpBatch`, `cHttpResponse`, and
`cadesHelpers` (user-info and app-state helpers). `ThisWorkbook` needs the
`Workbook_SheetCalculate` handler (drains the UDF auto-fetch queue) and
`bridge_stop` calling `modCms.CMS_StopWatchdog`.

---

## 2. Quick start

```vb
' Sheet load: register + fetch every curve on the sheet, async, in parallel.
' Headers of D9:H450 are Ticker, Ccy, DebtClass, Product, QuoteConvention.
Dim b As cCmsBatch
Set b = CMS_GetCurvesAsync([Curves!D9:H450], "LIVE", , _
                           "MyMod.OnCurve", "MyMod.OnAllDone")
' returns immediately; Excel stays live; callbacks fire as responses land

' ...later, anywhere, by ticker only:
?CMS_Quote("AEP", "5Y")
?CMS_CurveDiff("AEP")                          ' 5s10s
Set b = CMS_SetCurve5yAsync("AEP", 127.5)      ' re-mark off a new 5Y

' or in a cell:
'   =CURVE("AEP","10Y")     -> updates itself when data arrives
```

```vb
' The callbacks (plain Public Subs in a STANDARD module):
Public Sub OnCurve(curve As cCmsCurve, batch As cCmsBatch)
    If Not curve.IsOk Then Debug.Print curve.Key & " FAILED: " & curve.ErrorMsg
End Sub

Public Sub OnAllDone(batch As cCmsBatch)
    Debug.Print "done: " & batch.Count & " curves, " & batch.FailedCount & _
                " failed, " & batch.WallMs & "ms"
End Sub
```

---

## 3. Configuration & transports

Two transports carry the **same Rosetta XML**; the default is the proven
legacy SOAP path.

| Transport | Endpoint | Body |
|---|---|---|
| `SOAP` (default) | `http://cms-lxp.lehman.com/CreditMarkingService/` | rpc/encoded envelope, `<string>` = user-metadata, `<string0>` = Rosetta |
| `REST` | `<base>/api/v1/getMarketData` \| `setMarketData` | `{"requestHeader":{...},"request":"<cms-request>…"}` |

```vb
CMS_Configure "https://host:25551", "REST"     ' switch transport explicitly
CMS_Configure , , "someuser"                   ' impersonate (user-id field)
?CMS_EndpointUrl, CMS_Transport
```

Or set the optional named ranges `CmsEndpointUrl` / `CmsTransport` in the
workbook — read once on first use. The REST facade is the modernized version
of the same CreditMarkingService (verified: identical user-metadata and
Rosetta payloads); test it against DEV before switching production flows.

---

## 4. Core concepts

- **The store** (`CMS_Store()`): one in-memory `cCmsCurve` per fourTuple key
  (`TICKER.CCY.DEBTCLASS.PRODUCT`), holding the latest known state. Every
  successful GET replaces the entry; every successful SET merges the echo.
  Cells are optional — dump the store to a sheet only if you want to see it.
- **The ticker registry**: one curve per unique ticker. Registering
  (explicitly or implicitly via any full-tuple GET/SET) maps
  `TICKER -> key`, after which every call accepts just the ticker.
- **Batches** (`cCmsBatch`): a bulk GET/SET in flight. One HTTP request per
  curve by default (maximum parallelism); per-curve callbacks as responses
  land; one all-complete callback at the end.
- **Tenor grids**: GETs always use the 17-standard
  (`0M 3M 6M 9M 1Y 2Y 3Y 4Y 5Y 6Y 7Y 8Y 9Y 10Y 15Y 20Y 30Y`); SETs always
  the 12-standard (`0M 3M 6M 9M 1Y 2Y 3Y 4Y 5Y 6Y 7Y 10Y` — never
  8Y/9Y/15Y/20Y/30Y).
- **Tags**: `"LIVE"` for today, `"NYOISCLOSE"` + a `QuoteDate` for
  historical closes.
- **Reactivity**: sheet cells subscribe to curves and recalc automatically
  when the store changes.

---

## 5. Registration & ticker-first access

```vb
' One at a time:
register_curve "AEP", "USD", "SENIOR_NORE_14", "SNAC100", "QUOTED_SPREADS"

' Or a whole block (same 5 columns; blank rows skipped):
register_curve_by_range [Curves!D9:H450]
```

Registration stores the identity (without clobbering quotes already fetched)
and indexes the ticker. After that:

```vb
?CMS_GetStored("AEP").Quote("5Y")       ' store access by ticker
?CMS_ResolveKey("AEP")                  ' -> "AEP.USD.SENIOR_NORE_14.SNAC100"
?CMS_RegisteredTickers()(0)
```

**Workbook-open bootstrap** (recommended):

```vb
Private Sub Workbook_Open()
    Call http_start
    register_curve_by_range [Curves.CurveNames]   ' 5 columns!
    Application.CalculateFull   ' evaluate saved CURVE cells once so they
                                ' subscribe + queue their fetches; after this
                                ' the reactive chain is fully self-driving
End Sub
```

Also define a workbook name **`CmsRegisterRange`** referring to the same
5-tuple block: if a VBE state reset ever wipes the registry mid-session, the
engine silently re-registers from it on the next lookup (and the auto-fetch
chain refills quotes cell by cell) instead of erroring with "not registered".

Rules:

- Anything containing `.` is treated as a full fourTuple key and passes
  through unchanged.
- A GET/SET with full tuples **registers implicitly** — explicit
  registration is optional.
- A ticker-only call on an unregistered ticker **raises**, and for bulk
  inputs the resolution happens before anything is sent (no partial
  batches).
- Re-registering a ticker against a different tuple remaps it (warning in
  the Immediate window) — one curve per ticker is enforced by last-write.

---

## 6. GET — every flavor

All GETs share: `Tag` (`LIVE`/`NYOISCLOSE`), optional `QuoteDate`
(historical), callbacks, `CurvesPerRequest` (1 = one request per curve, the
default; raise to group N marketSets per request), `TimeoutMs`
(default 120s).

### Async (preferred — returns the batch immediately)

```vb
' 5-column tuple range or array (registers identities as a side effect):
Set b = CMS_GetCurvesAsync([D9:H450], "LIVE", , "M.OnCurve", "M.OnAllDone")

' 1-column ticker range (must be registered):
Set b = CMS_GetCurvesAsync([A4:A184], "LIVE")

' Single ticker / key string:
Set b = CMS_GetCurvesAsync("AEP")

' Ticker list in any shape (1D array, row, or column):
Set b = CMS_GetTickersAsync(Array("AEP", "XYZ", "ABC"))

' Per-part arguments (parts optional once registered):
Set b = CMS_GetCurveAsync("AEP", "USD", "SENIOR_NORE_14", "SNAC100", "QUOTED_SPREADS")
Set b = CMS_GetCurveAsync("AEP")                    ' registered ticker

' Historical (T-1 business day close):
Set b = CMS_GetCurvesAsync(rng, "NYOISCLOSE", Date - 1)
```

### Sync (blocks the calling VBA only; I/O still parallel underneath)

```vb
a = CMS_GetCurveQuoteData("AEP", "USD", "SENIOR_NORE_14", "SNAC100", _
                          "LIVE", Now(), "QUOTED_SPREADS")   ' legacy signature
a = CMS_GetCurveQuoteData("AEP")                             ' ticker-only
a2 = CMS_GetBulkCurveQuoteData([D9:H450], "LIVE")            ' 2D, row-per-input-row
a2 = CMS_GetBulkCurveQuoteData([A4:A30])                     ' ticker column
```

Result rows are 25 columns: the 17 tenors then `Recovery, LastMarkedOn,
LastMarkedBy, Owner, IsParent, CurveType, Status, ErrorMessage`
(`CMS_GetCurveQuoteDataHeader17()` gives the header).

---

## 7. SET — every flavor

Quotes are passed in the units the GET handed you (see
[Units](#18-units--conventions)). `TermQuotes`/`Quotes` rows may be 12-wide
(SET-tenor order) or 17-wide (full grid; the never-set tenors are ignored).
Recovery rates are only sent when `> 0`.

### Async

```vb
' Bulk - the CMS_SetBulkCurveQuoteData that never existed, parallel per curve:
Set b = CMS_SetCurvesAsync([D9:H450], quotes2D, recoveries, "M.OnSet", "M.OnAllSet")
Set b = CMS_SetCurvesAsync([A4:A30], quotes2D)        ' ticker column

' Single curve, per-part or ticker-only:
Set b = CMS_SetCurveAsync("AEP", quotesRow, "USD", "SENIOR_NORE_14", "SNAC100")
Set b = CMS_SetCurveAsync("AEP", quotesRow)           ' registered ticker
```

### Sync

```vb
r = CMS_SetCurveQuoteData("AEP", "USD", "SENIOR_NORE_14", "SNAC100", quotesRow)  ' legacy
r = CMS_SetTickerQuoteData("AEP", quotesRow)          ' ticker-only
r2 = CMS_SetBulkCurveQuoteData([D9:H450], quotes2D)   ' blocking bulk
```

SET responses echo the sent values; on success the echo is merged over the
computed curve in the store, so the cache stays exact.

---

## 8. The 5Y perturbation workflow

The transform (`CMS_ApplyNew5y`): front end `0M..4Y` keeps the **ratio** to
the old 5Y (`old/old5y * new5y`); long end `6Y..30Y` keeps the
**difference** (`new5y + (old − old5y)`); `5Y` becomes the new mark. All 17
points are computed and cached; only the 12 SET tenors go to CMS.

```vb
' One curve (cached quotes drive the transform - GET first):
Set b = CMS_SetCurve5yAsync("AEP", 127.5, "M.OnSet", "M.OnAllSet")

' Many curves: N x 2 array/range of (ticker-or-key, new 5Y):
Set b = CMS_SetCurves5yAsync([K4:L30], "M.OnSet", "M.OnAllSet")
Set b = CMS_SetCurves5yAsync(Array("AEP", 127.5))     ' single row shorthand

' Pure math, no I/O:
newQuotes17 = CMS_ApplyNew5y(CMS_GetStored("AEP").Quotes, 127.5)
```

Per-row failure handling: rows whose curve isn't in the store, whose level
isn't numeric, or whose cached curve has no/zero 5Y are FAILED individually
(callback fires with the reason) without sinking the rest of the batch.

---

## 9. Pending levels: stage on edit, mark on button

Staged amends live **on the cached curves** — no second cache, no cell
re-reads at mark time.

```vb
' ----- in the change handler (hardened: never leaves EnableEvents off,
'       never raises mid-event, uses Sh/Target's sheet not ActiveSheet) -----
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo SafeExit
    If Not NamedRangeExistsIn("cms_col", Sh) Then Exit Sub
    If Target.Column <> Sh.Range("cms_col").Column Then Exit Sub
    If Target.Row < 4 Or Target.Row > Sh.Range("last_row").Row Then Exit Sub

    Dim t As String: t = CStr(Sh.Cells(Target.Row, "A").Value2)
    If Not CMS_IsRegistered(t) Then Exit Sub      ' blank row / wiped store

    If IsEmpty(Target.Value2) Or Target.Value2 = "" Then
        Application.EnableEvents = False
        Target.Value2 = CMS_ClearPending5y(t)     ' restore cached original
    ElseIf IsNumeric(Target.Value2) Then
        CMS_StagePending5y t, CDbl(Target.Value2)
        CMS_SubscribeRange t, Sh.Cells(Target.Row, "A").Resize(1, 70)
    End If

SafeExit:
    Application.EnableEvents = True               ' ALWAYS restore
End Sub

' ----- the mark button -----
Public Sub mark_button()
    Dim b As cCmsBatch
    Set b = CMS_MarkPendingAsync("MyMod.OnSet", "MyMod.OnAllSet")
    If b Is Nothing Then MsgBox "Nothing staged."
End Sub
```

`CMS_MarkPendingAsync` drains: reads every staged level, clears it, computes
each full curve from the cache, and launches the parallel SETs. A failed SET
leaves the store unchanged so the user can re-stage. Inspection:

```vb
?CMS_PendingCount(), Join(CMS_PendingKeys(), ", ")
?CMS_Pending5y("AEP"), CMS_HasPending("AEP")
CMS_ClearAllPending          ' abandon everything staged
```

`=CURVE("AEP","Pending5y")` cells update live as levels are staged/cleared.

---

## 10. Curve analytics

Pure cache reads — **no hidden I/O** (the sheet UDFs add auto-fetch; the VBA
functions never fetch). Unregistered tickers raise; missing quotes return
`Empty`.

```vb
?CMS_Quote("AEP", "7Y")
?CMS_CurveDiff("AEP")                  ' 5s10s steepness: 10Y - 5Y
?CMS_CurveDiff("AEP", "1Y", "5Y")      ' 1s5s
?CMS_CurveRatio("AEP", "5Y", "10Y")    ' 10Y / 5Y
?CMS_TickerDiff("AEP", "XYZ", "5Y")    ' AEP - XYZ at 5Y
?CMS_TickerRatio("AEP", "XYZ")         ' AEP / XYZ at 5Y
```

Conventions: tenor pairs read *second minus/over first*; ticker pairs read
*first minus/over second*.

---

## 11. Reactive sheet functions (UDFs)

```vb
=CURVE("AEP")                  ' 5Y quote (default field)
=CURVE("AEP", "10Y")           ' any tenor
=CURVE("AEP", "Recovery")      ' any field: Status, Owner, CurveType, Error,
                               '   LastMarkedOn/By, IsParent, Pending5y,
                               '   FetchedAt, Ccy, Product, Key, ...
=CURVEDIFF("AEP")              ' 5s10s
=CURVERATIO("AEP","5Y","10Y")
=TICKERDIFF("AEP","XYZ","5Y")
=TICKERRATIO("AEP","XYZ")
```

Historical-close fields (see [below](#12b-historical-business-day-closes)):
`Prev5Y` (T-1), `Prev2_5Y` (T-2), … and `PrevDate` / `Prev2_Date` /
`PrevFetchedAt` — so `=CURVEDIFF("AEP","Prev5Y","5Y")` is the day change at
5Y, and `=CURVEDIFF("AEP","Prev2_5Y","5Y")` is the 2-day change.

- **Self-subscribing**: each call registers its own cell against the
  curve(s) it reads. Not volatile — only affected cells are touched.
  Because they're not volatile, **F9/Shift+F9 never re-evaluates a clean
  CURVE cell** — the subscription notify is what dirties/recalcs them.
- **Notify modes** (`CMS_SetNotifyMode`): how a store change reaches
  subscribed cells. `CMS_NOTIFY_CALC` (default) recalculates them
  immediately — `#Pending` becomes the level with no user action.
  `CMS_NOTIFY_DIRTY` only marks them (picked up at the next recalc);
  `CMS_NOTIFY_OFF` leaves cells alone. If CALC ever feels laggy, check
  `ThisWorkbook.ForceFullCalculation` in `CMS_Diag` before blaming it.
- **Auto-fetch**: a quote read on a *registered* but unfetched curve shows
  **`#Pending`**, queues the key, and arms the OnTime watchdog — the GET
  launches within ~1s with no user action. Derived functions
  (`CURVEDIFF`, …) propagate `#Pending` instead of erroring. 30s cooldown;
  FAILED curves show `#N/A` and are **not** auto-retried
  (`=CURVE(t,"Error")` shows why — `CMS_RefreshFailedAsync` retries);
  unregistered tickers show `#N/A` and are never fetched. Pass `FALSE` as
  the last argument to disable.
- Unknown field names return `#NAME?`.

---

## 12. Cell subscriptions

The mechanism under the UDFs, available directly for anything else:

```vb
CMS_SubscribeRange "AEP", ws.Range("B17:BT17")  ' recalc this row on any change
CMS_UnsubscribeRange "AEP", ws.Range("B17:BT17")
CMS_UnsubscribeAll "AEP"                        ' one curve
CMS_UnsubscribeAll                              ' everything
?CMS_SubscriptionCount()
```

Fires on: GET arrival, SET confirm, SET **failure** (status cells stay
honest), pending stage/clear. In automatic calc mode ranges are marked
`Dirty` (Excel recalcs dependents after the delivery macro); in manual mode
the subscribed ranges are calculated directly. Subscriptions to
unregistered tickers are held by name and start firing once the ticker
exists. Stale references (deleted sheets/cells) self-prune.

---

## 12b. Historical business-day closes

Each curve caches **one 17-grid per business date** (`NYOISCLOSE`,
weekend-aware via `WorkDay`; holidays too if a `CmsHolidays` named range of
dates exists), so T-1, T-2, … T-N coexist on the same store entry without
overwriting each other. Closes don't change intraday, so each (curve, date)
is fetched at most once per day.

- **Fresh live GET** auto-fetches T-1 through **T-`CMS_HistDepth`**
  (default **1 = T-1 only**). `CMS_SetHistDepth 2` to also pull T-2 on every
  fresh GET; `CMS_SetAutoPrevClose False` to disable auto-fetch entirely.
- **Deeper history is fetched on demand**, independent of the depth: a
  `=CURVE(t,"Prev2_5Y")` cell queues T-2 when it evaluates, or fetch
  explicitly with `CMS_GetHistCloseAsync(spec, 2)`.

```vb
?CMS_PrevQuote("AEP", "5Y")            ' cached T-1 close
?CMS_HistQuote("AEP", "5Y", 2)         ' cached T-2 close (Empty if not cached)
?CMS_DayChange("AEP", "5Y")            ' live - T-1
?CMS_HistChange("AEP", "5Y", 2)        ' live - T-2
?CMS_BizDay(-2)                        ' the T-2 business date
?CMS_GetStored("AEP").HistQuote("10Y", CMS_BizDay(-2))

' Explicit fetch (whole store or a spec) at an offset:
Set b = CMS_GetHistCloseAsync(, 2)             ' T-2 for everything
Set b = CMS_GetHistCloseAsync([A4:A30], 2)     ' T-2 for a subset
Set b = CMS_GetPrevCloseAsync()                ' T-1 alias, everything
```

Live quotes and every cached close live on the same `cCmsCurve` (`Quotes`
vs the date-keyed history; `HistQuotesFor(date)`, `HistDates()`), and all of
it survives refresh GETs (refreshes carry over the cached closes and any
staged pending level). In cells: `=CURVE("AEP","Prev5Y")` (T-1),
`=CURVE("AEP","Prev2_5Y")` (T-2), `=CURVE("AEP","Prev2_Date")`.

## 13. Whole-store refresh wrappers

All return the launched `cCmsBatch`, or `Nothing` when there is nothing to
do. All accept the usual `Tag, QuoteDate, OnCurve, OnAllDone,
CurvesPerRequest, TimeoutMs`.

```vb
Set b = CMS_RefreshAllAsync()             ' force-refresh EVERYTHING in the store
Set b = CMS_GetAllAsync()                 ' alias of the above
Set b = CMS_RefreshFailedAsync()          ' retry only FAILED curves
                                          ' (also re-arms UDF auto-fetch for them)
Set b = CMS_RefreshStaleAsync(15)         ' never-fetched or older than 15 min

?CMS_RegisteredCount()
?Join(CMS_RegisteredTickers(), ", ")
?Join(CMS_RegisteredKeys(), ", ")

CMS_WriteStoreToRange [Dump!A1]           ' whole store -> sheet, with header
```

Staleness uses `cCmsCurve.FetchedAt` (stamped on every successful server
round trip; also readable as `=CURVE("AEP","FetchedAt")`).

---

## 14. The store & the curve object

```vb
Dim cv As cCmsCurve
Set cv = CMS_GetStored("AEP")             ' by ticker (or StoreGet(key) by key)

cv.Quote("5Y")        ' selected quote per the curve's QuoteConvention
cv.Quotes             ' Variant(0..16), the full 17-grid
cv.Spreads            ' raw creditSpread points   (bps on GETs)
cv.Upfronts           ' raw creditUpfront points  (% on GETs)
cv.Field("Owner")     ' dynamic access to any field by name
cv.FiveYear()         ' Quotes(8)
cv.ApplyNew5y(127.5)  ' transformed 17-grid (pure function)
cv.ToRow()            ' the 25-column result row
cv.Recovery / .LastMarkedOn / .LastMarkedBy / .Owner / .IsParent / .CurveType
cv.Status / .ErrorMsg / .IsOk() / .HasQuotes()
cv.Pending5y          ' staged amend (Empty = none)
cv.FetchedAt / .ElapsedMs / .HttpStatus / .RequestId
cv.Key                ' "AEP.USD.SENIOR_NORE_14.SNAC100"
```

Store-level: `CMS_Store()` (the dictionary itself), `CMS_StoreToArray()`
(2D dump incl. a trailing "Pending 5Y" column), `CMS_StoreClear`,
`CMS_WriteBatchToRange b, [A1]` (a batch's results, row-per-input-row).

Registered-but-unfetched curves sit in the store with `Empty` quotes and
`Status = ""` until their first GET.

---

## 15. Generic shared store

`modSharedStore.bas` — namespaced key-value cache for anything that must
survive across modules, event handlers, and async callbacks:

```vb
Shared_Set "axes", "AEP", anythingIncludingObjects
v = Shared_Get("axes", "AEP")                  ' Empty if missing
v = Shared_GetOrDefault("axes", "AEP", 0)
?Shared_Has("axes", "AEP"), Shared_Count("axes")
Shared_Remove "axes", "AEP"
Shared_Clear "axes"
For Each k In Shared_Keys("axes"): ... : Next

Set snapshot = Shared_Drain("axes")  ' ATOMIC take-all + clear: iterate the
                                     ' snapshot while new entries land in a
                                     ' fresh namespace - ideal for
                                     ' stage-then-commit button flows
```

Session-scoped: a VBE reset wipes it, like all module-level state.

---

## 16. Callbacks contract

Plain `Public Sub`s in a **standard module** (not a sheet/ThisWorkbook
module), passed by name (`"Module.Sub"` preferred):

```vb
Public Sub OnCurve(curve As cCmsCurve, batch As cCmsBatch)   ' per response
Public Sub OnAllDone(batch As cCmsBatch)                     ' once, at the end
```

- Both run on Excel's main thread — touching sheets is safe.
- Exceptions inside callbacks are swallowed and logged to the Immediate
  window; they never break batch accounting.
- Rows failed at launch (bad input, unregistered, no cached curve) fire the
  per-curve callback synchronously *during* the launching call.
- The batch object stays alive and inspectable until complete (and
  `CMS_LastBatch` keeps the most recent one) even if your local variable
  goes out of scope.

Useful batch members: `Count`, `CompletedCurves`, `FailedCount`,
`PendingRequests`, `IsDone`, `Action`, `Keys`, `Curve(key)`, `WallMs`,
`MaxRequestMs`, `WaitAll([ms])`, `ToArray([header])`.

---

## 17. Delivery model — when callbacks fire

The .NET engine does all HTTP I/O on worker threads — always running. The
final hop into VBA is queued via `QueueAsMacro` and executes when Excel's
main thread can run a macro:

| Excel state | Delivery |
|---|---|
| Idle at the grid (normal use) | automatic, immediate |
| Your VBA running | deferred until `DoEvents` (what `WaitAll`/`CMS_Pump` pump) |
| VBE **break mode** (breakpoint / *Debug* on an error) | blocked until you resume |
| Cell edit mode / modal dialog | blocked until you leave it |

**Watchdog**: while any batch is in flight, a 1-second `Application.OnTime`
tick stays armed; its `DoEvents` flushes any delivery stuck waiting for a
macro context (this is what makes fire-and-forget SET-only flows complete
hands-off). It disarms when the last batch finishes; `bridge_stop` cancels
it on close. Break mode still blocks everything — nothing can run there.

**State resets wipe everything**: *End* on an error dialog, Stop/Reset, or
editing code in break mode clears the store, registry, subscriptions, AND
modBridge's request router — in-flight responses arriving after a reset are
dropped. Relaunch after any reset.

---

## 18. Units & conventions

Identical to the legacy add-in:

| Direction | creditSpread | creditUpfront |
|---|---|---|
| GET response | raw × 10000 → **bps** | raw × 100 → **%** |
| SET request | sent **verbatim** | sent verbatim |
| SET response | echoed verbatim | echoed verbatim |

So: SET in the same units the GET handed you and everything is consistent
(the 5Y/pending workflows do this automatically since they compute from the
cache). Quote selection per `QuoteConvention`: `SPREADS`/`QUOTED_SPREADS` →
creditSpread; `UPFRONT` → creditUpfront; `QUOTED` → upfront, falling back to
spread; `RUNNING` → spread, falling back to upfront. Standard-contract
products (`SNAC*`, `STEC*`, `STE*`, `STA*`, `BLCDS*`, `SUKU*`) get
`<contractualSpread>-777</contractualSpread>` on SET points, exactly as the
old flow sent.

---

## 19. Wire formats

**GET** (N marketSets per request — 1 by default):

```xml
<Rosetta version="5.0.15"><market><marketData><action>GET</action>
<date>yyyymmdd</date>
<marketSet><location>NYC</location><currency>CCY</currency>
  <dataSource>TAG</dataSource><version>CLOSE</version><type>creditSpread</type>
  <label>TICKER.CCY.DEBTCLASS.PRODUCT</label>
  <genericKeys>
    <nameValuePair name="tag">TAG</nameValuePair>
    <nameValuePair name="returnCreditCurveType">CONV</nameValuePair>
  </genericKeys>
</marketSet>…</marketData></market></Rosetta>
```

**SET** (one curve per request; mirrors `xmlBats.create_cds_xml`):

```xml
<Rosetta version="5.0.15"><market><marketData><action>SET</action>
<date>yyyymmdd</date>
<marketSet><location>NYC</location><logicalTime>LIVE</logicalTime>
  <point><label>4TUPLE</label>
    <creditSpread>
      <issuerTicker/><debtClass/><currency/><creditCurveType>PRODUCT</creditCurveType>
      <periodMultiplier>5</periodMultiplier><period>Y</period>
      <contractualSpread>-777</contractualSpread>   <!-- standard-contract only -->
    </creditSpread>
    <value>123.45</value>
  </point>…</marketSet></marketData></market></Rosetta>
```

User-metadata matches the legacy `CompileUserMetaData` field-for-field
(user-id, user-domain, application-id/-version, application-batch-id, id,
host-id). Bodies over 30K chars are staged into the XLL in chunks
automatically.

---

## 20. Failure semantics

- **Transport errors** (timeout, non-2xx, SOAP fault, REST FAILURE) fail
  every curve on that request with the reason; other requests in the batch
  are unaffected.
- **Curve-level CMS errors** arrive as batch `errorMessageContent` keyed by
  fourTuple inside the text (legacy contract) and are mapped per curve.
- A GET curve with overall `OK` but no points and no recovery → FAILED
  ("no data") — the legacy bulk rule.
- A SET whose response status isn't `OK` → FAILED regardless of echoed
  points.
- Failed SETs leave the store untouched; failed GETs leave the previous
  good entry in place (the store only updates on success — status/error are
  visible on the batch's curve and via subscriptions).

---

## 21. Diagnostics & debugging

```vb
?modCms.CMS_Diag
'  transport=SOAP; endpoint=http://...
'  bridge loaded=True; last dispatch error=
'  store curves=412; pending 5Y=3; deliveries=841; watchdog flushes=0
'  cell subscriptions=57; autofetch queued=0
'  active batches=1
'    [1] GET: 380/412 curves done, 32 request(s) pending, 0 failed
'  last batch: SET done=True curves=3 failed=0 wall=420ms slowest request=391ms
```

- `watchdog flushes > 0` → deliveries genuinely sat stuck until the watchdog
  freed them (environment-dependent; harmless but worth knowing).
- `wall ≈ slowest request` → delivery immediate; `wall ≫ slowest request` →
  responses waited for a macro context.
- `modBridge.gHttpVerbose = True` → per-request tracing.
- From the Immediate window: `CMS_Pump 5000` pumps until active batches
  finish (nothing pumps between Immediate statements otherwise);
  `CMS_LastBatch` inspects the most recent launch.

---

## 22. Gotchas

- **Ticker lists must be vertical** (1 column) for the generic entry points
  — a 1×N row is indistinguishable from a malformed tuple row. Use
  `CMS_GetTickersAsync` for arbitrary shapes.
- The legacy `CMS_SetCurveQuoteData` signature keeps `Ccy/DebtClass/Product`
  required (VBA forbids required-after-optional) — use
  `CMS_SetTickerQuoteData` for ticker-only sync SETs.
- VBA analytics return `Empty` for missing quotes; only the sheet UDFs
  auto-fetch.
- Registered-but-unfetched curves: staging a pending level is allowed, but
  the mark will fail that curve ("no 5Y mark") — GET before staging.
- UDF names (`CURVE`, …) are global to the workbook; rename in
  `modCmsUdf.bas` if they collide with anything.
- Everything module-level (store, registry, subscriptions, staged levels)
  is session state — a VBE reset wipes it.

---

## 23. Migration from the legacy add-in

| Legacy | Now |
|---|---|
| `CMS_GetCurveQuoteData(t,c,d,p,tag,date,conv)` | same call (17-grid rows; also ticker-only) |
| `CMS_GetBulkCurveQuoteData(range5,tag,date)` | same call (parallel under the hood; also ticker columns) |
| `CMS_SetCurveQuoteData(t,c,d,p,quotes,…)` | same call (also `CMS_SetTickerQuoteData`) |
| *(didn't exist)* | `CMS_SetBulkCurveQuoteData` / `CMS_SetCurvesAsync` |
| `CMS_CompileFourTuple` / `CMS_CompileDebtClass…` | same names in `modCms` |
| MSSOAP `SoapClient30`, `MSXML2` references | gone — late-bound, reference-free |
| `xmlBats.create_cds_xml` + winhttpjs.bat | `BuildRosettaSet` + ExcelBridge (internal) |
| 11/12/13-tenor standards, ad-hoc tenor arrays, EXTENDED 2-row mode, SetFlat, recovery-only setters, authorization/impersonation UI, menus, registry/session plumbing | dropped intentionally — GETs are always 17-standard, SETs always the 12-standard, per this workbook's usage |
