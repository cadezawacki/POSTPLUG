# CMS Async Curve Engine

Replaces the legacy CMS Data Access add-in (`modCmsGetCurveSpreads`,
`modCmsSetCurveSpreads`, `CMarketData`, `CWebServiceHelper`, `CArrayHelper`,
`modConst`, `modInternalUtils`, `modPublicUtils`, `modWindowsAPI`,
`modSpecificReferenceCurves`, `modControlMenu`) and the `xmlBats` /
winhttpjs.bat SET hack with three files built on the ExcelBridge async HTTP
engine:

| File | Role |
|---|---|
| `modCms.bas` | Wire formats, transports, parsing, curve math, public API, in-memory store |
| `cCmsCurve.cls` | One curve: 5-tuple identity, 17-grid quotes, scalars, status |
| `cCmsBatch.cls` | A bulk GET/SET in flight: per-curve + all-complete callbacks |

`cmsMarker.bas` (the Ctrl+Shift+M publish flow) was rewritten on top of it.
Everything is late-bound — no MSXML/MSSOAP/Outlook references needed.

## Is the `/api/v1` doc the same service?

Yes. `/api/v1/getMarketData` and `/api/v1/setMarketData` are the modernized
JSON facade over the same CreditMarkingService: the JSON `request` field
carries `<cms-request><operation name="getMarketData">` wrapping the **same
user-metadata and Rosetta XML** the old SOAP add-in sent. The user-metadata
format in the migration note matches the legacy `CompileUserMetaData`
field-for-field, and the note explicitly says SET calls are unaffected.
Both transports are implemented:

- **SOAP** (default — the proven path the workbook uses today): rpc/encoded
  envelope, `<string>` = user metadata, `<string0>` = Rosetta, POSTed to
  `http://cms-lxp.lehman.com/CreditMarkingService/`, `SOAPAction: ""`.
- **REST**: `{"requestHeader":{...},"request":"<cms-request>..."}` POSTed to
  `<base>/api/v1/getMarketData|setMarketData`; response is
  `{"responseHeader":{code,...},"response":"<Rosetta.../>"}`.

Switch with `CMS_Configure "https://host:25551", "REST"` or the optional
`CmsEndpointUrl` / `CmsTransport` named ranges. Everything downstream of the
transport (Rosetta build + parse) is shared.

## Wire formats kept byte-compatible

**GET** (one `<marketSet>` per curve, any number per request):

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
</marketSet>...</marketData></market></Rosetta>
```

**SET** (per curve; mirrors `xmlBats.create_cds_xml` exactly):

```xml
<Rosetta version="5.0.15"><market><marketData><action>SET</action>
<date>yyyymmdd</date>
<marketSet><location>NYC</location><logicalTime>LIVE</logicalTime>
  <point><label>4TUPLE</label>
    <creditSpread>
      <issuerTicker/><debtClass/><currency/><creditCurveType>PRODUCT</creditCurveType>
      <periodMultiplier>5</periodMultiplier><period>Y</period>
      <contractualSpread>-777</contractualSpread>   <!-- standard-contract products only -->
    </creditSpread>
    <value>123.45</value>
  </point> ... </marketSet></marketData></market></Rosetta>
```

## Conventions baked in

- **GET tenors**: always the 17-standard
  `0M 3M 6M 9M 1Y 2Y 3Y 4Y 5Y 6Y 7Y 8Y 9Y 10Y 15Y 20Y 30Y`.
- **SET tenors**: the 12 we mark — `0M 3M 6M 9M 1Y 2Y 3Y 4Y 5Y 6Y 7Y 10Y`
  (8Y/9Y/15Y/20Y/30Y are computed/stored but never sent).
- **Tags**: `LIVE` today, `NYOISCLOSE` for historical dates
  (`QuoteDate` parameter → Rosetta `<date>`).
- **Units** (identical to the legacy add-in): GET responses are scaled —
  creditSpread ×10000 → bps, creditUpfront ×100 → %. SET requests send your
  values verbatim and the SET response echoes them verbatim. So: SET in the
  same units the GET handed you, and everything is consistent.
- **5Y perturbation** (`CMS_ApplyNew5y`): front end `0M..4Y` keeps the *ratio*
  to the old 5Y (`old/old5y * new5y`); long end `6Y..30Y` keeps the
  *difference* (`new5y + (old - old5y)`); `5Y` = the new mark.

## API

### Async (preferred)

```vb
' Sheet load: bulk GET of D9:H450 (Ticker,Ccy,DebtClass,Product,QuoteConvention)
Dim b As cCmsBatch
Set b = CMS_GetCurvesAsync(Sheet1.Range("D9:H450"), "LIVE", , _
            "MyMod.OnCurve", "MyMod.OnAllDone")           ' returns immediately

' T-1 close
Set b = CMS_GetCurvesAsync(rng, "NYOISCLOSE", Date - 1, ...)

' Perturb: new 5Y levels for changed curves only; full curve computed from
' the store (populated by the GET), 12 tenors SET per curve, all parallel
Set b = CMS_SetCurves5yAsync(Array2D_of_key_and_new5y, "MyMod.OnSet", "MyMod.OnAllSet")

' Raw bulk SET (quotes 12-wide in SET order, or 17-wide full grid)
Set b = CMS_SetCurvesAsync(fiveTuples, quotes2D, recoveries, "MyMod.OnSet", "MyMod.OnAllSet")
```

Callback signatures (plain module subs, run on the main thread):

```vb
Public Sub OnCurve(curve As cCmsCurve, batch As cCmsBatch)  ' each response as it lands
Public Sub OnAllDone(batch As cCmsBatch)                    ' once, after the last one
```

`CurvesPerRequest` (GET only) groups N marketSets per HTTP request —
default 1 = maximum parallelism; raise it if the server prefers fewer,
fatter requests. SETs are always one curve per request.

### Sync compatibility (legacy signatures, 25-column rows: 17 tenors + Recovery,
LastMarkedOn, LastMarkedBy, Owner, IsParent, CurveType, Status, Error)

```vb
CMS_GetCurveQuoteData ticker, ccy, dc, product, "LIVE", , "SPREADS"   ' 1 row
CMS_GetBulkCurveQuoteData fiveTuplesRange, "LIVE"                     ' N rows
CMS_SetCurveQuoteData ticker, ccy, dc, product, termQuotes            ' echoed row
CMS_SetBulkCurveQuoteData fiveTuples, quotes2D                        ' N echoed rows
```

These block the calling VBA (DoEvents pump) but the I/O is still parallel
underneath — a 400-curve "bulk get" is N concurrent requests, not one
3-minute SOAP call.

### Tickers as first-class handles

One curve per unique ticker. Register identities once — explicitly, or
implicitly by GET/SETting with full tuples — and every call afterwards
accepts just the ticker:

```vb
register_curve "AEP", "USD", "SENIOR_NORE_14", "SNAC100", "QUOTED_SPREADS"
register_curve_by_range [Curves!D9:H450]      ' bulk; blank rows skipped

CMS_GetCurveQuoteData "AEP"                   ' sync, ticker-only
Set b = CMS_GetCurvesAsync([A4:A184], ...)    ' 1-col ticker range
Set b = CMS_GetCurvesAsync("AEP", ...)        ' single ticker string
Set b = CMS_GetTickersAsync(Array("AEP","XYZ"))  ' any-shape ticker list
Set b = CMS_SetCurveAsync("AEP", quotesRow)   ' ticker-only async SET
CMS_SetTickerQuoteData "AEP", quotesRow       ' ticker-only sync SET
Set b = CMS_SetCurve5yAsync("AEP", 125.5)     ' 5Y re-mark off the cache
```

Every get/set, sync and async, accepts: per-part arguments, a 5-column tuple
range, a 1-column ticker range, or a single ticker. Unregistered tickers
**raise** (before anything is sent). Ticker *lists* must be a single column
(or use `CMS_GetTickersAsync`, which normalizes any shape).

### Curve analytics (off the store, no I/O)

```vb
CMS_Quote("AEP", "5Y")                ' cached quote
CMS_CurveDiff("AEP")                  ' 5s10s steepness: 10Y - 5Y (tenors overridable)
CMS_CurveRatio("AEP", "5Y", "10Y")    ' 10Y / 5Y
CMS_TickerDiff("AEP", "XYZ", "5Y")    ' AEP - XYZ at 5Y
CMS_TickerRatio("AEP", "XYZ", "5Y")   ' AEP / XYZ at 5Y
```

Unregistered tickers raise; missing quotes return `Empty`.

### Pending 5Y workflow (stage on edit, mark on button)

Staged amends live on the cached curves themselves — no separate cache to
reconcile, no cell re-reads at mark time:

```vb
' sheet-change handler: user typed a new 5Y over the live one
CMS_StagePending5y ticker, newLevel

' user deleted their amend: un-stage and get the original back for the cell
cell.Value2 = CMS_ClearPending5y(ticker)

' mark button: drain every staged level into one parallel async SET
Set b = CMS_MarkPendingAsync("MyMod.OnSet", "MyMod.OnAllSet")
If b Is Nothing Then MsgBox "Nothing staged"
```

Also: `CMS_Pending5y`, `CMS_HasPending`, `CMS_PendingKeys`, `CMS_PendingCount`,
`CMS_ClearAllPending`. Staged levels are consumed (cleared) when the mark
drains; a failed SET leaves the store unchanged so the user can re-stage.
`CMS_StoreToArray` includes a trailing "Pending 5Y" column.

### Reactive sheet functions (`modCmsUdf.bas`)

```vb
=CURVE("AEP")                  ' 5Y quote (default)
=CURVE("AEP", "10Y")           ' any tenor
=CURVE("AEP", "Recovery")      ' any field: Status, Owner, CurveType, Error,
                               '   LastMarkedOn/By, IsParent, Pending5y, Key, ...
=CURVEDIFF("AEP")              ' 5s10s (10Y - 5Y); tenors overridable
=CURVERATIO("AEP","5Y","10Y")
=TICKERDIFF("AEP","XYZ","5Y")  ' AEP - XYZ at 5Y
=TICKERRATIO("AEP","XYZ")
```

Each call **subscribes its own cell** to the curve(s) it reads
(`Application.Caller`), so the cell recalcs automatically whenever the store
changes — GET arrives, SET confirms/fails, pending level staged or cleared.
Not volatile: only affected cells recalc, marked `Dirty` per curve as the
deliveries land (in manual calc mode the subscribed cells are calculated
directly).

**Auto-fetch**: a `CURVE` quote on a *registered* but unfetched curve shows
`#N/A`, queues the key, and the GET launches the moment calculation ends
(`Workbook_SheetCalculate` → `CMS_FlushAutoFetch`; the watchdog drains it
too). When the response lands the cell updates itself. 30s cooldown prevents
refetch storms; FAILED curves are not auto-retried (`=CURVE(t,"Error")` says
why); unregistered tickers are never fetched (no identity) — register first.
Pass `FALSE` as the last argument to disable. The VBA analytics
(`CMS_CurveDiff` etc.) stay pure cache reads — no hidden I/O.

Manual subscriptions for anything that isn't a `CURVE` cell — e.g. recalc a
row of formulas when a pending mark confirms:

```vb
CMS_SubscribeRange "AEP", ws.Range("B17:BT17")   ' fires on every store change
CMS_UnsubscribeRange "AEP", ws.Range("B17:BT17")
CMS_UnsubscribeAll "AEP"                          ' or all: CMS_UnsubscribeAll
```

### Generic shared store (`modSharedStore.bas`)

Namespaced key-value cache for anything that must survive across modules,
event handlers, and async callbacks (session-scoped — a VBE reset wipes it,
like all module state):

```vb
Shared_Set "axes", "AEP", someValueOrObject
v = Shared_Get("axes", "AEP")                 ' Empty if missing
Shared_GetOrDefault "axes", "AEP", 0
Shared_Has / Shared_Remove / Shared_Clear / Shared_Keys / Shared_Count
Set snapshot = Shared_Drain("axes")           ' atomic take-all + clear
```

### Store (curves in memory, cells optional)

Every successful GET upserts `CMS_Store()` (key = `TICKER.CCY.DEBTCLASS.PRODUCT`);
every successful SET merges the echo over the computed curve. Sheet output is
opt-in:

```vb
CMS_StoreToArray()                          ' everything, with header
CMS_WriteBatchToRange b, [A1], True         ' a batch, row-per-input-row
StoreGet("IBM.USD.SENIOR.DERIV").Quote("5Y")
```

## Delivery model — when do callbacks actually fire?

The .NET engine does the HTTP I/O on worker threads — that part is always
running in the background. The **final hop into VBA** (`EB_OnHttpBatch` →
`cHttpBatch` → `cCmsBatch` → your callbacks) is queued via Excel-DNA's
`QueueAsMacro` and executes only when Excel's main thread can run a macro:

| Excel state | Delivery |
|---|---|
| Idle at the grid (normal use) | automatic, immediate |
| Your VBA still running | deferred until `DoEvents` (this is what `WaitAll` pumps) |
| VBE **break mode** (breakpoint, or you clicked *Debug* on an error) | **blocked entirely** until you resume/reset |
| Cell edit mode / modal dialog | blocked until you leave it |

So in production (button fires `CMS_GetCurvesAsync`, the sub returns, the
user keeps working) callbacks fire by themselves as responses land. In a
*debugging session* nothing arrives while you sit at a breakpoint — and the
Immediate window doesn't pump between statements, so after launching from
there call `CMS_Pump 2000` to flush deliveries.

**Watchdog.** While any batch is in flight, modCms keeps a 1-second
`Application.OnTime` tick armed (`CmsWatchdogTick`); entering the tick gives
Excel a macro context and its `DoEvents` flushes any delivery that got stuck
waiting for one. This is what makes fire-and-forget SET-only flows complete
hands-off even when nothing else ever pumps. It disarms itself when the last
batch finishes; `bridge_stop` cancels it on workbook close. (Break mode still
blocks everything — OnTime can't fire there either.)

**State resets wipe everything.** Pressing *End* on an error dialog, the
VBE Stop/Reset button, or editing code in break mode clears every
module-level variable: the curve store, the active-batch registry, **and
modBridge's request router** — in-flight responses arriving after a reset
are silently dropped. If you reset mid-test, relaunch the batch.

### Debugging recipe (Immediate window)

```vb
?modCms.CMS_Diag                       ' transport, endpoint, store size, live batches,
                                       ' and the XLL's last Application.Run failure
Set b = CMS_GetCurvesAsync([D9:H450], "LIVE", , "MyMod.OnCurve", "MyMod.OnAllDone")
CMS_Pump 5000                          ' pump until done (or timeout)
?CMS_LastBatch.IsDone, CMS_LastBatch.FailedCount, CMS_Store().Count
?StoreGet("AEP.USD.SENIOR_NORE_14.SNAC100").Quote("5Y")   ' Quote returns a VALUE - no Set
```

Launched batches are held in a registry until complete (and `CMS_LastBatch`
keeps the most recent one), so they stay alive and inspectable even after
your local `batch` variable goes out of scope.

Callback name gotchas: the callback subs must be `Public` in a **standard
module** (not a sheet/ThisWorkbook module) with the exact signatures
`(curve As cCmsCurve, batch As cCmsBatch)` / `(batch As cCmsBatch)`.
If `Application.Run` can't resolve or invoke them, the failure is printed
to the Immediate window by `cCmsBatch`.

Yes — the sync wrappers write to the store too: they run the same async
parse path under a `WaitAll` pump, so a successful `CMS_GetCurveQuoteData`
leaves the curve in `CMS_Store()` (until the next state reset clears it).

## Failure semantics

- Transport errors (timeout, non-2xx, SOAP fault, REST FAILURE) fail every
  curve on that request with the reason; other requests are unaffected.
- Curve-level CMS errors arrive as batch `errorMessageContent` keyed by
  fourTuple inside the text (legacy contract) and are mapped per curve.
- A curve with `Status = "OK"` but no points and no recovery is marked FAILED
  ("no data") — same rule the legacy bulk path applied.
- Callbacks never break the batch: exceptions are swallowed and logged to the
  Immediate window. Set `modBridge.gHttpVerbose = True` for request tracing.

## Things intentionally dropped from the legacy add-in

User authorization gating (`GetCMSUsersByQuery`), impersonation UI, menus,
registry/session plumbing, Outlook mail, ad-hoc tenor arrays, 11/12/13-tenor
standards, `SetFlat`/recovery-only/term-quote single-cell UDF variants, and
the EXTENDED two-row mode. GETs are always 17-standard, SETs always the
12-standard, per this workbook's usage.
