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

### Store (curves in memory, cells optional)

Every successful GET upserts `CMS_Store()` (key = `TICKER.CCY.DEBTCLASS.PRODUCT`);
every successful SET merges the echo over the computed curve. Sheet output is
opt-in:

```vb
CMS_StoreToArray()                          ' everything, with header
CMS_WriteBatchToRange b, [A1], True         ' a batch, row-per-input-row
StoreGet("IBM.USD.SENIOR.DERIV").Quote("5Y")
```

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
