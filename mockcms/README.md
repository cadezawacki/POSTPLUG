# Mock CMS Server

A local, async stand-in for the production **CreditMarkingService** that the
ExcelBridge CMS engine (`modCms.bas`) talks to. It lets you develop and test
GET/SET flows against a server that behaves like CMS — without marking any real
curves.

It speaks the **exact wire contract** the VBA client expects, on both
transports:

| Transport | What the client does | What this server accepts |
|---|---|---|
| **SOAP** (default) | POSTs an rpc/encoded SOAP 1.1 envelope to `…/CreditMarkingService/` with `<string>` = user-metadata and `<string0>` = Rosetta XML | any POST path (e.g. `/CreditMarkingService/`) |
| **REST** | POSTs JSON `{requestHeader, request}` carrying the same Rosetta | `POST /api/v1/getMarketData`, `POST /api/v1/setMarketData` |

GET vs SET is decided by the Rosetta `<action>` element, exactly as the real
service routes it.

## What it simulates

- **First GET of an unknown curve** → returns a freshly generated,
  random-but-realistically-shaped CDS curve: strictly positive, upward sloping
  across the 17-standard grid, with the **5Y as a rounded integer in bps**. The
  curve is then cached, so repeat GETs are stable within a server session.
- **SET** → stores the marked levels and echoes them back **verbatim** (just
  like production). The next GET reflects them — so you can mark and re-mark all
  day without touching real data.
- **Units** match `modCms.bas` exactly: GET `creditSpread × 10000 → bps`,
  `creditUpfront × 100 → %`; SET values are verbatim. The server converts in
  both directions so values round-trip cleanly.
- Per-curve metadata (`owner`, `marked`, `lastMarkedBy`, `isParent`,
  `latestCurveMarkedType`), a recovery point (default 0.40), and the
  `statusCode = OK` envelope are all emitted in the shape
  `ParseRosettaResponse` walks.
- **Async / concurrent**: built on `aiohttp`, so many GET/SET requests are
  served concurrently on one event loop. An optional simulated latency makes
  the parallelism behave like the real network.

State is **in-memory only** — restart the server for a clean slate. Curves are
keyed by `(label, tag, date)`, so LIVE and historical (`NYOISCLOSE`) snapshots
are independent.

## 1. Install dependencies

Requires **Python 3.8+**. The only third-party dependency is `aiohttp`.

```bash
cd mockcms

# (recommended) isolated environment
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install -r requirements.txt  # just: aiohttp>=3.9
```

## 2. Boot the server

```bash
python server.py                 # SOAP + REST on http://0.0.0.0:8080
```

Options:

```
--host HOST        bind host (default 0.0.0.0)
--port PORT        bind port (default 8080)
--latency SECONDS  max simulated per-request latency (default 0.05; 0 disables)
--seed N           RNG seed, for reproducible generated curves
-q / --quiet       log warnings only
```

Examples:

```bash
python server.py --port 25551                 # match a REST-style host:port
python server.py --latency 0                   # no artificial delay
python server.py --seed 42                      # deterministic curves
```

You'll see a request log line per call, e.g.:

```
15:05:54  INFO   generated new curve  AEP.USD.SENIOR_NORE_14.SNAC100  (5Y=44.0 creditSpread)
15:05:54  INFO   SOAP GET   1 curve(s)  ->  getMarketData
```

Health check: `curl http://localhost:8080/health`

## 3. Point the Excel/VBA client at it

In the VBA Immediate window (or `Workbook_Open`):

```vb
' SOAP transport (default) - note the trailing slash, like the real endpoint:
CMS_Configure "http://localhost:8080/CreditMarkingService/", "SOAP"

' or REST transport - pass the BASE url; the client appends /api/v1/...:
CMS_Configure "http://localhost:8080", "REST"

?CMS_EndpointUrl, CMS_Transport
```

Then use the engine exactly as in production:

```vb
?CMS_GetCurveQuoteData("AEP","USD","SENIOR_NORE_14","SNAC100","LIVE",Now(),"QUOTED_SPREADS")
Set b = CMS_SetCurve5yAsync("AEP", 127.5)
```

> Tip: this server ignores the host in the URL and serves SOAP on **any** POST
> path, so any `http://localhost:PORT/anything/` works for the SOAP transport.

## 4. Smoke test (no Excel needed)

With the server running, run the included async test client. It builds
byte-compatible SOAP envelopes, fires several GETs concurrently, does a SET, and
confirms the mark round-trips through a subsequent GET:

```bash
python test_client.py                                   # default localhost:8080
python test_client.py http://localhost:25551/CreditMarkingService/
```

Expected output:

```
=== concurrent first GETs (random curves generated) ===
AEP.USD.SENIOR_NORE_14.SNAC100          5Y=  44.00  10Y=  51.28  1Y=  18.32  (bps)
...
=== SET a new curve on AEP (mark 5Y=200) ===
SET echo                                5Y= 200.00  10Y= 226.00  (bps, verbatim)
=== GET AEP again - reflects the mark ===
AEP.USD.SENIOR_NORE_14.SNAC100          5Y= 200.00  10Y= 226.00  (bps)
OK - SET round-tripped through a subsequent GET.
```

## Files

| File | Purpose |
|---|---|
| `server.py` | the mock server (SOAP + REST, async, stateful) |
| `test_client.py` | async smoke-test client mimicking the VBA wire format |
| `requirements.txt` | `aiohttp` |
