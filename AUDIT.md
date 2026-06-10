# POSTPLUG Codebase Audit — Async HTTP/WS Bridge

Scope: `Bridge.cs` (COM bridge), `modHttp.bas`, `modBridge.bas`, `cBridgeHost.cls`,
`cHttpBatch.cls`, `cHttpResponse.cls`, `cHttpCallback.cls`, `modSocketRouter.bas`,
`socket_filter.bas`, plus the HTTP-relevant parts of `cadesHelpers.bas` and the
workbook lifecycle in `ThisWorkbook.bas`. Functions referenced but not present in
the repo (`User_GetUserName`, `ExtractTimeEastern`, `frmProgress`, `sdrFilter.*`,
`clear_hub`, etc.) are assumed to exist elsewhere and were not flagged.

Items marked **[FIXED]** were changed in this commit. Items marked **[NOTE]** are
recommendations or accepted risks left for a follow-up.

---

## Pass 1 — Deep Bug Scan (correctness, races, leaks, silent corruption)

### 1.1 [FIXED] Silent loss of HTTP completions when Excel rejects the COM call
`Bridge.Raise()` swallowed *every* exception from the event sink. Events are
raised from threadpool threads into Excel's STA; whenever the user is editing a
cell or a modal dialog is open, COM rejects the incoming call with
`RPC_E_CALL_REJECTED` / `RPC_E_SERVERCALL_RETRYLATER`. The old code ate that
exception, so an `OnHttpResponse` could vanish: the batch's `mPending` never
decremented and `WaitAll` hung until its timeout. **This was the most likely
root cause of "requests randomly never complete."** `Raise` now retries those
two specific HRESULTs with backoff (up to ~2.7 s) before logging and giving up.

### 1.2 [FIXED] `Start()`/`Stop()` race — two connect loops fighting over `_ws`
`Start()` did not wait for (or fence) the previous connect loop. The old loop's
`finally` block did `_ws?.Dispose(); _ws = null;`, which could dispose and null
out the **new** loop's freshly connected socket — a classic cause of
reconnect flapping ("reconnections are causing issues"). Fixes:
- a `_generation` counter incremented on every `Start`/`Stop`; stale loops
  detect it and exit without touching shared state;
- each loop uses a *local* `ClientWebSocket` and clears the shared `_ws` field
  with `Interlocked.CompareExchange` only if it still owns it;
- `ReceiveLoopAsync`/`FlushOutboxAsync` operate on a captured socket reference
  instead of re-reading the mutable `_ws` field (which could become null
  mid-reconnect and throw `NullReferenceException` inside the send path).

### 1.3 [FIXED] `EnsureBridge` orphaned in-flight HTTP requests
When the WS loop had exited, `EnsureBridge` did `Set gHost = Nothing` and built
a brand-new host + Bridge object. Any HTTP request still in flight would raise
its completion on the **old, unreferenced** Bridge — the event was lost forever
and batches hung. `EnsureBridge` now restarts the WS loop **in place** on the
same Bridge object (`gHost.StartWs`), preserving the event sink.

### 1.4 [FIXED] Bridge restart wiped all registered socket routes
`StartBridge` unconditionally called `modSocketRouter.InitRouter`, so every
auto-restart (via `EnsureBridge`) destroyed all `RouteOnValue`/`RouteOnPresence`
registrations — messages then silently fell through to nothing. Now uses an
idempotent `EnsureRouter`; `InitRouter` remains available for an explicit reset.

### 1.5 [FIXED] Duplicate heartbeat `OnTime` chains
`StartHeartbeat` scheduled a new `Application.OnTime` chain without cancelling
the pending one, and `StopHeartbeat` can only cancel the *latest*
`gHeartbeatNextRun`. Every `StartBridge` re-entry therefore multiplied ping
chains (and each orphan chain re-armed itself forever, also keeping the
workbook "busy"). `StartHeartbeat` now cancels the pending tick first.

### 1.6 [FIXED] HTTP timeout did not cover the response body (net48)
`DoHttpAsync` used `HttpCompletionOption.ResponseHeadersRead` and then read the
body with `ReadAsByteArrayAsync()`, which does not observe the cancellation
token on .NET Framework. A server that returned headers and then trickled the
body bypassed the timeout entirely, leaking the request task and its `Task.Run`
thread slot. The body read is now aborted via `ct.Register(... resp.Dispose())`
and surfaces as a proper timeout.

### 1.7 [FIXED] Response bodies decoded as UTF-8 unconditionally
Manual `Encoding.UTF8.GetString(bytes)` corrupted any non-UTF-8 response
(e.g. `charset=iso-8859-1`). Replaced with `ReadAsStringAsync()`, which honors
the response charset. (Also removes one full buffer copy — see Pass 2.)

### 1.8 [FIXED] Unbounded WS outbox
`EnqueueSend` queued forever while disconnected — a slow/never reconnect grew
`_outbox` without limit (memory leak with a side of multi-minute replay storms
on reconnect). Now capped at 10,000 entries; overflow drops the oldest message
and raises `OnError` so the drop is visible.

### 1.9 [FIXED] `cHttpBatch.WaitAll` breaks across midnight
`Timer` wraps at midnight, so `(Timer - startTick)` went negative and the
overall timeout never fired (or fired instantly). Elapsed time now corrects for
rollover.

### 1.10 [FIXED] `cHttpBatch.Init` did not reset `mSimpleCallback`
Re-`Init`-ing a reused batch object kept the previous simple callback firing.

### 1.11 [FIXED] `cHttpCallback.Handle(r As HttpResponse)` — wrong type
The class is `cHttpResponse`; with no `HttpResponse` type in the project this
fails to compile, and even if such a type exists elsewhere, `cHttpBatch` passes
a `cHttpResponse` via `CallByName` → runtime type mismatch. Fixed (and the
matching doc comments in `cHttpBatch`/`modHttp`, including a missing `Set` in
the usage example).

### 1.12 [FIXED] `Application.enableEvents` clobbered on the HTTP log path
`b_OnHttpResponse` set `enableEvents = True` unconditionally on exit, silently
re-enabling events that an outer caller had deliberately disabled. Now
saves/restores the prior value.

### 1.13 [FIXED] `socket_filter`: `getaddrinfo` chain never freed; dead code
Every `Login`/`ConnectToWs` leaked the addrinfo allocation (`freeaddrinfo` was
never declared or called). Also removed the unreachable statements after
`Exit Function` in `ConnectToWs`'s failure branch (the `ConnectToWs = False`
assignments never executed).

### 1.14 [NOTE] `socket_filter` remaining issues (legacy module — recommend deprecation)
- `Login` declares a local `CONNECTED` that shadows the module-level flag, so
  `am_i_connected()` is wrong after `Login`.
- `Send` is declared `ByVal buf As String` → implicit ANSI conversion; any
  non-ASCII payload is corrupted on the wire. Use a byte-array overload.
- `WSAStartup`/`WSACleanup` pairing is unbalanced across error paths
  (`WSACleanup` decrements a refcount; mismatches can kill Winsock for the
  whole process).
- `MAX_BUF_SIZE` is set but unused; module-level `retVal` shared across calls.
This module overlaps zero percent with the `ExcelBridge` WS stack; if it only
serves the legacy webhook login, isolate or retire it.

### 1.15 [NOTE] `b_OnFullGrid` splits rows on the literal `"],["`
Any cell value containing `],[` corrupts the whole grid parse. Acceptable if
the server escapes strings (the splitter handles `\"` correctly), since `],[`
inside a JSON string would appear escaped — but a raw `],[` in a value is
passed through verbatim by `SplitJsonStrings`. Worth a server-side guarantee or
a depth-aware split like `RouteBatch` uses.

### 1.16 [NOTE] `ExtractJsonValue` matches keys at any depth
`"key":` is found anywhere in the payload, including inside nested objects and
inside *string values*. For the short, flat, server-controlled messages in use
this is fine; do not feed it arbitrary JSON.

### 1.17 [NOTE] `cHttpBatch.OnComplete` swallows callback errors
`On Error Resume Next` around `CallByName`/`Application.Run` means a buggy user
callback fails silently. Deliberate (a throwing callback must not poison the
event pump), but consider routing to a visible error sink.

---

## Pass 2 — Memory Optimization

### 2.1 [FIXED] `SplitJsonStrings` O(n²) string append
`buf = buf & ch` reallocates the accumulator once per character — for a
200-column row that's tens of thousands of allocations per row, per grid push.
Rewritten with a single preallocated scratch buffer written via `Mid$`
assignment: one allocation per row + one per emitted cell. This is the
dominant VBA allocation site on the full-grid path (easily >90% reduction
there).

### 2.2 [FIXED] Double-buffering of every HTTP response body
`ReadAsByteArrayAsync` + `Encoding.UTF8.GetString` held two full copies of
every body (bytes + string). `ReadAsStringAsync` decodes the stream without
retaining the intermediate byte array.

### 2.3 [FIXED] Re-serialization of every non-typed WS message
`DispatchObject` always did `root.ToString(Formatting.None)` — a full second
copy of every payload that Newtonsoft had just parsed. When the wire text is
already compact (the normal case for machine-generated JSON), the original
string is now passed through untouched; re-serialization only happens when the
payload contains formatting that VBA's compact-JSON scanners can't handle.

### 2.4 [FIXED] Unbounded outbox retention (see 1.8).

### 2.5 [NOTE] `cHttpBatch.mResults` retains every body until the batch dies
By design (`GetResult` needs them), but for large fan-outs where the per-request
callback already consumed the body, results could be released eagerly. If you
fire 500 POSTs with 1 MB responses, the batch pins ~500 MB until it goes out of
scope. Consider an opt-in `DiscardBodiesAfterCallback` flag.

### 2.6 [NOTE] Dispatch parses with `JToken.Parse` even for typed messages
A hand-rolled `type` sniff (like VBA's `ExtractJsonValue`) before full parse
would skip Newtonsoft's token tree entirely for messages that go straight to
`EnqueueBatch`. Combined with 2.3 the batch path would become zero-parse on the
.NET side. Left undone because the typed handlers (`cell`, `cell_batch`,
`full`) genuinely need the parse and correctness of the sniff must be proven
first.

### 2.7 [NOTE] `JsonToDict`/`JsonKeys` rescan the payload per key
`JsonToDict` is O(keys × len). Fine for short messages; avoid on large
payloads — `cadesHelpers.JSON_ToObject` is the proper parser there.

---

## Pass 3 — CPU / Throughput Optimization

### 3.1 [FIXED] Per-response sheet logging on the hot path
`b_OnHttpResponse` did an `xlUp` scan over the entire log column plus five
single-cell writes and a `Debug.Print` of every body **per completion** —
each one a cross-thread COM round-trip into Excel plus a sheet recalc risk.
For "numerous parallel POSTs" this serialized completions behind sheet I/O.
Diagnostics are now opt-in via `modBridge.gHttpVerbose` (default off). With
logging off, a completion is: construct `cHttpResponse` → dictionary lookup →
user callback. (Order-of-magnitude reduction in per-completion overhead.)

### 3.2 [FIXED] Reconnect storms burning CPU and connections (see 1.2/1.3/1.5)
Duplicate connect loops, duplicate heartbeat chains, and full-object teardown
per `EnsureBridge` all multiplied connection work. Single-loop ownership +
restart-in-place removes that entire class of churn.

### 3.3 [NOTE] HTTP parallelism is already sound on the .NET side
`HttpSendAsync` is truly fire-and-forget: static shared `HttpClient`,
`DefaultConnectionLimit = 100`, Nagle off, Expect100 off, per-request CTS,
completions on threadpool. The throughput ceiling is *delivery into Excel*
(one STA → all events serialized through the message pump), not the I/O. See
4.2 for the structural answer.

### 3.4 [NOTE] `WaitAll`'s `DoEvents` + `Sleep 5` poll
Adequate; `DoEvents` is what actually delivers the events. Reducing the sleep
helps latency marginally but raises CPU. Prefer the new
`SetAllCompleteCallback` (fully async, no pump loop at all) wherever the caller
doesn't truly need to block.

### 3.5 [NOTE] `RouteBatch`/`ExtractJsonValue` scan with per-character `Mid$`
Each `Mid$(s, i, 1)` allocates a 1-char string. A `Byte()` array scan
(`StrConv`) would be ~10× faster on large batches. Worth doing only if batch
payloads exceed a few hundred KB; current message sizes don't justify the
complexity.

### 3.6 [FIXED] `cadesHelpers.HttpRequest` (WinHttp, synchronous) removed
It duplicated `modHttp`, blocked Excel's UI thread for the full request
duration, and divided the timeout by 4 naively. Deprecated and deleted; it had
no callers in the repo. All HTTP now goes through the single
`modHttp`/`cHttpBatch` stack (use `HttpRequestSync` where a blocking call is
genuinely needed). Its private helpers (`AsString`, `NewDictionary`,
`JSON_Stringify`, `JSON_ToObject`) remain in use by other utilities.

---

## Pass 4 — Architectural Wins

### 4.1 [FIXED] HTTP lifecycle decoupled from WS lifecycle (see also Pass 5)
`cBridgeHost.StartUp` was split into `EnsureObject` (COM object only — HTTP
ready) and `StartWs` (socket loop). `modBridge.EnsureHttp` is the new HTTP-only
entry point. This is the structural change that makes everything in Pass 5
true.

### 4.2 [NOTE — biggest remaining win] Coalesce HTTP completions into batched events
Every completion is its own cross-thread COM call into the STA (~0.1–1 ms+
each, serialized, worse when Excel is busy). The bridge already proves the
pattern with `OnMessageBatch`/`_batchTimer`: add an `OnHttpResponseBatch(object
results)` raised on the same 16 ms window, delivering an `object[,]` of
`{id, status, body, headersJson, elapsedMs}`. VBA routes the whole array in
one event. For 200 parallel POSTs that's 200 STA transitions → ~12. Not done
here because it extends the COM event interface (new DispId ⇒ typelib
regeneration and VBA re-reference) — mechanical, but should be its own change.

### 4.3 [NOTE] One `Bridge` class carries three concerns
WS transport, HTTP client, and batch buffering live in one ~800-line COM class.
The HTTP side is already `static` — it could be extracted to an `ExcelHttp`
co-class behind its own interface, letting WS-free workbooks reference only
that. Defer until 4.2 lands (it touches the same interface).

### 4.4 [NOTE] String-shaped JSON as the universal interchange
.NET parses JSON → re-serializes → VBA re-scans character-by-character. For the
typed paths the bridge already does the right thing (`OnCellBatch` ships an
`object[,]`). Extend that idea: any high-volume message family deserves a typed
handler in `_handlers` + a typed COM event, leaving the string router for
low-volume traffic only.

### 4.5 [NOTE] `ThisWorkbook.Workbook_Open` runs the error path unconditionally
`On Error GoTo ErrorHandler` with no `Exit Sub` before the label means the
"handler" is really a finally-block. It works, but rename the label (`Finale:`)
or add the conventional structure before someone "fixes" it into a real bug.

---

## Pass 5 — Independence (HTTP without WS) & Reconnection Health

### 5.1 [FIXED] REST no longer requires — or even touches — the WebSocket
Before: `HttpPost(...)` → `HttpSendFF` → `EnsureBridge` → `StartBridge` →
`b.Start "ws://cds-sn-api-dev..."`. Every REST call from a fresh session opened
(or worse, endlessly retried) a socket to the hard-coded dev endpoint, spun up
the batch timer, the heartbeat, and the reconnect loop — to send one POST.

After: `HttpSendFF`/`HttpRequestSync` call `modBridge.EnsureHttp`, which only
instantiates the COM object. On the .NET side `HttpSendAsync`/`HttpSendSync`
were verified to share no state with the socket (`_http` is static; no method
on the HTTP path reads `_ws`, `_cts`, `_outbox`, or the batch timer). You can
hit a REST endpoint with the socket never started, stopped, or mid-reconnect.

The reverse also holds: `StartBridge`/WS messaging never invokes the HTTP path.
The two libraries now meet only at the shared COM object and event sink.

### 5.2 [FIXED] Reconnection no longer destabilizes HTTP (or itself)
The specific reconnect pathologies found and fixed:
- stale connect loop disposing the new loop's socket (1.2) — flapping;
- `EnsureBridge` tearing down the host mid-flight (1.3) — lost HTTP events;
- route table wipe on every restart (1.4) — silent message loss after recovery;
- heartbeat chain multiplication on every restart (1.5) — ping storms;
- unbounded outbox replay after long outages (1.8).
Remaining behavior: while `IsRunning` and reconnecting, `EnsureBridge` leaves
the loop alone (correct); backoff is exponential, capped by
`MaxReconnectDelayMs`, with jitter (correct).

### 5.3 [NOTE] `cHttpBatch` ergonomics for the parallel-POST use case
Added `SetAllCompleteCallback` so the "N POSTs → one callback when all done"
pattern is fully event-driven (no `WaitAll` pump needed):

```vba
Dim batch As New cHttpBatch
batch.Init
batch.SetSimpleCallback "modMyMod.OnOnePost"        ' optional: per-request
batch.SetAllCompleteCallback "modMyMod.OnAllPosts"  ' fired once at the end
Dim i As Long
For i = 1 To 50
    batch.AddPost endpoints(i), bodies(i)
Next i
' return immediately; OnAllPosts(batch) fires when the last one lands
```

Note the batch must stay referenced (module-level variable) until completion —
if it's a local that goes out of scope, `modBridge.gBatchRouter` still holds it
(strong ref), so results are not lost, but keep a reference if you intend to
read `GetResult` afterwards.

### 5.4 [FIXED] Startup no longer opens a socket; observer zones live from open
`Workbook_Open` called `bridge_start` → `StartBridge` → WS connect loop, so the
app could never run HTTP-only. It now calls `http_start` (`EnsureHttp` +
`SetupObserver`); `bridge_start` remains for explicitly enabling the socket
feed. The heartbeat is only armed by `StartBridge`, so HTTP-only mode has no
timer that could indirectly trigger a socket connect.

Separately, observer zones were registered **only** inside
`Workbook_SheetChange` — selection-driven callbacks (the intersection → async
POST flow) did nothing until the first cell *edit* on a `cms_col` sheet, and
`SetupObserver` was an empty stub. Registration now lives in one routine
(`RegisterObserverZones`), called from startup, `Workbook_SheetActivate`
(zones are bound per sheet name and previously kept pointing at the old sheet
until an edit), and `Workbook_SheetChange` as before. `IntersectionObserver`
gained an idempotent `EnsureInitialized`, and the range cache upserts instead
of throwing on duplicate registration.

### 5.5 [NOTE] The hard-coded WS URL lives in code
`BridgeWsUrl()` (extracted in this change) still embeds the dev endpoint.
Move it to a named range / Settings sheet alongside `SDR.Host` /
`Webhook.Port`, which already follow that pattern.
