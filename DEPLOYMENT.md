# ExcelBridge 2.0 — XLL Deployment Guide

ExcelBridge is now an **Excel-DNA XLL add-in**. There is no COM registration:
no `regasm`, no `.tlb`, no registry, no admin rights. Deployment is copying one
file.

## 1. Build

Requires the .NET SDK (any recent version; the target is `net48`).

```
dotnet build ExcelBridge.csproj -c Release
```

Outputs (in `bin\Release\net48\`):

| File | Use |
|---|---|
| `ExcelBridge-AddIn64-packed.xll` | **Ship this** for 64-bit Office (self-contained: ExcelBridge.dll + Newtonsoft.Json packed inside) |
| `ExcelBridge-AddIn-packed.xll` | Same, for 32-bit Office |
| `ExcelBridge-AddIn64.xll` / `ExcelBridge-AddIn.xll` | Unpacked variants (need the DLLs beside them) — dev use |

You may rename the packed file to `ExcelBridge-AddIn64.xll`; the VBA loader
(`modBridge.EnsureAddin`) probes both names.

## 2. Roll out

Per user/desktop, choose one:

- **Zero-touch (recommended):** put the XLL in the same folder as the workbook.
  `Workbook_Open` → `http_start` → `EnsureAddin` finds and loads it via
  `Application.RegisterXLL` automatically.
- **Central share:** put the XLL on a network path and set that full path in a
  workbook named range called `BridgeXllPath`.
- **Persistent add-in:** File → Options → Add-ins → Go… → Browse to the XLL
  (loads in every Excel session; `EnsureAddin` then no-ops).

No installer, no admin, versioning is file-replacement (close Excel first).

## 3. Workbook (VBA project) migration from 1.x

1. **Remove** the `cBridgeHost` class module.
2. **Remove** the `ExcelBridge` reference (Tools → References) — no longer used.
3. **Re-import** the updated `modHttp.bas`, `modBridge.bas`.
4. **Import** the new `modBridgeEvents.bas`.
5. Everything else is unchanged: `cHttpBatch`, `cHttpResponse`,
   `modSocketRouter`, `IntersectionObserver`, `ThisWorkbook` handlers,
   `cadesHelpers`.

Public VBA API is unchanged: `HttpPost`/`HttpGetSync`/`cHttpBatch`/
`StartBridge`/`StopBridge`/`EnsureHttp` all keep their signatures.

## 4. How it works now

```
VBA  --Application.Run("EB_*")-->  XLL (BridgeApi)  -->  BridgeEngine (HttpClient / ClientWebSocket)
VBA  <--Application.Run sink----  VbaDispatcher (QueueAsMacro, main thread, batched)
```

- **HTTP needs no socket.** `EnsureHttp` loads + attaches the XLL; only
  `StartBridge` opens the WebSocket.
- **Callbacks are delivered on Excel's main thread** via Excel-DNA's
  `QueueAsMacro`. The old failure mode (busy Excel rejecting cross-thread COM
  events → lost HTTP completions) is structurally gone.
- **HTTP completions are coalesced**: a burst of parallel POST completions
  arrives as ONE `EB_OnHttpBatch` call with a 2-D array, not N separate
  event transitions.
- The XLL delivers callbacks to **one attached workbook** per Excel session
  (`EB_Attach`, called by `EnsureHttp`; last attach wins). `bridge_stop`
  detaches on close.

### Callback contract (fixed names in `modBridgeEvents`)

| Sink | Payload |
|---|---|
| `EB_OnHttpBatch(results)` | 2-D rows: id, status, body, headersJson, elapsedMs, errorMsg |
| `EB_OnMessage(type, json)` | generic WS message (when enabled) |
| `EB_OnMessageBatch(jsonArray)` | coalesced WS messages |
| `EB_OnCellUpdate / EB_OnCellBatch / EB_OnFullGrid` | typed grid feed |
| `EB_OnWsStatus(state, detail)` | connected / disconnected / reconnecting |
| `EB_OnLog(level, message)` | diagnostics |

### String-size limits (the one XLL constraint)

Excel's evaluator caps **VBA → XLL** string arguments at 32,767 chars.
`modHttp` handles this transparently:

- Request bodies > 30,000 chars are staged via `EB_BodyBegin`/`EB_BodyAppend`.
- *Sync* response bodies > 30,000 chars are fetched via `EB_ResultChunk`.
- The **async** path (XLL → VBA) goes over COM `Application.Run` and has no
  such cap — large response bodies arrive whole.

## 5. Troubleshooting

- `?Application.Run("EB_Version")` in the Immediate window → `2.0.0` means the
  XLL is loaded.
- `?Application.Run("EB_LastDispatchError")` → last failed callback dispatch
  (e.g. a renamed sink or a VBA error inside one).
- `modBridge.gHttpVerbose = True` → per-completion `Debug.Print` + HttpLog
  sheet logging (off by default; it is the slow path).
- WS state: `?Application.Run("EB_WsIsRunning")`, `?Application.Run("EB_WsIsConnected")`.
