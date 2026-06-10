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
(`modBridge.EnsureAddin`) probes the packed name first, then the plain name.

> **Ship ONLY the packed file.** The unpacked XLL loads `ExcelBridge.dll` from
> disk and resolves `Newtonsoft.Json.dll` lazily — if either is missing beside
> it, `EB_Version` still works and the failure surfaces later, on the first
> HTTP call ("Could not load file or assembly 'Newtonsoft.Json…'").
> `?Application.Run("EB_Diag")` reports the loaded XLL path and whether the
> dependency resolves.

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

### Entry-point argument types

All `EB_*` HTTP entry points take coercive `object` parameters, so any VBA
type that sensibly converts is accepted (String/Long/Double/Boolean; omitted
arguments fall back to defaults). The canonical call shape:

```vba
' returns requestId (String); timeoutMs 0 = use default
id = Application.Run("EB_HttpSendAsync", method$, url$, body$, headers$, timeoutMs&)
```

Failures return the string `"#ERR <reason>"` (async) or an error message in
the result row (sync) — `modHttp` converts both into descriptive VBA errors.
A raw **"Type mismatch" (13)** when assigning a `Run` result means the
function itself returned an Excel error value: almost always a stale 1.x XLL
still loaded, or a renamed/missing entry point. Check
`?Application.Run("EB_Version")` first.

### String-size limits (the one XLL constraint)

Excel's evaluator caps **VBA → XLL** string arguments at 32,767 chars.
`modHttp` handles this transparently:

- Request bodies > 30,000 chars are staged via `EB_BodyBegin`/`EB_BodyAppend`.
- *Sync* response bodies > 30,000 chars are fetched via `EB_ResultChunk`.
- The **async** path (XLL → VBA) goes over COM `Application.Run` and has no
  such cap — large response bodies arrive whole.

## 5. Troubleshooting

**First stop:** run `modBridge.BridgeSelfTest "https://your-endpoint"` in the
Immediate window — it checks load state, version, attach, and both HTTP paths
with readable output.

**Error 424 "Object required" (or 1004) on `Application.Run("EB_*", ...)`**
means the name didn't resolve to a registered function: the XLL is **not
loaded in this Excel instance**. Causes, in rough order of likelihood:

1. Nothing loaded it yet — calling `EB_*` directly via `Run` bypasses
   `EnsureAddin`. Go through `modHttp`/`http_start`, or load explicitly:
   `?Application.RegisterXLL("C:\full\path\ExcelBridge-AddIn64-packed.xll")`
   (must print `True`).
2. **Mark-of-the-Web**: an XLL copied from a share/browser/email is blocked by
   Windows and Excel refuses it *silently* (`RegisterXLL` returns `False`).
   Right-click the file → Properties → **Unblock** (or
   `Unblock-File` in PowerShell). The most common rollout failure.
3. **Bitness mismatch**: 64-bit Office needs `ExcelBridge-AddIn64*.xll`; the
   32-bit flavor loads as nothing. `?Application.OperatingSystem` ending in
   `64-bit` = use the 64 file.
4. The XLL was double-clicked and opened in a **different Excel instance**
   than the one hosting your workbook.

Other checks:

- `?Application.Run("EB_Version")` → `2.0.0` means the correct build is loaded.
- `?Application.Run("EB_LastDispatchError")` → last failed callback dispatch
  (e.g. a renamed sink or a VBA error inside one).
- `modBridge.gHttpVerbose = True` → per-completion `Debug.Print` + HttpLog
  sheet logging (off by default; it is the slow path).
- WS state: `?Application.Run("EB_WsIsRunning")`, `?Application.Run("EB_WsIsConnected")`.
