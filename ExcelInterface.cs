
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Threading;
using ExcelDna.Integration;

namespace ExcelBridge
{
    // =========================================================
    //  Add-in lifecycle
    // =========================================================
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen() { }

        public void AutoClose()
        {
            try { BridgeApi.Shutdown(); } catch { }
        }
    }

    // =========================================================
    //  Entry points callable from VBA via Application.Run.
    //
    //  Contract notes:
    //  - VBA -> XLL string arguments are capped at 32,767 chars by
    //    Excel's evaluator (XLOPER12). Large request bodies are
    //    staged with EB_BodyBegin/EB_BodyAppend; large sync response
    //    bodies are fetched with EB_ResultChunk. The async callback
    //    path (XLL -> VBA via COM Application.Run) has no such cap.
    //  - All callbacks are delivered to ONE attached workbook
    //    (EB_Attach), into fixed-name sinks in modBridgeEvents.
    // =========================================================
    public static class BridgeApi
    {
        internal static readonly BridgeEngine Engine = new BridgeEngine();
        private static int _wired;

        // Keep each inline string comfortably under the 32,767 XLOPER12 cap.
        internal const int ChunkMax = 30000;

        private static readonly ConcurrentDictionary<string, StringBuilder> _stagedBodies
            = new ConcurrentDictionary<string, StringBuilder>();
        private static readonly ConcurrentDictionary<string, Tuple<string, DateTime>> _stagedResults
            = new ConcurrentDictionary<string, Tuple<string, DateTime>>();

        internal static void Shutdown()
        {
            Engine.Stop();
            VbaDispatcher.Detach();
        }

        private static void EnsureWired()
        {
            if (Interlocked.CompareExchange(ref _wired, 1, 0) != 0) return;

            Engine.OnHttpResponse += (id, status, body, headers, ms) =>
                VbaDispatcher.EnqueueHttp(id, status, body, headers, ms, "");
            Engine.OnHttpError += (id, msg, ms) =>
                VbaDispatcher.EnqueueHttp(id, 0, "", "{}", ms, msg);

            Engine.OnMessage += (t, j) => VbaDispatcher.EnqueueMacro("EB_OnMessage", t, j);
            Engine.OnMessageBatch += j => VbaDispatcher.EnqueueMacro("EB_OnMessageBatch", j);
            Engine.OnCellUpdate += (r, c, v) => VbaDispatcher.EnqueueMacro("EB_OnCellUpdate", r, c, v);
            Engine.OnCellBatch += cells => VbaDispatcher.EnqueueMacro("EB_OnCellBatch", cells);
            Engine.OnFullGrid += g => VbaDispatcher.EnqueueMacro("EB_OnFullGrid", g);

            Engine.OnConnected += () => VbaDispatcher.EnqueueMacro("EB_OnWsStatus", "connected", "");
            Engine.OnDisconnected += r => VbaDispatcher.EnqueueMacro("EB_OnWsStatus", "disconnected", r ?? "");
            Engine.OnReconnecting += (a, d) => VbaDispatcher.EnqueueMacro("EB_OnWsStatus", "reconnecting", a + "|" + d);
            Engine.OnError += m => VbaDispatcher.EnqueueMacro("EB_OnLog", "error", m ?? "");
            Engine.OnLog += (l, m) => VbaDispatcher.EnqueueMacro("EB_OnLog", l ?? "info", m ?? "");
        }

        // ---------- lifecycle / diagnostics ----------

        [ExcelFunction(IsHidden = true)]
        public static string EB_Version() { return "2.0.0"; }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_Attach(string workbookName)
        {
            EnsureWired();
            VbaDispatcher.Attach(workbookName);
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_Detach()
        {
            VbaDispatcher.Detach();
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static string EB_LastDispatchError() { return VbaDispatcher.LastError; }

        // ---------- HTTP ----------
        //
        // Entry points reached via Application.Run take object parameters and
        // coerce: a typed int/string parameter turns an omitted argument
        // (ExcelMissing) into a silent #VALUE!, which VBA then surfaces as a
        // bare "Type mismatch" on assignment. Failures return "#ERR <reason>"
        // so the VBA wrapper can raise something actionable instead.

        [ExcelFunction(IsHidden = true)]
        public static string EB_HttpSendAsync(object method, object url, object body, object headers, object timeoutMs)
        {
            try
            {
                EnsureWired();
                string u = ToStr(url);
                if (u.Length == 0) return "#ERR url is required";
                return Engine.HttpSendAsync(ToStr(method), u, ToStr(body), ToStr(headers), ToInt(timeoutMs));
            }
            catch (Exception ex) { return "#ERR " + ex.GetBaseException().Message; }
        }

        [ExcelFunction(IsHidden = true)]
        public static string EB_HttpSendAsyncBody(object method, object url, object bodyToken, object headers, object timeoutMs)
        {
            try
            {
                EnsureWired();
                string u = ToStr(url);
                if (u.Length == 0) return "#ERR url is required";
                return Engine.HttpSendAsync(ToStr(method), u, TakeStagedBody(ToStr(bodyToken)), ToStr(headers), ToInt(timeoutMs));
            }
            catch (Exception ex) { return "#ERR " + ex.GetBaseException().Message; }
        }

        [ExcelFunction(IsHidden = true)]
        public static object EB_HttpSendSync(object method, object url, object body, object headers, object timeoutMs)
        {
            try
            {
                EnsureWired();
                return PackSyncResult(Engine.HttpSendSync(ToStr(method), ToStr(url), ToStr(body), ToStr(headers), ToInt(timeoutMs)));
            }
            catch (Exception ex) { return SyncErrorRow(ex); }
        }

        [ExcelFunction(IsHidden = true)]
        public static object EB_HttpSendSyncBody(object method, object url, object bodyToken, object headers, object timeoutMs)
        {
            try
            {
                EnsureWired();
                return PackSyncResult(Engine.HttpSendSync(ToStr(method), ToStr(url), TakeStagedBody(ToStr(bodyToken)), ToStr(headers), ToInt(timeoutMs)));
            }
            catch (Exception ex) { return SyncErrorRow(ex); }
        }

        private static object SyncErrorRow(Exception ex)
        {
            return new object[] { 0, "", "{}", 0, ex.GetBaseException().Message, "", 0 };
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_HttpCancel(string requestId)
        {
            Engine.HttpCancel(requestId);
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_HttpSetDefaultTimeout(object timeoutMs)
        {
            int t = ToInt(timeoutMs);
            if (t > 0) Engine.HttpDefaultTimeoutMs = t;
            return true;
        }

        // ---------- large-string staging (VBA -> XLL) ----------

        [ExcelFunction(IsHidden = true)]
        public static string EB_BodyBegin()
        {
            var token = Guid.NewGuid().ToString("N");
            _stagedBodies[token] = new StringBuilder();
            return token;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_BodyAppend(string token, string chunk)
        {
            StringBuilder sb;
            if (!_stagedBodies.TryGetValue(token, out sb)) return false;
            sb.Append(chunk);
            return true;
        }

        private static string TakeStagedBody(string token)
        {
            StringBuilder sb;
            if (token != null && _stagedBodies.TryRemove(token, out sb)) return sb.ToString();
            return "";
        }

        // ---------- large-string retrieval (XLL -> VBA, sync path only) ----------

        [ExcelFunction(IsHidden = true)]
        public static string EB_ResultChunk(string token, object index)
        {
            Tuple<string, DateTime> entry;
            if (token == null || !_stagedResults.TryGetValue(token, out entry)) return "";
            int start = ToInt(index) * ChunkMax;
            if (start < 0 || start >= entry.Item1.Length) return "";
            int len = Math.Min(ChunkMax, entry.Item1.Length - start);
            return entry.Item1.Substring(start, len);
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_ResultRelease(string token)
        {
            Tuple<string, DateTime> removed;
            if (token != null) _stagedResults.TryRemove(token, out removed);
            return true;
        }

        // Sync results return through Excel's evaluator, so each element is
        // subject to the 32,767-char cap. Bodies that don't fit inline are
        // stashed and fetched chunk-wise by VBA.
        // Row shape: { status, inlineBody, headersJson, elapsedMs, error, bodyToken, bodyLength }
        private static object PackSyncResult(object engineResult)
        {
            var r = (object[])engineResult; // { status, body, headersJson, elapsedMs, error }
            string body = (string)r[1];
            string token = "";
            string inline = body;

            if (body.Length > ChunkMax)
            {
                token = Guid.NewGuid().ToString("N");
                PurgeStaleResults();
                _stagedResults[token] = Tuple.Create(body, DateTime.UtcNow);
                inline = "";
            }

            string headersJson = (string)r[2];
            if (headersJson.Length > ChunkMax) headersJson = headersJson.Substring(0, ChunkMax);

            return new object[] { r[0], inline, headersJson, r[3], r[4], token, body.Length };
        }

        // ---------- argument coercion (Application.Run marshaling) ----------

        private static string ToStr(object v)
        {
            if (v == null || v is ExcelMissing || v is ExcelEmpty || v is ExcelError) return "";
            var s = v as string;
            if (s != null) return s;
            return Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) ?? "";
        }

        private static int ToInt(object v, int fallback = 0)
        {
            if (v == null || v is ExcelMissing || v is ExcelEmpty || v is ExcelError) return fallback;
            if (v is double) return (int)(double)v;
            if (v is int) return (int)v;
            if (v is bool) return (bool)v ? 1 : 0;
            var s = v as string;
            if (s != null)
            {
                int r;
                return int.TryParse(s, out r) ? r : fallback;
            }
            try { return Convert.ToInt32(v); } catch { return fallback; }
        }

        private static bool ToBool(object v, bool fallback = false)
        {
            if (v == null || v is ExcelMissing || v is ExcelEmpty || v is ExcelError) return fallback;
            if (v is bool) return (bool)v;
            if (v is double) return (double)v != 0;
            var s = v as string;
            if (s != null)
            {
                bool r;
                return bool.TryParse(s, out r) ? r : fallback;
            }
            try { return Convert.ToBoolean(v); } catch { return fallback; }
        }

        private static void PurgeStaleResults()
        {
            var cutoff = DateTime.UtcNow.AddMinutes(-10);
            foreach (var kv in _stagedResults)
            {
                if (kv.Value.Item2 < cutoff)
                {
                    Tuple<string, DateTime> removed;
                    _stagedResults.TryRemove(kv.Key, out removed);
                }
            }
        }

        // ---------- WebSocket ----------

        [ExcelFunction(IsHidden = true)]
        public static bool EB_WsStart(string url)
        {
            EnsureWired();
            Engine.Start(url);
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_WsStop()
        {
            Engine.Stop();
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_WsSend(string msgType, string jsonPayload)
        {
            Engine.Send(msgType, jsonPayload ?? "");
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_WsIsConnected() { return Engine.IsConnected; }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_WsIsRunning() { return Engine.IsRunning; }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_WsConfig(object autoReconnect, object maxReconnectDelayMs,
                                       object raiseGenericOnMessage, object batchWindowMs, object batchMaxCount)
        {
            Engine.AutoReconnect = ToBool(autoReconnect, true);
            int v = ToInt(maxReconnectDelayMs);
            if (v > 0) Engine.MaxReconnectDelayMs = v;
            Engine.RaiseGenericOnMessage = ToBool(raiseGenericOnMessage);
            v = ToInt(batchWindowMs);
            if (v > 0) Engine.BatchWindowMs = v;
            v = ToInt(batchMaxCount);
            if (v > 0) Engine.BatchMaxCount = v;
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_PushCell(object row, object col, string value)
        {
            Engine.PushCell(ToInt(row), ToInt(col), value ?? "");
            return true;
        }

        [ExcelFunction(IsHidden = true)]
        public static bool EB_PushCells(object cells)
        {
            Engine.PushCells(cells);
            return true;
        }
    }

    // =========================================================
    //  Main-thread dispatcher: engine events (worker threads) are
    //  queued here and delivered into VBA via QueueAsMacro +
    //  Application.Run — always on Excel's main thread, when Excel
    //  is ready. Contiguous HTTP completions are coalesced into a
    //  single EB_OnHttpBatch call (one 2-D array per drain).
    // =========================================================
    internal static class VbaDispatcher
    {
        private sealed class HttpItem
        {
            public string Id, Body, Headers, Error;
            public int Status, ElapsedMs;
        }

        private sealed class MacroItem
        {
            public string Name;
            public object[] Args;
        }

        private static readonly ConcurrentQueue<object> _queue = new ConcurrentQueue<object>();
        private static int _scheduled;
        private static volatile string _macroPrefix;   // "'Book1.xlsm'!modBridgeEvents."
        private static readonly SendOrPostCallback _drain = _ => Drain();

        internal static string LastError = "";

        internal static void Attach(string workbookName)
        {
            if (string.IsNullOrEmpty(workbookName)) return;
            _macroPrefix = "'" + workbookName.Replace("'", "''") + "'!modBridgeEvents.";
            Kick(); // deliver anything buffered before attach
        }

        internal static void Detach()
        {
            _macroPrefix = null;
        }

        internal static void EnqueueHttp(string id, int status, string body, string headers, int elapsedMs, string error)
        {
            _queue.Enqueue(new HttpItem
            {
                Id = id, Status = status, Body = body ?? "",
                Headers = headers ?? "{}", ElapsedMs = elapsedMs, Error = error ?? ""
            });
            Kick();
        }

        internal static void EnqueueMacro(string name, params object[] args)
        {
            _queue.Enqueue(new MacroItem { Name = name, Args = args });
            Kick();
        }

        private static void Kick()
        {
            if (_macroPrefix == null) return;   // buffered until EB_Attach
            if (_queue.IsEmpty) return;
            if (Interlocked.CompareExchange(ref _scheduled, 1, 0) == 0)
                ExcelAsyncUtil.QueueAsMacro(_drain, null);
        }

        private static void Drain()
        {
            try
            {
                var prefix = _macroPrefix;
                if (prefix == null) return;     // detached; Attach will kick again

                object app = ExcelDnaUtil.Application;
                var http = new List<HttpItem>();

                object item;
                while (_queue.TryDequeue(out item))
                {
                    var h = item as HttpItem;
                    if (h != null) { http.Add(h); continue; }

                    FlushHttp(app, prefix, http);
                    var m = (MacroItem)item;
                    Run(app, prefix + m.Name, m.Args);
                }
                FlushHttp(app, prefix, http);
            }
            catch (Exception ex)
            {
                LastError = Stamp("drain: " + ex.GetBaseException().Message);
            }
            finally
            {
                Interlocked.Exchange(ref _scheduled, 0);
                Kick(); // anything enqueued while draining
            }
        }

        private static void FlushHttp(object app, string prefix, List<HttpItem> http)
        {
            if (http.Count == 0) return;

            var data = new object[http.Count, 6];
            for (int i = 0; i < http.Count; i++)
            {
                data[i, 0] = http[i].Id;
                data[i, 1] = http[i].Status;
                data[i, 2] = http[i].Body;
                data[i, 3] = http[i].Headers;
                data[i, 4] = http[i].ElapsedMs;
                data[i, 5] = http[i].Error;
            }
            Run(app, prefix + "EB_OnHttpBatch", new object[] { data });
            http.Clear();
        }

        private static void Run(object app, string macro, object[] args)
        {
            try
            {
                var full = new object[args.Length + 1];
                full[0] = macro;
                Array.Copy(args, 0, full, 1, args.Length);
                app.GetType().InvokeMember("Run", BindingFlags.InvokeMethod, null, app, full);
            }
            catch (Exception ex)
            {
                LastError = Stamp(macro + ": " + ex.GetBaseException().Message);
            }
        }

        private static string Stamp(string msg)
        {
            return DateTime.Now.ToString("HH:mm:ss") + " " + msg;
        }
    }
}
