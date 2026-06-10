
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.WebSockets;
using System.Runtime.InteropServices;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading;
using System.Threading.Tasks;

[assembly: Guid("5B1C2E77-6A1D-4C9E-9E7A-7C4F0C1A1000")]
[assembly: ComVisible(true)]

namespace ExcelBridge
{
    // =========================================================
    //  COM event interface (source interface for VBA WithEvents)
    // =========================================================
    [ComVisible(true)]
    [Guid("5B1C2E77-6A1D-4C9E-9E7A-7C4F0C1A1001")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IBridgeEvents
    {
        [DispId(1)]  void OnCellUpdate(int row, int col, string value);
        [DispId(2)]  void OnCellBatch(object cells);
        [DispId(3)]  void OnFullGrid(string jsonGrid);
        [DispId(4)]  void OnConnected();
        [DispId(5)]  void OnDisconnected(string reason);
        [DispId(6)]  void OnReconnecting(int attempt, int delayMs);
        [DispId(7)]  void OnError(string message);
        [DispId(8)]  void OnLog(string level, string message);
        [DispId(9)]  void OnMessage(string msgType, string jsonPayload);
        [DispId(10)] void OnHttpResponse(string requestId, int statusCode, string body, string headersJson, int elapsedMs);
        [DispId(11)] void OnHttpError(string requestId, string message, int elapsedMs);
        [DispId(12)] void OnMessageBatch(string jsonArray);
    }

    // =========================================================
    //  COM method interface
    // =========================================================
    [ComVisible(true)]
    [Guid("5B1C2E77-6A1D-4C9E-9E7A-7C4F0C1A1002")]
    public interface IBridge
    {
        // WebSocket
        [DispId(1)] void   Start(string url);
        [DispId(2)] void   Stop();
        [DispId(3)] void   PushCell(int row, int col, string value);
        [DispId(4)] void   PushCells(object cells);
        [DispId(5)] void   Send(string msgType, string jsonPayload);
        [DispId(6)] bool   IsConnected { get; }
        [DispId(7)] bool   AutoReconnect { get; set; }
        [DispId(8)] int    MaxReconnectDelayMs { get; set; }
        [DispId(9)] bool   RaiseGenericOnMessage { get; set; }

        // HTTP
        [DispId(10)] string HttpSendAsync(string method, string url, string body, string headers, int timeoutMs);
        [DispId(11)] object HttpSendSync(string method, string url, string body, string headers, int timeoutMs);
        [DispId(12)] void   HttpCancel(string requestId);
        [DispId(13)] int    HttpDefaultTimeoutMs { get; set; }

        // Batch buffering
        [DispId(14)] int    BatchWindowMs { get; set; }
        [DispId(15)] int    BatchMaxCount { get; set; }

        // State
        [DispId(16)] bool   IsRunning { get; }
    }

    // =========================================================
    //  Implementation
    // =========================================================
    [ComVisible(true)]
    [Guid("5B1C2E77-6A1D-4C9E-9E7A-7C4F0C1A1003")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(IBridgeEvents))]
    [ProgId("ExcelBridge.Bridge")]
    public class Bridge : IBridge
    {
        // ---------- Event delegates ----------
        public delegate void CellUpdateHandler(int row, int col, string value);
        public delegate void CellBatchHandler(object cells);
        public delegate void FullGridHandler(string jsonGrid);
        public delegate void ConnectedHandler();
        public delegate void DisconnectedHandler(string reason);
        public delegate void ReconnectingHandler(int attempt, int delayMs);
        public delegate void ErrorHandler(string message);
        public delegate void LogHandler(string level, string message);
        public delegate void MessageHandler(string msgType, string jsonPayload);
        public delegate void HttpResponseHandler(string requestId, int statusCode, string body, string headersJson, int elapsedMs);
        public delegate void HttpErrorHandler(string requestId, string message, int elapsedMs);
        public delegate void MessageBatchHandler(string jsonArray);

        public event CellUpdateHandler    OnCellUpdate;
        public event CellBatchHandler     OnCellBatch;
        public event FullGridHandler      OnFullGrid;
        public event ConnectedHandler     OnConnected;
        public event DisconnectedHandler  OnDisconnected;
        public event ReconnectingHandler  OnReconnecting;
        public event ErrorHandler         OnError;
        public event LogHandler           OnLog;
        public event MessageHandler       OnMessage;
        public event HttpResponseHandler  OnHttpResponse;
        public event HttpErrorHandler     OnHttpError;
        public event MessageBatchHandler  OnMessageBatch;

        // ---------- WebSocket state ----------
        private readonly Dictionary<string, Action<JObject>> _handlers;
        private ClientWebSocket _ws;
        private CancellationTokenSource _cts;
        private Task _loopTask;
        private Uri _uri;
        private volatile bool _shuttingDown;
        private int _generation;            // bumped on every Start/Stop to fence stale connect loops

        private readonly object _outboxLock = new object();
        private readonly LinkedList<string> _outbox = new LinkedList<string>();
        private readonly SemaphoreSlim _sendSemaphore = new SemaphoreSlim(1, 1);

        private static readonly Random _rng = new Random();

        public bool AutoReconnect { get; set; } = true;
        public int  MaxReconnectDelayMs { get; set; } = 30000;
        public bool RaiseGenericOnMessage { get; set; } = false;
        public bool IsConnected => _ws != null && _ws.State == WebSocketState.Open;
        public bool IsRunning => !_shuttingDown && _loopTask != null && !_loopTask.IsCompleted;

        // ---------- Batch buffering state ----------
        private readonly ConcurrentQueue<string> _batchQueue = new ConcurrentQueue<string>();
        private Timer _batchTimer;
        private int _batchCount;
        private readonly object _flushLock = new object();

        public int BatchWindowMs { get; set; } = 16;
        public int BatchMaxCount { get; set; } = 200;

        // ---------- HTTP state ----------
        private static readonly HttpClient _http = CreateHttpClient();
        private static readonly ConcurrentDictionary<string, CancellationTokenSource> _inflight
            = new ConcurrentDictionary<string, CancellationTokenSource>();

        public int HttpDefaultTimeoutMs { get; set; } = 30000;

        private static HttpClient CreateHttpClient()
        {
            // Tune global connection management before building the handler
            ServicePointManager.DefaultConnectionLimit = 100;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.UseNagleAlgorithm = false;

            // Enable TLS 1.2 (and 1.3 if the runtime supports it)
            try
            {
                ServicePointManager.SecurityProtocol |=
                    SecurityProtocolType.Tls12 | (SecurityProtocolType)12288 /* Tls13 */;
            }
            catch
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }

            var handler = new HttpClientHandler
            {
                AllowAutoRedirect = true,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                UseCookies = false
            };

            var client = new HttpClient(handler, disposeHandler: true)
            {
                // Per-request timeout is enforced via CancellationToken
                Timeout = Timeout.InfiniteTimeSpan
            };
            client.DefaultRequestHeaders.ExpectContinue = false;
            client.DefaultRequestHeaders.ConnectionClose = false;
            return client;
        }

        // =====================================================
        //  Constructor — wire up message dispatch table
        // =====================================================
        public Bridge()
        {
            _handlers = new Dictionary<string, Action<JObject>>(StringComparer.OrdinalIgnoreCase)
			{
				["cell"] = root =>
				{
					int r = root.Value<int>("row");
					int c = root.Value<int>("col");
					string v = root.Value<string>("val") ?? "";
					Raise(() => OnCellUpdate?.Invoke(r, c, v));
				},
				["cell_batch"] = root =>
				{
					var arr = (JArray)root["cells"];
					if (arr == null || arr.Count == 0) return;
					var data = new object[arr.Count, 3];
					for (int i = 0; i < arr.Count; i++)
					{
						var cell = (JObject)arr[i];
						data[i, 0] = cell.Value<int>("row");
						data[i, 1] = cell.Value<int>("col");
						data[i, 2] = cell.Value<string>("val") ?? "";
					}
					Raise(() => OnCellBatch?.Invoke(data));
				},
				["full"] = root =>
				{
					string g = root["grid"]?.ToString(Formatting.None) ?? "[]";
					Raise(() => OnFullGrid?.Invoke(g));
				},
			};
        }

        // =====================================================
        //  Public WebSocket API
        // =====================================================
        public void Start(string url)
        {
            Stop();
            _shuttingDown = false;
            _uri = new Uri(url);
            _cts = new CancellationTokenSource();
            var token = _cts.Token;
            int gen = Interlocked.Increment(ref _generation);
            int window = Math.Max(4, BatchWindowMs);
            _batchTimer = new Timer(_ => { try { FlushBatch(); } catch { } }, null, window, window);
            _loopTask = Task.Run(() => ConnectLoopAsync(gen, token));
        }

        public void Stop()
        {
            try
            {
                _shuttingDown = true;
                Interlocked.Increment(ref _generation);   // fence any connect loop still running
                var cts = _cts;
                cts?.Cancel();

                var ws = _ws;
                if (ws != null && ws.State == WebSocketState.Open)
                {
                    try
                    {
                        ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "bye",
                            CancellationToken.None).Wait(500);
                    }
                    catch { }
                }
            }
            catch { }
            finally
            {
                try { _batchTimer?.Dispose(); } catch { }
                _batchTimer = null;
                FlushBatch();
                var oldWs = Interlocked.Exchange(ref _ws, null);
                try { oldWs?.Dispose(); } catch { }
                try { _cts?.Dispose(); } catch { }
                _cts = null;
            }
        }

        public void PushCell(int row, int col, string value)
		{
			var obj = new JObject
			{
				["type"] = "push",
				["row"] = row,
				["col"] = col,
				["val"] = value ?? ""
			};
			EnqueueSend(obj.ToString(Formatting.None));
		}

        public void PushCells(object cells)
		{
			if (!(cells is object[,] arr))
			{
				RaiseError("PushCells: expected 2-D array");
				return;
			}

			int n = arr.GetLength(0);
			int lb0 = arr.GetLowerBound(0);
			int lb1 = arr.GetLowerBound(1);

			var cellsArr = new JArray();
			for (int i = 0; i < n; i++)
			{
				cellsArr.Add(new JObject
				{
					["row"] = Convert.ToInt32(arr[lb0 + i, lb1 + 0]),
					["col"] = Convert.ToInt32(arr[lb0 + i, lb1 + 1]),
					["val"] = arr[lb0 + i, lb1 + 2]?.ToString() ?? ""
				});
			}

			var obj = new JObject
			{
				["type"] = "push_batch",
				["cells"] = cellsArr
			};
			EnqueueSend(obj.ToString(Formatting.None));
		}

		public void Send(string msgType, string jsonPayload)
		{
			JObject obj;
			try
			{
				if (string.IsNullOrWhiteSpace(jsonPayload))
					obj = new JObject();
				else
					obj = JObject.Parse(jsonPayload);
			}
			catch (Exception ex)
			{
				RaiseError("Send: invalid jsonPayload: " + ex.Message);
				return;
			}
			obj["type"] = msgType ?? "";
			EnqueueSend(obj.ToString(Formatting.None));
		}

        // =====================================================
        //  Connect / receive loops
        // =====================================================
        private async Task ConnectLoopAsync(int gen, CancellationToken ct)
        {
            int attempt = 0;
            while (!ct.IsCancellationRequested && !_shuttingDown)
            {
                var ws = new ClientWebSocket();
                ws.Options.KeepAliveInterval = TimeSpan.FromSeconds(10);

                // A newer Start()/Stop() owns the bridge now — bow out without
                // touching shared state.
                if (gen != Volatile.Read(ref _generation)) { try { ws.Dispose(); } catch { } return; }
                _ws = ws;

                try
                {
                    Log("info", $"connecting to {_uri} (attempt {attempt + 1})");
                    await ws.ConnectAsync(_uri, ct).ConfigureAwait(false);
                    attempt = 0;
                    Raise(() => OnConnected?.Invoke());

                    _ = Task.Run(() => FlushOutboxAsync(ct));

                    await ReceiveLoopAsync(ws, ct).ConfigureAwait(false);
                    Raise(() => OnDisconnected?.Invoke("closed"));
                }
                catch (OperationCanceledException)
                {
                    Raise(() => OnDisconnected?.Invoke("cancelled"));
                    return;
                }
                catch (Exception ex)
                {
                    RaiseError("connect: " + ex.Message);
                    Raise(() => OnDisconnected?.Invoke("error"));
                }
                finally
                {
                    // Clear the shared field only if it is still our socket;
                    // a newer loop may have replaced it already.
                    Interlocked.CompareExchange(ref _ws, null, ws);
                    try { ws.Dispose(); } catch { }
                }

                if (_shuttingDown || !AutoReconnect || ct.IsCancellationRequested ||
                    gen != Volatile.Read(ref _generation)) return;

                attempt++;
                int delay = Math.Min(MaxReconnectDelayMs,
                                     (int)(500 * Math.Pow(2, Math.Min(attempt, 10))));
                int jitter;
                lock (_rng) jitter = _rng.Next(0, 250);
                delay += jitter;

                Raise(() => OnReconnecting?.Invoke(attempt, delay));
                try { await Task.Delay(delay, ct).ConfigureAwait(false); }
                catch (OperationCanceledException) { return; }
            }
        }

        private async Task ReceiveLoopAsync(ClientWebSocket ws, CancellationToken ct)
        {
            var buffer = new byte[64 * 1024];
            var seg = new ArraySegment<byte>(buffer);
            MemoryStream ms = null;

            while (!ct.IsCancellationRequested && ws.State == WebSocketState.Open)
            {
                WebSocketReceiveResult result;
                try
                {
                    result = await ws.ReceiveAsync(seg, ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException) { return; }
                catch { return; }

                if (result.MessageType == WebSocketMessageType.Close)
                {
                    try
                    {
                        await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "ack",
                            CancellationToken.None).ConfigureAwait(false);
                    }
                    catch { }
                    return;
                }

                if (result.EndOfMessage)
                {
                    var text = Encoding.UTF8.GetString(buffer, 0, result.Count);
                    try { Dispatch(text); } catch { }
                    continue;
                }

                if (ms == null) ms = new MemoryStream();
                ms.SetLength(0);
                ms.Write(buffer, 0, result.Count);
                do
                {
                    try
                    {
                        result = await ws.ReceiveAsync(seg, ct).ConfigureAwait(false);
                    }
                    catch (OperationCanceledException) { return; }
                    catch { return; }
                    if (result.MessageType == WebSocketMessageType.Close) return;
                    ms.Write(buffer, 0, result.Count);
                } while (!result.EndOfMessage);

                try { Dispatch(Encoding.UTF8.GetString(ms.GetBuffer(), 0, (int)ms.Length)); } catch { }
            }
        }

        private void Dispatch(string json)
		{
			try
			{
				var token = JToken.Parse(json);

				if (token.Type == JTokenType.Array)
				{
					var arr = (JArray)token;
					Log("debug", $"dispatch array: {arr.Count} items");
					foreach (var item in arr)
					{
						if (item.Type == JTokenType.Object)
							DispatchObject((JObject)item);
					}
					return;
				}

				if (token.Type == JTokenType.Object)
				{
					// Hand the original wire text through when it is already in the
					// compact form VBA's scanners expect, so DispatchObject can skip
					// a full re-serialization of the payload.
					DispatchObject((JObject)token, LooksCompact(json) ? json : null);
					return;
				}

				Raise(() => OnMessage?.Invoke("unknown", json));
			}
			catch (Exception ex)
			{
				RaiseError("dispatch: " + ex.Message);
			}
		}

		private void DispatchObject(JObject root, string raw = null)
		{
			var type = root.Value<string>("type") ?? "";

			if (_handlers.TryGetValue(type, out var handler))
			{
				handler(root);
				return;
			}

			if (raw == null) raw = root.ToString(Formatting.None);

			if (RaiseGenericOnMessage)
				Raise(() => OnMessage?.Invoke(type, raw));

			EnqueueBatch(raw);
		}

		// VBA-side routing (ExtractJsonValue / RouteBatch) scans for compact
		// "key":value text. Only skip re-serialization when the wire text is
		// already in that shape; false negatives just fall back to ToString.
		private static bool LooksCompact(string json)
		{
			return json.IndexOf('\n') < 0 && json.IndexOf('\r') < 0 &&
			       json.IndexOf("\": ", StringComparison.Ordinal) < 0 &&
			       json.IndexOf("\" :", StringComparison.Ordinal) < 0;
		}

        // =====================================================
        //  Batch buffer for non-typed messages
        // =====================================================
        private void EnqueueBatch(string json)
        {
            _batchQueue.Enqueue(json);
            if (Interlocked.Increment(ref _batchCount) >= BatchMaxCount)
                Task.Run(() => FlushBatch());
        }

        private void FlushBatch()
        {
            if (_batchQueue.IsEmpty) return;
            lock (_flushLock)
            {
                if (_batchQueue.IsEmpty) return;
                var sb = new StringBuilder(4096);
                sb.Append('[');
                bool first = true;
                int count = 0;
                while (_batchQueue.TryDequeue(out var item))
                {
                    if (!first) sb.Append(',');
                    sb.Append(item);
                    first = false;
                    count++;
                }
                sb.Append(']');
                Interlocked.Exchange(ref _batchCount, 0);
                var payload = sb.ToString();
                Log("info", $"flush batch: {count} items, {payload.Length} chars");
                Raise(() => OnMessageBatch?.Invoke(payload));
            }
        }

        // =====================================================
        //  Send plumbing
        // =====================================================
        private const int MaxOutboxCount = 10000;

        private void EnqueueSend(string body)
        {
            bool dropped = false;
            lock (_outboxLock)
            {
                // Bound the queue: while disconnected, callers can otherwise
                // grow this list without limit.
                if (_outbox.Count >= MaxOutboxCount) { _outbox.RemoveFirst(); dropped = true; }
                _outbox.AddLast(body);
            }
            if (dropped) RaiseError("outbox overflow: dropped oldest queued message");
            var cts = _cts;
            if (IsConnected && cts != null && !cts.IsCancellationRequested)
                _ = Task.Run(() => FlushOutboxAsync(cts.Token));
        }

        private async Task FlushOutboxAsync(CancellationToken ct)
        {
            try
            {
                await _sendSemaphore.WaitAsync(ct).ConfigureAwait(false);
            }
            catch (OperationCanceledException) { return; }

            try
            {
                while (!ct.IsCancellationRequested)
                {
                    var ws = _ws;
                    if (ws == null || ws.State != WebSocketState.Open) return;

                    string body;
                    lock (_outboxLock)
                    {
                        if (_outbox.Count == 0) return;
                        body = _outbox.First.Value;
                        _outbox.RemoveFirst();
                    }

                    try
                    {
                        var bytes = Encoding.UTF8.GetBytes(body);
                        await ws.SendAsync(new ArraySegment<byte>(bytes),
                            WebSocketMessageType.Text, true, ct).ConfigureAwait(false);
                    }
                    catch (OperationCanceledException)
                    {
                        lock (_outboxLock) _outbox.AddFirst(body);
                        return;
                    }
                    catch (Exception ex)
                    {
                        lock (_outboxLock) _outbox.AddFirst(body);
                        RaiseError("send: " + ex.Message);
                        return;
                    }
                }
            }
            finally
            {
                try { _sendSemaphore.Release(); } catch { }
            }
        }

        // =====================================================
        //  HTTP API
        // =====================================================

        /// <summary>Fire-and-forget. Returns a request ID immediately. Completion comes via OnHttpResponse / OnHttpError.</summary>
        public string HttpSendAsync(string method, string url, string body, string headers, int timeoutMs)
        {
            var requestId = Guid.NewGuid().ToString("N");
            int effectiveTimeout = timeoutMs > 0 ? timeoutMs : HttpDefaultTimeoutMs;
            var cts = new CancellationTokenSource(effectiveTimeout);
            _inflight[requestId] = cts;

            _ = Task.Run(async () =>
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                try
                {
                    var result = await DoHttpAsync(method, url, body, headers, cts.Token)
                                     .ConfigureAwait(false);
                    sw.Stop();
                    int elapsed = (int)Math.Min(sw.ElapsedMilliseconds, int.MaxValue);
                    Raise(() => OnHttpResponse?.Invoke(
                        requestId, result.status, result.body, result.headersJson, elapsed));
                }
                catch (OperationCanceledException)
                {
                    sw.Stop();
                    int elapsed = (int)Math.Min(sw.ElapsedMilliseconds, int.MaxValue);
                    Raise(() => OnHttpError?.Invoke(requestId, "timeout or cancelled", elapsed));
                }
                catch (Exception ex)
                {
                    sw.Stop();
                    int elapsed = (int)Math.Min(sw.ElapsedMilliseconds, int.MaxValue);
                    Raise(() => OnHttpError?.Invoke(requestId, ex.GetBaseException().Message, elapsed));
                }
                finally
                {
                    if (_inflight.TryRemove(requestId, out var removed))
                    {
                        try { removed.Dispose(); } catch { }
                    }
                }
            });

            return requestId;
        }

        /// <summary>Blocking call. Returns object[] { status, body, headersJson, elapsedMs, errorOrEmpty }.</summary>
        public object HttpSendSync(string method, string url, string body, string headers, int timeoutMs)
        {
            int effectiveTimeout = timeoutMs > 0 ? timeoutMs : HttpDefaultTimeoutMs;
            var sw = System.Diagnostics.Stopwatch.StartNew();
            CancellationTokenSource cts = null;
            try
            {
                cts = new CancellationTokenSource(effectiveTimeout);
                var token = cts.Token;
                var task = Task.Run(() => DoHttpAsync(method, url, body, headers, token));
                var result = task.GetAwaiter().GetResult();
                sw.Stop();
                int elapsed = (int)Math.Min(sw.ElapsedMilliseconds, int.MaxValue);
                return new object[] { result.status, result.body, result.headersJson, elapsed, "" };
            }
            catch (OperationCanceledException)
            {
                sw.Stop();
                int elapsed = (int)Math.Min(sw.ElapsedMilliseconds, int.MaxValue);
                return new object[] { 0, "", "{}", elapsed, "timeout or cancelled" };
            }
            catch (Exception ex)
            {
                sw.Stop();
                int elapsed = (int)Math.Min(sw.ElapsedMilliseconds, int.MaxValue);
                return new object[] { 0, "", "{}", elapsed, ex.GetBaseException().Message };
            }
            finally
            {
                try { cts?.Dispose(); } catch { }
            }
        }

        public void HttpCancel(string requestId)
        {
            if (string.IsNullOrEmpty(requestId)) return;
            if (_inflight.TryGetValue(requestId, out var cts))
            {
                try { cts.Cancel(); } catch { }
            }
        }

        private async Task<(int status, string body, string headersJson)>
            DoHttpAsync(string method, string url, string body, string headers, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(method)) method = "GET";
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException("url is required");

            var httpMethod = new HttpMethod(method.ToUpperInvariant());
            using var req = new HttpRequestMessage(httpMethod, url);

            // Parse headers: "Key: Value\nKey2: Value2"
            string contentType = null;
            var deferredContentHeaders = new List<KeyValuePair<string, string>>();
            if (!string.IsNullOrEmpty(headers))
            {
                var lines = headers.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var raw in lines)
                {
                    var line = raw.Trim();
                    if (line.Length == 0) continue;
                    int colon = line.IndexOf(':');
                    if (colon <= 0) continue;
                    var name = line.Substring(0, colon).Trim();
                    var value = line.Substring(colon + 1).Trim();
                    if (name.Length == 0) continue;

                    if (name.Equals("Content-Type", StringComparison.OrdinalIgnoreCase))
                    {
                        contentType = value;
                        continue;
                    }
                    if (name.Equals("Content-Length", StringComparison.OrdinalIgnoreCase) ||
                        name.Equals("Content-Encoding", StringComparison.OrdinalIgnoreCase))
                    {
                        deferredContentHeaders.Add(new KeyValuePair<string, string>(name, value));
                        continue;
                    }

                    if (!req.Headers.TryAddWithoutValidation(name, value))
                    {
                        // If it wasn't a request header, it's probably a content header — defer
                        deferredContentHeaders.Add(new KeyValuePair<string, string>(name, value));
                    }
                }
            }

            bool allowsBody =
                httpMethod == HttpMethod.Post ||
                httpMethod == HttpMethod.Put  ||
                httpMethod.Method == "PATCH"  ||
                httpMethod == HttpMethod.Delete;

            if (allowsBody && !string.IsNullOrEmpty(body))
            {
                var contentBytes = Encoding.UTF8.GetBytes(body);
                var content = new ByteArrayContent(contentBytes);
                content.Headers.TryAddWithoutValidation(
                    "Content-Type", contentType ?? "application/json; charset=utf-8");
                foreach (var kv in deferredContentHeaders)
                    content.Headers.TryAddWithoutValidation(kv.Key, kv.Value);
                req.Content = content;
            }

            using var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct)
                                        .ConfigureAwait(false);

            // ReadAsStringAsync honors the response charset (the old manual UTF-8
            // decode corrupted non-UTF-8 bodies). On net48 the body read does not
            // observe the CancellationToken, so a server that returns headers and
            // then trickles the body would bypass the timeout entirely — register
            // a dispose to abort the read when the token fires.
            string text;
            try
            {
                using (ct.Register(s => { try { ((HttpResponseMessage)s).Dispose(); } catch { } }, resp))
                {
                    text = await resp.Content.ReadAsStringAsync().ConfigureAwait(false);
                }
            }
            catch (Exception) when (ct.IsCancellationRequested)
            {
                throw new OperationCanceledException(ct);
            }

            // Headers -> compact JSON
            var headersObj = new JObject();
			foreach (var h in resp.Headers)
				headersObj[h.Key] = string.Join(",", h.Value);
			if (resp.Content != null)
			{
				foreach (var h in resp.Content.Headers)
					headersObj[h.Key] = string.Join(",", h.Value);
			}
			string headersJson = headersObj.ToString(Formatting.None);

            return ((int)resp.StatusCode, text, headersJson);
        }

        // =====================================================
        //  Raise helpers
        // =====================================================
        private const uint RPC_E_CALL_REJECTED        = 0x80010001;
        private const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;

        private void Raise(Action a)
        {
            // Events are raised from worker threads into Excel's STA. While the
            // user is editing a cell or a dialog is open, Excel rejects incoming
            // COM calls; swallowing that would silently drop the event (e.g. an
            // HTTP completion, hanging any batch waiting on it). Retry instead.
            for (int attempt = 0; ; attempt++)
            {
                try { a(); return; }
                catch (COMException ex) when (attempt < 10 &&
                    ((uint)ex.HResult == RPC_E_CALL_REJECTED ||
                     (uint)ex.HResult == RPC_E_SERVERCALL_RETRYLATER))
                {
                    Thread.Sleep(50 * (attempt + 1));
                }
                catch (Exception ex)
                {
                    try { OnLog?.Invoke("error", "handler threw: " + ex.Message); } catch { }
                    return;
                }
            }
        }

        private void RaiseError(string msg)
        {
            try { OnError?.Invoke(msg); } catch { }
        }

        private void Log(string level, string msg)
        {
            try { OnLog?.Invoke(level, msg); } catch { }
        }
    }
}


