using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace HRToolkit.Wpf.Services
{
    public class JsonRpcClient : IDisposable
    {
        private readonly StreamReader _reader;
        private readonly StreamWriter _writer;
        private readonly SemaphoreSlim _writeLock = new(1, 1);
        private readonly ConcurrentDictionary<long, TaskCompletionSource<JsonElement>> _pendingRequests = new();
        private readonly CancellationTokenSource _cts = new();
        private long _nextId = 1;

        public event Action<string, string>? OnLogReceived; // level, message
        public event Action<double, string>? OnProgressReceived; // ratio, message
        public event Action<JsonElement>? OnProjectChanged;
        public event Action? OnWorkspaceChanged;
        public event Action<bool, string>? OnToolFinished; // success, message/error

        public JsonRpcClient(StreamReader reader, StreamWriter writer)
        {
            _reader = reader;
            _writer = writer;
            Task.Run(ListenLoopAsync);
        }

        public void RaiseLog(string level, string message)
        {
            OnLogReceived?.Invoke(level, message);
        }

        public async Task<T?> SendRequestAsync<T>(string method, object? parameters = null, CancellationToken ct = default)
        {
            long id = Interlocked.Increment(ref _nextId);
            var tcs = new TaskCompletionSource<JsonElement>(TaskCreationOptions.RunContinuationsAsynchronously);
            _pendingRequests[id] = tcs;

            var requestObj = new
            {
                jsonrpc = "2.0",
                id = id,
                method = method,
                @params = parameters ?? new { }
            };

            string line = JsonSerializer.Serialize(requestObj);

            // Default 15-second timeout for RPC requests if not specified
            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(ct, timeoutCts.Token);

            await _writeLock.WaitAsync(linkedCts.Token);
            try
            {
                await _writer.WriteLineAsync(line.AsMemory(), linkedCts.Token);
                await _writer.FlushAsync(linkedCts.Token);
            }
            finally
            {
                _writeLock.Release();
            }

            using var registration = linkedCts.Token.Register(() =>
            {
                if (timeoutCts.IsCancellationRequested && !ct.IsCancellationRequested)
                {
                    tcs.TrySetException(new TimeoutException($"RPC 请求 '{method}' 超时（15秒内未收到响应）。"));
                }
                else
                {
                    tcs.TrySetCanceled();
                }
            });

            JsonElement resultElement = await tcs.Task;

            if (resultElement.ValueKind == JsonValueKind.Undefined || resultElement.ValueKind == JsonValueKind.Null)
            {
                return default;
            }

            return JsonSerializer.Deserialize<T>(resultElement.GetRawText());
        }

        private async Task ListenLoopAsync()
        {
            try
            {
                while (!_cts.Token.IsCancellationRequested)
                {
                    string? line = await _reader.ReadLineAsync(_cts.Token);
                    if (line == null)
                    {
                        // EOF detected: Python backend process closed stdout or terminated
                        var ex = new IOException("Python 后端进程通信中断（标准输出流已关闭，进程可能已退出）。");
                        foreach (var kvp in _pendingRequests)
                        {
                            kvp.Value.TrySetException(ex);
                        }
                        _pendingRequests.Clear();
                        OnLogReceived?.Invoke("error", "与 Python 后端引擎通信断开。");
                        break;
                    }
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    ProcessIncomingMessage(line);
                }
            }
            catch (OperationCanceledException) { }
            catch (Exception ex)
            {
                OnLogReceived?.Invoke("error", $"IPC 通信异常: {ex.Message}");
            }
        }

        private void ProcessIncomingMessage(string jsonLine)
        {
            try
            {
                using var doc = JsonDocument.Parse(jsonLine);
                var root = doc.RootElement.Clone();

                // 1. Response to request
                if (root.TryGetProperty("id", out var idProp) && idProp.ValueKind == JsonValueKind.Number)
                {
                    long id = idProp.GetInt64();
                    if (_pendingRequests.TryRemove(id, out var tcs))
                    {
                        if (root.TryGetProperty("error", out var errorProp))
                        {
                            string errMsg = errorProp.TryGetProperty("message", out var msgProp) ? msgProp.GetString() ?? "RPC Error" : "Unknown RPC Error";
                            tcs.TrySetException(new InvalidOperationException(errMsg));
                        }
                        else if (root.TryGetProperty("result", out var resultProp))
                        {
                            tcs.TrySetResult(resultProp);
                        }
                        else
                        {
                            tcs.TrySetResult(default);
                        }
                    }
                    return;
                }

                // 1.1 Global error response without request ID (e.g. parse error)
                if (root.TryGetProperty("error", out var globalErrProp))
                {
                    string errMsg = globalErrProp.TryGetProperty("message", out var msgProp) ? msgProp.GetString() ?? "RPC Error" : "Unknown RPC Error";
                    OnLogReceived?.Invoke("error", $"Python 核心报错: {errMsg}");
                    var firstKey = _pendingRequests.Keys.FirstOrDefault();
                    if (firstKey != 0 && _pendingRequests.TryRemove(firstKey, out var pendingTcs))
                    {
                        pendingTcs.TrySetException(new InvalidOperationException(errMsg));
                    }
                    return;
                }

                // 2. Notification Event from server
                if (root.TryGetProperty("method", out var methodProp) && methodProp.GetString() == "event")
                {
                    if (root.TryGetProperty("params", out var paramsProp))
                    {
                        string eventName = paramsProp.GetProperty("name").GetString() ?? "";
                        var data = paramsProp.GetProperty("data");

                        switch (eventName)
                        {
                            case "log":
                                string msg = data.GetProperty("message").GetString() ?? "";
                                string lvl = data.TryGetProperty("level", out var l) ? l.GetString() ?? "info" : "info";
                                OnLogReceived?.Invoke(lvl, msg);
                                break;
                            case "progress":
                                double ratio = data.GetProperty("ratio").GetDouble();
                                string pMsg = data.TryGetProperty("message", out var pm) ? pm.GetString() ?? "" : "";
                                OnProgressReceived?.Invoke(ratio, pMsg);
                                break;
                            case "project_changed":
                                OnProjectChanged?.Invoke(data);
                                break;
                            case "workspace_changed":
                                OnWorkspaceChanged?.Invoke();
                                break;
                            case "finished":
                                bool success = data.GetProperty("success").GetBoolean();
                                string finishedMsg = data.TryGetProperty("error", out var err) ? err.GetString() ?? "执行失败" : "执行完成";
                                OnToolFinished?.Invoke(success, finishedMsg);
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string preview = jsonLine.Length > 120 ? jsonLine.Substring(0, 120) + "..." : jsonLine;
                OnLogReceived?.Invoke("warning", $"解析 IPC 消息失败: {ex.Message} (内容: {preview})");
            }
        }

        public void Dispose()
        {
            _cts.Cancel();
            _writeLock.Dispose();
        }
    }
}
