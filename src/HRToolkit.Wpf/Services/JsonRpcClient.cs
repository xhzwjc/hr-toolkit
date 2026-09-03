using System;
using System.Collections.Concurrent;
using System.IO;
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

            await _writeLock.WaitAsync(ct);
            try
            {
                await _writer.WriteLineAsync(line.AsMemory(), ct);
                await _writer.FlushAsync(ct);
            }
            finally
            {
                _writeLock.Release();
            }

            using var registration = ct.Register(() => tcs.TrySetCanceled());
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
                    if (line == null) break;
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
                OnLogReceived?.Invoke("warning", $"解析 IPC 消息失败: {ex.Message}");
            }
        }

        public void Dispose()
        {
            _cts.Cancel();
            _writeLock.Dispose();
        }
    }
}
