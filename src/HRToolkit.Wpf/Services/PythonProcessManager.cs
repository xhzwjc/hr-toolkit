using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace HRToolkit.Wpf.Services
{
    public class PythonProcessManager : IDisposable
    {
        private Process? _process;
        private readonly StringBuilder _stderrBuffer = new();
        public JsonRpcClient? Client { get; private set; }
        public string LastStderr => _stderrBuffer.ToString();
        public string ResolvedPythonPath { get; private set; } = string.Empty;
        public string ResolvedWorkingDir { get; private set; } = string.Empty;

        public void StartEngine()
        {
            ResolvedWorkingDir = ResolveWorkingDirectory();
            ResolvedPythonPath = ResolvePythonExecutable(ResolvedWorkingDir);

            var psi = new ProcessStartInfo
            {
                FileName = ResolvedPythonPath,
                Arguments = "-m hr_toolkit --ipc",
                WorkingDirectory = ResolvedWorkingDir,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            psi.EnvironmentVariables["PYTHONUTF8"] = "1";
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";
            psi.EnvironmentVariables["PYTHONPATH"] = ResolvedWorkingDir;

            _process = new Process { StartInfo = psi };
            _process.Start();

            _process.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    _stderrBuffer.AppendLine(e.Data);
                    Debug.WriteLine($"[Python Stderr] {e.Data}");
                    Client?.RaiseLog("warning", e.Data);
                }
            };
            _process.BeginErrorReadLine();

            Client = new JsonRpcClient(_process.StandardOutput, _process.StandardInput);
        }

        private string ResolvePythonExecutable(string workingDir)
        {
            // 1. Working dir .venv or venv
            string venvWin = Path.Combine(workingDir, ".venv", "Scripts", "python.exe");
            if (File.Exists(venvWin)) return venvWin;

            string venvWin2 = Path.Combine(workingDir, "venv", "Scripts", "python.exe");
            if (File.Exists(venvWin2)) return venvWin2;

            string venvPosix = Path.Combine(workingDir, ".venv", "bin", "python");
            if (File.Exists(venvPosix)) return venvPosix;

            string venvPosix2 = Path.Combine(workingDir, "venv", "bin", "python");
            if (File.Exists(venvPosix2)) return venvPosix2;

            // 2. Base dir embedded or relative
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string embedded = Path.Combine(baseDir, "python", "python.exe");
            if (File.Exists(embedded)) return embedded;

            string standalone = Path.Combine(baseDir, "hr_toolkit_core.exe");
            if (File.Exists(standalone)) return standalone;

            string relVenvWin = Path.Combine(baseDir, "..", "..", "..", "..", ".venv", "Scripts", "python.exe");
            if (File.Exists(relVenvWin)) return Path.GetFullPath(relVenvWin);

            // 3. Fallback to system python
            return "python";
        }

        private string ResolveWorkingDirectory()
        {
            // 1. Current working directory (if run via dotnet run from repo root)
            string cwd = Directory.GetCurrentDirectory();
            if (Directory.Exists(Path.Combine(cwd, "hr_toolkit")))
            {
                return Path.GetFullPath(cwd);
            }

            // 2. Base directory relative traversal
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string repoRoot = Path.Combine(baseDir, "..", "..", "..", "..");
            if (Directory.Exists(Path.Combine(repoRoot, "hr_toolkit")))
            {
                return Path.GetFullPath(repoRoot);
            }

            return baseDir;
        }

        public void Dispose()
        {
            Client?.Dispose();
            if (_process != null && !_process.HasExited)
            {
                try
                {
                    _process.Kill(true);
                }
                catch { }
                _process.Dispose();
            }
        }
    }
}
