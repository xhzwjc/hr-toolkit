using System;
using System.Diagnostics;
using System.IO;

namespace HRToolkit.Wpf.Services
{
    public class PythonProcessManager : IDisposable
    {
        private Process? _process;
        public JsonRpcClient? Client { get; private set; }

        public void StartEngine()
        {
            string pythonExe = ResolvePythonExecutable();
            string workingDir = ResolveWorkingDirectory();

            var psi = new ProcessStartInfo
            {
                FileName = pythonExe,
                Arguments = "-m hr_toolkit --ipc",
                WorkingDirectory = workingDir,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            psi.EnvironmentVariables["PYTHONUTF8"] = "1";
            psi.EnvironmentVariables["PYTHONIOENCODING"] = "utf-8";

            _process = new Process { StartInfo = psi };
            _process.Start();

            // Direct stderr to debug / error logging
            _process.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    Debug.WriteLine($"[Python Stderr] {e.Data}");
                }
            };
            _process.BeginErrorReadLine();

            Client = new JsonRpcClient(_process.StandardOutput, _process.StandardInput);
        }

        private string ResolvePythonExecutable()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // 1. Packaged standalone python executable or embedded runtime
            string embedded = Path.Combine(baseDir, "python", "python.exe");
            if (File.Exists(embedded)) return embedded;

            string standalone = Path.Combine(baseDir, "hr_toolkit_core.exe");
            if (File.Exists(standalone)) return standalone;

            // 2. Developer local venv
            string venvPy = Path.Combine(baseDir, "..", "..", "..", "..", ".venv", "Scripts", "python.exe");
            if (File.Exists(venvPy)) return Path.GetFullPath(venvPy);

            string venvPyPosix = Path.Combine(baseDir, "..", "..", "..", "..", ".venv", "bin", "python");
            if (File.Exists(venvPyPosix)) return Path.GetFullPath(venvPyPosix);

            // 3. System fallback
            return "python";
        }

        private string ResolveWorkingDirectory()
        {
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
