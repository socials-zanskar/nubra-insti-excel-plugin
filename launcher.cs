using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Net.NetworkInformation;

internal static class Program
{
    private static readonly object LogLock = new object();

    private static void AppendLog(string logPath, string message)
    {
        lock (LogLock)
        {
            File.AppendAllText(
                logPath,
                "[" + DateTime.Now.ToString("s") + "] " + message + Environment.NewLine,
                Encoding.UTF8
            );
        }
    }

    private static string ResolveNodeExe(string baseDir)
    {
        var candidates = new[]
        {
            Path.Combine(baseDir, "runtime", "node", "node.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "nodejs", "node.exe")
        };

        foreach (var candidate in candidates)
        {
            if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
            {
                return candidate;
            }
        }

        return "node";
    }

    private static int RunProcess(
        string fileName,
        string arguments,
        string workingDirectory,
        string logPath,
        bool waitForExit,
        bool suppressOutput)
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            WorkingDirectory = workingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = false
        };

        using (var process = new Process())
        {
            process.StartInfo = psi;
            process.OutputDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                {
                    if (!suppressOutput)
                    {
                        Console.WriteLine(e.Data);
                    }
                    AppendLog(logPath, "[stdout] " + e.Data);
                }
            };
            process.ErrorDataReceived += (sender, e) =>
            {
                if (!string.IsNullOrWhiteSpace(e.Data))
                {
                    if (!suppressOutput)
                    {
                        Console.Error.WriteLine(e.Data);
                    }
                    AppendLog(logPath, "[stderr] " + e.Data);
                }
            };

            AppendLog(logPath, "Running: " + fileName + " " + arguments);
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            if (!waitForExit)
            {
                return 0;
            }

            process.WaitForExit();
            AppendLog(logPath, "Exit code: " + process.ExitCode);
            return process.ExitCode;
        }
    }

    private static bool IsPortListening(int port)
    {
        try
        {
            var listeners = IPGlobalProperties.GetIPGlobalProperties().GetActiveTcpListeners();
            foreach (var endpoint in listeners)
            {
                if (endpoint.Port == port)
                {
                    return true;
                }
            }

            return false;
        }
        catch
        {
            return false;
        }
    }

    private static void EnsureServerRunning(string nodeExe, string baseDir, string logPath)
    {
        if (IsPortListening(3000))
        {
            AppendLog(logPath, "Dev server already reachable.");
            return;
        }

        var serverScript = Path.Combine(baseDir, "dev-server.js");
        if (!File.Exists(serverScript))
        {
            throw new InvalidOperationException("Missing file: " + serverScript);
        }

        RunProcess(nodeExe, "\"" + serverScript + "\"", baseDir, logPath, false, true);

        for (var i = 0; i < 20; i++)
        {
            if (IsPortListening(3000))
            {
                AppendLog(logPath, "Dev server became ready.");
                return;
            }
            Thread.Sleep(1000);
        }

        throw new InvalidOperationException(
            "Dev server did not become ready. If this machine was not set up earlier, localhost certificate/loopback setup is still required once."
        );
    }

    private static void RunNodeCli(string nodeExe, string cliScript, string arguments, string baseDir, string logPath)
    {
        if (!File.Exists(cliScript))
        {
            throw new InvalidOperationException("Missing CLI script: " + cliScript);
        }

        var exitCode = RunProcess(nodeExe, "\"" + cliScript + "\" " + arguments, baseDir, logPath, true, false);
        if (exitCode != 0)
        {
            throw new InvalidOperationException("Command failed: " + Path.GetFileName(cliScript) + " " + arguments);
        }
    }

    private static int Main()
    {
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        var logPath = Path.Combine(baseDir, "launcher.log");

        try
        {
            var manifestPath = Path.Combine(baseDir, "manifest.xml");
            var devSettingsCli = Path.Combine(baseDir, "node_modules", "office-addin-dev-settings", "cli.js");
            var nodeExe = ResolveNodeExe(baseDir);

            AppendLog(logPath, "Launcher start");

            if (!File.Exists(manifestPath))
            {
                Console.Error.WriteLine("Missing file: " + manifestPath);
                return 1;
            }

            Console.WriteLine("[launcher] Starting Nubra Insti Excel Plugin (test mode)...");
            Console.WriteLine("[launcher] Folder: " + baseDir);

            EnsureServerRunning(nodeExe, baseDir, logPath);
            RunNodeCli(nodeExe, devSettingsCli, "register \"" + manifestPath + "\"", baseDir, logPath);
            RunNodeCli(nodeExe, devSettingsCli, "sideload \"" + manifestPath + "\" desktop --app excel", baseDir, logPath);

            Console.WriteLine("[launcher] Completed.");
            AppendLog(logPath, "Launcher completed");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("[launcher] Fatal error: " + ex.Message);
            AppendLog(logPath, "fatal: " + ex);
            Console.WriteLine("Press Enter to close...");
            Console.ReadLine();
            return 1;
        }
    }
}
