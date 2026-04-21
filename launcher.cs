using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

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

    private static void RequireFile(string path, string helpText)
    {
        if (!File.Exists(path))
        {
            throw new InvalidOperationException("Missing file: " + path + Environment.NewLine + helpText);
        }
    }

    private static bool IsTaskpaneReady(string logPath)
    {
        try
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            ServicePointManager.ServerCertificateValidationCallback = TrustLocalhostCertificate;
            var request = (HttpWebRequest)WebRequest.Create("https://localhost:3000/taskpane.html");
            request.Method = "GET";
            request.Timeout = 2000;
            request.ReadWriteTimeout = 2000;

            using (var response = (HttpWebResponse)request.GetResponse())
            {
                var ok = (int)response.StatusCode >= 200 && (int)response.StatusCode < 300;
                if (!ok)
                {
                    AppendLog(logPath, "Taskpane health check returned HTTP " + (int)response.StatusCode);
                }
                return ok;
            }
        }
        catch (Exception ex)
        {
            AppendLog(logPath, "Taskpane health check failed: " + ex.Message);
            return false;
        }
    }

    private static bool TrustLocalhostCertificate(
        object sender,
        X509Certificate certificate,
        X509Chain chain,
        SslPolicyErrors sslPolicyErrors)
    {
        var request = sender as HttpWebRequest;
        if (request != null && request.RequestUri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return sslPolicyErrors == SslPolicyErrors.None;
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
            try
            {
                process.Start();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Could not start '" + fileName + "'. Install Node.js or keep the bundled runtime folder with this launcher. " + ex.Message,
                    ex
                );
            }
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

    private static void EnsureServerRunning(string nodeExe, string baseDir, string logPath)
    {
        if (IsTaskpaneReady(logPath))
        {
            AppendLog(logPath, "Taskpane already reachable.");
            return;
        }

        var serverScript = Path.Combine(baseDir, "dev-server.js");
        RequireFile(serverScript, "Keep dev-server.js in the same folder as NubraInstiExcelLauncher.exe.");

        RunProcess(nodeExe, "\"" + serverScript + "\"", baseDir, logPath, false, true);

        for (var i = 0; i < 20; i++)
        {
            if (IsTaskpaneReady(logPath))
            {
                AppendLog(logPath, "Taskpane became ready.");
                return;
            }
            Thread.Sleep(1000);
        }

        throw new InvalidOperationException(
            "The Nubra plugin server did not become ready at https://localhost:3000/taskpane.html. " +
            "Run setup-local.ps1 once, make sure port 3000 is free, and keep all plugin files together."
        );
    }

    private static void RunNodeCli(string nodeExe, string cliScript, string arguments, string baseDir, string logPath)
    {
        RequireFile(
            cliScript,
            "The launcher needs node_modules. Run setup-local.ps1 once, or ship the node_modules folder with this launcher."
        );

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

            RequireFile(manifestPath, "Keep manifest.xml in the same folder as NubraInstiExcelLauncher.exe.");

            Console.WriteLine("[launcher] Starting Nubra Insti Excel Plugin...");
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
