// TARGET:nvector-agent-prepare
// START_IN:

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Constants;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

public class nVector_Desktop_Prepare : ScriptBase
{
    void Execute()
    {
        // version 1.0.1
        // =====================================================
        // Configurable
        // =====================================================
        string tempDir = Environment.GetEnvironmentVariable("TEMP");
        if (string.IsNullOrEmpty(tempDir))
            tempDir = Path.GetTempPath();

        string fileName    = "nvector-agent.exe";
        string filePath    = Path.Combine(tempDir, fileName);
        string processName = "nvector-agent";
        bool forceCopy     = true;  // copy from appliance even if file exists

        // =====================================================
        // Paths
        // =====================================================
        string nvidiaRoot     = Path.Combine(tempDir, "nvidia");
        string screenshotPath = Path.Combine(nvidiaRoot, "SSIM_screenshots");
        string logFilePath    = Path.Combine(nvidiaRoot, "agent.log");

        try
        {
            // ----- Skip launch if already running in THIS session -----
            int mySession = Process.GetCurrentProcess().SessionId;
            bool alreadyRunning = Process.GetProcessesByName(processName).Any(p => p.SessionId == mySession);
            if (alreadyRunning)
            {
                Log($"{processName}.exe is already running in this session; skipping launch.");
                return;
            }

            // ----- Ensure needed directories exist -----
            Directory.CreateDirectory(nvidiaRoot);
            Directory.CreateDirectory(screenshotPath);

            // ----- Copy nvector-agent.exe from Appliance ScriptContent if needed -----
            if (forceCopy || !FileExists(filePath))
            {
                Log("Copying nvector-agent.exe from Appliance ScriptContent");
                CopyFile(
                    sourcePath      : UrnBaseForFiles.UrnBase + "nvector-agent.exe",
                    destinationPath : filePath
                );
            }
            else
            {
                Log("nvector-agent.exe already exists and forceCopy is false");
            }

            if (!FileExists(filePath))
            {
                Log("nvector-agent.exe not found after copy; aborting.");
                return;
            }

            // ----- Launch the agent (quiet) -----
            var psi = new ProcessStartInfo
            {
                FileName         = filePath,
                Arguments        = $"-r desktop -s \"{screenshotPath}\" -l \"{logFilePath}\"",
                WorkingDirectory = tempDir,
                UseShellExecute  = false,
                CreateNoWindow   = true,
                WindowStyle      = ProcessWindowStyle.Minimized
            };

            Process.Start(psi);
            Log($"{fileName} started with '-r desktop', '-s {screenshotPath}', '-l {logFilePath}'.");

            // ----- Verify process -----
            Wait(3);
            bool isRunning = Process.GetProcessesByName(processName).Any(p => p.SessionId == mySession);
            Log(isRunning ? $"{processName}.exe is running." : $"{processName}.exe is not running.");
        }
        catch (Exception ex)
        {
            Log(ex.ToString());
        }
    }
}