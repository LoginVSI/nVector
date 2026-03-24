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
        // version 1.1.0

        // =====================================================
        // Configurable
        // =====================================================
        string fileName    = "nvector-agent.exe";
        string processName = "nvector-agent";
        bool forceCopy     = true;  // copy from appliance even if file already exists

        // =====================================================
        // Paths
        // All artifacts (exe, screenshots, log) written to:
        //   C:\ProgramData\NVIDIA Corporation\nVector\
        //
        // NOTE: If the target session user cannot write to C:\ProgramData\
        // due to environment policy (FSLogix, Citrix UPM, GPO restrictions),
        // update nvidiaRoot to an accessible, static path and re-upload this workload
        // to the Login Enterprise appliance.

        // =====================================================
        string nvidiaRoot     = @"C:\ProgramData\NVIDIA Corporation\nVector";
        string filePath       = Path.Combine(nvidiaRoot, fileName);
        string logFilePath    = Path.Combine(nvidiaRoot, "agent.log");

        // Screenshots dir: agent 2.0+ automatically creates batch1, batch2...
        // subfolders here per screenshot round. Point at the root dir only —
        // do not point at a subfolder. The nVector Client Role integration watches
        // for these batch subfolders to trigger SSIM processing.
        string screenshotPath = nvidiaRoot;

        try
        {
            // ----- Skip launch if already running in THIS session -----
            int mySession = Process.GetCurrentProcess().SessionId;
            bool alreadyRunning = Process.GetProcessesByName(processName)
                                         .Any(p => p.SessionId == mySession);
            if (alreadyRunning)
            {
                Log($"{processName}.exe is already running in this session; skipping launch.");
                return;
            }

            // ----- Ensure all required directories exist -----
            // Directory.CreateDirectory is safe to call if directory already exists.
            Directory.CreateDirectory(nvidiaRoot);
            Log($"Ensured directory exists: {nvidiaRoot}");

            // ----- Copy nvector-agent.exe from Appliance ScriptContent -----
            if (forceCopy || !FileExists(filePath))
            {
                Log("Copying nvector-agent.exe from Appliance ScriptContent...");
                CopyFile(
                    sourcePath      : UrnBaseForFiles.UrnBase + "nvector-agent.exe",
                    destinationPath : filePath
                );
                Log($"Copied to: {filePath}");
            }
            else
            {
                Log("nvector-agent.exe already exists and forceCopy is false — skipping copy.");
            }

            if (!FileExists(filePath))
            {
                Log("ERROR: nvector-agent.exe not found after copy attempt. Aborting.");
                return;
            }

            // ----- Launch the agent in desktop role -----
            // All flags set explicitly so values are easy to find and adjust.
            //
            // IMPORTANT: --screenshots-per-batch and --max-screenshot-rounds
            // MUST match the values set in the nVector Agent Client role on the
            // Login Enterprise Launcher. Mismatched values will cause screenshot
            // synchronization failures and unreliable SSIM scores.
            var psi = new ProcessStartInfo
            {
                FileName         = filePath,
                Arguments        = "--role desktop"                           +
                                   $" --screenshots-dir \"{screenshotPath}\"" +
                                   $" --log-file \"{logFilePath}\""           +
                                   " --screenshots-per-batch 4"               +  // nvector-agent 2.0 default; increase for more screenshots per SSIM round
                                   " --max-screenshot-rounds 0",                 // nvector-agent 2.0 default; 0 = unlimited rounds for the duration of the test
                WorkingDirectory = nvidiaRoot,
                UseShellExecute  = false,
                CreateNoWindow   = true,
                WindowStyle      = ProcessWindowStyle.Minimized
            };

            Log("Launching nvector-agent.exe in desktop role...");
            Process.Start(psi);
            Log($"  --role:                  desktop");
            Log($"  --screenshots-dir:       {screenshotPath}");
            Log($"  --log-file:              {logFilePath}");
            Log($"  --screenshots-per-batch: 4");
            Log($"  --max-screenshot-rounds: 0 (unlimited)");

            // ----- Verify the process started -----
            Wait(3);
            bool isRunning = Process.GetProcessesByName(processName)
                                    .Any(p => p.SessionId == mySession);
            if (isRunning)
                Log($"{processName}.exe is confirmed running.");
            else
                Log($"WARNING: {processName}.exe does not appear to be running. Check log: {logFilePath}");
        }
        catch (Exception ex)
        {
            Log($"EXCEPTION in nVector_Desktop_Prepare: {ex}");
        }
    }
}
