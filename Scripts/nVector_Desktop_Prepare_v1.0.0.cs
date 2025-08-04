// TARGET:ping
// START_IN:

using LoginPI.Engine.ScriptBase;
using LoginPI.Engine.ScriptBase.Constants;
using System;
using System.Diagnostics;
using System.IO;

public class nVector_Desktop_Prepare : ScriptBase
{
    void Execute()
    {
        // =====================================================
        // Configurable Variables
        // =====================================================
        string tempDir         = Environment.GetEnvironmentVariable("TEMP");
        string fileName        = "nvector-agent.exe";
        string filePath        = Path.Combine(tempDir, fileName);
        string processName     = "nvector-agent";
        bool forceCopy         = true;  // Set to true to copy from appliance even if file exists

        // =====================================================
        // Other Paths/Params (now using tempDir)
        // =====================================================
        string screenshotPath  = Path.Combine(tempDir, "nvidia", "SSIM_screenshots");
        string logFilePath     = Path.Combine(tempDir, "nvidia", "agent.log");

        try
        {
            // ----- Ensure needed directories exist -----
            // Creates both the parent "nvidia" folder and the SSIM_screenshots subfolder
            Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));
            Directory.CreateDirectory(screenshotPath);

            // ----- Copy nvector-agent.exe from Appliance scriptcontent if needed -----
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

            // ----- Skip launch if already running -----
            if (Process.GetProcessesByName(processName).Length > 0)
            {
                Log($"{processName}.exe is already running; skipping launch.");
                return;
            }

            // ----- Launch the agent -----
            var process = new Process();
            process.StartInfo.FileName      = filePath;
            process.StartInfo.Arguments     = $"-r desktop -s \"{screenshotPath}\" -l \"{logFilePath}\"";
            process.StartInfo.WindowStyle   = ProcessWindowStyle.Minimized;   // start minimized
            process.Start();

            Log($"{fileName} started with '-r desktop', '-s {screenshotPath}', '-l {logFilePath}'.");

            // ----- Verify process -----
            Wait(1);
            bool isRunning = Process.GetProcessesByName(processName).Length > 0;
            Log(isRunning
                ? $"{processName}.exe is running."
                : $"{processName}.exe is not running.");
        }
        catch (Exception ex)
        {
            Log($"An error occurred: {ex.Message}");
        }
    }
}
