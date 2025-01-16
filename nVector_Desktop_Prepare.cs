// TARGET:ping
// START_IN:

using LoginPI.Engine.ScriptBase;
using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;

public class nVector_Desktop_Prepare : ScriptBase
{
    void Execute()
    {
        // Define variables for paths and process name
        string tempDir = Environment.GetEnvironmentVariable("TEMP"); // Path to the %temp% directory
        string fileName = "nvector-agent.exe";
        string filePath = Path.Combine(tempDir, fileName); // Full path to the executable
        string downloadUrl = "https://myDomain.LoginEnterprise.com/contentDelivery/content/nvidia/nvector-agent.exe"; // Placeholder URL
        string processName = "nvector-agent";

        // Define additional command line parameters for the agent
        string screenshotPath = Path.Combine(tempDir, @"nvidia\SSIM_screenshots"); // Directory for screenshots
        string logFilePath = Path.Combine(tempDir, @"nvidia\agent.log"); // Path to the log file

        try
        {
            // Ensure the directories exist
            Directory.CreateDirectory(screenshotPath);

            // Step 1: Disable SSL certificate validation
            ServicePointManager.ServerCertificateValidationCallback = delegate (
                object sender,
                X509Certificate certificate,
                X509Chain chain,
                SslPolicyErrors sslPolicyErrors)
            {
                return true; // Always accept the certificate
            };

            // Step 2: Download the nvector-agent.exe file to the %temp% directory
            using (WebClient client = new WebClient())
            {
                client.DownloadFile(downloadUrl, filePath);
                Log($"File downloaded successfully to: {filePath}");
            }

            // Step 3: Start the downloaded file with the required arguments
            // Arguments: -r desktop (role), -s screenshotPath, -l logFilePath
            Process process = new Process();
            process.StartInfo.FileName = filePath;
            process.StartInfo.Arguments = $"-r desktop -s \"{screenshotPath}\" -l \"{logFilePath}\"";
            process.Start();
            Log($"{fileName} started with arguments '-r desktop', '-s {screenshotPath}', and '-l {logFilePath}'.");

            // Step 4: Wait for 1 second before verifying if the process is running
            Wait(1);

            // Step 5: Verify if the nvector-agent process is running
            bool isRunning = Process.GetProcessesByName(processName).Length > 0;

            if (isRunning)
            {
                Log($"{processName}.exe is running.");
            }
            else
            {
                Log($"{processName}.exe is not running.");
            }
        }
        catch (Exception ex)
        {
            // Log any errors encountered during execution
            Log($"An error occurred: {ex.Message}");
        }
    }
}
