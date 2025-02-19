# nVector Metrics Uploader Script

This script, created on **14 January 2025**, automates the monitoring of a results CSV file for visual latency metrics and uploads the data to the Login Enterprise API. It integrates NVIDIA nVector with Login VSI to provide end-to-end performance insights into desktop environments.

---

## Overview

- **Visual Latency Monitoring**: Measures latency between the display and client-side rendering using NVIDIA nVector’s watermarking technique.  
- **Data Integration**: Automatically uploads latency metrics to Login Enterprise for centralized analysis.  
- **Time Offset Adjustments**: Adjusts timestamps for local time zones, ensuring accurate data interpretation.  
- **Robust Error Handling**: Logs errors and ensures continuity during execution.  
- **Scalable and Configurable**: Suitable for a variety of setups with minimal configuration.

---

## Use Case

This script is designed for monitoring latency in business-critical applications to ensure that end-users experience optimal performance. By measuring visual latency—the time deviation between desktop rendering and client display—the script helps IT teams:

- **Identify Latency Bottlenecks**: Proactively address performance issues before they impact users.  
- **Quantify Visual Latency**: Provide objective metrics for troubleshooting and optimization.  
- **Maintain Productivity**: Ensure smooth operation of graphics-intensive workloads, especially under load.

---

## Script Workflow

1. **Terminate Running `nvector-agent` Instances**  
   Checks for existing `nvector-agent` processes and terminates them to avoid conflicts.

2. **Start `nvector-agent`**  
   - Verifies the existence of the executable (`nvector-agent.exe`).  
   - Starts the agent with the following arguments:  
     - `-r client`  
     - `-m <CsvFilePath>`  
     - `-p <NvectorAgentCheckIntervalMs>`  
     - `-s <NvectorScreenshotDir>`  
     - `-l <NvectorLogFile>`

3. **Ensure Launcher Process is Running**  
   Verifies and starts the `LoginEnterprise.Launcher.UI` process if not already running.

4. **Monitor CSV File for New Lines**  
   - Continuously reads the metrics CSV file to identify unprocessed lines.  
   - Updates the processed line count to avoid duplication.

5. **Build and Upload Metrics to the API**  
   - Parses CSV lines for timestamps and latency values.  
   - Adjusts timestamps using the configured offset.  
   - Filters out invalid or excessive values.  
   - Constructs JSON payloads for new metrics.  
   - Sends data to the Login Enterprise API endpoint.

---

## Configurable Variables

### Paths

- **`NvectorAgentExePath`**  
  Path to the `nvector-agent.exe` executable.  
  ```powershell
  $NvectorAgentExePath = "C:\temp\nVector-agent\nvector-agent.exe"
  ```

- **`CsvFilePath`**  
  Path to the metrics CSV file.
  ```powershell
  $CsvFilePath = "C:\temp\nvector\latency_metrics.csv"
  ```

- **`NvectorScreenshotDir`**  
  Directory for screenshots (required by nVector for some tests).
  ```powershell
  $NvectorScreenshotDir = "C:\temp\nvidia\SSIM_screenshots"
  ```

- **`NvectorLogFile`**  
  Log file for the nvector-agent.
  ```powershell
  $NvectorLogFile = "C:\temp\nvidia\agent.log"
  ```

- **`ScriptLogFile`**  
  Path for this script’s logs.
  ```powershell
  $ScriptLogFile = "C:\temp\nVectorAgent_ClientMeasurements_Uploader.txt"
  ```

### API Configuration

- **`ConfigurationAccessToken`**  
  Token for Login Enterprise API authentication.
  
  To obtain:
  1. Log into Login Enterprise.
  2. Navigate to **External Notifications → Public API**.
  3. Click **New System Access Token**.
  4. Provide a name and select **Configuration** as the Access Level.
  5. Save and copy the token.

- **`BaseUrl`**  
  Base URL of the Login Enterprise instance.
  ```powershell
  $BaseUrl = "https://myLoginEnterprise.myDomain.com/"
  ```

- **`ApiEndpoint`**  
  API endpoint for metrics upload.
  ```powershell
  $ApiEndpoint = "publicApi/v7-preview/platform-metrics"
  ```

- **`EnvironmentId`**  
  Unique identifier for the environment.

### Time Offset

- **`TimeOffset`**  
  Adjusts timestamps relative to UTC.
  
  Examples:
  - `"0:00"`: UTC.
  - `"-7:00"`: Pacific Standard Time (PST).
  - `"+2:00"`: Central European Summer Time (CEST).

### Intervals and Thresholds

- **`NvectorAgentCheckIntervalMs`**  
  Frequency for nvector-agent polling, in milliseconds.  
  *Default: 5000*

- **`PollingInterval`**  
  Script’s CSV monitoring frequency, in seconds.  
  *Default: 10*

- **`MaxLatencyThreshold`**  
  Maximum allowable latency, in milliseconds, before discarding data.  
  *Default: 10000*

### CSV Existence Check

- **`CsvCheckTimeoutSeconds`**  
  How many seconds to wait for the CSV file to appear.  
  *Default: 5*

- **`CsvCheckIntervalSeconds`**  
  How often (in seconds) to re-check for the CSV file within that timeout period.  
  *Default: 1*

### Prerequisites

- **NVIDIA nVector Installation**: Ensure the `nvector-agent` is installed on both the desktop and client machines.
- **Login Enterprise**: Tested with Login Enterprise version 5.13.6 or later.
- **Access Tokens and IDs**: Obtain a configuration access token and environment ID from your Login Enterprise instance.

### Benefits

- **End-User Experience Monitoring**: Accurately measures visual latency that impacts user productivity.
- **Proactive Performance Management**: Detects and quantifies latency issues before they escalate.
- **Data Integration**: Consolidates latency metrics for analysis in Login Enterprise.
- **Scalability**: Configurable for diverse setups and environments.

### Examples of Latency Results

#### Hourly View

*Hourly View Placeholder*  
Figure: Latency metrics displayed over an hour timeframe.

#### Daily View

*Daily View Placeholder*  
Figure: Latency metrics displayed over a day timeframe.

### Setup Steps

1. **Upload the Target Workload**  
   *(Placeholder: Additional information to be added if needed.)*

2. **Place the Script**  
   - Save the script as `nVector Prepare.ps1`.
   - Configure the script to run automatically on the Login Enterprise Launcher upon user login (or another preferred schedule).

### Script Reference

For the full PowerShell script, see the file `nVector Prepare.ps1` in your environment. This script:

- Terminates any running `nvector-agent` processes.
- Starts `nvector-agent` with the specified arguments.
- Ensures the Login Enterprise launcher process is running.
- Checks for the CSV file and its header.
- Continuously monitors the CSV for new data lines and uploads valid latency metrics to the Login Enterprise API.

### Additional Information

- **Best Practices**:
  - Use clear and concise variable configurations.
  - Follow error-handling and logging best practices.
  
- **Notes**:
  - Modify the variables at the top of the script to suit your environment.
  - Ensure both the desktop and client machines meet NVIDIA nVector requirements.

### Change Log

**v1.0 (14 January 2025)**

- Initial release of the nVector Metrics Uploader script.

© 2025 NVIDIA nVector and Login VSI. All rights reserved. For support, contact [support@loginvsi.com](mailto:support@loginvsi.com).
